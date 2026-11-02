"""
Extraction des transactions depuis un relevé de compte LCL (PDF).

Le tableau "ECRITURES DE LA PERIODE" n'est pas détecté par
pdfplumber.extract_tables() : il n'a pas de vraies lignes verticales
continues, seulement de courts traits (des "amorces" de colonne) au
niveau de la ligne d'en-tête et de la ligne de totaux. On récupère donc
les 4 bornes de colonnes (DATE | LIBELLE | VALEUR | DEBIT | CREDIT) à
partir de ces traits, puis on reconstruit chaque ligne du tableau en
regroupant les mots de la page par position verticale (top) et en les
répartissant dans les colonnes par position horizontale (x0).

Autres difficultés traitées :
- montant en colonne DEBIT ou CREDIT distingué uniquement par position x
- écritures étalées sur plusieurs lignes (ville + montant répété dans la
  devise d'origine, ou bloc de prélèvement SEPA sur 5 lignes) -> fusionnées
  avec la ligne principale
- format numérique français ("1 300,00")
- lignes qui ne sont pas des écritures (ANCIEN SOLDE, SOLDE INTERMEDIAIRE,
  TOTAUX, SOLDE EN EUROS) -> exclues des transactions mais gardées pour
  vérifier que la somme des écritures extraites retombe sur les totaux
  imprimés par la banque
"""
import bisect
import re
from dataclasses import dataclass, field
from datetime import date

import pdfplumber

PERIOD_RE = re.compile(r"du (\d{2})\.(\d{2})\.(\d{4}) au (\d{2})\.(\d{2})\.(\d{4})")

MOIS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
           "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]


def dominant_month(transactions, period_start, period_end):
    """Détermine le mois 'principal' d'un relevé qui chevauche deux mois
    civils : celui qui compte le plus d'écritures (à égalité, celui du
    début de période)."""
    counts = {}
    for t in transactions:
        d = t.get("date_complete")
        if d is None:
            continue
        counts[(d.year, d.month)] = counts.get((d.year, d.month), 0) + 1

    if not counts:
        ref = period_end or period_start
        return (ref.year, ref.month) if ref else (None, None)

    best = max(counts.items(), key=lambda kv: (kv[1], kv[0] == (period_start.year, period_start.month)))
    return best[0]


def month_name_fr(month: int) -> str:
    return MOIS_FR[month - 1]

AMOUNT_RE = re.compile(r"^-?\d[\d ]*,\d{2}$")
DATE_RE = re.compile(r"^\d{2}\.\d{2}$")
NON_TRANSACTION_LABELS = (
    "ANCIEN SOLDE",
    "SOLDE INTERMEDIAIRE",
    "TOTAUX",
    "SOLDE EN EUROS",
    "NOUVEAU SOLDE",
)
FINAL_ROW_LABELS = ("SOLDE EN EUROS", "NOUVEAU SOLDE")
# pied de page répété sur chaque page (mentions légales) : tout ce qui suit
# n'appartient plus au tableau des écritures
PAGE_FOOTER_RE = re.compile(r"Cr[ée]dit Lyonnais|SIREN|^Page\s+\d+\s*/\s*\d+", re.IGNORECASE)

REQUIRED_HEADERS = ("DATE", "LIBELLE", "VALEUR", "DEBIT", "CREDIT")


def parse_amount(s: str) -> float:
    return float(s.replace(" ", "").replace("\xa0", "").replace(",", "."))


@dataclass
class Row:
    date: str = ""
    libelle_parts: list = field(default_factory=list)
    valeur: str = ""
    debit: str = ""
    credit: str = ""

    @property
    def libelle(self):
        return " ".join(self.libelle_parts).strip()


def get_columns(page):
    """Repère l'en-tête DATE/LIBELLE/VALEUR/DEBIT/CREDIT sur la page et les
    4 bornes de colonnes correspondantes (petits traits verticaux relevés
    au niveau de cette ligne)."""
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    header = {w["text"]: w for w in words if w["text"] in REQUIRED_HEADERS}
    if not set(REQUIRED_HEADERS) <= header.keys():
        return None, words

    header_top = min(w["top"] for w in header.values())
    header_bottom = max(w["bottom"] for w in header.values())
    x_left = header["DATE"]["x1"]
    x_right = header["CREDIT"]["x0"]

    ticks = set()
    for l in page.lines:
        if abs(l["x0"] - l["x1"]) < 1.5 and l["top"] >= header_top - 2:
            x = round((l["x0"] + l["x1"]) / 2, 1)
            if x_left < x < x_right:
                ticks.add(x)
    bounds = sorted(ticks)
    if len(bounds) != 4:
        return None, words

    # 4 séparateurs internes -> 5 colonnes : DATE | LIBELLE | VALEUR | DEBIT | CREDIT
    return {"bounds": bounds, "header_bottom": header_bottom}, words


def assign_column(x0, bounds):
    # bisect sur les 4 bornes internes -> renvoie un index de colonne 0..4
    return bisect.bisect_right(bounds, x0)


def extract_rows_from_page(page):
    cols, words = get_columns(page)
    if cols is None:
        return []
    bounds = cols["bounds"]
    words = [w for w in words if w["top"] > cols["header_bottom"] + 1]

    lines = {}
    for w in words:
        key = round(w["top"], 0)
        lines.setdefault(key, []).append(w)
    line_tops = sorted(lines.keys())

    order = ["DATE", "LIBELLE", "VALEUR", "DEBIT", "CREDIT"]
    rows = []
    current = None
    for top in line_tops:
        line_words = sorted(lines[top], key=lambda w: w["x0"])
        full_line_txt = " ".join(w["text"] for w in line_words)
        if PAGE_FOOTER_RE.search(full_line_txt):
            # pied de page légal répété sur chaque page : on arrête là, le
            # reste de la page ne contient plus d'écritures
            break

        by_col = {name: [] for name in order}
        for w in line_words:
            idx = assign_column(w["x0"], bounds)
            by_col[order[idx]].append(w["text"])

        date_txt = " ".join(by_col["DATE"]).strip()
        libelle_txt = " ".join(by_col["LIBELLE"]).strip()
        is_summary = any(libelle_txt.upper().startswith(lbl) for lbl in NON_TRANSACTION_LABELS)
        is_new_row = bool(DATE_RE.match(date_txt)) or is_summary

        if is_new_row:
            if current is not None:
                rows.append(current)
            current = Row(date=date_txt)
            if libelle_txt:
                current.libelle_parts.append(libelle_txt)
            if by_col["VALEUR"]:
                current.valeur = " ".join(by_col["VALEUR"])
            if by_col["DEBIT"]:
                current.debit = " ".join(by_col["DEBIT"])
            if by_col["CREDIT"]:
                current.credit = " ".join(by_col["CREDIT"])
            if any(libelle_txt.upper().startswith(lbl) for lbl in FINAL_ROW_LABELS):
                rows.append(current)
                current = None
                break  # tout ce qui suit sur la page est du texte légal, pas des écritures
        else:
            if current is None:
                continue
            if libelle_txt:
                current.libelle_parts.append(libelle_txt)
            if by_col["VALEUR"] and not current.valeur:
                current.valeur = " ".join(by_col["VALEUR"])
            if by_col["DEBIT"] and not current.debit and AMOUNT_RE.match("".join(by_col["DEBIT"])):
                current.debit = " ".join(by_col["DEBIT"])
            if by_col["CREDIT"] and not current.credit and AMOUNT_RE.match("".join(by_col["CREDIT"])):
                current.credit = " ".join(by_col["CREDIT"])
    if current is not None:
        rows.append(current)
    return rows


def extract_period(pdf):
    """Lit 'du DD.MM.YYYY au DD.MM.YYYY' sur la première page pour connaître
    l'année réelle des dates (qui ne sont imprimées qu'en JJ.MM dans le
    tableau des écritures)."""
    text = pdf.pages[0].extract_text() or ""
    m = PERIOD_RE.search(text)
    if not m:
        return None, None
    d1, m1, y1, d2, m2, y2 = m.groups()
    return date(int(y1), int(m1), int(d1)), date(int(y2), int(m2), int(d2))


def resolve_year(day_month: str, period_start, period_end):
    """'DD.MM' -> date complète, en déduisant l'année depuis la période du
    relevé (une même période peut chevaucher deux mois de deux années
    différentes, ex. décembre/janvier)."""
    try:
        d, m = (int(x) for x in day_month.split("."))
    except ValueError:
        return None
    if period_start is None:
        return None
    for candidate_year in {period_start.year, period_end.year}:
        try:
            cand = date(candidate_year, m, d)
        except ValueError:
            continue
        if period_start <= cand <= period_end:
            return cand
    # repli : année du début de période
    try:
        return date(period_start.year, m, d)
    except ValueError:
        return None


def extract_statement(pdf_path):
    """Retourne (transactions, meta). meta contient les soldes/totaux lus
    dans le PDF, pour vérifier que l'extraction est complète."""
    transactions = []
    meta = {"ancien_solde": None, "solde_final": None, "total_debit": None, "total_credit": None,
            "periode_debut": None, "periode_fin": None}

    with pdfplumber.open(pdf_path) as pdf:
        period_start, period_end = extract_period(pdf)
        meta["periode_debut"], meta["periode_fin"] = period_start, period_end
        for page in pdf.pages:
            for row in extract_rows_from_page(page):
                libelle = row.libelle.strip()
                label_upper = libelle.upper()

                if any(label_upper.startswith(lbl) for lbl in NON_TRANSACTION_LABELS):
                    amt = None
                    if row.debit and AMOUNT_RE.match(row.debit):
                        amt = -parse_amount(row.debit)
                    elif row.credit and AMOUNT_RE.match(row.credit):
                        amt = parse_amount(row.credit)
                    if label_upper.startswith("ANCIEN SOLDE"):
                        meta["ancien_solde"] = amt
                    elif any(label_upper.startswith(l) for l in FINAL_ROW_LABELS):
                        meta["solde_final"] = amt
                    elif label_upper.startswith("TOTAUX"):
                        if row.debit and AMOUNT_RE.match(row.debit):
                            meta["total_debit"] = parse_amount(row.debit)
                        if row.credit and AMOUNT_RE.match(row.credit):
                            meta["total_credit"] = parse_amount(row.credit)
                    continue

                if not libelle:
                    continue

                debit = parse_amount(row.debit) if row.debit and AMOUNT_RE.match(row.debit) else None
                credit = parse_amount(row.credit) if row.credit and AMOUNT_RE.match(row.credit) else None
                if debit is None and credit is None:
                    continue

                montant = -debit if debit is not None else credit
                transactions.append({
                    "date": row.date,
                    "date_complete": resolve_year(row.date, period_start, period_end),
                    "libelle": libelle,
                    "valeur": row.valeur,
                    "montant": round(montant, 2),
                    "sens": "debit" if debit is not None else "credit",
                })

    return transactions, meta


if __name__ == "__main__":
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "releve1.pdf"
    txs, meta = extract_statement(path)
    print(f"{len(txs)} transactions extraites\n")
    for t in txs:
        print(f"{t['date']}  {t['montant']:>10.2f}  {t['libelle'][:80]}")

    print("\n--- Vérification ---")
    print("meta:", meta)
    somme_debit = -sum(t["montant"] for t in txs if t["montant"] < 0)
    somme_credit = sum(t["montant"] for t in txs if t["montant"] > 0)
    print(f"Somme débits extraits : {somme_debit:.2f}  (PDF: {meta['total_debit']})")
    print(f"Somme crédits extraits: {somme_credit:.2f}  (PDF: {meta['total_credit']})")
    if meta["ancien_solde"] is not None and meta["solde_final"] is not None:
        solde_calc = meta["ancien_solde"] - somme_debit + somme_credit
        print(f"Solde recalculé : {solde_calc:.2f}  (PDF: {meta['solde_final']})")
