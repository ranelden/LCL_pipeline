#!/usr/bin/env python3
"""
Génère un Excel de suivi des dépenses à partir d'un ou plusieurs relevés de
compte LCL en PDF. Ne dépend d'aucune IA :
- extraction par position dans le PDF (extract.py)
- nettoyage des libellés (clean.py)
- catégorisation décidée par toi et mémorisée dans un JSON (categorize.py)

Un fichier Excel est généré PAR RELEVÉ (un relevé peut chevaucher deux
mois), nommé "AAAA MM JJ - <Mois principal>.xlsx" et déposé par défaut dans
le dossier ./output (créé automatiquement s'il n'existe pas).

Chaque fichier a deux pages :
- "Travail"  : le détail de toutes les écritures, à corriger si besoin
               (tu peux changer une catégorie directement dans cette page).
- "Total"    : le total par catégorie, en formules (SUMIF), donc mis à
               jour automatiquement dès que tu modifies la page Travail.

Relancer le script sur le même relevé (même fichier de sortie) NE PERD PAS
tes corrections : les écritures déjà présentes gardent la catégorie que tu
leur as donnée dans la page Travail, et cette catégorie est mémorisée dans
categories.json pour que les mois suivants la reconnaissent automatiquement.

Usage :
    python generer_rapport.py releve_fevrier.pdf [releve_mars.pdf ...]
    python generer_rapport.py releve.pdf --output-dir output --store categories.json
"""
import argparse
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from categorize import SKIP_CATEGORY, load_store, remember_correction, resolve_category, save_store
from clean import clean_label
from extract import dominant_month, extract_statement, month_name_fr

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT_NAME, color="FFFFFF", bold=True)
BASE_FONT = Font(name=FONT_NAME, size=10)
REVIEW_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = "#,##0.00 €;-#,##0.00 €"

TX_HEADERS = ["Date", "Mois", "Libellé", "Libellé complet", "Catégorie", "Montant (€)", "Relevé source"]
EXCLUDED_FROM_SPENDING = {"Revenus", "Virements internes / Épargne"}


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def output_filename(period_start, period_end, transactions) -> str:
    y, m = dominant_month(transactions, period_start, period_end)
    ref = period_end or period_start
    return f"{ref.year:04d} {ref.month:02d} {ref.day:02d} - {month_name_fr(m)}.xlsx"


def load_existing_travail(path: Path):
    """Relit la page Travail d'un fichier déjà généré : renvoie
    {(date_iso, libellé_nettoyé, montant): catégorie}."""
    if not path.exists():
        return {}
    wb = load_workbook(path)
    if "Travail" not in wb.sheetnames:
        return {}
    ws = wb["Travail"]
    existing = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        date_val, _mois, libelle, _complet, categorie, montant, _source = row[:7]
        date_iso = date_val.strftime("%Y-%m-%d") if hasattr(date_val, "strftime") else str(date_val)
        key = (date_iso, libelle, round(float(montant), 2) if montant is not None else None)
        existing[key] = categorie
    return existing


def build_workbook(transactions, out_path: Path, store: dict, interactive: bool):
    existing = load_existing_travail(out_path)

    # resynchronise les corrections manuelles déjà présentes dans le fichier
    for (_date_iso, libelle, _montant), categorie in existing.items():
        remember_correction(libelle, categorie, store)

    rows = []
    for t in sorted(transactions, key=lambda x: (x["date_complete"] or "", x["libelle"])):
        if t["date_complete"] is None:
            continue
        clean = clean_label(t["libelle"])
        date_iso = t["date_complete"].strftime("%Y-%m-%d")
        key = (date_iso, clean, t["montant"])

        if key in existing:
            categorie = existing[key]
        else:
            categorie = resolve_category(clean, t["libelle"], t["montant"], store, interactive=interactive)

        rows.append({
            "date": t["date_complete"],
            "libelle": clean,
            "libelle_complet": t["libelle"],
            "categorie": categorie,
            "montant": t["montant"],
            "source": t["source"],
        })

    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Travail")
    for c, h in enumerate(TX_HEADERS, start=1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, len(TX_HEADERS))
    ws.freeze_panes = "A2"

    for i, r in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=r["date"]).number_format = "dd/mm/yyyy"
        ws.cell(row=i, column=2, value=f'=TEXT(A{i},"yyyy-mm")')
        ws.cell(row=i, column=3, value=r["libelle"])
        ws.cell(row=i, column=4, value=r["libelle_complet"])
        cat_cell = ws.cell(row=i, column=5, value=r["categorie"])
        if r["categorie"] == SKIP_CATEGORY:
            cat_cell.fill = REVIEW_FILL
        ws.cell(row=i, column=6, value=r["montant"]).number_format = MONEY_FMT
        ws.cell(row=i, column=7, value=r["source"])
        for c in range(1, len(TX_HEADERS) + 1):
            ws.cell(row=i, column=c).font = BASE_FONT
            ws.cell(row=i, column=c).border = BORDER

    autosize(ws, [12, 10, 32, 45, 26, 14, 30])
    last_row = ws.max_row

    # Page Total : une ligne par catégorie, en SUMIF sur la page Travail
    total = wb.create_sheet("Total")
    total.cell(row=1, column=1, value="Catégorie")
    total.cell(row=1, column=2, value="Total (€)")
    style_header(total, 2)

    categories_used = list(store["categories"]) + [SKIP_CATEGORY]
    seen = set()
    categories_used = [c for c in categories_used if not (c in seen or seen.add(c))]

    for i, cat in enumerate(categories_used, start=2):
        total.cell(row=i, column=1, value=cat).font = BASE_FONT
        formula = f'=SUMIF(Travail!$E$2:$E${last_row},$A{i},Travail!$F$2:$F${last_row})'
        c = total.cell(row=i, column=2, value=formula)
        c.number_format = MONEY_FMT
        c.font = BASE_FONT

    total_row = len(categories_used) + 2
    total.cell(row=total_row, column=1, value="Total").font = Font(name=FONT_NAME, bold=True)
    c = total.cell(row=total_row, column=2, value=f"=SUM(B2:B{total_row - 1})")
    c.number_format = MONEY_FMT
    c.font = Font(name=FONT_NAME, bold=True)

    dep_row = total_row + 2
    included = [i for i, cat in enumerate(categories_used, start=2) if cat not in EXCLUDED_FROM_SPENDING]
    total.cell(row=dep_row, column=1, value="Dépenses réelles (hors virements internes / revenus)").font = Font(
        name=FONT_NAME, bold=True, italic=True)
    if included:
        terms = "+".join(f"B{r}" for r in included)
        c = total.cell(row=dep_row, column=2, value=f"=-({terms})")
        c.number_format = MONEY_FMT
        c.font = Font(name=FONT_NAME, bold=True, italic=True)

    autosize(total, [42, 16])

    wb.save(out_path)
    pending = sorted({r["libelle"] for r in rows if r["categorie"] == SKIP_CATEGORY})
    return len(rows), pending


def main():
    parser = argparse.ArgumentParser(description="Génère un Excel de suivi des dépenses depuis des relevés LCL PDF.")
    parser.add_argument("pdfs", nargs="+", help="Un ou plusieurs relevés LCL au format PDF")
    parser.add_argument("--output-dir", default="output", help="Dossier de sortie (créé si absent, défaut: output)")
    parser.add_argument("--store", default="categories.json", help="Fichier JSON des catégories et de l'apprentissage")
    parser.add_argument("--non-interactif", action="store_true",
                         help="Ne pose aucune question en console ; les libellés inconnus restent 'À catégoriser'")
    args = parser.parse_args()

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    store_path = Path(args.store)
    store = load_store(store_path)

    for pdf in args.pdfs:
        txs, meta = extract_statement(pdf)
        if not txs:
            print(f"{pdf} : aucune transaction extraite, vérifie le fichier.")
            continue
        for t in txs:
            t["source"] = Path(pdf).name

        somme_debit = -sum(t["montant"] for t in txs if t["montant"] < 0)
        somme_credit = sum(t["montant"] for t in txs if t["montant"] > 0)
        if meta["total_debit"] is not None and abs(somme_debit - meta["total_debit"]) > 0.01:
            print(f"  ATTENTION {pdf} : total débit extrait {somme_debit:.2f} € "
                  f"≠ total imprimé {meta['total_debit']:.2f} € — vérifie la page Travail.")
        if meta["ancien_solde"] is not None and meta["solde_final"] is not None:
            solde_calc = meta["ancien_solde"] - somme_debit + somme_credit
            if abs(solde_calc - meta["solde_final"]) > 0.01:
                print(f"  ATTENTION {pdf} : solde recalculé {solde_calc:.2f} € "
                      f"≠ solde imprimé {meta['solde_final']:.2f} € — vérifie la page Travail.")

        filename = output_filename(meta["periode_debut"], meta["periode_fin"], txs)
        out_path = out_dir / filename

        nb, pending = build_workbook(txs, out_path, store, interactive=not args.non_interactif)
        save_store(store, store_path)

        print(f"{pdf} -> {out_path} ({nb} écritures)")
        if pending:
            print(f"  À catégoriser plus tard ({len(pending)}) : {', '.join(sorted(set(pending)))}")


if __name__ == "__main__":
    main()
