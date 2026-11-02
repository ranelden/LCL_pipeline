"""
Nettoyage des libellés de transactions LCL pour obtenir un nom de commerçant
court et stable (utilisé à la fois pour l'affichage et comme clé
d'apprentissage des catégories dans categorize.py).

Sans nettoyage, un même commerçant apparaît sous un libellé différent
chaque mois à cause de la date embarquée ("CB BOURSORAMA 07/02/26" puis
"CB BOURSORAMA 03/03/26") : la catégorisation apprise ne pourrait jamais
reconnaître deux fois le même commerçant. On retire donc :
- le préfixe "CB "
- tout ce qui suit un champ technique (REF.CLIENT:, REF.MANDAT:,
  NATURE DU PAIEMENT:, ID.CREANCIER:, LIBELLE:) sur les virements/prélèvements
- les blocs "<ville ou host> EUR <montant>" qui répètent le montant dans la
  devise d'origine
- les dates JJ/MM ou JJ/MM/AA embarquées dans le libellé
"""
import re

_CUT_MARKERS = ("REF.CLIENT:", "REF.MANDAT:", "NATURE DU PAIEMENT:", "ID.CREANCIER:", "LIBELLE:")
_EUR_CHUNK_RE = re.compile(r"\S+\s+EUR\s+[\d\s]+,\d{2}", re.IGNORECASE)
_DATE_RE = re.compile(r"\b\d{2}[./]\d{2}(?:[./]\d{2,4})?\b")
_CB_PREFIX_RE = re.compile(r"^\s*CB\s+", re.IGNORECASE)
_SPACES_RE = re.compile(r"\s{2,}")


def clean_label(raw: str) -> str:
    label = raw.strip()

    cut_at = min((label.find(m) for m in _CUT_MARKERS if m in label), default=None)
    if cut_at is not None:
        label = label[:cut_at]

    label = _CB_PREFIX_RE.sub("", label)
    label = _EUR_CHUNK_RE.sub("", label)
    label = _DATE_RE.sub("", label)
    label = _SPACES_RE.sub(" ", label).strip(" -")

    return label or raw.strip()
