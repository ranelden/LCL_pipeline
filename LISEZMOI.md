# Pipeline dépenses LCL (v2)

Quatre fichiers Python, aucune IA requise pour l'usage courant :

- `extract.py` — lit un relevé LCL en PDF et en extrait les transactions (date, libellé, montant), plus la période du relevé.
- `clean.py` — raccourcit les libellés (retire "CB", les dates, les répétitions de montant en devise, les champs techniques des virements/prélèvements).
- `categorize.py` — catégorisation décidée par toi, mémorisée dans `categories.json`.
- `generer_rapport.py` — assemble le tout et écrit un Excel par relevé.

## Installation (une seule fois)

```
pip install pdfplumber openpyxl
```

## Utilisation

```
python generer_rapport.py releve_fevrier.pdf
```

Par défaut, le fichier est déposé dans un dossier `output/` (créé automatiquement si absent — tu peux aussi le créer toi-même) et nommé automatiquement, par exemple :

```
output/2026 03 06 - Février.xlsx
```

(la date reprend celle de fin de relevé, le nom du mois est celui qui compte le plus d'écritures sur la période, puisqu'un relevé LCL chevauche presque toujours deux mois civils).

Tu peux donner plusieurs PDF d'un coup pour traiter tout ton historique : un fichier Excel est généré par relevé.

```
python generer_rapport.py releve_janvier.pdf releve_fevrier.pdf releve_mars.pdf
```

## Catégorisation

La première fois qu'un commerçant apparaît, le script pose la question directement dans la console :

```
Catégorie pour : BOURSORAMA  (-400.00 €)
  1) Alimentation
  2) Logement
  3) Transport
  ...
  0) Passer (à catégoriser plus tard)
  n) Nouvelle catégorie
Choix :
```

Tape le numéro, `0` pour laisser en "À catégoriser" et y revenir plus tard, ou `n` pour créer une nouvelle catégorie à la volée. Ta réponse est mémorisée dans `categories.json` : le mois suivant, ce commerçant sera reconnu automatiquement, sans nouvelle question.

Tu peux aussi éditer `categories.json` toi-même à tout moment (ajouter/renommer/réordonner les catégories, ou corriger un mapping).

Ce fichier de mémorisation n'est utile qu'en usage interactif (dans un vrai terminal). Si tu lances le script depuis un script planifié ou sans terminal, ajoute `--non-interactif` : les commerçants inconnus resteront simplement "À catégoriser" sans bloquer.

## Le fichier Excel généré

Deux pages :

- **Travail** : le détail de toutes les écritures du relevé (libellé court + libellé complet, catégorie, montant). C'est la page où tu vérifies et corriges — tu peux changer une catégorie directement dans une cellule.
- **Total** : le total par catégorie (formules `SUMIF`), plus une ligne "Dépenses réelles" qui exclut virements internes et revenus.

**Relancer le script sur le même relevé ne perd jamais tes corrections.** Le script relit d'abord la page Travail existante : toute écriture déjà présente garde la catégorie que tu lui as donnée (même si tu l'as changée à la main dans Excel), et cette catégorie est resynchronisée dans `categories.json`. Ensuite la page Total est entièrement reconstruite à partir de la page Travail, donc elle reflète toujours tes dernières corrections. Concrètement : tu corriges une catégorie dans Excel -> tu relances `python generer_rapport.py releve.pdf` -> la page Total se met à jour, et ce commerçant sera reconnu automatiquement les mois suivants.

## Limites connues

- Le script suppose que LCL garde la même mise en page de relevé ; si LCL change de modèle, l'extraction peut nécessiter un ajustement.
- Certains libellés très techniques (ex. "CDMB") ne veulent rien dire sans que tu saches ce que c'est : c'est justement pour ça qu'on te demande plutôt que de deviner.
