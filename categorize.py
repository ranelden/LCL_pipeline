"""
Catégorisation des transactions, décidée par toi et mémorisée dans un
fichier JSON (par défaut categories.json à côté du script) :

- "categories" : la liste de tes catégories, dans l'ordre d'affichage du
  menu. Modifiable directement dans le JSON (ajouter/renommer/réordonner).
- "mappings" : association "libellé nettoyé (MAJUSCULES)" -> catégorie,
  remplie au fur et à mesure que tu réponds aux questions posées en
  console, et aussi resynchronisée depuis tes corrections manuelles dans
  la page de travail de l'Excel (voir generer_rapport.py).

Aucune IA, aucune règle de mots-clés cachée : la toute première fois qu'un
commerçant apparaît, le script demande ; ensuite il s'en souvient.
"""
import json
import sys
from pathlib import Path

DEFAULT_CATEGORIES = [
    "Alimentation",
    "Logement",
    "Transport",
    "Abonnements",
    "Loisirs",
    "Santé",
    "Shopping",
    "Admin",
    "Frais bancaires",
    "Virements internes / Épargne",
    "Revenus",
]

SKIP_CATEGORY = "À catégoriser"


def load_store(path: Path) -> dict:
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            store = json.load(f)
        store.setdefault("categories", list(DEFAULT_CATEGORIES))
        store.setdefault("mappings", {})
        return store
    return {"categories": list(DEFAULT_CATEGORIES), "mappings": {}}


def save_store(store: dict, path: Path) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2, sort_keys=False)


def _print_menu(store: dict) -> None:
    for i, cat in enumerate(store["categories"], start=1):
        print(f"  {i}) {cat}")
    print("  0) Passer (à catégoriser plus tard)")
    print("  n) Nouvelle catégorie")


def prompt_category(clean: str, raw: str, montant: float, store: dict) -> str:
    """Pose la question en console. Retourne la catégorie choisie
    (SKIP_CATEGORY si l'utilisateur passe)."""
    print(f"\nCatégorie pour : {clean}  ({montant:+.2f} €)")
    if raw != clean:
        print(f"  (libellé complet : {raw})")
    while True:
        _print_menu(store)
        choice = input("Choix : ").strip().lower()
        if choice == "0" or choice == "":
            return SKIP_CATEGORY
        if choice == "n":
            new_cat = input("Nom de la nouvelle catégorie : ").strip()
            if not new_cat:
                continue
            if new_cat not in store["categories"]:
                store["categories"].append(new_cat)
            return new_cat
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(store["categories"]):
                return store["categories"][idx - 1]
        print("Choix non reconnu, réessaie.")


def resolve_category(clean: str, raw: str, montant: float, store: dict, interactive: bool = True) -> str:
    key = clean.upper()
    if key in store["mappings"]:
        return store["mappings"][key]

    if interactive and sys.stdin.isatty():
        category = prompt_category(clean, raw, montant, store)
        if category != SKIP_CATEGORY:
            store["mappings"][key] = category
        return category

    return SKIP_CATEGORY


def remember_correction(clean: str, category: str, store: dict) -> None:
    """Enregistre une catégorie déjà choisie (ex. relue depuis une
    correction manuelle faite dans l'Excel), sans redemander en console."""
    if category and category != SKIP_CATEGORY:
        store["mappings"][clean.upper()] = category
