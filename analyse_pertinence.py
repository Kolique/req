#!/usr/bin/env python3
"""
Analyse automatique de fichiers Excel de trames LoRaWAN.

Pour chaque fichier (traité séparément) :
  1. Suppression des trames invalides (SF < 7 n'existe pas en LoRaWAN).
  2. Suppression des doublons de DevEUI (on garde la trame la plus récente).
  3. Classification de chaque capteur selon sa pertinence :
       - Indispensable   : Redondance = 1 (quel que soit le SF)
       - Pertinence ++   : Redondance = 2 et SF dans {7, 8, 9}
       - Pertinence +    : Redondance = 2 et SF > 9
       - Non pertinent   : Redondance > 5
       - À définir       : tout ce qui ne rentre dans aucune règle (ex. Redondance 3 à 5)
  4. Export d'un fichier Excel de résultats avec une feuille par catégorie
     et une feuille de synthèse.

Utilisation :
    python3 analyse_pertinence.py fichier1.xlsx fichier2.xlsx ...
    python3 analyse_pertinence.py dossier/          # traite tous les .xlsx du dossier
    python3 analyse_pertinence.py                   # traite tous les .xlsx du dossier courant
"""

import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Ordre d'affichage des catégories dans la synthèse et les feuilles
CATEGORIES = ["Indispensable", "Pertinence ++", "Pertinence +", "Non pertinent", "À définir"]


def classer(redondance, sf) -> str:
    """Applique les règles de pertinence à une trame."""
    if redondance > 5:
        return "Non pertinent"
    if redondance == 1:
        return "Indispensable"
    if redondance == 2 and sf in (7, 8, 9):
        return "Pertinence ++"
    if redondance == 2 and sf > 9:
        return "Pertinence +"
    return "À définir"


def analyser_fichier(chemin: Path) -> pd.DataFrame:
    """Lit un fichier Excel, dédoublonne les DevEUI et classe chaque capteur."""
    df = pd.read_excel(chemin)

    colonnes_requises = {"DevEUI", "Redondance", "SF"}
    manquantes = colonnes_requises - set(df.columns)
    if manquantes:
        raise ValueError(f"Colonnes manquantes dans {chemin.name} : {', '.join(sorted(manquantes))}")

    nb_avant = len(df)

    # Les trames avec SF < 7 sont invalides (le SF LoRaWAN va de 7 à 12)
    invalides = df["SF"] < 7
    if invalides.any():
        print(f"  {invalides.sum()} trame(s) invalide(s) ignorée(s) (SF < 7)")
        df = df[~invalides]

    # Dédoublonnage : une seule ligne par DevEUI, on garde la trame la plus récente
    if "Heure" in df.columns:
        df = df.sort_values("Heure")
    df = df.drop_duplicates(subset="DevEUI", keep="last").reset_index(drop=True)

    print(f"  {nb_avant} trames -> {len(df)} DevEUI uniques ({nb_avant - len(df)} doublons supprimés)")

    df["Pertinence"] = [classer(r, s) for r, s in zip(df["Redondance"], df["SF"])]
    return df


def exporter_resultats(df: pd.DataFrame, sortie: Path) -> None:
    """Écrit le fichier Excel de résultats : synthèse + une feuille par catégorie."""
    synthese = (
        df["Pertinence"]
        .value_counts()
        .reindex(CATEGORIES, fill_value=0)
        .rename_axis("Pertinence")
        .reset_index(name="Nombre de capteurs")
    )
    synthese["%"] = (synthese["Nombre de capteurs"] / len(df) * 100).round(1)

    with pd.ExcelWriter(sortie, engine="openpyxl") as writer:
        synthese.to_excel(writer, sheet_name="Synthèse", index=False)
        df.to_excel(writer, sheet_name="Tous les capteurs", index=False)
        for cat in CATEGORIES:
            sous_df = df[df["Pertinence"] == cat]
            if not sous_df.empty:
                # Les noms de feuille Excel n'acceptent ni '+' répétés ni > 31 caractères
                nom = cat.replace("++", "plus-plus").replace("+", "plus")[:31]
                sous_df.to_excel(writer, sheet_name=nom, index=False)

    print(f"  Résultats écrits dans : {sortie}")
    for _, ligne in synthese.iterrows():
        print(f"    {ligne['Pertinence']:<15} {ligne['Nombre de capteurs']:>5}  ({ligne['%']} %)")


def main() -> None:
    args = sys.argv[1:]

    # Constitution de la liste des fichiers à traiter
    fichiers: list[Path] = []
    cibles = [Path(a) for a in args] if args else [Path(".")]
    for cible in cibles:
        if cible.is_dir():
            fichiers.extend(sorted(cible.glob("*.xlsx")))
        elif cible.is_file():
            fichiers.append(cible)
        else:
            print(f"Ignoré (introuvable) : {cible}")

    # On ne retraite pas nos propres fichiers de sortie
    fichiers = [f for f in fichiers if not f.stem.endswith("_resultats")]

    if not fichiers:
        print("Aucun fichier .xlsx à traiter.")
        sys.exit(1)

    for fichier in fichiers:
        print(f"\nTraitement de : {fichier.name}")
        try:
            df = analyser_fichier(fichier)
        except Exception as e:
            print(f"  ERREUR : {e}")
            continue
        sortie = fichier.with_name(f"{fichier.stem}_resultats.xlsx")
        exporter_resultats(df, sortie)


if __name__ == "__main__":
    main()
