#!/usr/bin/env python3
"""
Analyse automatique de fichiers Excel de trames LoRaWAN.

Pour chaque fichier (traité séparément) :
  1. Suppression des trames invalides (SF < 7 n'existe pas en LoRaWAN).
  2. Suppression des doublons de DevEUI (on garde la trame la plus récente).
  3. Classification de chaque capteur selon sa pertinence :
       - Indispensable   : Redondance = 1 (quel que soit le SF)
       - Pertinence +++  : Redondance = 2 et SF dans {7, 8, 9}
       - Pertinence ++   : Redondance = 2 et SF > 9
       - Pertinence +    : Redondance 3 ou 4 (le SF n'est pas pris en compte)
       - Non pertinent   : Redondance > 5
       - À définir       : tout ce qui ne rentre dans aucune règle (ex. Redondance 5)
  4. Export d'un fichier Excel de résultats :
       - Feuille "Synthèse"      : période des données, chiffres clés,
                                   répartition des pertinences + graphique
       - Feuille "Statistiques"  : distribution des SF + graphique,
                                   qualité du signal (RSSI/SNR) par catégorie,
                                   activité des capteurs (nb de trames)
       - Feuille "Tous les capteurs" + une feuille par catégorie

Utilisation :
    python3 analyse_pertinence.py fichier1.xlsx fichier2.xlsx ...
    python3 analyse_pertinence.py dossier/          # traite tous les .xlsx du dossier
    python3 analyse_pertinence.py                   # traite tous les .xlsx du dossier courant
"""

import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Ordre d'affichage des catégories et couleur associée (vert -> rouge, gris = à définir)
CATEGORIES = ["Indispensable", "Pertinence +++", "Pertinence ++", "Pertinence +",
              "Non pertinent", "À définir"]
COULEURS = {
    "Indispensable": "0CA30C",
    "Pertinence +++": "FAB219",
    "Pertinence ++": "EC835A",
    "Pertinence +": "D03B3B",
    "Non pertinent": "8B1A1A",
    "À définir": "8C8C8C",
}
# Noms de feuille Excel (pas de '+' ni plus de 31 caractères)
NOMS_FEUILLE = {
    "Indispensable": "Indispensable",
    "Pertinence +++": "Pertinence plus-plus-plus",
    "Pertinence ++": "Pertinence plus-plus",
    "Pertinence +": "Pertinence plus",
    "Non pertinent": "Non pertinent",
    "À définir": "À définir",
}

GRAS = Font(bold=True)
TITRE = Font(bold=True, size=14)


def classer(redondance, sf) -> str:
    """Applique les règles de pertinence à une trame."""
    if redondance > 5:
        return "Non pertinent"
    if redondance == 1:
        return "Indispensable"
    if redondance == 2 and sf in (7, 8, 9):
        return "Pertinence +++"
    if redondance == 2 and sf > 9:
        return "Pertinence ++"
    if redondance in (3, 4):
        return "Pertinence +"
    return "À définir"


def analyser_fichier(chemin: Path):
    """Lit un fichier Excel, dédoublonne les DevEUI et classe chaque capteur.

    Retourne (df dédoublonné et classé, dictionnaire de métadonnées).
    """
    df = pd.read_excel(chemin)

    colonnes_requises = {"DevEUI", "Redondance", "SF"}
    manquantes = colonnes_requises - set(df.columns)
    if manquantes:
        raise ValueError(f"Colonnes manquantes dans {chemin.name} : {', '.join(sorted(manquantes))}")

    meta = {"Fichier source": chemin.name,
            "Date de l'analyse": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Trames lues": len(df)}

    # Les trames avec SF < 7 sont invalides (le SF LoRaWAN va de 7 à 12)
    df = df[df["SF"] >= 7]

    # Période couverte par les données
    if "Heure" in df.columns:
        debut, fin = df["Heure"].min(), df["Heure"].max()
        meta["Période des données"] = f"du {debut:%d/%m/%Y %H:%M} au {fin:%d/%m/%Y %H:%M}"
        df = df.sort_values("Heure")

    # Nombre de trames émises par capteur (avant dédoublonnage) : mesure d'activité
    nb_trames = df.groupby("DevEUI").size().rename("Nb trames")

    # Dédoublonnage : une seule ligne par DevEUI, on garde la trame la plus récente
    nb_avant = len(df)
    df = df.drop_duplicates(subset="DevEUI", keep="last").reset_index(drop=True)
    meta["Doublons DevEUI supprimés"] = nb_avant - len(df)
    meta["Capteurs uniques (DevEUI)"] = len(df)

    df = df.merge(nb_trames, on="DevEUI")
    df["Pertinence"] = [classer(r, s) for r, s in zip(df["Redondance"], df["SF"])]

    print(f"  {meta['Trames lues']} trames -> {len(df)} DevEUI uniques "
          f"({meta['Doublons DevEUI supprimés']} doublons supprimés)")
    return df, meta


def ecrire_tableau(ws, ligne, titre, table: pd.DataFrame) -> int:
    """Écrit un titre + un DataFrame dans la feuille à partir de `ligne` (1-indexé).

    Retourne la première ligne libre après le tableau.
    """
    ws.cell(row=ligne, column=1, value=titre).font = GRAS
    ligne += 1
    for j, col in enumerate(table.columns, start=1):
        ws.cell(row=ligne, column=j, value=col).font = GRAS
    for i, (_, valeurs) in enumerate(table.iterrows(), start=1):
        for j, v in enumerate(valeurs, start=1):
            ws.cell(row=ligne + i, column=j, value=v)
    return ligne + len(table) + 2


def feuille_synthese(wb, df: pd.DataFrame, meta: dict) -> None:
    ws = wb.create_sheet("Synthèse", 0)

    ws["A1"] = "Analyse de pertinence des capteurs LoRaWAN"
    ws["A1"].font = TITRE

    # Bloc d'informations générales (dont la période des données)
    ligne = 3
    for cle, valeur in meta.items():
        ws.cell(row=ligne, column=1, value=cle).font = GRAS
        ws.cell(row=ligne, column=2, value=valeur)
        ligne += 1

    # Tableau de répartition des pertinences
    compte = df["Pertinence"].value_counts().reindex(CATEGORIES, fill_value=0)
    table = pd.DataFrame({
        "Pertinence": compte.index,
        "Nombre de capteurs": compte.values,
        "%": (compte.values / len(df) * 100).round(1),
    })
    debut_table = ligne + 1
    ecrire_tableau(ws, debut_table, "Répartition par pertinence", table)

    # Graphique en secteurs de la répartition
    pie = PieChart()
    pie.title = "Répartition des capteurs par pertinence"
    data = Reference(ws, min_col=2, min_row=debut_table + 1, max_row=debut_table + 1 + len(table))
    labels = Reference(ws, min_col=1, min_row=debut_table + 2, max_row=debut_table + 1 + len(table))
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    serie = pie.series[0]
    for i, cat in enumerate(CATEGORIES):
        point = DataPoint(idx=i)
        point.graphicalProperties.solidFill = COULEURS[cat]
        serie.data_points.append(point)
    pie.height, pie.width = 9, 13
    ws.add_chart(pie, "E3")

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 8


def feuille_statistiques(wb, df: pd.DataFrame) -> None:
    ws = wb.create_sheet("Statistiques")

    # Distribution des SF
    sf = df["SF"].value_counts().sort_index()
    table_sf = pd.DataFrame({"SF": sf.index, "Nombre de capteurs": sf.values})
    ligne = ecrire_tableau(ws, 1, "Distribution des SF (Spreading Factor)", table_sf)

    bar = BarChart()
    bar.type = "col"
    bar.title = "Nombre de capteurs par SF"
    bar.legend = None
    data = Reference(ws, min_col=2, min_row=2, max_row=2 + len(table_sf))
    labels = Reference(ws, min_col=1, min_row=3, max_row=2 + len(table_sf))
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(labels)
    bar.series[0].graphicalProperties.solidFill = "4472C4"
    bar.height, bar.width = 8, 12
    ws.add_chart(bar, "E1")

    # Qualité du signal par catégorie de pertinence
    if {"RSSI", "SNR"} <= set(df.columns):
        stats = (
            df.groupby("Pertinence")
            .agg(**{
                "Nb capteurs": ("DevEUI", "count"),
                "RSSI moyen (dBm)": ("RSSI", "mean"),
                "RSSI min": ("RSSI", "min"),
                "RSSI max": ("RSSI", "max"),
                "SNR moyen (dB)": ("SNR", "mean"),
            })
            .reindex([c for c in CATEGORIES if c in df["Pertinence"].values])
            .round(1)
            .reset_index()
        )
        ligne = ecrire_tableau(ws, max(ligne, 18), "Qualité du signal par pertinence", stats)

    # Activité des capteurs (nombre de trames émises sur la période)
    actifs = df.nlargest(10, "Nb trames")[["DevEUI", "Nb trames", "SF", "Pertinence"]]
    ws.cell(row=ligne, column=1,
            value=f"Trames par capteur : moyenne {df['Nb trames'].mean():.1f}, "
                  f"médiane {df['Nb trames'].median():.0f}, max {df['Nb trames'].max()}").font = GRAS
    ecrire_tableau(ws, ligne + 1, "Top 10 des capteurs les plus actifs", actifs)

    for col, largeur in zip("ABCDEF", (26, 18, 12, 12, 12, 16)):
        ws.column_dimensions[col].width = largeur


def exporter_resultats(df: pd.DataFrame, meta: dict, sortie: Path) -> None:
    """Écrit le fichier Excel de résultats : synthèse, statistiques et détail."""
    with pd.ExcelWriter(sortie, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Tous les capteurs", index=False)
        for cat in CATEGORIES:
            sous_df = df[df["Pertinence"] == cat]
            if not sous_df.empty:
                sous_df.to_excel(writer, sheet_name=NOMS_FEUILLE[cat], index=False)

        feuille_synthese(writer.book, df, meta)
        feuille_statistiques(writer.book, df)

    print(f"  Résultats écrits dans : {sortie}")
    compte = df["Pertinence"].value_counts().reindex(CATEGORIES, fill_value=0)
    for cat, nb in compte.items():
        print(f"    {cat:<15} {nb:>5}  ({nb / len(df) * 100:.1f} %)")


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
    fichiers = [f for f in fichiers if not f.stem.endswith(("-analyse", "_resultats"))]

    if not fichiers:
        print("Aucun fichier .xlsx à traiter.")
        sys.exit(1)

    for fichier in fichiers:
        print(f"\nTraitement de : {fichier.name}")
        try:
            df, meta = analyser_fichier(fichier)
        except Exception as e:
            print(f"  ERREUR : {e}")
            continue
        # Nom de sortie : <fichier>-<date des données JJMM>-analyse.xlsx
        if "Heure" in df.columns and df["Heure"].notna().any():
            suffixe = f"-{df['Heure'].max():%d%m}-analyse"
        else:
            suffixe = "-analyse"
        sortie = fichier.with_name(f"{fichier.stem}{suffixe}.xlsx")
        exporter_resultats(df, meta, sortie)


if __name__ == "__main__":
    main()
