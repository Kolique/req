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
Dès que plusieurs fichiers sont traités ensemble (les antennes d'un même
contrat), la pertinence est calculée par recoupement entre antennes :
1 fichier = 1 antenne, et pour chaque DevEUI de chaque antenne on cherche
ce DevEUI dans les AUTRES antennes du contrat. La redondance réelle est le
nombre d'antennes du contrat qui reçoivent le capteur : s'il n'apparaît
dans aucune autre antenne, c'est un très bon signal -> Indispensable
(Note 1), même si la colonne Redondance du fichier est > 1 (elle compte
seulement le nombre de fois où l'antenne l'a entendu dans la journée).
Cette pertinence recalculée figure dans le rapport de chaque antenne
(avec les colonnes "Nb antennes" et "Vue aussi par"), et un rapport global
du contrat est produit en plus. Avec un seul fichier, la colonne Redondance
du fichier est utilisée telle quelle.

  4. Export d'un fichier Excel de résultats :
       - Feuille "Synthèse"      : période des données, chiffres clés,
                                   répartition des pertinences + graphique
       - Feuille "Statistiques"  : distribution des SF + graphique,
                                   qualité du signal (RSSI/SNR) par catégorie,
                                   activité des capteurs (nb de trames)
       - Feuille "Tous les capteurs" + une feuille par catégorie

Suivi dans le temps : chaque exécution est historisée dans deux fichiers CSV
(<contrat>-historique-syntheses.csv et <contrat>-historique-capteurs.csv) et
le fichier <contrat>-suivi.xlsx est régénéré avec la courbe d'évolution des
pertinences et la liste des capteurs ayant changé (nouveau, disparu,
amélioration, dégradation) depuis l'analyse précédente. Relancer sur les
mêmes données (même date) remplace la ligne du jour au lieu de la dupliquer.

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
# Note chiffrée associée (1 = très bon signal)
NOTES = {
    "Indispensable": 1,
    "Pertinence +++": 2,
    "Pertinence ++": 3,
    "Pertinence +": 4,
    "Non pertinent": 5,
    "À définir": None,
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
    df["Note"] = [NOTES[p] for p in df["Pertinence"]]

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


def analyse_globale(resultats: list) -> pd.DataFrame:
    """Analyse croisée multi-antennes : 1 fichier = 1 antenne.

    Pour chaque DevEUI, la redondance réelle est le nombre d'antennes
    différentes qui le reçoivent. Un capteur vu par une seule antenne
    est indispensable. Le SF retenu est le meilleur (le plus bas)
    observé parmi les antennes.
    """
    trames = []
    for chemin, df in resultats:
        d = df.copy()
        d["Antenne"] = chemin.stem
        trames.append(d)
    tout = pd.concat(trames, ignore_index=True)

    agregats = {
        "Nb antennes": ("Antenne", "nunique"),
        "Antennes": ("Antenne", lambda s: ", ".join(sorted(set(s)))),
        "SF": ("SF", "min"),
        "Nb trames": ("Nb trames", "sum"),
    }
    if "Heure" in tout.columns:
        agregats["Dernière trame"] = ("Heure", "max")

    capteurs = tout.groupby("DevEUI").agg(**agregats).reset_index()
    capteurs["Pertinence"] = [classer(n, s) for n, s in zip(capteurs["Nb antennes"], capteurs["SF"])]
    capteurs["Note"] = [NOTES[p] for p in capteurs["Pertinence"]]
    return capteurs


def exporter_globale(resultats: list, dossier: Path) -> None:
    capteurs = analyse_globale(resultats)

    meta = {
        "Type d'analyse": "Globale multi-antennes (redondance = nb d'antennes recevant le capteur)",
        "Date de l'analyse": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Antennes analysées": ", ".join(chemin.stem for chemin, _ in resultats),
        "Nombre d'antennes": len(resultats),
        "Capteurs uniques (DevEUI)": len(capteurs),
        "Capteurs vus par une seule antenne": int((capteurs["Nb antennes"] == 1).sum()),
    }
    if "Dernière trame" in capteurs.columns:
        heures = pd.concat([df["Heure"] for _, df in resultats if "Heure" in df.columns])
        meta["Période des données"] = f"du {heures.min():%d/%m/%Y %H:%M} au {heures.max():%d/%m/%Y %H:%M}"
        suffixe = f"-{heures.max():%d%m}"
    else:
        suffixe = ""

    nom_contrat = dossier.resolve().name or "contrat"
    sortie = dossier / f"{nom_contrat}-analyse-globale{suffixe}.xlsx"
    print(f"\nAnalyse globale du contrat '{nom_contrat}' ({len(resultats)} antennes)")
    exporter_resultats(capteurs, meta, sortie)
    return capteurs


def mettre_a_jour_suivi(capteurs: pd.DataFrame, dossier: Path, date_donnees) -> None:
    """Historise l'analyse du jour et régénère le fichier de suivi du contrat.

    Deux fichiers CSV servent de mémoire entre les exécutions :
      - <contrat>-historique-syntheses.csv : une ligne par analyse (répartition)
      - <contrat>-historique-capteurs.csv  : la pertinence de chaque DevEUI à chaque analyse
    Le fichier <contrat>-suivi.xlsx est régénéré à chaque fois : courbe
    d'évolution des pertinences + capteurs ayant changé depuis l'analyse précédente.
    Relancer le script sur les mêmes données (même date) remplace la ligne du jour.
    """
    from openpyxl import Workbook
    from openpyxl.chart import LineChart

    nom = dossier.resolve().name or "contrat"
    date_str = f"{date_donnees:%Y-%m-%d}" if date_donnees is not None else datetime.now().strftime("%Y-%m-%d")

    # 1. Historique des synthèses (une ligne par date de données)
    compte = capteurs["Pertinence"].value_counts().reindex(CATEGORIES, fill_value=0)
    ligne = {"Date": date_str, "Capteurs": len(capteurs), **{c: int(n) for c, n in compte.items()}}
    chemin_synth = dossier / f"{nom}-historique-syntheses.csv"
    if chemin_synth.exists():
        historique = pd.read_csv(chemin_synth)
        historique = historique[historique["Date"] != date_str]
        historique = pd.concat([historique, pd.DataFrame([ligne])], ignore_index=True)
    else:
        historique = pd.DataFrame([ligne])
    historique = historique.sort_values("Date").reset_index(drop=True)
    historique.to_csv(chemin_synth, index=False)

    # 2. Historique par capteur (pour détecter les changements de pertinence)
    detail = capteurs[["DevEUI", "Pertinence", "Note"]].copy()
    detail.insert(0, "Date", date_str)
    chemin_capt = dossier / f"{nom}-historique-capteurs.csv"
    if chemin_capt.exists():
        hist_capteurs = pd.read_csv(chemin_capt)
        hist_capteurs = hist_capteurs[hist_capteurs["Date"] != date_str]
        hist_capteurs = pd.concat([hist_capteurs, detail], ignore_index=True)
    else:
        hist_capteurs = detail
    hist_capteurs.to_csv(chemin_capt, index=False)

    # 3. Fichier de suivi : évolution + changements
    wb = Workbook()

    ws = wb.active
    ws.title = "Évolution"
    ws["A1"] = f"Suivi de la pertinence — contrat {nom}"
    ws["A1"].font = TITRE
    ecrire_tableau(ws, 3, "Historique des analyses", historique)

    graph = LineChart()
    graph.title = "Évolution du nombre de capteurs par pertinence"
    graph.y_axis.title = "Nombre de capteurs"
    # Colonnes : A=Date, B=Capteurs, C.. = catégories (lignes 4=en-têtes, 5..=données)
    data = Reference(ws, min_col=3, max_col=2 + len(CATEGORIES), min_row=4, max_row=4 + len(historique))
    graph.add_data(data, titles_from_data=True)
    graph.set_categories(Reference(ws, min_col=1, min_row=5, max_row=4 + len(historique)))
    for i, cat in enumerate(CATEGORIES):
        serie = graph.series[i]
        serie.graphicalProperties.line.solidFill = COULEURS[cat]
        serie.graphicalProperties.line.width = 25000  # ~2 pt
        serie.smooth = False
    graph.height, graph.width = 10, 22
    ws.add_chart(graph, f"A{6 + len(historique)}")
    ws.column_dimensions["A"].width = 14

    ws2 = wb.create_sheet("Changements")
    dates = sorted(hist_capteurs["Date"].unique())
    if len(dates) < 2:
        ws2["A1"] = "Première analyse : les changements apparaîtront à partir de la prochaine exécution."
    else:
        avant, apres = dates[-2], dates[-1]
        p_avant = hist_capteurs[hist_capteurs["Date"] == avant].set_index("DevEUI")["Pertinence"]
        p_apres = hist_capteurs[hist_capteurs["Date"] == apres].set_index("DevEUI")["Pertinence"]
        lignes = []
        for eui in sorted(set(p_avant.index) | set(p_apres.index)):
            av, ap = p_avant.get(eui), p_apres.get(eui)
            if av is None:
                lignes.append({"DevEUI": eui, "Avant": "—", "Après": ap, "Changement": "Nouveau capteur"})
            elif ap is None:
                lignes.append({"DevEUI": eui, "Avant": av, "Après": "—", "Changement": "Capteur disparu"})
            elif av != ap:
                sens = "Amélioration" if (NOTES.get(ap) or 9) < (NOTES.get(av) or 9) else "Dégradation"
                lignes.append({"DevEUI": eui, "Avant": av, "Après": ap, "Changement": sens})
        ws2["A1"] = f"Changements entre le {avant} et le {apres}"
        ws2["A1"].font = GRAS
        if lignes:
            ecrire_tableau(ws2, 3, f"{len(lignes)} capteur(s) concerné(s)", pd.DataFrame(lignes))
        else:
            ws2["A3"] = "Aucun changement de pertinence."
        for col, largeur in zip("ABCD", (26, 16, 16, 20)):
            ws2.column_dimensions[col].width = largeur

    sortie = dossier / f"{nom}-suivi.xlsx"
    wb.save(sortie)
    print(f"  Suivi mis à jour : {sortie} ({len(historique)} analyse(s) dans l'historique)")


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
    fichiers = [f for f in fichiers
                if not f.stem.endswith(("-analyse", "_resultats", "-suivi"))
                and "-analyse-globale" not in f.stem]

    if not fichiers:
        print("Aucun fichier .xlsx à traiter.")
        sys.exit(1)

    # 1re passe : lecture, nettoyage et dédoublonnage de chaque antenne
    resultats = []
    for fichier in fichiers:
        print(f"\nTraitement de : {fichier.name}")
        try:
            df, meta = analyser_fichier(fichier)
        except Exception as e:
            print(f"  ERREUR : {e}")
            continue
        resultats.append([fichier, df, meta])

    # Avec plusieurs antennes (même contrat) : pour chaque DevEUI de chaque
    # antenne, on cherche ce DevEUI dans les autres antennes du contrat.
    # La pertinence est alors recalculée avec cette redondance réelle
    # (1 seule antenne = très bon signal = Indispensable, Note 1).
    if len(resultats) >= 2:
        vu_par: dict = {}
        for fichier, df, _ in resultats:
            for eui in df["DevEUI"]:
                vu_par.setdefault(eui, set()).add(fichier.stem)

        for element in resultats:
            fichier, df, meta = element
            autres = [", ".join(sorted(vu_par[e] - {fichier.stem})) for e in df["DevEUI"]]
            df["Nb antennes"] = [len(vu_par[e]) for e in df["DevEUI"]]
            df["Vue aussi par"] = [a if a else "Aucune autre antenne" for a in autres]
            # La colonne Redondance du fichier n'est pas utilisée ici : elle
            # compte le nombre de fois où l'antenne a entendu le capteur dans
            # la journée, pas le nombre d'antennes. Seul le nombre d'antennes
            # du contrat qui voient le capteur fait la pertinence.
            df["Pertinence"] = [classer(n, s) for n, s in zip(df["Nb antennes"], df["SF"])]
            df["Note"] = [NOTES[p] for p in df["Pertinence"]]
            # La pertinence en dernière colonne pour rester lisible
            df = df[[c for c in df.columns if c not in ("Pertinence", "Note")] + ["Pertinence", "Note"]]
            element[1] = df
            meta["Antennes du contrat"] = ", ".join(f.stem for f, _, _ in resultats)
            meta["Capteurs vus uniquement par cette antenne"] = int((df["Nb antennes"] == 1).sum())

    # 2e passe : export du rapport de chaque antenne
    for fichier, df, meta in resultats:
        # Nom de sortie : <fichier>-<date des données JJMM>-analyse.xlsx
        if "Heure" in df.columns and df["Heure"].notna().any():
            suffixe = f"-{df['Heure'].max():%d%m}-analyse"
        else:
            suffixe = "-analyse"
        sortie = fichier.with_name(f"{fichier.stem}{suffixe}.xlsx")
        print(f"\nRapport de l'antenne : {fichier.stem}")
        exporter_resultats(df, meta, sortie)

    # Rapport global du contrat en plus des rapports par antenne,
    # puis mise à jour du suivi dans le temps
    if len(resultats) >= 2:
        dossier = resultats[0][0].parent
        capteurs = exporter_globale([(f, df) for f, df, _ in resultats], dossier)
        heures = [df["Heure"].max() for _, df, _ in resultats if "Heure" in df.columns]
        mettre_a_jour_suivi(capteurs, dossier, max(heures) if heures else None)
    elif resultats:
        fichier, df, _ = resultats[0]
        date_donnees = df["Heure"].max() if "Heure" in df.columns else None
        mettre_a_jour_suivi(df, fichier.parent, date_donnees)


if __name__ == "__main__":
    main()
