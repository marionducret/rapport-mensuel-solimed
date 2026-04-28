#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSAR Tool - Génération automatique du rapport PDF mensuel SSR
Streamlit Cloud version : toutes les données sont chargées via load_data().
"""
#%%
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
import matplotlib.backends.backend_pdf as pdf_backend
from pathlib import Path
from datetime import datetime
import numpy as np
import re
import zipfile
import tempfile
import io
import os
from PIL import Image
import textwrap
from matplotlib import font_manager

BASE_DIR = Path(__file__).parent

#%%
# ══════════════════════════════════════════════════════════════════════════════
#  SECTION CONFIGURATION — tout ce qui est paramétrable est ici
# ══════════════════════════════════════════════════════════════════════════════

OUTPUT_PDF  = "rapport_mensuel.pdf"

# ── Templates Canva ───────────────────────────────────────────────────────────
# Déposer les PNG exportés depuis Canva dans ./design/
CANVA_COVER_PATH = "design/page_garde_all.png"
CANVA_COVER_PATH_HC = "design/page_garde_HC.png"
CANVA_PAGE_HC_PATH  = "design/page_graph_HC_pays.png"
CANVA_PAGE_HTP_PATH   = "design/page_graph_HTP_pays.png"

AUTEUR = "Dr Nathalie DUCRET"
DATE_RAPPORT = datetime.today().strftime("%d/%m/%Y")
SERVICE = "Rapport évolution mensuelle SMR"

#à automatiser
OBJECTIFS = {
    "obj_AM_mois": 0,
    "obj_BR_mois": 0
}

KPI_CONFIG = [
    ("recette_BR_period",   "Recette Base Remboursement cumulée", "{:.0f} €",  "obj_BR_mois"),
    ("montantAM_valorise_HC",   "Recette Assurance Maladie cumulée", "{:.0f} €",  "obj_AM_mois"),
    ("effectif_transmis_HC",  "Séjours HC transmis",      "{:.0f}",     None),
    ("effectif_transmis_HTP",  "Jours HTP transmis",      "{:.0f}",     None),
    ("recette_BR_moy_sej",    "Recette Base Remboursement moyenne par jour (HC)", "{:.0f} €",  None),
    ("recette_BR_moy_jour",    "Recette Base Remboursement moyenne par jour (HTP)", "{:.0f} €",  None),
    ("taux_valorisation_HC",  "Taux de valorisation séjours HC",  "{:.1f} %",   None),
    ("taux_valorisation_HTP",  "Taux de valorisation jours HTP",  "{:.1f} %",   None),
  
]

# KPI_CONFIG_HC = [
#     ("recette_BR_period",   "Recette Base Remboursement cumulée", "{:.0f} €",  "obj_BR_mois"),
#     ("montantAM_valorise_HC",   "Recette Assurance Maladie cumulée", "{:.0f} €",  "obj_AM_mois"),
#     ("effectif_transmis_HC",  "Séjours HC transmis",      "{:.0f}",     None),
#     ("recette_BR_moy_jour",    "Recette Base Remboursement moyenne par jour (HC)", "{:.0f} €",  None),
#     ("taux_valorisation_HC",  "Taux de valorisation séjours HC",  "{:.1f} %",   None),
# ]


KPI_CONFIG_HC = [
    (
        "recette_BR_cumule_total",
        "Recette BR cumulée",
        "{:,.0f} €",
        "recette_BR_mois_total",
        "{:,.0f} € sur le mois",
        "obj_BR_mois",
    ),
    (
        "montantAM_valorise_HC",
        "Recette AM cumulée",
        "{:,.0f} €",
        "montantAM_mois_HC",
        "{:,.0f} € sur le mois",
        "obj_AM_mois",
    ),
    (
        "effectif_transmis_HC",
        "Séjours HC transmis cumulés",
        "{:,.0f}",
        "sejours_transmis_mois_HC",
        "{:,.0f} sur le mois",
        None,
    ),
    (
        "recette_BR_moy_jour_cumule_HC",
        "BR moyen / jour cumulé",
        "{:,.0f} €",
        "recette_BR_moy_jour_mois_HC",
        "{:,.0f} € sur le mois",
        None,
    ),
    (
        "taux_valorisation_cumule_HC",
        "Taux de valorisation cumulé",
        "{:.1f} %",
        "taux_valorisation_mois_HC",
        "{:.1f} % sur le mois",
        None,
    ),
]

THEMES = {
    "HC ": {
        "plots": [
            {
                "type": "multi",
                "series": [
                    (
                        "sejours_valorises_mois_HC",
                        "Séjours valorisés"
                    ),
                    (
                        "sejours_transmis_mois_HC",
                        "Séjours transmis"
                    ),
                ],
                "title": "Activité du mois : séjours",
            },
            {
                "type": "bar",
                "series": [
                    (
                        "taux_valorisation_mois_HC",
                        "Taux du mois"
                    ),
                    (
                        "taux_valorisation_cumule_HC",
                        "Taux cumulé"
                    ),
                ],
                "title": "Taux de valorisation (séjours valorisés/séjour transmis)",
            },
            {
                "type": "single_hlines",
                "objectif": None,
                "series": [
                    (
                        "recette_BR_moy_jour_cumule_HC",
                        "BR cumulé / jours valorisés cumulés"
                    ),
                ],
                "title": "Recette Base Remboursement moyenne par jour (cumul sur la période)",
            },
        ]
    },
    "HTP ": {
        "plots": [
              {
                "type": "multi",
                "series": [("jour_valo_supp", "Jour valorisé supplémentaire par rapport à M-1"),
                           ("jour_tot_supp",  "Jour supplémentaire par rapport à M-1")],
                "title": "Evolution de l'activité (jours)",
            },
            {
                "type": "bar",
                "series": [("taux_valorisation_HTP", "Taux de valorisation")],
                "title": "Taux de Valorisation",
            },
            {
                "type": "single_hlines",
                "objectif": None,
                "series": [("recette_BR_moy_jour", "Evolution de la recette brute moyenne par jour")],
                "title": "Evolution de la recette Base Remboursement moyenne par jour",
            },
        ]
    },}


COLORS     = ["#2563EB", "#16A34A", "#16A34A", "#E11D48", "#E11D48"]
BLEU_FONCE = "#1E3A5F"
BLEU       = "#2563EB"
GRIS_TEXTE = "#6B7280"
ROUGE      = "#E11D48"
VERT       = "#16A34A"
BLANC      = "#FFFFFF"
TEAL       = "#028181"
VIOLET     = "#7C3AED"
ORANGE     = "#F09516"

def format_fr(val, fmt="{:,.0f}"):
    try:
        return fmt.format(float(val)).replace(",", " ")
    except Exception:
        return "N/A"

#%%
# ══════════════════════════════════════════════════════════════════════════════
#  UTILITAIRES CANVA
# ══════════════════════════════════════════════════════════════════════════════
def _charger_bg(path: str):
    candidates = [
        path,
        str(Path(__file__).parent / path),
        str(Path(__file__).parent / Path(path).name),
        str(Path(os.getcwd()) / path),
        str(Path(os.getcwd()) / Path(path).name),
        str(Path(os.getcwd()) / "design" / Path(path).name),
    ]
    for p in candidates:
        try:
            img = Image.open(p).convert("RGB")
            return np.array(img)
        except Exception:
            continue
    return None

def _appliquer_bg(fig: plt.Figure, bg_img) -> None:
    if bg_img is None:
        return
    ax_bg = fig.add_axes([0, 0, 1, 1], zorder=0)
    ax_bg.imshow(
        bg_img,
        aspect="auto",
        extent=[0, 1, 0, 1],
        zorder=0,
    )
    ax_bg.set_xlim(0, 1)
    ax_bg.set_ylim(0, 1)
    ax_bg.axis("off")
    ax_bg.set_navigate(False)


# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT DES DONNÉES
# ══════════════════════════════════════════════════════════════════════════════

def load_data(uploaded_zip, uploaded_excel):

    tmp = tempfile.TemporaryDirectory()
    tmp_path = Path(tmp.name)

    if hasattr(uploaded_zip, "read"):
        with zipfile.ZipFile(io.BytesIO(uploaded_zip.read()), "r") as zf:
            zf.extractall(tmp_path)
    else:
        with zipfile.ZipFile(uploaded_zip, "r") as zf:
            zf.extractall(tmp_path)

    valo_excel = pd.read_excel(uploaded_excel)

    def extract_month(folder_name):
        match = re.search(r"(202\d)_M(\d+)$", folder_name)
        if match:
            return f"{match.group(1)}_M{match.group(2)}"
        match = re.search(r"M(\d+)$", folder_name)
        if match:
            return f"2025_M{match.group(1)}"
        return None

    def month_key(m):
        year, month = m.split("_M")
        return (int(year), int(month))

    month_dirs = []
    for p in tmp_path.rglob("*"):
        if not p.is_dir():
            continue
        if "__MACOSX" in str(p):
            continue
        m = extract_month(p.name)
        if m:
            month_dirs.append((m, p))

    if not month_dirs:
        raise ValueError("❌ Aucun dossier mois détecté dans le ZIP")

    month_dirs_dict = {}
    for m, p in month_dirs:
        month_dirs_dict[m] = p

    sorted_months = sorted(month_dirs_dict.keys(), key=month_key)

    data = {}
    for month in sorted_months:
        folder = month_dirs_dict[month]
        html_files = list(folder.glob("*.html"))
        raev = next((f for f in html_files if "raev" in f.name), None)
        sv   = next((f for f in html_files if "sv"   in f.name), None)
        if not raev or not sv:
            print(f"⚠️ Mois {month} ignoré (fichiers manquants)")
            continue
        try:
            data[month] = {
                "raev": pd.read_html(raev)[1],
                "sv":   pd.read_html(sv)[0],
            }
        except Exception as e:
            print(f"⚠️ Erreur lecture {month}: {e}")
            continue

    if not data:
        raise ValueError("❌ Aucun mois exploitable (HTML non reconnus)")

    evol_rows = []
    for curr_mois in sorted(data.keys(), key=month_key):
        curr = data[curr_mois]["raev"]
        value_AM = curr.loc[
            curr["Zone de valorisation"].str.contains("TOTAL activité valorisée"),
            "Montant AM",
        ].iloc[0]
        value_AM = float(str(value_AM).replace(" ", "").replace(",", "."))

        curr2 = data[curr_mois]["sv"]
        curr2 = curr2.iloc[[0, 11]].copy()
        col_ssrha_br = [c for c in curr2.columns if "SSRHA" in c and "Montant BR" in c][0]
        col_htp_br   = [c for c in curr2.columns if "HTP"   in c and "Montant BR" in c][0]
        curr2 = curr2.rename(columns={
            col_ssrha_br: "SSRHA en HC - Montant BR",
            col_htp_br:   "Journées en HTP - Montant BR",
        })
        for col in ["SSRHA en HC - Montant BR", "Journées en HTP - Montant BR"]:
            curr2[col] = (
                curr2[col].astype(str)
                .str.replace(" ", "", regex=False)
                .str.replace(",", ".", regex=False)
            )
            curr2[col] = pd.to_numeric(curr2[col], errors="coerce")
        curr2.loc[
            curr2["Type d'activité"] == "Activité valorisée",
            "SSRHA en HC - Montant AM",
        ] = value_AM
        curr2["Mois"] = curr_mois

        df_month = curr2.pivot(index="Mois", columns="Type d'activité")
        df_month.columns = [f"{metric}_{act}" for metric, act in df_month.columns]
        df_month.columns = [
            "effectif_transmis_HC",
            "effectif_valorise_HC",
            "montantBR_transmis_HC",
            "montantBR_valorise_HC",
            "effectif_transmis_HTP",
            "effectif_valorise_HTP",
            "montantBR_transmis_HTP",
            "montantBR_valorise_HTP",
            "montantAM_transmis_HC",
            "montantAM_valorise_HC",
        ]
        jours_valo_HC = valo_excel[valo_excel["mois"] == curr_mois]["jours_valo"].values[0]
        df_month["jour_valo_HC"] = jours_valo_HC
        evol_rows.append(df_month)

    if not evol_rows:
        raise ValueError("❌ Aucun mois valide après traitement")

    evol_df = pd.concat(evol_rows)
    evol_df["taux_valorisation_HC"] = evol_df["effectif_valorise_HC"] / evol_df["effectif_transmis_HC"] * 100
    evol_df["recette_BR_moy_sej"]   = evol_df["montantBR_valorise_HC"] / evol_df["effectif_valorise_HC"]
    evol_df["recette_BR_moy_jour"]  = evol_df["montantBR_valorise_HC"] / evol_df["jour_valo_HC"]
    evol_df["ecart_valo"]           = evol_df["montantBR_valorise_HC"].diff()
    evol_df["sejour_supp"]          = evol_df["effectif_transmis_HC"].diff()
    evol_df["sejour_valo_supp"]     = evol_df["effectif_valorise_HC"].diff()
    evol_df["jour_valo_supp"]       = evol_df["jour_valo_HC"].diff()
    evol_df["recette_BR_moy_mois"]  = evol_df["montantBR_valorise_HC"].diff()
    evol_df["recette_AM_moy_mois"]  = evol_df["montantAM_valorise_HC"].diff()
    evol_df.loc[evol_df.index[0], "recette_BR_moy_mois"] = evol_df["montantBR_valorise_HC"].iloc[0]
    evol_df.loc[evol_df.index[0], "recette_AM_moy_mois"] = evol_df["montantAM_valorise_HC"].iloc[0]
    evol_df = evol_df.reset_index()
    evol_df["jour_tot_supp"] = 0

    PERIODE = f"{evol_df['Mois'].iloc[0]} → {evol_df['Mois'].iloc[-1]}"

    return {
        "evol_df":  evol_df,
        "PERIODE":  PERIODE,
        "_tmp_dir": tmp,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT INCRÉMENTAL — un seul mois, colonnes brutes uniquement
# ══════════════════════════════════════════════════════════════════════════════

def _calc_jours_valo(csv_file) -> float:
    """
    Calcule le nombre de jours valorisés à partir du CSV VisualValoSejours.
    Règles :
      - Filtre HOSP = C
      - Pour les lignes où NBJV_GMT = 90 ET MNT_BR_GMT = 0 ou vide :
        on neutralise NBJV_GMT (mis à 0) mais on conserve NBJV_GMTH
      - Retourne la somme NBJV_GMT + NBJV_GMTH
    """
    df = pd.read_csv(csv_file, sep=None, engine="python")
    for col in ["NBJV_GMT", "MNT_BR_GMT", "NBJV_GMTH"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df[df["HOSP"] == "C"].copy()
    masque_exclusion = (df["NBJV_GMT"] == 90) & (
        df["MNT_BR_GMT"].isna() | (df["MNT_BR_GMT"] == 0)
    )
    df.loc[masque_exclusion, "NBJV_GMT"] = 0   # neutralise les 90j GMT non facturés
    return float(df["NBJV_GMT"].sum() + df["NBJV_GMTH"].sum())


def load_data_brut(uploaded_zip, uploaded_csv):
    """
    Identique à load_data() mais retourne uniquement les colonnes BRUTES,
    sans les colonnes dérivées par .diff() (ecart_valo, sejour_supp, etc.).
    Utilisé en mode incrémental : on fusionne d'abord les brutes de tous les
    mois, puis on appelle recalculer_derives() sur la série complète.

    uploaded_csv : fichier CSV VisualValoSejours (remplace l'Excel de valorisation).
    """
    tmp = tempfile.TemporaryDirectory()
    tmp_path = Path(tmp.name)

    if hasattr(uploaded_zip, "read"):
        with zipfile.ZipFile(io.BytesIO(uploaded_zip.read()), "r") as zf:
            zf.extractall(tmp_path)
    else:
        with zipfile.ZipFile(uploaded_zip, "r") as zf:
            zf.extractall(tmp_path)

    jours_valo_mois = _calc_jours_valo(uploaded_csv)

    def extract_month(folder_name):
        match = re.search(r"(202\d)_M(\d+)$", folder_name)
        if match:
            return f"{match.group(1)}_M{match.group(2)}"
        match = re.search(r"M(\d+)$", folder_name)
        if match:
            return f"2025_M{match.group(1)}"
        return None

    def month_key(m):
        year, month = m.split("_M")
        return (int(year), int(month))

    month_dirs_dict = {}
    for p in tmp_path.rglob("*"):
        if not p.is_dir() or "__MACOSX" in str(p):
            continue
        m = extract_month(p.name)
        if m:
            month_dirs_dict[m] = p

    if not month_dirs_dict:
        # Fallback : fichiers à la racine du ZIP, détection depuis les noms de fichiers
        for f in tmp_path.glob("*.html"):
            match = re.search(r"\.(202\d)\.(\d+)\.", f.name)
            if match:
                month_dirs_dict[f"{match.group(1)}_M{match.group(2)}"] = tmp_path
                break

    if not month_dirs_dict:
        raise ValueError("❌ Aucun dossier mois détecté dans le ZIP")

    data = {}
    for month in sorted(month_dirs_dict.keys(), key=month_key):
        folder     = month_dirs_dict[month]
        html_files = list(folder.glob("*.html"))
        raev = next((f for f in html_files if "raev" in f.name), None)
        sv   = next((f for f in html_files if "sv"   in f.name), None)
        if not raev or not sv:
            print(f"⚠️ Mois {month} ignoré (fichiers manquants)")
            continue
        try:
            data[month] = {"raev": pd.read_html(raev)[1], "sv": pd.read_html(sv)[0]}
        except Exception as e:
            print(f"⚠️ Erreur lecture {month}: {e}")

    if not data:
        raise ValueError("❌ Aucun mois exploitable (HTML non reconnus)")

    evol_rows = []
    for curr_mois in sorted(data.keys(), key=month_key):
        curr     = data[curr_mois]["raev"]
        value_AM = curr.loc[
            curr["Zone de valorisation"].str.contains("TOTAL activité valorisée"),
            "Montant AM",
        ].iloc[0]
        value_AM = float(str(value_AM).replace(" ", "").replace(",", "."))

        curr2        = data[curr_mois]["sv"]
        curr2        = curr2.iloc[[0, 11]].copy()
        col_ssrha_br = [c for c in curr2.columns if "SSRHA" in c and "Montant BR" in c][0]
        col_htp_br   = [c for c in curr2.columns if "HTP"   in c and "Montant BR" in c][0]
        curr2        = curr2.rename(columns={
            col_ssrha_br: "SSRHA en HC - Montant BR",
            col_htp_br:   "Journées en HTP - Montant BR",
        })
        for col in ["SSRHA en HC - Montant BR", "Journées en HTP - Montant BR"]:
            curr2[col] = (
                curr2[col].astype(str)
                .str.replace(" ", "", regex=False)
                .str.replace(",", ".", regex=False)
            )
            curr2[col] = pd.to_numeric(curr2[col], errors="coerce")
        curr2.loc[
            curr2["Type d'activité"] == "Activité valorisée",
            "SSRHA en HC - Montant AM",
        ] = value_AM
        curr2["Mois"] = curr_mois

        df_month = curr2.pivot(index="Mois", columns="Type d'activité")
        df_month.columns = [f"{metric}_{act}" for metric, act in df_month.columns]
        df_month.columns = [
            "effectif_transmis_HC",
            "effectif_valorise_HC",
            "montantBR_transmis_HC",
            "montantBR_valorise_HC",
            "effectif_transmis_HTP",
            "effectif_valorise_HTP",
            "montantBR_transmis_HTP",
            "montantBR_valorise_HTP",
            "montantAM_transmis_HC",
            "montantAM_valorise_HC",
        ]
        df_month["jour_valo_HC"] = jours_valo_mois
        evol_rows.append(df_month)

    if not evol_rows:
        raise ValueError("❌ Aucun mois valide après traitement")

    brut_df = pd.concat(evol_rows).reset_index()
    brut_df["taux_valorisation_HC"] = (
        brut_df["effectif_valorise_HC"] / brut_df["effectif_transmis_HC"] * 100
    )
    brut_df["taux_valorisation_HTP"] = (
        brut_df["effectif_valorise_HTP"] / brut_df["effectif_transmis_HTP"] * 100
    )
    brut_df["recette_BR_moy_sej"]  = brut_df["montantBR_valorise_HC"] / brut_df["effectif_valorise_HC"]
    brut_df["recette_BR_moy_jour"] = brut_df["montantBR_valorise_HC"] / brut_df["jour_valo_HC"]
    brut_df["recette_BR_period"] = brut_df["montantBR_valorise_HC"].fillna(0) + brut_df["montantBR_valorise_HTP"].fillna(0)

    return {"brut_df": brut_df, "_tmp_dir": tmp}

def recalculer_derives(brut_df):
    """
    Les fichiers M sont des cumuls depuis le début de l'année.
    Exemple :
      M1 = cumul 01/01 → 01/02
      M2 = cumul 01/01 → 01/03

    Pour le rapport mensuel, on calcule les valeurs DU MOIS par différence :
      mois M2 = cumul M2 - cumul M1

    On conserve aussi les valeurs cumulées pour les KPI secondaires
    et pour la recette BR moyenne par jour sur la période totale.
    """
    df = brut_df.copy().reset_index(drop=True)

    if df.empty:
        raise ValueError("❌ Aucune donnée à traiter.")

    # ── Valeurs du mois = différence entre deux périodes cumulées ─────
    df["sejours_transmis_mois_HC"] = df["effectif_transmis_HC"].diff()
    df["sejours_valorises_mois_HC"] = df["effectif_valorise_HC"].diff()

    df["montantBR_mois_HC"] = df["montantBR_valorise_HC"].diff()
    df["montantAM_mois_HC"] = df["montantAM_valorise_HC"].diff()
    df["jours_valorises_mois_HC"] = df["jour_valo_HC"].diff()

    # HTP si présent
    df["jours_transmis_mois_HTP"] = df["effectif_transmis_HTP"].diff()
    df["jours_valorises_mois_HTP"] = df["effectif_valorise_HTP"].diff()
    df["montantBR_mois_HTP"] = df["montantBR_valorise_HTP"].diff()

    # ── Premier mois : pas de période précédente, donc M1 = valeur cumulée M1 ──
    first = df.index[0]

    df.loc[first, "sejours_transmis_mois_HC"] = df.loc[first, "effectif_transmis_HC"]
    df.loc[first, "sejours_valorises_mois_HC"] = df.loc[first, "effectif_valorise_HC"]

    df.loc[first, "montantBR_mois_HC"] = df.loc[first, "montantBR_valorise_HC"]
    df.loc[first, "montantAM_mois_HC"] = df.loc[first, "montantAM_valorise_HC"]
    df.loc[first, "jours_valorises_mois_HC"] = df.loc[first, "jour_valo_HC"]

    df.loc[first, "jours_transmis_mois_HTP"] = df.loc[first, "effectif_transmis_HTP"]
    df.loc[first, "jours_valorises_mois_HTP"] = df.loc[first, "effectif_valorise_HTP"]
    df.loc[first, "montantBR_mois_HTP"] = df.loc[first, "montantBR_valorise_HTP"]

    # ── Indicateurs DU MOIS ───────────────────────────────────────────
    df["taux_valorisation_mois_HC"] = (
        df["sejours_valorises_mois_HC"] / df["sejours_transmis_mois_HC"] * 100
    )

    df["recette_BR_moy_jour_mois_HC"] = (
        df["montantBR_mois_HC"] / df["jours_valorises_mois_HC"]
    )

    # ── Indicateurs CUMULÉS sur toute la période ──────────────────────
    df["taux_valorisation_cumule_HC"] = (
        df["effectif_valorise_HC"] / df["effectif_transmis_HC"] * 100
    )

    df["recette_BR_moy_jour_cumule_HC"] = (
        df["montantBR_valorise_HC"] / df["jour_valo_HC"]
    )

    df["recette_BR_mois_total"] = (
        df["montantBR_mois_HC"].fillna(0) + df["montantBR_mois_HTP"].fillna(0)
    )

    df["recette_BR_cumule_total"] = (
        df["montantBR_valorise_HC"].fillna(0) + df["montantBR_valorise_HTP"].fillna(0)
    )

    # ── Compatibilité avec tes anciens noms si utilisés ailleurs ──────
    df["sejour_supp"] = df["sejours_transmis_mois_HC"]
    df["sejour_valo_supp"] = df["sejours_valorises_mois_HC"]
    df["jour_valo_supp"] = df["jours_valorises_mois_HC"]
    df["ecart_valo"] = df["montantBR_mois_HC"]

    return df

#%%
# ══════════════════════════════════════════════════════════════════════════════
#  MOYENNES ANNÉE PRÉCÉDENTE
# ══════════════════════════════════════════════════════════════════════════════

def load_annee_precedente(uploaded_zip, uploaded_csv_m12):
   
    tmp = tempfile.TemporaryDirectory()
    tmp_path = Path(tmp.name)

    # ── Extraction ZIP ───────────────────────────────────────────────
    if hasattr(uploaded_zip, "read"):
        with zipfile.ZipFile(io.BytesIO(uploaded_zip.read()), "r") as zf:
            zf.extractall(tmp_path)
    else:
        with zipfile.ZipFile(uploaded_zip, "r") as zf:
            zf.extractall(tmp_path)

    # ── Recherche des fichiers HTML dans tout le ZIP ─────────────────
    html_files = list(tmp_path.rglob("*.html"))

    sv = next((f for f in html_files if "sv" in f.name.lower()), None)

    if not sv:
        raise ValueError("❌ Fichier SV introuvable dans le ZIP M12")

    # ── Lecture du SV M12 ────────────────────────────────────────────
    try:
        curr2 = pd.read_html(sv)[0]
    except Exception as e:
        raise ValueError(f"❌ Erreur lecture SV M12 : {e}")

    # Lignes Activité transmise + Activité valorisée
    curr2 = curr2.iloc[[0, 11]].copy()

    col_ssrha_br = [
        c for c in curr2.columns
        if "SSRHA" in c and "Montant BR" in c
    ][0]

    curr2 = curr2.rename(columns={
        col_ssrha_br: "SSRHA en HC - Montant BR"
    })

    curr2["SSRHA en HC - Montant BR"] = pd.to_numeric(
        curr2["SSRHA en HC - Montant BR"]
        .astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce",
    )

    # ── Récupérer uniquement l'activité valorisée ────────────────────
    ligne_valorisee = curr2[
        curr2["Type d'activité"] == "Activité valorisée"
    ]

    if ligne_valorisee.empty:
        raise ValueError("❌ Ligne 'Activité valorisée' introuvable dans le SV M12")

    montantBR_valorise_HC = float(
        ligne_valorisee["SSRHA en HC - Montant BR"].iloc[0]
    )

    # ── VisualValo M12 cumulé : jours valorisés HC ───────────────────
    jours_valo_HC = _calc_jours_valo(uploaded_csv_m12)

    if jours_valo_HC == 0 or pd.isna(jours_valo_HC):
        raise ValueError("❌ Nombre de jours valorisés nul ou invalide dans le VisualValo M12")

    # ── Moyenne annuelle par jour HC ─────────────────────────────────
    moyennes = {
        "recette_BR_moy_jour": float(montantBR_valorise_HC / jours_valo_HC)
    }

    return moyennes

# ══════════════════════════════════════════════════════════════════════════════
#  FONCTIONS GRAPHIQUES
# ══════════════════════════════════════════════════════════════════════════════

barlow_bold = font_manager.FontProperties(
    fname=BASE_DIR / "design" / "Barlow-Bold.ttf",
    size=14)

# def style_xticklabels(ax, x_vals, y_vals):
#     ax.set_xticks(range(len(x_vals)))
#     ax.set_xticklabels(x_vals)
#     for i, label in enumerate(ax.get_xticklabels()):
#         if i > 0 and y_vals.iloc[i] < y_vals.iloc[i - 1]:
#             label.set_color(VIOLET)
#         else:
#             label.set_color(GRIS_TEXTE)

def annoter_tous_les_points(ax, x_vals, y_vals, fmt="{:,.0f}", couleur=BLEU):
    y_vals = y_vals.reset_index(drop=True)
    for i, val in enumerate(y_vals):
        try:
            v = float(val)
        except (TypeError, ValueError):
            continue
        if np.isnan(v):
            continue
        try:
            label = fmt.format(v)
        except (ValueError, TypeError):
            label = str(v)
        ax.annotate(
            label,
            xy=(i, v),
            xytext=(0, 12),
            textcoords="offset points",
            fontsize=9, fontweight="bold", color=couleur,
            ha="center", va="bottom",
            bbox=dict(boxstyle="round,pad=0.2", facecolor=BLANC,
                      edgecolor=couleur, alpha=0.85, linewidth=0.7),
        )

def _style_ax(ax):
    ax.patch.set_facecolor("white")
    ax.patch.set_alpha(0.90)
    ax.grid(True, axis="y", linestyle="--", alpha=0.4, color="#9CA3AF")
    ax.grid(False, axis="x")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="x", rotation=0, labelsize=12, pad=10)
    ax.tick_params(axis="y", labelsize=12, colors=GRIS_TEXTE)
    ax.yaxis.set_tick_params(pad=5)
    ax.xaxis.set_tick_params(pad=5)

def make_ax_hlines(ax, col, title, objectif, evol_df, fmt="{:,.0f}", moy_annuelle=None):
    x_vals = list(evol_df["Mois"])
    y_vals = evol_df[col].reset_index(drop=True)
    ax.plot(x_vals, y_vals, linewidth=2.5, color=BLEU,
            marker="o", markersize=5, markerfacecolor="white", markeredgewidth=2)
    moyenne = y_vals.mean()
    ax.axhline(moyenne, color="#9CA3AF", linestyle="--", linewidth=1.5,
               label=f"Moyenne période ({format_fr(moyenne)})")
    if objectif is not None:
        ax.axhline(objectif, color=ORANGE, linestyle="--", linewidth=1.5,
                   label=f"Objectif mensuel ({objectif:,.0f})")
    if moy_annuelle is not None:
        ax.axhline(moy_annuelle, color=VIOLET, linestyle="--", linewidth=1.5,
                   label=f"Moy. année préc. ({format_fr(moy_annuelle)})")
    ax.set_title(title, pad=10,  fontproperties=barlow_bold)
    ax.legend(fontsize=10, framealpha=0.9, loc="best")
    _style_ax(ax)
    #style_xticklabels(ax, x_vals, y_vals)
    annoter_tous_les_points(ax, x_vals, y_vals, fmt=fmt)

def make_ax_bar(ax, series, title, evol_df, fmt="{:.1f} %"):
    x_vals = list(evol_df["Mois"])
    x = np.arange(len(x_vals))

    n_series = len(series)
    width = 0.35 if n_series == 2 else 0.55

    for i, (col, label) in enumerate(series):
        y_vals = evol_df[col].reset_index(drop=True)

        offset = (i - (n_series - 1) / 2) * width

        bars = ax.bar(
            x + offset,
            y_vals,
            width=width,
            label=label,
            alpha=0.85,
            zorder=3,
        )

        for bar, val in zip(bars, y_vals):
            if pd.isna(val):
                continue

            label_txt = format_fr(val, fmt)

            ax.text(
                bar.get_x() + bar.get_width() / 2,
                val + max(y_vals.max() * 0.02, 1),
                label_txt,
                ha="center",
                va="bottom",
                fontsize=9,
                fontweight="bold",
                color=GRIS_TEXTE,
            )

    ax.axhline(0, color=GRIS_TEXTE, linewidth=0.8)
    ax.set_title(title, pad=10, fontproperties=barlow_bold)

    _style_ax(ax)

    ax.set_xticks(x)
    ax.set_xticklabels(x_vals)
    ax.legend(fontsize=10, framealpha=0.9, loc="best")

def make_ax_multi(ax, plots, title, evol_df, moy_annuelle=None):
    x_vals = list(evol_df["Mois"])
    for i, (col, label) in enumerate(plots):
        y_vals = evol_df[col].reset_index(drop=True)
        ax.plot(x_vals, y_vals, linewidth=2.5, color=COLORS[i % len(COLORS)],
                marker="o", markersize=5, markerfacecolor="white",
                markeredgewidth=2, label=label)
        annoter_tous_les_points(ax, x_vals, y_vals, couleur=COLORS[i % len(COLORS)])
        if moy_annuelle is not None and col in moy_annuelle and moy_annuelle[col] is not None:
            ax.axhline(moy_annuelle[col], color=COLORS[i % len(COLORS)],
                       linestyle=":", linewidth=1.5,
                       label=f"Moy. année préc. — {label.split(' ')[0]} ({moy_annuelle[col]:,.0f})")
    ax.set_title(title, pad=10, fontproperties=barlow_bold)
    ax.legend(fontsize=10, framealpha=0.9, loc="best")
    _style_ax(ax)
    ax.set_xticks(range(len(x_vals)))
    ax.set_xticklabels(x_vals)
    for i, label in enumerate(ax.get_xticklabels()):
        if i > 0:
            all_down = all(evol_df[col].iloc[i] < evol_df[col].iloc[i - 1] for col, _ in plots)
            label.set_color(VIOLET if all_down else GRIS_TEXTE)
        else:
            label.set_color(GRIS_TEXTE)

# ══════════════════════════════════════════════════════════════════════════════
#  MISE EN FORME PDF — pages avec template Canva
# ══════════════════════════════════════════════════════════════════════════════
#
# ── Page garde ────────────────────────────────────────────────────────────────
COVER_ETAB_Y        = 0.702   # centre vertical de la box "Centre Médical de"
COVER_ETAB_X        = 0.500   # centré horizontalement
 
# Grand bloc KPI (zone teal pointillée)
KPI_POS_ALL = {
    "recette_BR_period":       (0.160, 0.495),
    "montantAM_valorise_HC":   (0.385, 0.495),
    "effectif_transmis_HC":    (0.610, 0.495),
    "effectif_transmis_HTP":   (0.835, 0.495),

    "recette_BR_moy_sej":      (0.160, 0.235),
    "recette_BR_moy_jour":     (0.385, 0.235),
    "taux_valorisation_HC":    (0.610, 0.235),
    "taux_valorisation_HTP":   (0.835, 0.235),
}

# KPI_POS_HC = {
#     "recette_BR_period":       (0.275, 0.430),
#     "montantAM_valorise_HC":   (0.500, 0.430),
#     "effectif_transmis_HC":    (0.725, 0.430),

#     "recette_BR_moy_sej":      (0.390, 0.176),
#     "taux_valorisation_HC":    (0.610, 0.176),
# }

KPI_POS_HC = {
    "recette_BR_cumule_total":           (0.275, 0.430),
    "montantAM_valorise_HC":             (0.500, 0.430),
    "effectif_transmis_HC":              (0.725, 0.430),

    "recette_BR_moy_jour_cumule_HC":     (0.390, 0.176),
    "taux_valorisation_cumule_HC":       (0.610, 0.176),
}

 
# ── Pages graphiques HC / HTP ─────────────────────────────────────────────────
# Graphique haut gauche
GRAPH_LEFT_L  = 0.050 #horizontal
GRAPH_LEFT_B  = 0.595 #vertical
GRAPH_LEFT_W  = 0.410 #largeur
GRAPH_LEFT_H  = 0.235 #hauteur

# Commentaire haut gauche
COMMENT_SMALL_L_L = 0.050
COMMENT_SMALL_L_B = 0.450
COMMENT_SMALL_L_W = 0.420
COMMENT_SMALL_L_H = 0.090

# Graphique haut droit
GRAPH_RIGHT_L = 0.530
GRAPH_RIGHT_B = 0.595
GRAPH_RIGHT_W = 0.410
GRAPH_RIGHT_H = 0.235

# Commentaire haut droit
COMMENT_SMALL_R_L = 0.530
COMMENT_SMALL_R_B = 0.450
COMMENT_SMALL_R_W = 0.420
COMMENT_SMALL_R_H = 0.090

# Grand graphique bas À GAUCHE
GRAPH_BIG_L = 0.050
GRAPH_BIG_B = 0.105
GRAPH_BIG_W = 0.410
GRAPH_BIG_H = 0.235

# Commentaire bas À DROITE
COMMENT_BIG_L = 0.530
COMMENT_BIG_B = 0.180
COMMENT_BIG_W = 0.420
COMMENT_BIG_H = 0.090
 
# Pied de page
PAGE_NUM_Y          = 0.020
PAGE_NUM_X          = 0.970

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE DE GARDE 
# ══════════════════════════════════════════════════════════════════════════════
def _page_garde_with_data(nom_etablissement, nom_etablissement_layout, periode,
                          dernier, avant_dernier, evol_df, inclure_htp=True):
    """
    Page de garde avec background Canva + KPIs.
    Appelée uniquement depuis generate_pdf().
    """
    fig = plt.figure(figsize=(17, 12))
    fig.patch.set_facecolor(BLANC)

    cover_path = CANVA_COVER_PATH if inclure_htp else CANVA_COVER_PATH_HC
    bg = _charger_bg(cover_path)
    if bg is not None:
        _appliquer_bg(fig, bg)

    ax = fig.add_axes([0, 0, 1, 1], zorder=2)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.patch.set_alpha(0)

    # ── Nom établissement ────────────────────────────────────────────
    barlow_title = font_manager.FontProperties(
        fname=BASE_DIR / "design" / "Barlow-Bold.ttf",
        size=34
    )

    ax.text(
        COVER_ETAB_X, COVER_ETAB_Y,
        nom_etablissement_layout,
        ha="center",
        va="center",
        color=TEAL,
        zorder=3,
        fontproperties=barlow_title
    )

    # ── Helpers internes ─────────────────────────────────────────────
    def _fleche(val, ref, fmt=None):
        try:
            if ref is None or pd.isna(ref):
                return "–", GRIS_TEXTE

            d = float(val) - float(ref)

            if fmt and "%" in fmt:
                unit = " %"
            elif fmt and "€" in fmt:
                unit = " €"
            else:
                unit = ""

            if d > 0:
                return f"▲ +{format_fr(d)}{unit}", VERT
            if d < 0:
                return f"▼ {format_fr(d)}{unit}", ROUGE

            return "= stable", GRIS_TEXTE

        except Exception:
            return "–", GRIS_TEXTE

    def _badge(val, objectif):
        try:
            if objectif is None or objectif <= 0:
                return "Objectif à définir", GRIS_TEXTE

            val = float(val)

            if val >= objectif:
                return f"✓ Objectif atteint ({format_fr(objectif)} €)", VERT

            pct = (1 - val / objectif) * 100
            return f"✗ -{pct:.1f}% de l'objectif ({format_fr(objectif)} €)", ROUGE

        except Exception:
            return None, None

    kpi_config = KPI_CONFIG if inclure_htp else KPI_CONFIG_HC
    kpi_pos = KPI_POS_ALL if inclure_htp else KPI_POS_HC

    # ── KPIs ─────────────────────────────────────────────────────────
    for item in kpi_config:
        if len(item) == 4:
            col, label, fmt, obj_key = item
            col_mois, fmt_mois = None, None
        else:
            col, label, fmt, col_mois, fmt_mois, obj_key = item

        if col not in kpi_pos:
            continue

        x, y = kpi_pos[col]

        val_cumul = dernier.get(col, float("nan"))
        val_mois = dernier.get(col_mois, float("nan")) if col_mois else None

        # Valeur principale = cumul
        ax.text(
            x, y,
            format_fr(val_cumul, fmt),
            ha="center",
            va="center",
            fontsize=18,
            fontweight="bold",
            color=GRIS_TEXTE,
            zorder=3,
        )

        # Valeur secondaire = mois
        if col_mois:
            ax.text(
                x, y - 0.030,
                format_fr(val_mois, fmt_mois),
                ha="center",
                va="center",
                fontsize=8.5,
                color=GRIS_TEXTE,
                zorder=3,
            )

        # ── Flèche = évolution du mois actuel vs mois précédent ───────
        if len(evol_df) >= 2:

            # ── 1. Cas TAUX → comparaison directe mois vs mois ─────────
            if "taux" in col:
                if col_mois:
                    # taux du mois
                    val_now = evol_df.iloc[-1][col_mois]
                    val_prev = evol_df.iloc[-2][col_mois]
                else:
                    # fallback cumul
                    val_now = evol_df.iloc[-1][col]
                    val_prev = evol_df.iloc[-2][col]

            # ── 2. Cas MOYENNE → comparaison directe cumul vs cumul ────
            elif "moy" in col:
                val_now = evol_df.iloc[-1][col]
                val_prev = evol_df.iloc[-2][col]

            # ── 3. Cas CUMUL → reconstruction du mois ──────────────────
            else:
                cumul_now = evol_df.iloc[-1][col]
                cumul_prev = evol_df.iloc[-2][col]

                # mois actuel
                val_now = cumul_now - cumul_prev

                # mois précédent
                if len(evol_df) == 2:
                    val_prev = cumul_prev
                else:
                    cumul_prev2 = evol_df.iloc[-3][col]
                    val_prev = cumul_prev - cumul_prev2

            fleche, couleur_fl = _fleche(val_now, val_prev, fmt)

        else:
            fleche, couleur_fl = "–", GRIS_TEXTE
        ax.text(
            x, y - 0.055,
            fleche,
            ha="center",
            va="center",
            fontsize=9.5,
            fontweight="bold",
            color=couleur_fl,
            zorder=3,
        )

        # ── Objectif ─────────────────────────────────────────────────
        if obj_key and OBJECTIFS.get(obj_key) is not None:
            valeur_obj = val_mois if val_mois is not None else val_cumul
            badge_txt, badge_col = _badge(valeur_obj, OBJECTIFS[obj_key])

            if badge_txt:
                ax.text(
                    x, y - 0.076,
                    badge_txt,
                    ha="center",
                    va="center",
                    fontsize=7.5,
                    color=badge_col,
                    style="italic",
                    zorder=3,
                )

    # ── Pied de page ─────────────────────────────────────────────────
    ax.text(
        0.03, PAGE_NUM_Y,
        f"{AUTEUR}  |  {nom_etablissement}  |  {DATE_RAPPORT}",
        ha="left",
        va="center",
        fontsize=11,
        color=GRIS_TEXTE,
        zorder=3
    )

    ax.text(
        PAGE_NUM_X, PAGE_NUM_Y,
        "Page 1",
        ha="right",
        va="center",
        fontsize=11,
        fontweight="bold",
        color=GRIS_TEXTE,
        zorder=3
    )

    return fig

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE GRAPHIQUE GÉNÉRIQUE  (HC ou HTP selon le background passé)
# ══════════════════════════════════════════════════════════════════════════════
def _build_page_graphique(fig, theme, config, evol_df, page_num,
                          NOM_ETAB, PERIODE, canva_path,
                          custom_comments=None, moy_annuelle=None):
    bg = _charger_bg(canva_path)
    if bg is not None:
        _appliquer_bg(fig, bg)

    # ── Création des 3 axes graphiques ────────────────────────────────
    INNER = 0.010 #marge interne supp
    PAD_L = 0.030
    PAD_R = 0.012

    ax_gl = fig.add_axes([
        GRAPH_LEFT_L + PAD_L,
        GRAPH_LEFT_B + INNER,
        GRAPH_LEFT_W - PAD_L - PAD_R,
        GRAPH_LEFT_H - 2 * INNER,
    ], zorder=3)

    ax_gr = fig.add_axes([
        GRAPH_RIGHT_L + PAD_L,
        GRAPH_RIGHT_B + INNER,
        GRAPH_RIGHT_W - PAD_L - PAD_R,
        GRAPH_RIGHT_H - 2 * INNER,
    ], zorder=3)

    ax_gb = fig.add_axes([
        GRAPH_BIG_L + PAD_L,
        GRAPH_BIG_B + INNER,
        GRAPH_BIG_W - PAD_L - PAD_R,
        GRAPH_BIG_H - 2 * INNER,
    ], zorder=3)

    # ── Création des 3 axes commentaires ─────────────────────────────
    CINNER = 0.012

    ax_cl = fig.add_axes([
        COMMENT_SMALL_L_L + CINNER,
        COMMENT_SMALL_L_B + CINNER,
        COMMENT_SMALL_L_W - 2 * CINNER,
        COMMENT_SMALL_L_H - 2 * CINNER,
    ], zorder=3)
 
    ax_cr = fig.add_axes([
        COMMENT_SMALL_R_L + CINNER,
        COMMENT_SMALL_R_B + CINNER,
        COMMENT_SMALL_R_W - 2 * CINNER,
        COMMENT_SMALL_R_H - 2 * CINNER,
    ], zorder=3)
 
    ax_cb = fig.add_axes([
        COMMENT_BIG_L + CINNER,
        COMMENT_BIG_B + CINNER,
        COMMENT_BIG_W - 2 * CINNER,
        COMMENT_BIG_H - 2 * CINNER,
    ], zorder=3)
 
    graph_axes   = [ax_gl, ax_gr, ax_gb]
    comment_axes = [ax_cl, ax_cr, ax_cb]

    # ── Dispatch par type pour chaque sous-graphe ─────────────────────
    for i, subplot in enumerate(config["plots"]):
        ax     = graph_axes[i]
        ax_c   = comment_axes[i]
        t      = subplot["type"]
        series = subplot["series"]
        title  = subplot["title"]

        if t == "bar":
            make_ax_bar(ax, series, title, evol_df)
        elif t == "single_hlines":
            col, _ = series[0]
            # mapping explicite pour moyenne année précédente
            if moy_annuelle:
                if col == "recette_BR_moy_jour_cumule_HC":
                    moy = moy_annuelle.get("recette_BR_moy_jour")
                else:
                    moy = moy_annuelle.get(col)
            else:
                moy = None

            make_ax_hlines(
                ax,
                col,
                title,
                subplot.get("objectif"),
                evol_df,
                moy_annuelle=moy
            )
        elif t == "multi":
            make_ax_multi(ax, series, title, evol_df, moy_annuelle=moy_annuelle)

        _draw_comment(ax_c, series, theme, evol_df, custom_comments)

    # ── Pied de page ─────────────────────────────────────────────────
    ax_n = fig.add_axes([0, 0, 1, 1], zorder=4)
    ax_n.set_xlim(0, 1); ax_n.set_ylim(0, 1)
    ax_n.axis("off"); ax_n.patch.set_alpha(0)
    ax_n.text(0.03, PAGE_NUM_Y,
              f"{AUTEUR}  |  {NOM_ETAB}  |  {DATE_RAPPORT}",
              ha="left", va="center", fontsize=11, color=GRIS_TEXTE, zorder=5)
    ax_n.text(PAGE_NUM_X, PAGE_NUM_Y, f"Page {page_num}",
              ha="right", va="center", fontsize=11,
              fontweight="bold", color=GRIS_TEXTE, zorder=5)
# ── Helpers graphiques internes ───────────────────────────────────────────────
 
def _draw_subplot(ax, plot_list, evol_df, moy_annuelle):
    """Trace une courbe simple (1 série) dans `ax`."""
    col, titre = plot_list[0]
    moy = moy_annuelle.get(col) if moy_annuelle else None
    # Détection auto format
    if "taux" in col:
        fmt = "{:.1f} %"
    elif "recette" in col or "ecart" in col or "montant" in col.lower():
        fmt = "{:.0f} €"
    else:
        fmt = "{:.0f}"
    make_ax_hlines(ax, col, titre, OBJECTIFS.get(col), evol_df,
                   fmt=fmt, moy_annuelle=moy)
 
 
def _draw_subplot_bar(ax, plot_list, evol_df):
    """Trace un graphique en barres (écarts) dans `ax`."""
    col, titre = plot_list[0]
    if "recette" in col or "ecart" in col or "montant" in col.lower():
        fmt = "{:.0f} €"
    else:
        fmt = "{:.0f}"
    make_ax_bar(ax, col, titre, evol_df, fmt=fmt)
 
def _draw_comment(ax, subplot_plots, theme, evol_df, custom_comments, fontsize=12):
    ax.axis("off")
    ax.patch.set_facecolor("#F9FAFB")
    ax.patch.set_alpha(0.95)
    texts = []
    for col, titre in subplot_plots:
        key = (theme, col)
        if custom_comments and key in custom_comments:
            texts.append(custom_comments[key])
        else:
            texts.append(generate_comment(col, titre, evol_df))
    full_text = "\n".join(texts)

    largeur = ax.get_position().width
    chars_par_ligne = int(largeur * 160) 
    lignes = textwrap.fill(full_text, width=max(chars_par_ligne, 30))

    ax.text(
        0.025, 0.92,
        lignes,
        fontsize=fontsize, 
        color="#374151", 
        va="top",
        transform=ax.transAxes,
        linespacing=1.3,
        clip_on=False,
        wrap=True
    )
 
# ══════════════════════════════════════════════════════════════════════════════
#  WRAPPERS HC / HTP  (rétrocompatibilité)
# ══════════════════════════════════════════════════════════════════════════════
 
def _build_page_graphique_HC(fig, theme, config, evol_df, page_num,
                              NOM_ETAB, PERIODE,
                              custom_comments=None, moy_annuelle=None):
    _build_page_graphique(fig, theme, config, evol_df, page_num,
                          NOM_ETAB, PERIODE,
                          canva_path=CANVA_PAGE_HC_PATH,
                          custom_comments=custom_comments,
                          moy_annuelle=moy_annuelle)
 
 
def _build_page_graphique_HTP(fig, theme, config, evol_df, page_num,
                               NOM_ETAB, PERIODE,
                               custom_comments=None, moy_annuelle=None):
    _build_page_graphique(fig, theme, config, evol_df, page_num,
                          NOM_ETAB, PERIODE,
                          canva_path=CANVA_PAGE_HTP_PATH,
                          custom_comments=custom_comments,
                          moy_annuelle=moy_annuelle)

# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION DES COMMENTAIRES
# ══════════════════════════════════════════════════════════════════════════════

def generate_comment(col, titre, evol_df):
    series = evol_df[col].dropna()

    if len(series) < 2:
        return "Données insuffisantes pour analyse."

    debut = series.iloc[0]
    fin = series.iloc[-1]

    if "taux" in col:
        return (
            f"{titre} : le taux passe de {debut:.1f} % à {fin:.1f} % "
            f"sur la période. La moyenne observée est de {series.mean():.1f} %."
        )

    trend = fin - debut
    trend_pct = (trend / debut) * 100 if debut != 0 else 0

    if trend > 0:
        tendance = "hausse"
    elif trend < 0:
        tendance = "baisse"
    else:
        tendance = "stabilité"

    return (
        f"{titre} : on observe une {tendance} de {trend_pct:.1f} % "
        f"sur la période. La valeur moyenne est de {format_fr(series.mean())}, "
        f"avec un minimum de {format_fr(series.min())} et un maximum de {format_fr(series.max())}."
    )

# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION DES FIGURES POUR STREAMLIT
# ══════════════════════════════════════════════════════════════════════════════

def generate_all_figures(evol_df, moy_annuelle=None, inclure_htp=True):
    figures = []
    for theme, config in THEMES.items():
        if not inclure_htp and "HTP" in theme.upper():
                    continue

        for i, subplot in enumerate(config["plots"]):
            fig, ax = plt.subplots(figsize=(10, 5))
            t      = subplot["type"]
            series = subplot["series"]

            if t == "bar":
                make_ax_bar(ax, series, subplot["title"], evol_df)
            elif t == "single_hlines":
                col, titre = series[0]
                moy = moy_annuelle.get(col) if moy_annuelle else None
                make_ax_hlines(ax, col, titre, subplot.get("objectif"),
                               evol_df, moy_annuelle=moy)
            elif t == "multi":
                make_ax_multi(ax, series, theme, evol_df, moy_annuelle=moy_annuelle)

            fig.tight_layout()
            figures.append((theme, f"Graphe {i+1}", fig, series))
    return figures

# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION DU PDF
# ══════════════════════════════════════════════════════════════════════════════

def generate_pdf(evol_df, NOM_ETAB, NOM_ETAB_LAYOUT, PERIODE,
                 custom_comments=None, moy_annuelle=None, inclure_htp=True):
   
    buf = io.BytesIO()
 
    dernier       = evol_df.iloc[-1].to_dict()
    avant_dernier = evol_df.iloc[-2].to_dict() if len(evol_df) > 1 else None
 
    with pdf_backend.PdfPages(buf) as pdf:
 
        # ── Page 1 : Garde + KPIs ────────────────────────────────────
        fig = _page_garde_with_data(
            nom_etablissement=NOM_ETAB,
            nom_etablissement_layout=NOM_ETAB_LAYOUT,
            periode=PERIODE,
            dernier=dernier,
            avant_dernier=avant_dernier,
            evol_df=evol_df,
            inclure_htp=inclure_htp
        )
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)
 
        # ── Pages graphiques ─────────────────────────────────────────
        page_num = 2
        for theme, config in THEMES.items():
            if not inclure_htp and "HTP" in theme.upper():
                continue
            else:
                fig = plt.figure(figsize=(17, 12))
                fig.patch.set_facecolor(BLANC)
                canva_path = CANVA_PAGE_HTP_PATH if "HTP" in theme.upper() \
                            else CANVA_PAGE_HC_PATH
                _build_page_graphique(
                    fig, theme, config, evol_df,
                    page_num, NOM_ETAB, PERIODE,
                    canva_path=canva_path,
                    custom_comments=custom_comments,
                    moy_annuelle=moy_annuelle,
                )
                pdf.savefig(fig, bbox_inches="tight")
                plt.close(fig)
                page_num += 1
 
        # Métadonnées
        d = pdf.infodict()
        d["Title"]        = f"Rapport mensuel – {NOM_ETAB}"
        d["Author"]       = AUTEUR
        d["Subject"]      = f"Évolution mensuelle SSR – {PERIODE}"
        d["CreationDate"] = datetime.today()
 
    buf.seek(0)
    return buf.read() 
     





 