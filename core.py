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
CANVA_COVER_PATH = "design/page_garde.png"
CANVA_PAGE_HC_PATH  = "design/page_graph_HC.png"
CANVA_PAGE_HTP_PATH   = "design/page_graph_HTP.png"

AUTEUR = "Dr Nathalie DUCRET"
DATE_RAPPORT = datetime.today().strftime("%d/%m/%Y")
SERVICE = "Rapport évolution mensuelle SMR"

#à automatiser
OBJECTIFS = {
    "obj_AM_mois": 0,
    "obj_BR_mois": 0
}

KPI_CONFIG = [
    ("recette_BR_period",   "Recette brute pour la période", "{:,.0f} €",  "obj_BR_mois"),
    ("montantAM_valorise_HC",   "Recette AM pour la période en HC", "{:,.0f} €",  "obj_AM_mois"),
    ("recette_BR_moy_sej",    "Recette brute par séjour", "{:,.0f} €",  None),
    ("taux_valorisation_HC",  "Taux de valorisation HC",  "{:.1f} %",   None),
    ("effectif_transmis_HC",  "Séjours transmis HC",      "{:.0f}",     None),
]

KPI_COULEURS = [
    ("#DCFCE7", "#16A34A"),
    ("#DCFCE7", "#16A34A"),
    ("#DCFCE7", "#16A34A"),
    ("#DBEAFE", "#2563EB"),
    ("#FEF9C3", "#E0CE09")
]

THEMES = {
    "HC ": {
        "plots": [
            {
                "type": "bar",
                "series": [("taux_valorisation_HC", "Taux de valorisation"),
                           ("ecart_valo", "Écart avec M-1")],
                "title": "Valorisation",
            },
            {
                "type": "single_hlines",
                "objectif": None,
                "series": [("recette_BR_moy_sej", "Evolution de la recette brute moyenne par séjour")],
                "title": "Evolution de la recette brute moyenne par séjour",
            },
            {
                "type": "multi",
                "series": [("sejour_valo_supp", "Séjour valorisé supplémentaire par rapport à M-1"),
                           ("sejour_supp",      "Séjour supplémentaire par rapport à M-1")],
                "title": "Evolution de l'activité (séjours)",
            },
        ]
    },
    "HTP ": {
        "plots": [
            {
                "type": "bar",
                "series": [("taux_valorisation_HTP", "Taux de valorisation"),
                           ("ecart_valo",             "Écart de valorisation avec M-1")],
                "title": "Valorisation",
            },
            {
                "type": "single_hlines",
                "objectif": None,
                "series": [("recette_BR_moy_jour", "Evolution de la recette brute moyenne par jour")],
                "title": "Evolution de la recette brute moyenne par jour",
            },
            {
                "type": "multi",
                "series": [("jour_valo_supp", "Jour valorisé supplémentaire par rapport à M-1"),
                           ("jour_tot_supp",  "Jour supplémentaire par rapport à M-1")],
                "title": "Evolution de l'activité (jours)",
            },
        ]
    },
}

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
    brut_df["recette_BR_period"] = brut_df["montantBR_valorise_HC"] + brut_df["montantBR_valorise_HTP"]

    return {"brut_df": brut_df, "_tmp_dir": tmp}

def recalculer_derives(brut_df):
    """
    Calcule toutes les colonnes dérivées par .diff() sur la série COMPLÈTE
    (historique + nouveau mois fusionnés et triés).
    Réplique exactement les lignes 287-296 de load_data().
    Retourne un evol_df prêt pour generate_all_figures() et generate_pdf().
    """
    df = brut_df.copy().reset_index(drop=True)
    if df.empty:
        raise ValueError("❌ Aucune donnée à traiter — vérifiez que le mois n'est pas dans MOIS_EXCLUS.")
    df["ecart_valo"]          = df["montantBR_valorise_HC"].diff()
    df["sejour_supp"]         = df["effectif_transmis_HC"].diff()
    df["sejour_valo_supp"]    = df["effectif_valorise_HC"].diff()
    df["jour_valo_supp"]      = df["jour_valo_HC"].diff()
    df["jour_tot_supp"] = 0 #à calculer
    return df

#%%
# ══════════════════════════════════════════════════════════════════════════════
#  MOYENNES ANNÉE PRÉCÉDENTE
# ══════════════════════════════════════════════════════════════════════════════

def load_annee_precedente(uploaded_zip):
    """
    Parse un ZIP contenant tous les dossiers mois d'une année passée
    (sans CSV de jours valo — on ne calcule que les colonnes disponibles).
    Debug : uploaded_zip = '/Users/marionducret/Desktop/SOLIMED/Rapport évolution mensuelle/Ceyrat/Ceyrat.zip'
    """
    tmp      = tempfile.TemporaryDirectory()
    tmp_path = Path(tmp.name)

    if hasattr(uploaded_zip, "read"):
        with zipfile.ZipFile(io.BytesIO(uploaded_zip.read()), "r") as zf:
            zf.extractall(tmp_path)
    else:
        with zipfile.ZipFile(uploaded_zip, "r") as zf:
            zf.extractall(tmp_path)

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
        raise ValueError("❌ Aucun dossier mois détecté dans le ZIP")

    data = {}
    for month in sorted(month_dirs_dict.keys(), key=month_key):
        folder     = month_dirs_dict[month]
        html_files = list(folder.glob("*.html"))
        sv   = next((f for f in html_files if "sv"   in f.name), None)
        if not sv:
            continue
        try:
            data[month] = {"sv": pd.read_html(sv)[0]}
        except Exception:
            continue

    if not data:
        raise ValueError("❌ Aucun mois exploitable")

    rows = []
    for curr_mois in sorted(data.keys(), key=month_key):
        curr2        = data[curr_mois]['sv']
        curr2        = curr2.iloc[[0, 11]].copy()
        col_ssrha_br = [c for c in curr2.columns if "SSRHA" in c and "Montant BR" in c][0]
        curr2        = curr2.rename(columns={
            col_ssrha_br: "Séjour en HC - Montant BR"
        })
        curr2["Séjour en HC - Montant BR"] = pd.to_numeric(
            curr2["Séjour en HC - Montant BR"].astype(str).str.replace(" ", "", regex=False).str.replace(",", ".", regex=False),
            errors="coerce",
        )
        curr2["Mois"] = curr_mois
        df_month = curr2.pivot(index="Mois", columns="Type d'activité")
        df_month.columns = [f"{metric}_{act}" for metric, act in df_month.columns]
        # Le sv de l'année précédente n'a pas de Montant AM — 8 colonnes seulement
        df_month.columns = [
            "effectif_transmis_HC", "effectif_valorise_HC",
            "montantBR_transmis_HC", "montantBR_valorise_HC",
            "effectif_transmis_HTP", "effectif_valorise_HTP",
            "montantBR_transmis_HTP", "montantBR_valorise_HTP",
        ]
        rows.append(df_month)

    if not rows:
        raise ValueError("❌ Aucun mois valide")

    df = pd.concat(rows).reset_index()

    df["recette_BR_moy_sej"] = df["montantBR_valorise_HC"]/df["effectif_valorise_HC"]

    # Moyennes mensuelles
    moyennes = {}
    moyennes["recette_BR_moy_sej"] = float(df["recette_BR_moy_sej"].mean())

    return moyennes

# ══════════════════════════════════════════════════════════════════════════════
#  FONCTIONS GRAPHIQUES
# ══════════════════════════════════════════════════════════════════════════════

barlow_bold = font_manager.FontProperties(
    fname=BASE_DIR / "design" / "Barlow-Bold.ttf",
    size=11)

def style_xticklabels(ax, x_vals, y_vals):
    ax.set_xticks(range(len(x_vals)))
    ax.set_xticklabels(x_vals)
    for i, label in enumerate(ax.get_xticklabels()):
        if i > 0 and y_vals.iloc[i] < y_vals.iloc[i - 1]:
            label.set_color(VIOLET)
        else:
            label.set_color(GRIS_TEXTE)

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
    ax.tick_params(axis="x", rotation=45, labelsize=8)
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TEXTE)
    ax.yaxis.set_tick_params(pad=1)
    ax.xaxis.set_tick_params(pad=1)

def make_ax_hlines(ax, col, title, objectif, evol_df, fmt="{:,.0f}", moy_annuelle=None):
    x_vals = list(evol_df["Mois"])
    y_vals = evol_df[col].reset_index(drop=True)
    ax.plot(x_vals, y_vals, linewidth=2.5, color=BLEU,
            marker="o", markersize=5, markerfacecolor="white", markeredgewidth=2)
    moyenne = y_vals.mean()
    ax.axhline(moyenne, color="#9CA3AF", linestyle="--", linewidth=1.5,
               label=f"Moyenne globale ({moyenne:,.0f})")
    if objectif is not None:
        ax.axhline(objectif, color=ORANGE, linestyle="--", linewidth=1.5,
                   label=f"Objectif mensuel ({objectif:,.0f})")
    if moy_annuelle is not None:
        ax.axhline(moy_annuelle, color=VIOLET, linestyle="--", linewidth=1.5,
                   label=f"Moy. année préc. ({moy_annuelle:,.0f})")
    ax.set_title(title, pad=10,  fontproperties=barlow_bold)
    ax.legend(fontsize=9, framealpha=0.9, loc="best")
    _style_ax(ax)
    style_xticklabels(ax, x_vals, y_vals)
    annoter_tous_les_points(ax, x_vals, y_vals, fmt=fmt)

def make_ax_bar(ax, col, title, evol_df, fmt="{:,.0f}"):
    x_vals   = list(evol_df["Mois"])
    y_vals   = evol_df[col].reset_index(drop=True)
    couleurs = [VERT if v >= 0 else ROUGE for v in y_vals]
    bars     = ax.bar(range(len(x_vals)), y_vals, color=couleurs, alpha=0.85, zorder=3)
    for bar, val in zip(bars, y_vals):
        if np.isnan(val):
            continue
        try:
            label = fmt.format(val)
        except (ValueError, TypeError):
            label = str(val)
        va     = "bottom" if val >= 0 else "top"
        offset = abs(y_vals.abs().max()) * 0.015 if y_vals.abs().max() != 0 else 1
        y_pos  = val + offset if val >= 0 else val - offset
        ax.text(
            bar.get_x() + bar.get_width() / 2, y_pos, label,
            ha="center", va=va, fontsize=9, fontweight="bold",
            color=VERT if val >= 0 else ROUGE,
        )
    ax.axhline(0, color=GRIS_TEXTE, linewidth=0.8, linestyle="-")
    ax.set_title(title, pad=10, fontproperties=barlow_bold)
    _style_ax(ax)
    ax.set_xticks(range(len(x_vals)))
    ax.set_xticklabels(x_vals)
    style_xticklabels(ax, x_vals, y_vals)

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
    ax.legend(fontsize=9, framealpha=0.9, loc="best")
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
COVER_ETAB_Y        = 0.742   # centre vertical de la box "Centre Médical de"
COVER_ETAB_X        = 0.500   # centré horizontalement
 
# Grand bloc KPI (zone teal pointillée)
KPI_BOX_LEFT        = 0.039
KPI_BOX_BOTTOM      = 0.062
KPI_BOX_WIDTH       = 0.921
KPI_BOX_HEIGHT      = 0.556
 
# ── Pages graphiques HC / HTP ─────────────────────────────────────────────────
# Bloc graphique GAUCHE (teal dashed, haut)
GRAPH_LEFT_L        = 0.044
GRAPH_LEFT_B        = 0.616
GRAPH_LEFT_W        = 0.450
GRAPH_LEFT_H        = 0.272
 
# Bloc graphique DROIT (teal dashed, haut)
GRAPH_RIGHT_L       = 0.516
GRAPH_RIGHT_B       = 0.618
GRAPH_RIGHT_W       = 0.448
GRAPH_RIGHT_H       = 0.270
 
# Blocs commentaire petit GAUCHE (gris, milieu)
COMMENT_SMALL_L_L   = 0.042
COMMENT_SMALL_L_B   = 0.492
COMMENT_SMALL_L_W   = 0.434
COMMENT_SMALL_L_H   = 0.086
 
# Blocs commentaire petit DROIT (gris, milieu)
COMMENT_SMALL_R_L   = 0.514
COMMENT_SMALL_R_B   = 0.492
COMMENT_SMALL_R_W   = 0.446
COMMENT_SMALL_R_H   = 0.086
 
# Grand bloc graphique BAS (teal dashed)
GRAPH_BIG_L         = 0.054
GRAPH_BIG_B         = 0.191
GRAPH_BIG_W         = 0.910
GRAPH_BIG_H         = 0.267
 
# Grand bloc commentaire BAS (gris)
COMMENT_BIG_L       = 0.042
COMMENT_BIG_B       = 0.074
COMMENT_BIG_W       = 0.918
COMMENT_BIG_H       = 0.071
 
# Pied de page
PAGE_NUM_Y          = 0.020
PAGE_NUM_X          = 0.970

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE DE GARDE 
# ══════════════════════════════════════════════════════════════════════════════

def _page_garde_with_data(nom_etablissement, nom_etablissement_layout, periode, dernier, avant_dernier):
    """
    Page de garde avec background Canva + KPIs.
    Appelée uniquement depuis generate_pdf().
    """
    fig = plt.figure(figsize=(12, 17))
    fig.patch.set_facecolor(BLANC)
 
    bg = _charger_bg(CANVA_COVER_PATH)
    if bg is not None:
        _appliquer_bg(fig, bg)
 
    ax = fig.add_axes([0, 0, 1, 1], zorder=2)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.patch.set_alpha(0)
 
    # Nom établissement dans la box
    barlow_title = font_manager.FontProperties(
        fname=BASE_DIR / "design" / "Barlow-Bold.ttf",
        size=34)
    
    ax.text(
        COVER_ETAB_X, COVER_ETAB_Y,
        nom_etablissement_layout,
        ha="center", va="center",
        color=TEAL, zorder=3, 
        fontproperties=barlow_title
    )
 
    def _fleche(val, ref):
        try:
            if ref is None or np.isnan(float(ref)):
                return "–", GRIS_TEXTE
            d = float(val) - float(ref)
            if d > 0:  return f"▲ +{d:,.0f}", VERT
            if d < 0:  return f"▼ {d:,.0f}", ROUGE
            return "= stable", GRIS_TEXTE
        except Exception:
            return "–", GRIS_TEXTE
 
    def _badge(val, objectif):
        try:
            val = float(val)
            if val >= objectif:
                return f"✓ Objectif atteint ({objectif:,.0f} €)", VERT
            pct = (1 - val / objectif) * 100
            return f"✗ -{pct:.1f}% de l'objectif ({objectif:,.0f} €)", ROUGE
        except Exception:
            return None, None
 
    n_kpi    = len(KPI_CONFIG)
    card_h   = KPI_BOX_HEIGHT / n_kpi
    card_gap = 0.006
 
    for i, (col, label, fmt, obj_key) in enumerate(KPI_CONFIG):
        card_top    = KPI_BOX_BOTTOM + KPI_BOX_HEIGHT - i * card_h
        card_bottom = card_top - card_h + card_gap
        card_cy     = (card_top + card_bottom) / 2
 
        couleur_fond, couleur_bord = KPI_COULEURS[i % len(KPI_COULEURS)]
        val = dernier.get(col, float("nan"))
        ref = avant_dernier.get(col) if avant_dernier else None
 
        # Fond de carte
        ax.add_patch(mpatches.FancyBboxPatch(
            (KPI_BOX_LEFT + 0.006, card_bottom + card_gap * 0.3),
            KPI_BOX_WIDTH - 0.012, card_h - card_gap * 1.8,
            boxstyle="round,pad=0.003", linewidth=1.2,
            edgecolor=couleur_bord, facecolor=couleur_fond,
            zorder=2, clip_on=True,
        ))
 
        # Label
        ax.text(KPI_BOX_LEFT + 0.020, card_cy, label,
                ha="left", va="center", fontsize=15, color=GRIS_TEXTE, zorder=3)
 
        # Valeur
        val_x = KPI_BOX_LEFT + KPI_BOX_WIDTH * 0.52
        try:    val_str = fmt.format(val)
        except: val_str = "N/A"
        ax.text(val_x, card_cy, val_str,
                ha="center", va="center", fontsize=19,
                fontweight="bold", color=BLEU_FONCE, zorder=3)
 
        # Flèche évolution
        fleche_x = KPI_BOX_LEFT + KPI_BOX_WIDTH * 0.82
        fleche, couleur_fl = _fleche(val, ref)
        ax.text(fleche_x, card_cy, fleche,
                ha="center", va="center", fontsize=16,
                fontweight="bold", color=couleur_fl, zorder=3)
 
        # Badge objectif
        if obj_key and OBJECTIFS.get(obj_key) is not None:
            badge_txt, badge_col = _badge(val, OBJECTIFS[obj_key])
            if badge_txt:
                ax.text(val_x, card_cy - card_h * 0.22, badge_txt,
                        ha="center", va="center", fontsize=11,
                        color=badge_col, style="italic", zorder=3)
 
    # Pied de page
    ax.text(0.03, PAGE_NUM_Y,
            f"{AUTEUR}  |  {nom_etablissement}  |  {DATE_RAPPORT}",
            ha="left", va="center", fontsize=11, color=GRIS_TEXTE, zorder=3)
    ax.text(PAGE_NUM_X, PAGE_NUM_Y, "Page 1",
            ha="right", va="center", fontsize=11,
            fontweight="bold", color=GRIS_TEXTE, zorder=3)
 
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
    INNER = 0.018 #marge interne supp
    PAD_L = 0.036
    PAD_R = 0.010

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
    CINNER = 0.008

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
            make_ax_bar(ax, series[0][0], series[0][1], title, evol_df)
        elif t == "single_hlines":
            col, _ = series[0]
            moy = moy_annuelle.get(col) if moy_annuelle else None
            make_ax_hlines(ax, col, title, subplot.get("objectif"),
                           evol_df, moy_annuelle=moy)
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
        fmt = "{:,.0f} €"
    else:
        fmt = "{:.0f}"
    make_ax_hlines(ax, col, titre, OBJECTIFS.get(col), evol_df,
                   fmt=fmt, moy_annuelle=moy)
 
 
def _draw_subplot_bar(ax, plot_list, evol_df):
    """Trace un graphique en barres (écarts) dans `ax`."""
    col, titre = plot_list[0]
    if "recette" in col or "ecart" in col or "montant" in col.lower():
        fmt = "{:,.0f} €"
    else:
        fmt = "{:.0f}"
    make_ax_bar(ax, col, titre, evol_df, fmt=fmt)
 
def _draw_comment(ax, subplot_plots, theme, evol_df, custom_comments, fontsize=11):
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
        0.01, 0.95,
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
    trend     = series.iloc[-1] - series.iloc[0]
    trend_pct = (trend / series.iloc[0]) * 100 if series.iloc[0] != 0 else 0
    if trend > 0:
        tendance = "hausse"
    elif trend < 0:
        tendance = "baisse"
    else:
        tendance = "stabilité"
    return (
        f"{titre} : On observe une {tendance} globale de {trend_pct:.1f}% "
        f"sur la période. La valeur moyenne est de {series.mean():,.0f}, "
        f"avec un minimum de {series.min():,.0f} et un maximum de {series.max():,.0f}."
    )

# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION DES FIGURES POUR STREAMLIT
# ══════════════════════════════════════════════════════════════════════════════

def generate_all_figures(evol_df, moy_annuelle=None):
    figures = []
    for theme, config in THEMES.items():
        for i, subplot in enumerate(config["plots"]):
            fig, ax = plt.subplots(figsize=(10, 5))
            t      = subplot["type"]
            series = subplot["series"]

            if t == "bar":
                # plusieurs barres sur le même axe
                make_ax_bar(ax, series[0][0], series[0][1], evol_df)
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
                 custom_comments=None, moy_annuelle=None):
   
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
        )
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)
 
        # ── Pages graphiques ─────────────────────────────────────────
        page_num = 2
        for theme, config in THEMES.items():
            fig = plt.figure(figsize=(12, 17))
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


for theme, config in THEMES.items():
    print(f"theme : {theme}")
    print(f"config : {config}")
    print("\n\n")

     
     





 