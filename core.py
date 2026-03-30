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
from matplotlib.image import imread
from pathlib import Path
from datetime import datetime
import numpy as np
import re
import zipfile
import tempfile
import io
#%%
# ══════════════════════════════════════════════════════════════════════════════
#  SECTION CONFIGURATION — tout ce qui est paramétrable est ici
# ══════════════════════════════════════════════════════════════════════════════

OUTPUT_PDF  = "rapport_mensuel.pdf"

# ── Templates Canva ───────────────────────────────────────────────────────────
# Déposer les PNG exportés depuis Canva dans ./assets/
CANVA_COVER_PATH = "./design/page_garde.png"
CANVA_PAGE_PATH  = "./design/page_graph.png"
CANVA_KPI_PATH   = "./design/page_kpi.png"

AUTEUR  = "SOLIMED"
SERVICE = "Rapport évolution mensuelle SSR"

OBJECTIFS = {
    "recette_AM_moy_mois": 392_400,
    "recette_BR_moy_mois": 360_000,
}

KPI_CONFIG = [
    ("taux_valorisation_HC",  "Taux de valorisation HC",          "{:.1f} %",   None),
    ("recette_BR_moy_mois",   "Recette mensuelle brute",              "{:,.0f} €",  "recette_BR_moy_mois"),
    ("recette_AM_moy_mois",   "Recette mensuelle AM",              "{:,.0f} €",  "recette_AM_moy_mois"),
    ("recette_BR_moy_sej",    "Recette brute par séjour",          "{:,.0f} €",  None),
    ("recette_BR_moy_jour",   "Recette brute par jour",   "{:,.0f} €",  None),
    ("effectif_transmis_HC",  "Séjours transmis HC",               "{:.0f}",     None),
]

KPI_COULEURS = [
    ("#DBEAFE", "#2563EB"),
    ("#DCFCE7", "#16A34A"),
    ("#FEF9C3", "#D97706"),
    ("#F3E8FF", "#7C3AED"),
    ("#FEE2E2", "#E11D48"),
    ("#F3F4F6", "#6B7280"),
]

THEMES = {
    "Valorisation": {
        "type": "bar",
        "plots": [
            ("ecart_valo", "Écart de valorisation avec M-1"),
        ],
    },
    "Recette brute journalière": {
        "type": "single_hlines",
        "objectif": [None],
        "plots": [
            ("recette_BR_moy_jour", "Recette brute journalière"),
        ],
    },
    "Recette brute par séjour": {
        "type": "single_hlines",
        "objectif": [None],
        "plots": [
            ("recette_BR_moy_sej", "Recette brute par séjour"),
        ],
    },
    "Activité : Séjours": {
        "type": "multi",
        "plots": [
            ("sejour_supp",       "Séjour supplémentaire par rapport à M-1"),
            ("sejour_valo_supp",  "Séjour valorisé supplémentaire par rapport à M-1"),
        ],
    },
    "Activité : Jours ": {
        "type": "multi",
        "plots": [
            ("jour_valo_supp",      "Jour valorisé supplémentaire par rapport à M-1"),
            ("jour_tot_supp", "Jour supplémentaire par rapport à M-1"),
        ],
    },
}

MOIS_EXCLUS = []  # à renseigner si certains mois doivent être exclus

COLORS     = ["#2563EB", "#16A34A", "#16A34A", "#E11D48", "#E11D48"]
BLEU_FONCE = "#1E3A5F"
BLEU       = "#2563EB"
BLEU_CLAIR = "#DBEAFE"
GRIS_TEXTE = "#6B7280"
GRIS_CLAIR = "#F3F4F6"
ROUGE      = "#E11D48"
VERT       = "#16A34A"
BLANC      = "#FFFFFF"
TEAL       = "#007B7B"
NOIR       = "#1A1A1A"
VIOLET     = "#7C3AED"
ORANGE     = "#F97316"

#%%
# ══════════════════════════════════════════════════════════════════════════════
#  UTILITAIRES CANVA
# ══════════════════════════════════════════════════════════════════════════════

def _charger_bg(path: str):
    """
    Charge un PNG/JPEG Canva comme background (robuste au format réel).
    Cherche dans : chemin absolu, dossier du script, dossier courant, ./design/.
    """
    import os as _os
    from pathlib import Path as _Path
    import numpy as _np
    from PIL import Image as _Image

    candidates = [
        path,
        str(_Path(__file__).parent / path),
        str(_Path(__file__).parent / _Path(path).name),
        str(_Path(_os.getcwd()) / path),
        str(_Path(_os.getcwd()) / _Path(path).name),
        str(_Path(_os.getcwd()) / "design" / _Path(path).name),
    ]
    for p in candidates:
        try:
            img = _Image.open(p).convert("RGB")
            return _np.array(img)
        except Exception:
            continue
    return None


def _appliquer_bg(fig: plt.Figure, bg_img) -> None:
    """Affiche bg_img en plein fond de la figure."""
    if bg_img is None:
        return
    ax_bg = fig.add_axes([0, 0, 1, 1], zorder=0)
    ax_bg.imshow(
        bg_img,
        aspect="auto",
        extent=[0, 1, 0, 1],
        transform=ax_bg.transAxes,
        zorder=0,
    )
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

    evol_df = evol_df[~evol_df["Mois"].isin(MOIS_EXCLUS)]#supp 2026

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
    brut_df["recette_BR_moy_sej"]  = brut_df["montantBR_valorise_HC"] / brut_df["effectif_valorise_HC"]
    brut_df["recette_BR_moy_jour"] = brut_df["montantBR_valorise_HC"] / brut_df["jour_valo_HC"]
    brut_df = brut_df[~brut_df["Mois"].isin(MOIS_EXCLUS)]

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
    df["recette_BR_moy_mois"] = df["montantBR_valorise_HC"].diff()
    df["recette_AM_moy_mois"] = df["montantAM_valorise_HC"].diff()
    # Premier mois : pas de M-1, on reprend la valeur brute (comme load_data)
    df.loc[df.index[0], "recette_BR_moy_mois"] = df["montantBR_valorise_HC"].iloc[0]
    df.loc[df.index[0], "recette_AM_moy_mois"] = df["montantAM_valorise_HC"].iloc[0]
    df["jour_tot_supp"] = 0
    return df

#%%
# ══════════════════════════════════════════════════════════════════════════════
#  MOYENNES ANNÉE PRÉCÉDENTE
# ══════════════════════════════════════════════════════════════════════════════

def load_annee_precedente(uploaded_zip):
    """
    Parse un ZIP contenant tous les dossiers mois d'une année passée
    (sans CSV de jours valo — on ne calcule que les colonnes disponibles).
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
        curr2["SSRHA en HC - Montant BR"] = pd.to_numeric(
            curr2["SSRHA en HC - Montant BR"].astype(str).str.replace(" ", "", regex=False).str.replace(",", ".", regex=False),
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

    # Calcul des colonnes dérivées nécessaires (sans recette_BR_moy_jour — pas de jours valo)
    df["recette_BR_moy_sej"] = df["montantBR_valorise_HC"]/df["effectif_valorise_HC"]

    # Moyennes mensuelles
    moyennes = {}
    moyennes["recette_BR_moy_sej"] = float(df["recette_BR_moy_sej"].mean())

    return moyennes


# ══════════════════════════════════════════════════════════════════════════════
#  FONCTIONS GRAPHIQUES
# ══════════════════════════════════════════════════════════════════════════════

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
            fontsize=10, fontweight="bold", color=couleur,
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
    ax.tick_params(axis="x", rotation=45, labelsize=13)
    ax.tick_params(axis="y", labelsize=13, colors=GRIS_TEXTE)


def make_ax(ax, col, titre, evol_df, fmt="{:,.0f}"):
    x_vals = list(evol_df["Mois"])
    y_vals = evol_df[col].reset_index(drop=True)
    ax.plot(x_vals, y_vals, linewidth=2.5, color=BLEU,
            marker="o", markersize=5, markerfacecolor="white", markeredgewidth=2)
    ax.set_title("", pad=0)
    _style_ax(ax)
    style_xticklabels(ax, x_vals, y_vals)
    annoter_tous_les_points(ax, x_vals, y_vals, fmt=fmt)


def make_ax_hlines(ax, col, titre, objectif, evol_df, fmt="{:,.0f}", moy_annuelle=None):
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
    ax.set_title("", pad=0)
    ax.legend(fontsize=12, framealpha=0.9, loc="best")
    _style_ax(ax)
    style_xticklabels(ax, x_vals, y_vals)
    annoter_tous_les_points(ax, x_vals, y_vals, fmt=fmt)


def make_ax_bar(ax, col, titre, evol_df, fmt="{:,.0f}"):
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
            ha="center", va=va, fontsize=10, fontweight="bold",
            color=VERT if val >= 0 else ROUGE,
        )
    ax.axhline(0, color=GRIS_TEXTE, linewidth=0.8, linestyle="-")
    ax.set_title("", pad=0)
    _style_ax(ax)
    ax.set_xticks(range(len(x_vals)))
    ax.set_xticklabels(x_vals)
    style_xticklabels(ax, x_vals, y_vals)


def make_ax_multi(ax, plots, theme_title, evol_df, moy_annuelle=None):
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
    ax.set_title("", pad=0)
    ax.legend(fontsize=12, framealpha=0.9, loc="best")
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

# ── Dimensions calibrées sur les PNG Canva 1414×2000 px ──────────────────────
#
# PAGE DE GARDE
# "Présenté par"  label : mpl_y ≈ 0.654  (le nom établissement s'écrit en-dessous)
# "Période"       label : mpl_y ≈ 0.579  (la période s'écrit en-dessous)
# Les labels sont à x ≈ 0.091 (début du texte teal)
# Valeurs dynamiques décalées légèrement vers le bas du label
COVER_PRES_LABEL_Y  = 0.654   # y du label "Etablissement"
COVER_NOM_ETAB_Y    = 0.624   # y de la valeur NOM_ETAB (sous le label)
COVER_PERI_LABEL_Y  = 0.579   # y du label "Période"
COVER_PERIODE_Y     = 0.510   # y de la valeur PERIODE
COVER_DATE_Y        = 0.170  # y de la kpi (dans le bloc teal bas-droite)
COVER_DATE_X        = 0.650   # x de la kpi (centre du bloc teal)
COVER_TEXT_X        = 0.091   # x de départ des textes dynamiques

# PAGE KPI (calibration pixel-perfect sur le PNG 1414×2000)
# Grand bloc unique : [left=0.060, bottom=0.057, w=0.880, h=0.794]
KPI_BOX_LEFT    = 0.060
KPI_BOX_BOTTOM  = 0.057
KPI_BOX_WIDTH   = 0.880
KPI_BOX_HEIGHT  = 0.794

# PAGE GRAPHIQUE (calibration pixel-perfect sur le PNG 1414×2000)
# Bandeau titre teal : mpl_y centre ≈ 0.944
PAGE_TITRE_X        = 0.030
PAGE_TITRE_Y        = 0.944
# Grand bloc graphique — marges internes de 1.5%
PAGE_GRAPH_LEFT     = 0.100
PAGE_GRAPH_BOTTOM   = 0.375
PAGE_GRAPH_WIDTH    = 0.840
PAGE_GRAPH_HEIGHT   = 0.460
# Petit bloc commentaire — marges internes de 1.5%
PAGE_COMMENT_LEFT   = 0.070
PAGE_COMMENT_BOTTOM = 0.075
PAGE_COMMENT_WIDTH  = 0.840
PAGE_COMMENT_HEIGHT = 0.220
# Numéro de page
PAGE_NUM_X          = 0.940
PAGE_NUM_Y          = 0.032


def page_garde(nom_etablissement: str, periode: str,
               date_generation: str | None = None,
               cover_kpi: dict | None = None) -> plt.Figure:
    if date_generation is None:
        date_generation = datetime.today().strftime("%d/%m/%Y")

    fig = plt.figure(figsize=(12, 17))
    fig.patch.set_facecolor(BLANC)

    bg = _charger_bg(CANVA_COVER_PATH)

    if bg is not None:
        # ── Mode Canva ─────────────────────────────────────────────────
        _appliquer_bg(fig, bg)

        ax = fig.add_axes([0, 0, 1, 1], zorder=1)
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        ax.patch.set_alpha(0)

        # Valeur "Etablissement" — bien en-dessous du label
        ax.text(
            COVER_TEXT_X, COVER_PRES_LABEL_Y - 0.0605,
            nom_etablissement,
            ha="left", va="center",
            fontsize=30, fontweight="bold", color=NOIR,
            zorder=2,
        )
        # Valeur "Période" — bien en-dessous du label
        ax.text(
            COVER_TEXT_X, COVER_PERI_LABEL_Y - 0.07,
            periode,
            ha="left", va="center",
            fontsize=22, fontweight="bold", color=NOIR,
            zorder=2,
        )
        # KPI Recette BR dans le carré teal bas-droite
        # Carré : left=0.557, bottom=0.047, width=0.398, height=0.259 -> centre x=0.756, y=0.176
        # On affiche la valeur du dernier mois si disponible via le titre (pas accès à evol_df ici)
        # → on passe la valeur en paramètre optionnel via cover_kpi
        if cover_kpi is not None:
            kpi_cx = 0.756
            kpi_cy = 0.140
            ax.text(kpi_cx, kpi_cy + 0.058,
                    "Recette brute mensuelle",
                    ha="center", va="center", fontsize=16, fontweight="bold",
                    color=BLANC, zorder=2)
            ax.text(kpi_cx, kpi_cy + 0.003,
                    cover_kpi["valeur"],
                    ha="center", va="center", fontsize=26,
                    fontweight="bold", color=BLANC, zorder=2)
            ax.text(kpi_cx, kpi_cy - 0.052,
                    cover_kpi["evolution"],
                    ha="center", va="center", fontsize=16,
                    fontweight="bold", color=BLANC, zorder=2)


    else:
        # ── Fallback matplotlib ────────────────────────────────────────
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        ax.add_patch(mpatches.FancyBboxPatch(
            (0, 0.72), 1, 0.28, boxstyle="square,pad=0",
            linewidth=0, facecolor=BLEU_FONCE,
        ))
        ax.axhline(y=0.72, xmin=0, xmax=1, color=BLEU, linewidth=4)
        ax.text(0.5, 0.88, "RAPPORT D'ÉVOLUTION MENSUELLE",
                ha="center", va="center", fontsize=28, fontweight="bold", color=BLANC)
        ax.add_patch(mpatches.FancyBboxPatch(
            (0.1, 0.50), 0.8, 0.16, boxstyle="round,pad=0.02",
            linewidth=1.5, edgecolor=BLEU, facecolor=BLEU_CLAIR,
        ))
        ax.text(0.5, 0.61, nom_etablissement,
                ha="center", va="center", fontsize=22, fontweight="bold", color=BLEU_FONCE)
        ax.text(0.5, 0.50, f"Période analysée : {periode}",
                ha="center", va="center", fontsize=16, color=GRIS_TEXTE)
        ax.plot([0.1, 0.9], [0.46, 0.46], color=GRIS_CLAIR, linewidth=1.5)
        ax.text(0.5, 0.38,
                "Suivi de la valorisation · Recettes BR/AM · Activité HC/HTP",
                ha="center", va="center", fontsize=11, color=GRIS_TEXTE, style="italic")
        for x, couleur, label in [(0.25, BLEU, "Valorisation"),
                                   (0.50, VERT, "Recettes"),
                                   (0.75, ROUGE, "Activités")]:
            ax.add_patch(plt.Circle((x, 0.28), 0.055, color=couleur, zorder=3))
            ax.text(x, 0.28, label[0], ha="center", va="center",
                    fontsize=14, fontweight="bold", color=BLANC, zorder=4)
            ax.text(x, 0.20, label, ha="center", va="center",
                    fontsize=9, color=GRIS_TEXTE)
        ax.axhline(y=0.08, xmin=0.05, xmax=0.95, color=GRIS_CLAIR, linewidth=1)

    return fig

def page_synthese(evol_df) -> plt.Figure:
    dernier       = evol_df.iloc[-1]
    avant_dernier = evol_df.iloc[-2] if len(evol_df) > 1 else None
    mois_label    = dernier["Mois"]

    fig = plt.figure(figsize=(12, 17))
    fig.patch.set_facecolor(BLANC)

    bg = _charger_bg(CANVA_KPI_PATH)
    if bg is not None:
        _appliquer_bg(fig, bg)

    # ── Axe principal transparent ─────────────────────────────────────
    ax = fig.add_axes([0, 0, 1, 1], zorder=1)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.patch.set_alpha(0)

    if bg is not None:
        # Titre dans le bandeau teal Canva
        ax.text(0.055, PAGE_TITRE_Y,
                "SYNTHÈSE — INDICATEURS CLÉS",
                ha="left", va="center",
                fontsize=20, fontweight="bold", color=BLANC, zorder=2)
    else:
        ax.patch.set_alpha(1)
        ax.add_patch(mpatches.FancyBboxPatch(
            (0, 0.88), 1, 0.12, boxstyle="square,pad=0",
            linewidth=0, facecolor=BLEU_FONCE,
        ))
        ax.axhline(y=0.88, xmin=0, xmax=1, color=BLEU, linewidth=3)
        ax.text(0.5, 0.95, "SYNTHÈSE — INDICATEURS CLÉS",
                ha="center", va="center", fontsize=20, fontweight="bold", color=BLANC)

    # Sous-titre mois (MARCHE PAS)
    ax.text(0.5, 0.826,
            f"Dernier mois disponible : {mois_label}",
            ha="center", va="center", fontsize=12,
            color=TEAL if bg is not None else BLEU_CLAIR,
            fontweight="bold", zorder=2)

    def fleche_et_couleur(val, ref):
        if ref is None or np.isnan(ref):
            return "–", GRIS_TEXTE
        delta = val - ref
        if delta > 0:   return f"▲ +{delta:,.0f}", VERT
        elif delta < 0: return f"▼ {delta:,.0f}", ROUGE
        return "= stable", GRIS_TEXTE

    def badge_objectif(val, objectif):
        if objectif is None: return None, None
        if val >= objectif:  return f"✓ Objectif atteint ({objectif:,.0f} €)", VERT
        pct = (1 - val / objectif) * 100
        return f"✗ -{pct:.1f}% de l'objectif ({objectif:,.0f} €)", ROUGE

    # ── Cartes KPI dans le grand bloc Canva ──────────────────────────
    MARGIN  = 0.018
    box_l   = KPI_BOX_LEFT   + MARGIN
    box_b   = KPI_BOX_BOTTOM + MARGIN
    box_w   = KPI_BOX_WIDTH  - 2 * MARGIN
    box_h   = KPI_BOX_HEIGHT - 2 * MARGIN

    n_kpi      = len(KPI_CONFIG)
    card_h     = box_h / n_kpi
    card_gap   = 0.006

    for i, (col, label, fmt, obj_key) in enumerate(KPI_CONFIG):
        # Coordonnées de la carte (de haut en bas)
        card_top    = box_b + box_h - i * card_h
        card_bottom = card_top - card_h + card_gap
        card_cy     = (card_top + card_bottom) / 2

        couleur_fond, couleur_bord = KPI_COULEURS[i % len(KPI_COULEURS)]
        val = dernier.get(col, float("nan"))
        ref = avant_dernier.get(col) if avant_dernier is not None else None

        # Fond de carte
        ax.add_patch(mpatches.FancyBboxPatch(
            (box_l + 0.004, card_bottom + card_gap * 0.3),
            box_w - 0.008, card_h - card_gap * 1.8,
            boxstyle="round,pad=0.003", linewidth=1.2,
            edgecolor=couleur_bord, facecolor=couleur_fond, zorder=2,
            clip_on=True,
        ))

        # Label
        ax.text(box_l + 0.018, card_cy,
                label,
                ha="left", va="center", fontsize=16, color=GRIS_TEXTE, zorder=3)

        # Valeur — centré dans la moitié gauche du cadre
        val_x = box_l + box_w * 0.52
        try:    val_str = fmt.format(val)
        except: val_str = "N/A"
        ax.text(val_x, card_cy,
                val_str,
                ha="center", va="center", fontsize=20,
                fontweight="bold", color=BLEU_FONCE, zorder=3)

        # Flèche évolution — dans le dernier tiers du cadre
        fleche_x = box_l + box_w * 0.82
        try:    fleche, couleur_fl = fleche_et_couleur(val, ref)
        except: fleche, couleur_fl = "–", GRIS_TEXTE
        ax.text(fleche_x, card_cy,
                fleche,
                ha="center", va="center", fontsize=18,
                fontweight="bold", color=couleur_fl, zorder=3)

        # Badge objectif
        if obj_key and OBJECTIFS.get(obj_key) is not None:
            try:
                badge_txt, badge_col = badge_objectif(val, OBJECTIFS[obj_key])
                if badge_txt:
                    ax.text(val_x, card_cy - card_h * 0.22,
                            badge_txt,
                            ha="center", va="center", fontsize=12,
                            color=badge_col, style="italic", zorder=3)
            except: pass

    # Numéro de page
    ax.text(PAGE_NUM_X, PAGE_NUM_Y, "Page 2",
            ha="right", va="center", fontsize=11,
            fontweight="bold", color=GRIS_TEXTE, zorder=2)

    return fig


def _build_page_graphique(fig: plt.Figure, theme: str, config: dict,
                          evol_df, page_num: int, NOM_ETAB: str,
                          PERIODE: str, custom_comments=None,
                          moy_annuelle=None) -> None:
    """
    Construit une page graphique sur la figure `fig` déjà créée,
    en utilisant le template Canva si disponible.
    Les graphiques vont dans PAGE_GRAPH_*, les commentaires dans PAGE_COMMENT_*.
    """
    bg = _charger_bg(CANVA_PAGE_PATH)
    plots = config["plots"]

    if bg is not None:
        _appliquer_bg(fig, bg)

    # ── Titre dans le bandeau teal ────────────────────────────────────
    if bg is not None:
        ax_titre = fig.add_axes([0, 0, 1, 1], zorder=2)
        ax_titre.set_xlim(0, 1)
        ax_titre.set_ylim(0, 1)
        ax_titre.axis("off")
        ax_titre.patch.set_alpha(0)
        ax_titre.text(
            0.055, PAGE_TITRE_Y,
            theme.upper(),
            ha="left", va="center",
            fontsize=20, fontweight="bold", color=BLANC, zorder=3,
        )

    else:
        # Entête matplotlib originale
        ax_h = fig.add_axes([0, 0.91, 1, 0.09])
        ax_h.set_xlim(0, 1)
        ax_h.set_ylim(0, 1)
        ax_h.axis("off")
        ax_h.add_patch(mpatches.FancyBboxPatch(
            (0, 0), 1, 1, boxstyle="square,pad=0",
            linewidth=0, facecolor=BLEU_FONCE,
        ))
        ax_h.axhline(y=0, xmin=0, xmax=1, color=BLEU, linewidth=3)
        ax_h.text(0.03, 0.55, theme.upper(),
                  ha="left", va="center", fontsize=14, fontweight="bold", color=BLANC)
        ax_h.text(0.97, 0.65, NOM_ETAB,
                  ha="right", va="center", fontsize=10, color=BLEU_CLAIR)
        ax_h.text(0.97, 0.30, PERIODE,
                  ha="right", va="center", fontsize=9, color=GRIS_TEXTE)

    # ── Zone graphique ────────────────────────────────────────────────
    n = len(plots)
    is_multi = config["type"] == "multi"

    # Marges internes pour rester bien dans le cadre Canva
    MARGIN = 0.012
    graph_left   = PAGE_GRAPH_LEFT   + MARGIN
    graph_bottom = PAGE_GRAPH_BOTTOM + MARGIN
    graph_width  = PAGE_GRAPH_WIDTH  - 2 * MARGIN
    graph_height = PAGE_GRAPH_HEIGHT - 2 * MARGIN

    if is_multi:
        # Un seul axe couvrant toute la zone graphique
        ax_g = fig.add_axes(
            [graph_left, graph_bottom, graph_width, graph_height],
            zorder=3,
        )
        make_ax_multi(ax_g, plots, theme, evol_df, moy_annuelle=moy_annuelle)
    else:
        hspace = 0.05
        h_plot = graph_height / n
        for i, (col, titre) in enumerate(plots):
            bottom = graph_bottom + (n - 1 - i) * h_plot + hspace / 2
            height = h_plot - hspace
            ax_g = fig.add_axes(
                [graph_left, bottom, graph_width, height],
                zorder=3,
            )
            if config["type"] == "bar":
                make_ax_bar(ax_g, col, titre, evol_df)
            elif config["type"] == "single_hlines":
                moy = moy_annuelle.get(col) if moy_annuelle else None
                make_ax_hlines(ax_g, col, titre, config["objectif"][i], evol_df, moy_annuelle=moy)
            else:
                make_ax(ax_g, col, titre, evol_df)

    # ── Zone commentaire ──────────────────────────────────────────────
    comment_texts = []
    for col, titre in plots:
        if custom_comments and (theme, col) in custom_comments:
            comment_texts.append(custom_comments[(theme, col)])
        else:
            comment_texts.append(generate_comment(col, titre, evol_df))
    full_comment = "\n\n".join(comment_texts)

    CMARGIN = 0.005
    ax_c = fig.add_axes(
        [PAGE_COMMENT_LEFT + CMARGIN,
         PAGE_COMMENT_BOTTOM + CMARGIN,
         PAGE_COMMENT_WIDTH  - 2 * CMARGIN,
         PAGE_COMMENT_HEIGHT - 2 * CMARGIN],
        zorder=3,
    )
    ax_c.axis("off")
    ax_c.patch.set_facecolor("#F9FAFB")
    ax_c.patch.set_alpha(0.95)

    if bg is None:
        ax_c.add_patch(mpatches.FancyBboxPatch(
            (0, 0), 1, 1,
            boxstyle="round,pad=0.02",
            facecolor="#F9FAFB", edgecolor="#E5E7EB",
        ))

    # Formater le texte avec textwrap pour éviter débordements
    import textwrap as _tw
    max_chars_per_line = 85
    wrapped_lines = []
    for para in full_comment.split("\n\n"):
        lines = _tw.wrap(para, width=max_chars_per_line)
        wrapped_lines.extend(lines)
        wrapped_lines.append("")
    # Max 6 lignes de contenu
    content_lines = [l for l in wrapped_lines if l][:6]
    display_comment = "\n".join(content_lines)
    ax_c.text(
        0.01, 0.80,
        "Analyse :\n\n" + display_comment,
        fontsize=15, color="#374151", va="top",
        transform=ax_c.transAxes,
        linespacing=1.5,
        clip_on=True,
    )

    # ── Numéro de page ────────────────────────────────────────────────
    ax_num = fig.add_axes([0, 0, 1, 1], zorder=4)
    ax_num.set_xlim(0, 1)
    ax_num.set_ylim(0, 1)
    ax_num.axis("off")
    ax_num.patch.set_alpha(0)
    ax_num.text(
        PAGE_NUM_X, PAGE_NUM_Y,
        f"Page {page_num}",
        ha="right", va="center", fontsize=11, color=GRIS_TEXTE, zorder=5,
    )


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
    """Retourne une liste de (theme, fig, plots) pour affichage Streamlit.
    moy_annuelle : dict optionnel {"col": valeur} issu de load_annee_precedente().
    """
    figures = []
    for theme, config in THEMES.items():
        plots = config["plots"]
        if config["type"] == "bar":
            fig = plt.figure(figsize=(8, 6))
            gs  = GridSpec(len(plots), 1, figure=fig)
            for i, (col, titre) in enumerate(plots):
                ax = fig.add_subplot(gs[i])
                make_ax_bar(ax, col, titre, evol_df)
        elif config["type"] == "single_hlines":
            fig = plt.figure(figsize=(8, 6))
            gs  = GridSpec(len(plots), 1, figure=fig)
            for i, (col, titre) in enumerate(plots):
                ax = fig.add_subplot(gs[i])
                moy = moy_annuelle.get(col) if moy_annuelle else None
                make_ax_hlines(ax, col, titre, config["objectif"][i], evol_df, moy_annuelle=moy)
        elif config["type"] == "multi":
            fig, ax = plt.subplots(figsize=(8, 6))
            make_ax_multi(ax, plots, theme, evol_df, moy_annuelle=moy_annuelle)
        else:
            fig = plt.figure(figsize=(8, 6))
            gs  = GridSpec(len(plots), 1, figure=fig)
            for i, (col, titre) in enumerate(plots):
                ax = fig.add_subplot(gs[i])
                make_ax(ax, col, titre, evol_df)
        figures.append((theme, fig, plots))
    return figures


# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION DU PDF
# ══════════════════════════════════════════════════════════════════════════════

def generate_pdf(evol_df, NOM_ETAB, PERIODE, custom_comments=None, moy_annuelle=None):
    """
    Génère le PDF et le retourne sous forme de bytes (pour st.download_button).
    """
    buf = io.BytesIO()

    with pdf_backend.PdfPages(buf) as pdf:

        # ── 1. Page de garde
        # Calcul KPI recette BR dernier mois pour la page de garde
        dernier = evol_df.iloc[-1]
        avant_dernier = evol_df.iloc[-2] if len(evol_df) > 1 else None
        try:
            val_br = dernier["recette_BR_moy_mois"]
            val_str = f"{val_br:,.0f} €"
            if avant_dernier is not None:
                delta = val_br - avant_dernier["recette_BR_moy_mois"]
                evo_str = f"{'▲' if delta >= 0 else '▼'} {delta:+,.0f} €"
            else:
                evo_str = ""
            cover_kpi = {"valeur": val_str, "evolution": evo_str}
        except Exception:
            cover_kpi = None
        fig = page_garde(NOM_ETAB, PERIODE, cover_kpi=cover_kpi)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # ── 2. Synthèse KPIs
        fig = page_synthese(evol_df)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # ── 3+. Pages graphiques
        page_num = 3
        for theme, config in THEMES.items():
            fig = plt.figure(figsize=(12, 17))
            fig.patch.set_facecolor(BLANC)
            _build_page_graphique(
                fig, theme, config, evol_df,
                page_num, NOM_ETAB, PERIODE, custom_comments,
                moy_annuelle=moy_annuelle,
            )
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)
            page_num += 1

        # Métadonnées PDF
        d = pdf.infodict()
        d["Title"]        = f"Rapport mensuel – {NOM_ETAB}"
        d["Author"]       = AUTEUR
        d["Subject"]      = f"Évolution mensuelle SSR – {PERIODE}"
        d["CreationDate"] = datetime.today()

    buf.seek(0)
    return buf.read()
