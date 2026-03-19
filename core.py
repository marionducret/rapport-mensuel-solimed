#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSAR Tool - Génération automatique du rapport PDF mensuel SSR
Streamlit Cloud version : toutes les données sont chargées via load_data().
"""

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

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION CONFIGURATION — tout ce qui est paramétrable est ici
# ══════════════════════════════════════════════════════════════════════════════

OUTPUT_PDF  = "rapport_mensuel.pdf"

# ── Templates Canva ───────────────────────────────────────────────────────────
# Déposer les PNG exportés depuis Canva dans ./assets/
CANVA_COVER_PATH = "./design/page_garde.png"
CANVA_PAGE_PATH  = "./design/page_graph.png"

AUTEUR  = "SOLIMED"
SERVICE = "Rapport évolution mensuelle SSR"

OBJECTIFS = {
    "recette_AM_moy_mois": 392_400,
    "recette_BR_moy_mois": 360_000,
}

KPI_CONFIG = [
    ("taux_valorisation_HC",  "Taux de valorisation HC",          "{:.1f} %",   None),
    ("recette_BR_moy_mois",   "Recette BR mensuelle",              "{:,.0f} €",  "recette_BR_moy_mois"),
    ("recette_AM_moy_mois",   "Recette AM mensuelle",              "{:,.0f} €",  "recette_AM_moy_mois"),
    ("recette_BR_moy_sej",    "Recette BR moy. / séjour",          "{:,.0f} €",  None),
    ("recette_BR_moy_jour",   "Recette BR moy. / jour valorisé",   "{:,.0f} €",  None),
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
    "Recette BR par jour": {
        "type": "single_hlines",
        "objectif": [None],
        "plots": [
            ("recette_BR_moy_jour", "Recette brute moyenne par jour valorisé"),
        ],
    },
    "Recette mensuelle BR": {
        "type": "single_hlines",
        "objectif": [OBJECTIFS["recette_BR_moy_mois"]],
        "plots": [
            ("recette_BR_moy_mois", "Recette BR mensuelle"),
        ],
    },
    "Séjours": {
        "type": "multi",
        "plots": [
            ("sejour_supp",       "Séjour supplémentaire par rapport à M-1"),
            ("sejour_valo_supp",  "Séjour valorisé supplémentaire par rapport à M-1"),
        ],
    },
    "Jours valorisés": {
        "type": "multi",
        "plots": [
            ("jour_valo_supp",      "Jour valorisé supplémentaire par rapport à M-1"),
            ("jour_valo_supp_test", "Jours NON valorisés (test)"),
        ],
    },
}

MOIS_EXCLUS = ["2026_M1"]

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


# ══════════════════════════════════════════════════════════════════════════════
#  UTILITAIRES CANVA
# ══════════════════════════════════════════════════════════════════════════════

def _charger_bg(path: str):
    """Charge un PNG Canva comme background. Retourne None si non trouvé."""
    try:
        return imread(path)
    except Exception:
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
    evol_df["jour_valo_supp_test"] = 0

    NOM_ETAB = "Extraction"
    PERIODE  = f"{evol_df['Mois'].iloc[0]} → {evol_df['Mois'].iloc[-1]}"

    return {
        "evol_df":  evol_df,
        "NOM_ETAB": NOM_ETAB,
        "PERIODE":  PERIODE,
        "_tmp_dir": tmp,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  FONCTIONS GRAPHIQUES
# ══════════════════════════════════════════════════════════════════════════════

def style_xticklabels(ax, x_vals, y_vals):
    ax.set_xticks(range(len(x_vals)))
    ax.set_xticklabels(x_vals)
    for i, label in enumerate(ax.get_xticklabels()):
        if i > 0 and y_vals.iloc[i] < y_vals.iloc[i - 1]:
            label.set_color(ROUGE)
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
            fontsize=7, fontweight="bold", color=couleur,
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


def make_ax(ax, col, titre, evol_df, fmt="{:,.0f}"):
    x_vals = list(evol_df["Mois"])
    y_vals = evol_df[col].reset_index(drop=True)
    ax.plot(x_vals, y_vals, linewidth=2.5, color=BLEU,
            marker="o", markersize=5, markerfacecolor="white", markeredgewidth=2)
    ax.set_title(titre, fontsize=11, fontweight="bold", pad=10, loc="left")
    _style_ax(ax)
    style_xticklabels(ax, x_vals, y_vals)
    annoter_tous_les_points(ax, x_vals, y_vals, fmt=fmt)


def make_ax_hlines(ax, col, titre, objectif, evol_df, fmt="{:,.0f}"):
    x_vals = list(evol_df["Mois"])
    y_vals = evol_df[col].reset_index(drop=True)
    ax.plot(x_vals, y_vals, linewidth=2.5, color=BLEU,
            marker="o", markersize=5, markerfacecolor="white", markeredgewidth=2)
    moyenne = y_vals.mean()
    ax.axhline(moyenne, color="#9CA3AF", linestyle="--", linewidth=1.5,
               label=f"Moyenne globale ({moyenne:,.0f})")
    if objectif is not None:
        ax.axhline(objectif, color=ROUGE, linestyle="--", linewidth=1.5,
                   label=f"Objectif mensuel ({objectif:,.0f})")
    ax.set_title(titre, fontsize=11, fontweight="bold", pad=10, loc="left")
    ax.legend(fontsize=9, framealpha=0.9, loc="best")
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
            ha="center", va=va, fontsize=8, fontweight="bold",
            color=VERT if val >= 0 else ROUGE,
        )
    ax.axhline(0, color=GRIS_TEXTE, linewidth=0.8, linestyle="-")
    ax.set_title(titre, fontsize=11, fontweight="bold", pad=10, loc="left")
    _style_ax(ax)
    ax.set_xticks(range(len(x_vals)))
    ax.set_xticklabels(x_vals)
    style_xticklabels(ax, x_vals, y_vals)


def make_ax_multi(ax, plots, theme_title, evol_df):
    x_vals = list(evol_df["Mois"])
    for i, (col, label) in enumerate(plots):
        y_vals = evol_df[col].reset_index(drop=True)
        ax.plot(x_vals, y_vals, linewidth=2.5, color=COLORS[i % len(COLORS)],
                marker="o", markersize=5, markerfacecolor="white",
                markeredgewidth=2, label=label)
        annoter_tous_les_points(ax, x_vals, y_vals, couleur=COLORS[i % len(COLORS)])
    ax.set_title("", fontsize=11, fontweight="bold", pad=10, loc="left")
    ax.legend(fontsize=9, framealpha=0.9, loc="best")
    _style_ax(ax)
    ax.set_xticks(range(len(x_vals)))
    ax.set_xticklabels(x_vals)
    for i, label in enumerate(ax.get_xticklabels()):
        if i > 0:
            all_down = all(evol_df[col].iloc[i] < evol_df[col].iloc[i - 1] for col, _ in plots)
            label.set_color(ROUGE if all_down else GRIS_TEXTE)
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
COVER_PRES_LABEL_Y  = 0.654   # y du label "Présenté par"
COVER_NOM_ETAB_Y    = 0.622   # y de la valeur NOM_ETAB (sous le label)
COVER_PERI_LABEL_Y  = 0.579   # y du label "Période"
COVER_PERIODE_Y     = 0.547   # y de la valeur PERIODE
COVER_DATE_Y        = 0.200   # y de la date (dans le bloc teal bas-droite)
COVER_DATE_X        = 0.650   # x de la date (centre du bloc teal)
COVER_TEXT_X        = 0.091   # x de départ des textes dynamiques

# PAGE GRAPHIQUE (calibration pixel-perfect sur le PNG 1414×2000)
# Bandeau titre teal : mpl_y centre ≈ 0.944, x 0.03 → 0.63
PAGE_TITRE_X        = 0.040
PAGE_TITRE_Y        = 0.944
# Grand bloc graphique   : [left=0.066, bottom=0.363, w=0.868, h=0.488]
PAGE_GRAPH_LEFT     = 0.066
PAGE_GRAPH_BOTTOM   = 0.363
PAGE_GRAPH_WIDTH    = 0.868
PAGE_GRAPH_HEIGHT   = 0.488
# Petit bloc commentaire : [left=0.066, bottom=0.062, w=0.868, h=0.247]
PAGE_COMMENT_LEFT   = 0.066
PAGE_COMMENT_BOTTOM = 0.062
PAGE_COMMENT_WIDTH  = 0.868
PAGE_COMMENT_HEIGHT = 0.247
# Numéro de page + infos (dans ou sous le bloc commentaire)
PAGE_NUM_X          = 0.940
PAGE_NUM_Y          = 0.030


def page_garde(nom_etablissement: str, periode: str,
               date_generation: str | None = None) -> plt.Figure:
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

        # Valeur "Présenté par"
        ax.text(
            COVER_TEXT_X, COVER_NOM_ETAB_Y,
            nom_etablissement,
            ha="left", va="center",
            fontsize=16, fontweight="bold", color=TEAL,
            zorder=2,
        )
        # Valeur "Période"
        ax.text(
            COVER_TEXT_X, COVER_PERIODE_Y,
            periode,
            ha="left", va="center",
            fontsize=14, color=TEAL,
            zorder=2,
        )


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
        ax.text(0.5, 0.54, f"Période analysée : {periode}",
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


def page_sommaire(themes: dict, page_depart: int = 4) -> plt.Figure:
    fig = plt.figure(figsize=(12, 17))
    fig.patch.set_facecolor(BLANC)

    bg = _charger_bg(CANVA_PAGE_PATH)
    if bg is not None:
        _appliquer_bg(fig, bg)

    # Axe principal transparent si Canva, sinon fond bleu
    ax = fig.add_axes([0, 0, 1, 1], zorder=1)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    if bg is not None:
        ax.patch.set_alpha(0)
        # Titre dans le bandeau teal
        ax.text(PAGE_TITRE_X, PAGE_TITRE_Y, "SOMMAIRE",
                ha="left", va="center",
                fontsize=18, fontweight="bold", color=BLANC, zorder=2)
    else:
        ax.add_patch(mpatches.FancyBboxPatch(
            (0, 0.88), 1, 0.12, boxstyle="square,pad=0",
            linewidth=0, facecolor=BLEU_FONCE,
        ))
        ax.axhline(y=0.88, xmin=0, xmax=1, color=BLEU, linewidth=3)
        ax.text(0.5, 0.94, "SOMMAIRE",
                ha="center", va="center", fontsize=24, fontweight="bold", color=BLANC)

    # Contenu sommaire — dans le grand bloc Canva (y: 0.851→0.363)
    entrees = [("Introduction & synthèse", "Indicateurs clés du dernier mois", 3, GRIS_TEXTE)]
    page = page_depart
    for i, (nom_theme, config) in enumerate(themes.items()):
        if not nom_theme:
            continue
        n_plots    = len(config["plots"])
        sous_titre = f"{n_plots} graphique{'s' if n_plots > 1 else ''}"
        couleur    = COLORS[i] if i < len(COLORS) else GRIS_TEXTE
        entrees.append((nom_theme, sous_titre, page, couleur))
        page += 1

    # Zone de liste dans le grand bloc Canva
    y_top      = 0.820
    y_bottom   = 0.380
    espacement = (y_top - y_bottom) / max(len(entrees), 1)

    for i, (titre, sous_titre, num_page, couleur) in enumerate(entrees):
        y = y_top - i * espacement

        ax.add_patch(plt.Circle((0.09, y), 0.018, color=couleur, zorder=3))
        ax.text(0.09, y, str(i + 1), ha="center", va="center",
                fontsize=9, fontweight="bold", color=BLANC, zorder=4)
        ax.text(0.13, y + 0.010, titre,
                ha="left", va="center", fontsize=12, fontweight="bold", color=BLEU_FONCE)
        ax.text(0.13, y - 0.012, sous_titre,
                ha="left", va="center", fontsize=9, color=GRIS_TEXTE)
        ax.annotate("", xy=(0.87, y), xytext=(0.60, y),
                    arrowprops=dict(arrowstyle="-", color=GRIS_CLAIR,
                                    linestyle="dotted", lw=1.5))
        ax.text(0.92, y, str(num_page),
                ha="center", va="center", fontsize=12,
                fontweight="bold", color=couleur)
        if i < len(entrees) - 1:
            ax.axhline(y=y - espacement * 0.45, xmin=0.07, xmax=0.95,
                       color=GRIS_CLAIR, linewidth=0.7)

    # Numéro de page bas
    ax.text(PAGE_NUM_X, PAGE_NUM_Y, "2",
            ha="right", va="center", fontsize=9,
            fontweight="bold", color=GRIS_TEXTE, zorder=2)

    return fig


def page_synthese(evol_df) -> plt.Figure:
    dernier       = evol_df.iloc[-1]
    avant_dernier = evol_df.iloc[-2] if len(evol_df) > 1 else None
    mois_label    = dernier["Mois"]

    fig = plt.figure(figsize=(12, 17))
    fig.patch.set_facecolor(BLANC)

    bg = _charger_bg(CANVA_PAGE_PATH)
    if bg is not None:
        _appliquer_bg(fig, bg)

    ax = fig.add_axes([0, 0, 1, 1], zorder=1)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    if bg is not None:
        ax.patch.set_alpha(0)
        ax.text(PAGE_TITRE_X, PAGE_TITRE_Y,
                "SYNTHÈSE — INDICATEURS CLÉS",
                ha="left", va="center",
                fontsize=16, fontweight="bold", color=BLANC, zorder=2)
    else:
        ax.add_patch(mpatches.FancyBboxPatch(
            (0, 0.88), 1, 0.12, boxstyle="square,pad=0",
            linewidth=0, facecolor=BLEU_FONCE,
        ))
        ax.axhline(y=0.88, xmin=0, xmax=1, color=BLEU, linewidth=3)
        ax.text(0.5, 0.95, "SYNTHÈSE — INDICATEURS CLÉS",
                ha="center", va="center", fontsize=20, fontweight="bold", color=BLANC)
        ax.text(0.5, 0.90, f"Dernier mois disponible : {mois_label}",
                ha="center", va="center", fontsize=13, color=BLEU_CLAIR)

    # Sous-titre mois dans le grand bloc
    ax.text(0.5, 0.838,
            f"Dernier mois disponible : {mois_label}",
            ha="center", va="center", fontsize=12,
            color=TEAL if bg is not None else BLEU_CLAIR,
            fontweight="bold", zorder=2)

    def fleche_et_couleur(val, ref):
        if ref is None or np.isnan(ref):
            return "–", GRIS_TEXTE
        delta = val - ref
        if delta > 0:
            return f"▲ +{delta:,.0f}", VERT
        elif delta < 0:
            return f"▼ {delta:,.0f}", ROUGE
        return "= stable", GRIS_TEXTE

    def badge_objectif(val, objectif):
        if objectif is None:
            return None, None
        if val >= objectif:
            return f"✓ Objectif atteint ({objectif:,.0f} €)", VERT
        pct = (1 - val / objectif) * 100
        return f"✗ -{pct:.1f}% de l'objectif ({objectif:,.0f} €)", ROUGE

    # Cartes KPI dans le grand bloc (y: 0.810 → 0.380)
    y_top      = 0.805
    card_h     = 0.068
    espacement = 0.073

    for i, (col, label, fmt, obj_key) in enumerate(KPI_CONFIG):
        y                       = y_top - i * espacement
        couleur_fond, couleur_bord = KPI_COULEURS[i % len(KPI_COULEURS)]
        val = dernier.get(col, float("nan"))
        ref = avant_dernier.get(col) if avant_dernier is not None else None

        ax.add_patch(mpatches.FancyBboxPatch(
            (0.068, y - card_h + 0.005), 0.864, card_h,
            boxstyle="round,pad=0.01", linewidth=1.5,
            edgecolor=couleur_bord, facecolor=couleur_fond, zorder=2,
        ))
        ax.text(0.11, y - 0.022, label,
                ha="left", va="center", fontsize=10, color=GRIS_TEXTE, zorder=3)
        try:
            val_str = fmt.format(val)
        except (ValueError, TypeError):
            val_str = "N/A"
        ax.text(0.55, y - 0.022, val_str,
                ha="center", va="center", fontsize=15,
                fontweight="bold", color=BLEU_FONCE, zorder=3)
        try:
            fleche, couleur_fl = fleche_et_couleur(val, ref)
        except (TypeError, ValueError):
            fleche, couleur_fl = "–", GRIS_TEXTE
        ax.text(0.80, y - 0.022, fleche,
                ha="center", va="center", fontsize=10,
                fontweight="bold", color=couleur_fl, zorder=3)
        if obj_key and OBJECTIFS.get(obj_key) is not None:
            try:
                badge_txt, badge_col = badge_objectif(val, OBJECTIFS[obj_key])
                if badge_txt:
                    ax.text(0.55, y - 0.055, badge_txt,
                            ha="center", va="center", fontsize=8,
                            color=badge_col, style="italic", zorder=3)
            except (TypeError, ValueError):
                pass

    ax.text(PAGE_NUM_X, PAGE_NUM_Y, "3",
            ha="right", va="center", fontsize=9,
            fontweight="bold", color=GRIS_TEXTE, zorder=2)

    return fig


def _build_page_graphique(fig: plt.Figure, theme: str, config: dict,
                          evol_df, page_num: int, NOM_ETAB: str,
                          PERIODE: str, custom_comments=None) -> None:
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
            PAGE_TITRE_X, PAGE_TITRE_Y,
            theme.upper(),
            ha="left", va="center",
            fontsize=16, fontweight="bold", color=BLANC, zorder=3,
        )
        ax_titre.text(
            0.96, PAGE_TITRE_Y + 0.015,
            NOM_ETAB,
            ha="right", va="center",
            fontsize=9, color=BLANC, zorder=3,
        )
        ax_titre.text(
            0.96, PAGE_TITRE_Y - 0.020,
            PERIODE,
            ha="right", va="center",
            fontsize=8, color=BLANC, zorder=3,
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
    h_plot = PAGE_GRAPH_HEIGHT / n
    hspace = 0.08

    axes_graph = []
    for i in range(n):
        bottom = PAGE_GRAPH_BOTTOM + (n - 1 - i) * h_plot + hspace / 2
        height = h_plot - hspace
        ax_g = fig.add_axes(
            [PAGE_GRAPH_LEFT, bottom, PAGE_GRAPH_WIDTH, height],
            zorder=3,
        )
        axes_graph.append(ax_g)

    for i, (col, titre) in enumerate(plots):
        ax_g = axes_graph[i]
        if config["type"] == "bar":
            make_ax_bar(ax_g, col, titre, evol_df)
        elif config["type"] == "single_hlines":
            make_ax_hlines(ax_g, col, titre, config["objectif"][i], evol_df)
        elif config["type"] == "multi":
            make_ax_multi(ax_g, plots, theme, evol_df)
            # multi utilise un seul ax pour toutes les séries
            break
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

    ax_c = fig.add_axes(
        [PAGE_COMMENT_LEFT, PAGE_COMMENT_BOTTOM,
         PAGE_COMMENT_WIDTH, PAGE_COMMENT_HEIGHT],
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

    ax_c.text(
        0.01, 0.95,
        "Analyse :\n\n" + full_comment,
        fontsize=8.5, color="#374151", va="top", wrap=True,
        transform=ax_c.transAxes,
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
        ha="right", va="center", fontsize=8, color=GRIS_TEXTE, zorder=5,
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

def generate_all_figures(evol_df):
    """Retourne une liste de (theme, fig, plots) pour affichage Streamlit."""
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
                make_ax_hlines(ax, col, titre, config["objectif"][i], evol_df)
        elif config["type"] == "multi":
            fig, ax = plt.subplots(figsize=(8, 6))
            make_ax_multi(ax, plots, theme, evol_df)
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

def generate_pdf(evol_df, NOM_ETAB, PERIODE, custom_comments=None):
    """
    Génère le PDF et le retourne sous forme de bytes (pour st.download_button).
    """
    buf = io.BytesIO()

    with pdf_backend.PdfPages(buf) as pdf:

        # ── 1. Page de garde
        fig = page_garde(NOM_ETAB, PERIODE)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # ── 2. Sommaire
        fig = page_sommaire(THEMES, page_depart=4)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # ── 3. Synthèse KPIs
        fig = page_synthese(evol_df)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        # ── 4+. Pages graphiques
        page_num = 4
        for theme, config in THEMES.items():
            fig = plt.figure(figsize=(12, 17))
            fig.patch.set_facecolor(BLANC)
            _build_page_graphique(
                fig, theme, config, evol_df,
                page_num, NOM_ETAB, PERIODE, custom_comments,
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
