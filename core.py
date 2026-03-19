#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSAR Tool - Génération automatique du rapport PDF mensuel SSR
Usage: python rapport_mensu.py
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
import glob
import subprocess

#pd.set_option('future.no_silent_downcasting', True)


# ══════════════════════════════════════════════════════════════════════════════
#  SECTION CONFIGURATION — tout ce qui est paramétrable est ici
# ══════════════════════════════════════════════════════════════════════════════

# ── Chemins ──────────────────────────────────────────────────────────────────

BASE_PATH = "./data/"
VALO_EXCEL = "/Users/marionducret/Desktop/SOLIMED/Rapport évolution mensuelle/Nb_jours_valo.xlsx"
OUTPUT_PDF = "rapport_mensuel.pdf"
LOGO_PATH = "./assets/logo_solimed.png"

# ── Identité du rapport ───────────────────────────────────────────────────────

AUTEUR  = "SOLIMED"
SERVICE = "Rapport évolution mensuelle SSR"

# ── Objectifs mensuels (en €) — modifiez selon l'établissement ───────────────

OBJECTIFS = {
    "recette_AM_moy_mois": 392_400,   # objectif recette AM mensuelle
    "recette_BR_moy_mois": 360_000,   # objectif recette BR mensuelle
}

# ── KPIs affichés sur la page de synthèse ────────────────────────────────────
# Chaque entrée : (colonne_evol_df, label_affiché, format, clé_objectif_ou_None)

KPI_CONFIG = [
    ("taux_valorisation_HC",  "Taux de valorisation HC",          "{:.1f} %",   None),
    ("recette_BR_moy_mois",   "Recette BR mensuelle",              "{:,.0f} €",  "recette_BR_moy_mois"),
    ("recette_AM_moy_mois",   "Recette AM mensuelle",              "{:,.0f} €",  "recette_AM_moy_mois"),
    ("recette_BR_moy_sej",    "Recette BR moy. / séjour",          "{:,.0f} €",  None),
    ("recette_BR_moy_jour",   "Recette BR moy. / jour valorisé",   "{:,.0f} €",  None),
    ("effectif_transmis_HC",  "Séjours transmis HC",               "{:.0f}",     None),
]

# ── Couleurs des cartes KPI (dans le même ordre que KPI_CONFIG) ───────────────
# (couleur_fond, couleur_bordure)

KPI_COULEURS = [
    ("#DBEAFE", "#2563EB"),   # taux valorisation  → bleu
    ("#DCFCE7", "#16A34A"),   # recette BR         → vert
    ("#FEF9C3", "#D97706"),   # recette AM         → jaune
    ("#F3E8FF", "#7C3AED"),   # BR/séjour          → violet
    ("#FEE2E2", "#E11D48"),   # BR/jour            → rouge
    ("#F3F4F6", "#6B7280"),   # séjours transmis   → gris
]

# ── Thèmes graphiques ─────────────────────────────────────────────────────────
# type : "bar" | "single" | "single_hlines" | "multi"
# Pour "single_hlines", "objectif" est une liste dans le même ordre que "plots".
#   → mettre None dans la liste pour afficher uniquement la moyenne globale (pas d'objectif fixe)

THEMES = {
    "Valorisation": {
        "type": "bar",
        "plots": [
            ("ecart_valo", "Écart de valorisation avec M-1"),
        ],
    },
    "Recette BR par jour": {
        "type": "single_hlines",
        "objectif": [None],           # None = moyenne globale uniquement, pas de ligne objectif
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
            ("jour_valo_supp_test", "Jours NON valorisés (test)"),   # ← variable test placeholder
        ],
    },
}

# ── Mois à exclure du rapport (ex : mois en cours incomplet) ─────────────────

MOIS_EXCLUS = ["2026_M1"]

# ── Palette graphique ─────────────────────────────────────────────────────────

COLORS          = ["#2563EB", "#16A34A", "#16A34A", "#E11D48", "#E11D48"]
BLEU_FONCE      = "#1E3A5F"
BLEU            = "#2563EB"
BLEU_CLAIR      = "#DBEAFE"
GRIS_TEXTE      = "#6B7280"
GRIS_CLAIR      = "#F3F4F6"
ROUGE           = "#E11D48"
VERT            = "#16A34A"
BLANC           = "#FFFFFF"


# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT DES DONNÉES
# ══════════════════════════════════════════════════════════════════════════════

result = subprocess.run(
    ["osascript", "-e",
     f'POSIX path of (choose folder with prompt "Sélectionnez un établissement"'
     f' default location POSIX file "{BASE_PATH}")'],
    capture_output=True, text=True,
)
ETAB_PATH = result.stdout.strip()
path = Path(ETAB_PATH)

# Nom affiché dans le rapport (dossier sélectionné)
NOM_ETAB = path.name


def extract_month(folder_name):
    match = re.search(r"(\d{4}_M\d+|M\d+)$", folder_name)
    return match.group(1) if match else None


def month_key(m):
    if "_" in m:
        year, month = m.split("_M")
        return (int(year), int(month))
    return (0, int(m[1:]))


def load_valo(periode, dossier):
    match_annee = re.match(r"(\d{4})_M(\d{1,2})", periode)
    match_mois  = re.match(r"M(\d{1,2})$", periode)
    if match_annee:
        annee = match_annee.group(1)
        mois  = match_annee.group(2).zfill(2)
    elif match_mois:
        annee = "2025"
        mois  = match_mois.group(1).zfill(2)
    else:
        raise ValueError(f"Format de période non reconnu : {periode}")
    pattern = f"{dossier}*.{annee}.{mois}.*VisualValoSejours.csv"
    fichiers = glob.glob(pattern)
    if not fichiers:
        raise FileNotFoundError(f"Aucun fichier trouvé pour {periode} ({pattern})")
    return pd.read_csv(fichiers[0], sep=";", encoding="latin-1")


# Lecture des fichiers HTML par mois
html = {}
for month_folder in path.iterdir():
    if month_folder.is_dir():
        month = extract_month(month_folder.name)
        if month is None:
            continue
        html_files  = list(month_folder.glob("*.ano-rha-sha.t1v1raev.html"))
        html_files2 = list(month_folder.glob("*.ano-rha-sha.t1v1sv.html"))
        html[month] = {
            "raev": html_files[0]  if html_files  else None,
            "sv":   html_files2[0] if html_files2 else None,
        }

# Chargement des tableaux
data = {}
for month, files in html.items():
    data[month] = {}
    if files["raev"]:
        data[month]["raev"] = pd.read_html(files["raev"])[1]
    if files["sv"]:
        data[month]["sv"] = pd.read_html(files["sv"])[0]

sorted_months = sorted(data.keys(), key=month_key)
valo_excel    = pd.read_excel(VALO_EXCEL)

# ══════════════════════════════════════════════════════════════════════════════
#  CALCULS
# ══════════════════════════════════════════════════════════════════════════════

evol_df = []

for curr_mois in sorted_months:

    # RAEV
    curr     = data[curr_mois]["raev"]
    value_AM = curr.loc[
        curr["Zone de valorisation"].str.contains("TOTAL activité valorisée"),
        "Montant AM",
    ].iloc[0]
    value_AM = float(value_AM.replace(" ", "").replace(",", "."))

    # SV
    curr2 = data[curr_mois]["sv"]
    curr2 = curr2.iloc[[0, 11]].copy()

    col_ssrha_br = [c for c in curr2.columns if "SSRHA" in c and "Montant BR" in c][0]
    col_htp_br   = [c for c in curr2.columns if "HTP"   in c and "Montant BR" in c][0]

    curr2 = curr2.rename(columns={
        col_ssrha_br: "SSRHA en HC - Montant BR",
        col_htp_br:   "Journées en HTP - Montant BR",
    })

    cols = ["SSRHA en HC - Montant BR", "Journées en HTP - Montant BR"]
    for col in cols:
        curr2[col] = (
            curr2[col]
            .replace(".", None)
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

    evol_df.append(df_month)

evol_df = pd.concat(evol_df, ignore_index=False)

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

# ── Variable test — jours non valorisés (placeholder à remplacer) ─────────────
evol_df["jour_valo_supp_test"] = 0   # remplacez par le vrai calcul quand disponible

# Exclusion des mois configurés
evol_df = evol_df[~evol_df["Mois"].isin(MOIS_EXCLUS)]

# Période pour les en-têtes
PERIODE = f"{evol_df['Mois'].iloc[0]} → {evol_df['Mois'].iloc[-1]}"


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
    """Bulle de valeur sur chaque point de la courbe."""
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


def make_ax(ax, col, titre, fmt="{:,.0f}"):
    x_vals = list(evol_df["Mois"])
    y_vals = evol_df[col].reset_index(drop=True)
    ax.plot(x_vals, y_vals, linewidth=2.5, color=BLEU,
            marker="o", markersize=5, markerfacecolor="white", markeredgewidth=2)
    ax.set_title(titre, fontsize=11, fontweight="bold", pad=10, loc="left")
    ax.grid(True, axis="y", linestyle="--", alpha=0.4, color="#9CA3AF")
    ax.grid(False, axis="x")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="x", rotation=45, labelsize=8)
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TEXTE)
    style_xticklabels(ax, x_vals, y_vals)
    annoter_tous_les_points(ax, x_vals, y_vals, fmt=fmt)


def make_ax_hlines(ax, col, titre, objectif, fmt="{:,.0f}"):
    """
    Courbe avec ligne de moyenne globale toujours affichée.
    Si objectif is not None, une ligne d'objectif est également tracée.
    """
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
    ax.grid(True, axis="y", linestyle="--", alpha=0.4, color="#9CA3AF")
    ax.grid(False, axis="x")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="x", rotation=45, labelsize=8)
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TEXTE)
    style_xticklabels(ax, x_vals, y_vals)
    annoter_tous_les_points(ax, x_vals, y_vals, fmt=fmt)


def make_ax_bar(ax, col, titre, fmt="{:,.0f}"):
    """
    Histogramme avec barres colorées selon signe (vert = positif, rouge = négatif)
    et valeurs chiffrées affichées au-dessus/dessous de chaque barre.
    """
    x_vals  = list(evol_df["Mois"])
    y_vals  = evol_df[col].reset_index(drop=True)
    couleurs = [VERT if v >= 0 else ROUGE for v in y_vals]
    bars = ax.bar(range(len(x_vals)), y_vals, color=couleurs, alpha=0.85, zorder=3)

    # Valeurs au-dessus/dessous de chaque barre
    for bar, val in zip(bars, y_vals):
        if np.isnan(val):
            continue
        try:
            label = fmt.format(val)
        except (ValueError, TypeError):
            label = str(val)
        va    = "bottom" if val >= 0 else "top"
        # Petit décalage pour que le texte ne colle pas à la barre
        offset = abs(y_vals.abs().max()) * 0.015 if y_vals.abs().max() != 0 else 1
        y_pos  = val + offset if val >= 0 else val - offset
        ax.text(
            bar.get_x() + bar.get_width() / 2, y_pos, label,
            ha="center", va=va, fontsize=8, fontweight="bold",
            color=VERT if val >= 0 else ROUGE,
        )

    ax.axhline(0, color=GRIS_TEXTE, linewidth=0.8, linestyle="-")
    ax.set_title(titre, fontsize=11, fontweight="bold", pad=10, loc="left")
    ax.grid(True, axis="y", linestyle="--", alpha=0.4, color="#9CA3AF")
    ax.grid(False, axis="x")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="x", rotation=45, labelsize=8)
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TEXTE)
    ax.set_xticks(range(len(x_vals)))
    ax.set_xticklabels(x_vals)
    style_xticklabels(ax, x_vals, y_vals)



def make_ax_multi(ax, plots, theme_title):
    x_vals = list(evol_df["Mois"])
    for i, (col, label) in enumerate(plots):
        y_vals = evol_df[col].reset_index(drop=True)
        ax.plot(x_vals, y_vals, linewidth=2.5, color=COLORS[i % len(COLORS)],
                marker="o", markersize=5, markerfacecolor="white",
                markeredgewidth=2, label=label)
        # Annotations sur tous les points de chaque courbe
        annoter_tous_les_points(ax, x_vals, y_vals, couleur=COLORS[i % len(COLORS)])
    ax.set_title("", fontsize=11, fontweight="bold", pad=10, loc="left")
    ax.legend(fontsize=9, framealpha=0.9, loc="best")
    ax.grid(True, axis="y", linestyle="--", alpha=0.4, color="#9CA3AF")
    ax.grid(False, axis="x")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="x", rotation=45, labelsize=8)
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TEXTE)
    ax.set_xticks(range(len(x_vals)))
    ax.set_xticklabels(x_vals)
    for i, label in enumerate(ax.get_xticklabels()):
        if i > 0:
            all_down = all(evol_df[col].iloc[i] < evol_df[col].iloc[i - 1] for col, _ in plots)
            label.set_color(ROUGE if all_down else GRIS_TEXTE)
        else:
            label.set_color(GRIS_TEXTE)


# ══════════════════════════════════════════════════════════════════════════════
#  FONCTIONS DE MISE EN FORME PDF
# ══════════════════════════════════════════════════════════════════════════════

def page_garde(nom_etablissement: str, periode: str,
               date_generation: str | None = None) -> plt.Figure:
    """Page de garde du rapport."""
    if date_generation is None:
        date_generation = datetime.today().strftime("%d/%m/%Y")

    fig = plt.figure(figsize=(12, 17))
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.patch.set_facecolor(BLANC)

    # Bande supérieure
    ax.add_patch(mpatches.FancyBboxPatch(
        (0, 0.72), 1, 0.28, boxstyle="square,pad=0",
        linewidth=0, facecolor=BLEU_FONCE,
    ))
    ax.axhline(y=0.72, xmin=0, xmax=1, color=BLEU, linewidth=4)

    ax.text(0.5, 0.88, "RAPPORT D'ÉVOLUTION MENSUELLE",
            ha="center", va="center", fontsize=28, fontweight="bold", color=BLANC)
    try:
        logo = imread(LOGO_PATH)
        ax_logo = fig.add_axes([0.35, 0.74, 0.30, 0.08])
        ax_logo.imshow(logo)
        ax_logo.axis("off")
    except FileNotFoundError:
        ax.text(0.5, 0.80, "MENSUELLE SSR",
                ha="center", va="center", fontsize=22, fontweight="bold", color=BLEU_CLAIR)

    # Bloc établissement
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

    # Pictos
    for x, couleur, label in [(0.25, BLEU, "Valorisation"),
                               (0.50, VERT, "Recettes"),
                               (0.75, ROUGE, "Activités")]:
        ax.add_patch(plt.Circle((x, 0.28), 0.055, color=couleur, zorder=3))
        ax.text(x, 0.28, label[0], ha="center", va="center",
                fontsize=14, fontweight="bold", color=BLANC, zorder=4)
        ax.text(x, 0.20, label, ha="center", va="center",
                fontsize=9, color=GRIS_TEXTE)

    ax.axhline(y=0.08, xmin=0.05, xmax=0.95, color=GRIS_CLAIR, linewidth=1)
    ax.text(0.5, 0.05, f"Généré le {date_generation} · {AUTEUR}",
            ha="center", va="center", fontsize=9, color=GRIS_TEXTE)

    return fig


def page_sommaire(themes: dict, page_depart: int = 4) -> plt.Figure:
    """Page de sommaire avec numérotation automatique."""
    fig = plt.figure(figsize=(12, 17))
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.patch.set_facecolor(BLANC)

    ax.add_patch(mpatches.FancyBboxPatch(
        (0, 0.88), 1, 0.12, boxstyle="square,pad=0",
        linewidth=0, facecolor=BLEU_FONCE,
    ))
    ax.text(0.5, 0.94, "SOMMAIRE",
            ha="center", va="center", fontsize=24, fontweight="bold", color=BLANC)
    ax.axhline(y=0.88, xmin=0, xmax=1, color=BLEU, linewidth=3)

    entrees = [("Introduction & synthèse", "Indicateurs clés du dernier mois", 3, GRIS_TEXTE)]

    page = page_depart
    for i, (nom_theme, config) in enumerate(themes.items()):
        if not nom_theme:
            continue
        n_plots   = len(config["plots"])
        sous_titre = f"{n_plots} graphique{'s' if n_plots > 1 else ''}"
        couleur = COLORS[i] if i < len(COLORS) else GRIS_TEXTE
        entrees.append((nom_theme, sous_titre, page, couleur))
        page += 1

    y_start    = 0.80
    espacement = 0.09

    for i, (titre, sous_titre, num_page, couleur) in enumerate(entrees):
        y = y_start - i * espacement

        ax.add_patch(plt.Circle((0.08, y), 0.025, color=couleur, zorder=3))
        ax.text(0.08, y, str(i + 1), ha="center", va="center",
                fontsize=11, fontweight="bold", color=BLANC, zorder=4)

        ax.text(0.15, y + 0.012, titre,
                ha="left", va="center", fontsize=13, fontweight="bold", color=BLEU_FONCE)
        ax.text(0.15, y - 0.015, sous_titre,
                ha="left", va="center", fontsize=10, color=GRIS_TEXTE)

        ax.annotate("", xy=(0.88, y), xytext=(0.58, y),
                    xycoords="axes fraction",
                    arrowprops=dict(arrowstyle="-", color=GRIS_CLAIR,
                                    linestyle="dotted", lw=1.5))

        ax.text(0.92, y, str(num_page),
                ha="center", va="center", fontsize=13,
                fontweight="bold", color=couleur)

        if i < len(entrees) - 1:
            ax.axhline(y=y - 0.04, xmin=0.05, xmax=0.95,
                       color=GRIS_CLAIR, linewidth=0.8)

    ax.axhline(y=0.04, xmin=0.05, xmax=0.95, color=GRIS_CLAIR, linewidth=1)
    ax.text(0.5, 0.02, "2", ha="center", va="center", fontsize=9, color=GRIS_TEXTE)

    return fig


def page_synthese(evol_df) -> plt.Figure:
    """
    Page de synthèse KPIs.
    Utilise KPI_CONFIG et KPI_COULEURS définis en section Configuration.
    """
    dernier      = evol_df.iloc[-1]
    avant_dernier = evol_df.iloc[-2] if len(evol_df) > 1 else None
    mois_label   = dernier["Mois"]

    fig = plt.figure(figsize=(12, 17))
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.patch.set_facecolor(BLANC)

    ax.add_patch(mpatches.FancyBboxPatch(
        (0, 0.88), 1, 0.12, boxstyle="square,pad=0",
        linewidth=0, facecolor=BLEU_FONCE,
    ))
    ax.text(0.5, 0.95, "SYNTHÈSE — INDICATEURS CLÉS",
            ha="center", va="center", fontsize=20, fontweight="bold", color=BLANC)
    ax.text(0.5, 0.90, f"Dernier mois disponible : {mois_label}",
            ha="center", va="center", fontsize=13, color=BLEU_CLAIR)
    ax.axhline(y=0.88, xmin=0, xmax=1, color=BLEU, linewidth=3)

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

    y_start    = 0.82
    card_h     = 0.10
    espacement = 0.115

    for i, (col, label, fmt, obj_key) in enumerate(KPI_CONFIG):
        y             = y_start - i * espacement
        couleur_fond, couleur_bord = KPI_COULEURS[i % len(KPI_COULEURS)]

        val = dernier.get(col, float("nan"))
        ref = avant_dernier.get(col) if avant_dernier is not None else None

        ax.add_patch(mpatches.FancyBboxPatch(
            (0.05, y - card_h + 0.01), 0.90, card_h,
            boxstyle="round,pad=0.01", linewidth=1.5,
            edgecolor=couleur_bord, facecolor=couleur_fond,
        ))

        ax.text(0.10, y - 0.025, label,
                ha="left", va="center", fontsize=11, color=GRIS_TEXTE)

        try:
            val_str = fmt.format(val)
        except (ValueError, TypeError):
            val_str = "N/A"
        ax.text(0.55, y - 0.025, val_str,
                ha="center", va="center", fontsize=16,
                fontweight="bold", color=BLEU_FONCE)

        try:
            fleche, couleur_fl = fleche_et_couleur(val, ref)
        except (TypeError, ValueError):
            fleche, couleur_fl = "–", GRIS_TEXTE
        ax.text(0.80, y - 0.025, fleche,
                ha="center", va="center", fontsize=11,
                fontweight="bold", color=couleur_fl)

        if obj_key and OBJECTIFS.get(obj_key) is not None:
            try:
                badge_txt, badge_col = badge_objectif(val, OBJECTIFS[obj_key])
                if badge_txt:
                    ax.text(0.55, y - 0.068, badge_txt,
                            ha="center", va="center", fontsize=9,
                            color=badge_col, style="italic")
            except (TypeError, ValueError):
                pass

    ax.axhline(y=0.04, xmin=0.05, xmax=0.95, color=GRIS_CLAIR, linewidth=1)
    ax.text(0.5, 0.02, "3", ha="center", va="center", fontsize=9, color=GRIS_TEXTE)

    return fig


def ajouter_entete_pied(fig: plt.Figure, titre_theme: str, num_page: int) -> None:
    """
    Ajoute un en-tête et un pied de page à une figure matplotlib existante.
    Modifie la figure en place.
    MODIFICATION : suppression de fig.subplots_adjust() ici — géré via
    fig.subplots_adjust() dans la boucle de génération pour éviter le conflit
    avec tight_layout sur les axes flottants d'en-tête/pied.
    """
    # En-tête
    ax_h = fig.add_axes([0, 0.91, 1, 0.09])
    ax_h.set_xlim(0, 1)
    ax_h.set_ylim(0, 1)
    ax_h.axis("off")
    ax_h.add_patch(mpatches.FancyBboxPatch(
        (0, 0), 1, 1, boxstyle="square,pad=0",
        linewidth=0, facecolor=BLEU_FONCE,
    ))
    ax_h.axhline(y=0, xmin=0, xmax=1, color=BLEU, linewidth=3)
    ax_h.text(0.03, 0.55, titre_theme.upper(),
              ha="left", va="center", fontsize=14, fontweight="bold", color=BLANC)
    ax_h.text(0.97, 0.65, NOM_ETAB,
              ha="right", va="center", fontsize=10, color=BLEU_CLAIR)
    ax_h.text(0.97, 0.30, PERIODE,
              ha="right", va="center", fontsize=9, color=GRIS_TEXTE)

    # Pied de page
    ax_f = fig.add_axes([0, 0, 1, 0.05])
    ax_f.set_xlim(0, 1)
    ax_f.set_ylim(0, 1)
    ax_f.axis("off")
    ax_f.axhline(y=0.85, xmin=0.02, xmax=0.98, color=GRIS_CLAIR, linewidth=0.8)
    ax_f.text(0.03, 0.35, f"{AUTEUR} · {SERVICE}",
              ha="left", va="center", fontsize=8, color=GRIS_TEXTE)
    ax_f.text(0.97, 0.35, f"Page {num_page}",
              ha="right", va="center", fontsize=8,
              fontweight="bold", color=GRIS_TEXTE)
    ax_f.text(0.5, 0.35, datetime.today().strftime("Généré le %d/%m/%Y"),
              ha="center", va="center", fontsize=8, color=GRIS_TEXTE)


# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION DES COMMENTAIRES
# ══════════════════════════════════════════════════════════════════════════════

def generate_comment(col, titre):
    data = evol_df[col].dropna()

    if len(data) < 2:
        return "Données insuffisantes pour analyse."

    trend = data.iloc[-1] - data.iloc[0]
    trend_pct = (trend / data.iloc[0]) * 100 if data.iloc[0] != 0 else 0

    if trend > 0:
        tendance = "hausse"
    elif trend < 0:
        tendance = "baisse"
    else:
        tendance = "stabilité"

    return (
        f"{titre} : On observe une {tendance} globale de {trend_pct:.1f}% "
        f"sur la période. La valeur moyenne est de {data.mean():,.0f}, "
        f"avec un minimum de {data.min():,.0f} et un maximum de {data.max():,.0f}."
    )

# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION DU PDF
# ══════════════════════════════════════════════════════════════════════════════
def generate_pdf(custom_comments=None):
    with pdf_backend.PdfPages(OUTPUT_PDF) as pdf:
    
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
            plots = config["plots"]
    
            if config["type"] == "bar":
                n   = len(plots)
                fig = plt.figure(figsize=(12, 12))
                fig.suptitle(theme, fontsize=18, fontweight="bold", color=BLEU_FONCE)
                # MODIFICATION : hspace retiré du GridSpec, géré par subplots_adjust ci-dessous
                gs  = GridSpec(n, 1, figure=fig)
                for i, (col, titre) in enumerate(plots):
                    ax = fig.add_subplot(gs[i])
                    make_ax_bar(ax, col, titre)
    
            elif config["type"] == "single_hlines":
                n         = len(plots)
                objectifs = config["objectif"]
                fig       = plt.figure(figsize=(12, 12))
                fig.suptitle(theme, fontsize=18, fontweight="bold", color=BLEU_FONCE)
                # MODIFICATION : hspace retiré du GridSpec, géré par subplots_adjust ci-dessous
                gs = GridSpec(n, 1, figure=fig)
                for i, (col, titre) in enumerate(plots):
                    ax = fig.add_subplot(gs[i])
                    make_ax_hlines(ax, col, titre, objectifs[i])
    
            elif config["type"] == "multi":
                fig, ax = plt.subplots(figsize=(12, 12))
                fig.suptitle(theme, fontsize=18, fontweight="bold", color=BLEU_FONCE)
                make_ax_multi(ax, plots, theme)
    
            else:  # "single"
                n   = len(plots)
                fig = plt.figure(figsize=(12, 12))
                fig.suptitle(theme, fontsize=18, fontweight="bold", color=BLEU_FONCE)
                # MODIFICATION : hspace retiré du GridSpec, géré par subplots_adjust ci-dessous
                gs = GridSpec(n, 1, figure=fig)
                for i, (col, titre) in enumerate(plots):
                    ax = fig.add_subplot(gs[i])
                    make_ax(ax, col, titre)
    
            ajouter_entete_pied(fig, theme or "Activité", page_num)
            fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.25, hspace=0.6)  
            
            #AJOUT COMMENTAIRE
            comment_texts = []
            
            for col, titre in plots:
                comment_texts.append(generate_comment(col, titre))
            
            full_comment = "\n\n".join(comment_texts)
            
            # encart visuel
            ax_comment = fig.add_axes([0.08, 0.05, 0.89, 0.15])
            ax_comment.axis("off")
            
            ax_comment.text(
                0, 1,
                "Analyse :\n\n" + full_comment,
                fontsize=9,
                color="#374151",
                va="top",
                wrap=True
            )
            
            # fond léger
            ax_comment.add_patch(
                mpatches.FancyBboxPatch(
                    (0, 0),
                    1,
                    1,
                    boxstyle="round,pad=0.02",
                    facecolor="#F9FAFB",
                    edgecolor="#E5E7EB"
                )
            )
                    
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)
            page_num += 1
    
        # ── Métadonnées PDF
        d = pdf.infodict()
        d["Title"]        = f"Rapport mensuel – {NOM_ETAB}"
        d["Author"]       = AUTEUR
        d["Subject"]      = f"Évolution mensuelle SSR – {PERIODE}"
        d["CreationDate"] = datetime.today()

# pour afficher les graphes dans Streamlit
def generate_all_figures():
    figures = []

    for theme, config in THEMES.items():
        plots = config["plots"]

        if config["type"] == "bar":
            fig = plt.figure(figsize=(8, 6))
            gs = GridSpec(len(plots), 1, figure=fig)

            for i, (col, titre) in enumerate(plots):
                ax = fig.add_subplot(gs[i])
                make_ax_bar(ax, col, titre)

        elif config["type"] == "single_hlines":
            fig = plt.figure(figsize=(8, 6))
            gs = GridSpec(len(plots), 1, figure=fig)

            for i, (col, titre) in enumerate(plots):
                ax = fig.add_subplot(gs[i])
                make_ax_hlines(ax, col, titre, config["objectif"][i])

        elif config["type"] == "multi":
            fig, ax = plt.subplots(figsize=(8, 6))
            make_ax_multi(ax, plots, theme)

        else:
            fig = plt.figure(figsize=(8, 6))
            gs = GridSpec(len(plots), 1, figure=fig)

            for i, (col, titre) in enumerate(plots):
                ax = fig.add_subplot(gs[i])
                make_ax(ax, col, titre)

        figures.append((theme, fig, plots))

    return figures