#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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
import zipfile
import tempfile
from io import BytesIO

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

OUTPUT_PDF = "rapport_mensuel.pdf"
LOGO_PATH = "./assets/logo_solimed.png"

AUTEUR = "SOLIMED"
SERVICE = "Rapport évolution mensuelle SSR"

MOIS_EXCLUS = ["2026_M1"]

BLEU = "#2563EB"
BLEU_FONCE = "#1E3A5F"
GRIS_TEXTE = "#6B7280"
GRIS_CLAIR = "#F3F4F6"
ROUGE = "#E11D48"
VERT = "#16A34A"
BLANC = "#FFFFFF"

# ─────────────────────────────────────────────
# UTILS ZIP
# ─────────────────────────────────────────────

def extract_zip(uploaded_zip):
    temp_dir = tempfile.TemporaryDirectory()

    with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
        zip_ref.extractall(temp_dir.name)

    root = Path(temp_dir.name)

    # gérer zip avec dossier racine unique
    subfolders = [f for f in root.iterdir() if f.is_dir()]
    if len(subfolders) == 1:
        root = subfolders[0]

    return root

# ─────────────────────────────────────────────
# UTILS DATA
# ─────────────────────────────────────────────

def extract_month(name):
    match = re.search(r"(\d{4}_M\d+|M\d+)$", name)
    return match.group(1) if match else None

def month_key(m):
    if "_" in m:
        y, m = m.split("_M")
        return (int(y), int(m))
    return (0, int(m[1:]))

# ─────────────────────────────────────────────
# CHARGEMENT DATA
# ─────────────────────────────────────────────

def load_data(path):
    html = {}

    for month_folder in path.iterdir():
        if month_folder.is_dir():
            month = extract_month(month_folder.name)
            if not month:
                continue

            html_files  = list(month_folder.glob("*.raev.html"))
            html_files2 = list(month_folder.glob("*.sv.html"))

            html[month] = {
                "raev": html_files[0] if html_files else None,
                "sv": html_files2[0] if html_files2 else None,
            }

    data = {}
    for month, files in html.items():
        data[month] = {}

        if files["raev"]:
            data[month]["raev"] = pd.read_html(files["raev"])[1]

        if files["sv"]:
            data[month]["sv"] = pd.read_html(files["sv"])[0]

    return data

# ─────────────────────────────────────────────
# CALCULS
# ─────────────────────────────────────────────

def compute_evol_df(data, valo_excel):

    sorted_months = sorted(data.keys(), key=month_key)
    evol_df = []

    for curr_mois in sorted_months:

        curr = data[curr_mois]["raev"]
        value_AM = curr.loc[
            curr["Zone de valorisation"].str.contains("TOTAL activité valorisée"),
            "Montant AM",
        ].iloc[0]

        value_AM = float(value_AM.replace(" ", "").replace(",", "."))

        curr2 = data[curr_mois]["sv"].iloc[[0, 11]].copy()

        col_ssrha = [c for c in curr2.columns if "SSRHA" in c][0]
        col_htp = [c for c in curr2.columns if "HTP" in c][0]

        curr2 = curr2.rename(columns={
            col_ssrha: "SSRHA",
            col_htp: "HTP"
        })

        for col in ["SSRHA", "HTP"]:
            curr2[col] = curr2[col].astype(str).str.replace(",", ".")
            curr2[col] = pd.to_numeric(curr2[col], errors="coerce")

        curr2["Mois"] = curr_mois

        df_month = curr2.groupby("Mois").sum()

        jours_valo = valo_excel[valo_excel["mois"] == curr_mois]["jours_valo"].values[0]
        df_month["jour_valo"] = jours_valo

        evol_df.append(df_month)

    evol_df = pd.concat(evol_df).reset_index()
    evol_df = evol_df[~evol_df["Mois"].isin(MOIS_EXCLUS)]

    return evol_df

# ─────────────────────────────────────────────
# FIGURES
# ─────────────────────────────────────────────

def generate_all_figures(uploaded_zip, uploaded_excel):

    path = extract_zip(uploaded_zip)
    NOM_ETAB = path.name

    valo_excel = pd.read_excel(uploaded_excel)

    data = load_data(path)
    evol_df = compute_evol_df(data, valo_excel)

    figures = []

    for col in evol_df.columns:
        if col == "Mois":
            continue

        fig, ax = plt.subplots()
        ax.plot(evol_df["Mois"], evol_df[col])
        ax.set_title(col)

        figures.append((col, fig, [(col, col)]))

    return figures

# ─────────────────────────────────────────────
# PDF
# ─────────────────────────────────────────────

def generate_pdf(figures):

    buffer = BytesIO()

    with pdf_backend.PdfPages(buffer) as pdf:
        for _, fig, _ in figures:
            pdf.savefig(fig)

    buffer.seek(0)
    return buffer