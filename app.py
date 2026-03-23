import streamlit as st
import core
import pandas as pd
import io
import base64
import requests

st.set_page_config(layout="wide")
st.title("Générateur de rapport mensuel SSR")

# ══════════════════════════════════════════════════════════════════════════════
#  GITHUB — lecture / écriture du fichier historique (colonnes brutes)
# ══════════════════════════════════════════════════════════════════════════════

GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]
GITHUB_REPO  = st.secrets["GITHUB_REPO"]
GITHUB_PATH  = st.secrets["GITHUB_PATH"]

GH_API = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_PATH}"
GH_HEADERS = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json",
}

def github_lire():
    r = requests.get(GH_API, headers=GH_HEADERS)
    if r.status_code == 404:
        return None, None
    r.raise_for_status()
    meta    = r.json()
    contenu = base64.b64decode(meta["content"])
    df      = pd.read_parquet(io.BytesIO(contenu))
    return df, meta["sha"]


def github_ecrire(df, sha, message):
    buf = io.BytesIO()
    df.to_parquet(buf, index=False)
    payload = {"message": message, "content": base64.b64encode(buf.getvalue()).decode()}
    if sha:
        payload["sha"] = sha
    requests.put(GH_API, headers=GH_HEADERS, json=payload).raise_for_status()


def month_key(m):
    try:
        annee, num = m.split("_M")
        return (int(annee), int(num))
    except Exception:
        return (9999, 9999)


# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT DE L'HISTORIQUE AU DÉMARRAGE
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="Récupération de l'historique sur GitHub…", ttl=60)
def recuperer_historique():
    try:
        return github_lire()
    except Exception as e:
        st.warning(f"⚠️ Impossible de lire l'historique GitHub : {e}")
        return None, None


hist_brut_df, hist_sha = recuperer_historique()

if hist_brut_df is not None:
    mois_connus = sorted(hist_brut_df["Mois"].unique(), key=month_key)
    st.info(
        f"📚 Historique chargé depuis GitHub — "
        f"**{len(mois_connus)} mois** : {' · '.join(mois_connus)}"
    )
else:
    st.info("📭 Aucun historique sur GitHub — premier chargement.")

# ══════════════════════════════════════════════════════════════════════════════
#  UPLOADS
# ══════════════════════════════════════════════════════════════════════════════

st.subheader("📂 Données à intégrer")

uploaded_excel = st.file_uploader("📊 Fichier Excel de valorisation (.xlsx)", type=["xlsx"])
uploaded_zip   = st.file_uploader(
    "📁 ZIP du nouveau mois à ajouter",
    type=["zip"],
    help="Peut contenir un ou plusieurs mois. Les doublons avec l'historique sont ignorés.",
)

if not uploaded_excel or not uploaded_zip:
    st.warning("Veuillez uploader le fichier ZIP et le fichier Excel.")
    st.stop()

NOM_ETAB = st.text_input("🏥 Nom de l'établissement", placeholder="ex : Ceyrat")
if not NOM_ETAB:
    st.warning("Veuillez saisir le nom de l'établissement.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT DU NOUVEAU MOIS (colonnes brutes)
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="Chargement des nouvelles données…")
def charger_brut(zip_bytes, excel_bytes):
    return core.load_data_brut(io.BytesIO(zip_bytes), io.BytesIO(excel_bytes))

nouveau = charger_brut(uploaded_zip.read(), uploaded_excel.read())
nouveau_brut_df = nouveau["brut_df"]

# ══════════════════════════════════════════════════════════════════════════════
#  FUSION BRUTES + RECALCUL DES DÉRIVÉES SUR LA SÉRIE COMPLÈTE
# ══════════════════════════════════════════════════════════════════════════════

if hist_brut_df is not None:
    mois_nouveaux = set(nouveau_brut_df["Mois"].unique())
    mois_hist     = set(hist_brut_df["Mois"].unique())
    doublons      = mois_nouveaux & mois_hist

    if doublons:
        st.warning(
            f"⚠️ Ces mois existent déjà dans l'historique et sont ignorés : "
            f"{', '.join(sorted(doublons, key=month_key))}"
        )
        nouveau_brut_df = nouveau_brut_df[~nouveau_brut_df["Mois"].isin(doublons)]

    if nouveau_brut_df.empty:
        st.warning("Aucun nouveau mois à ajouter — l'historique existant est affiché.")
        brut_complet = hist_brut_df
    else:
        brut_complet = pd.concat([hist_brut_df, nouveau_brut_df], ignore_index=True)
else:
    brut_complet = nouveau_brut_df

# Tri chronologique
brut_complet = brut_complet.iloc[
    brut_complet["Mois"].map(month_key).argsort()
].reset_index(drop=True)

# Recalcul des .diff() sur la série complète et triée
evol_df = core.recalculer_derives(brut_complet)

mois_tries = sorted(evol_df["Mois"].unique(), key=month_key)
PERIODE    = f"{mois_tries[0]} → {mois_tries[-1]}"

st.success(f"✅ Données prêtes — **{NOM_ETAB}** · {PERIODE}")
st.caption(f"Mois dans le rapport : {' · '.join(mois_tries)}")

# ══════════════════════════════════════════════════════════════════════════════
#  GRAPHES + COMMENTAIRES  (identique à l'original)
# ══════════════════════════════════════════════════════════════════════════════

comments = {}
figures  = core.generate_all_figures(evol_df)

for theme, fig, plots in figures:
    st.header(theme)
    col1, col2 = st.columns([2, 1])
    with col1:
        st.pyplot(fig)
    with col2:
        for col, titre in plots:
            auto_comment = core.generate_comment(col, titre, evol_df)
            edited = st.text_area(
                titre, value=auto_comment, height=120,
                key=f"{theme}_{col}",
            )
            comments[(theme, col)] = edited
    st.divider()

# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION PDF + SAUVEGARDE GITHUB
# ══════════════════════════════════════════════════════════════════════════════

st.subheader("📤 Export")

if st.button("📄 Générer le PDF et sauvegarder l'historique"):

    # 1. PDF — generate_pdf() reçoit le même evol_df qu'avant, rien ne change
    with st.spinner("Génération du PDF…"):
        pdf_bytes = core.generate_pdf(
            evol_df=evol_df,
            NOM_ETAB=NOM_ETAB,
            PERIODE=PERIODE,
            custom_comments=comments,
        )

    # 2. On sauvegarde les colonnes BRUTES sur GitHub (pas les dérivées)
    #    → au prochain mois, la fusion + recalcul sera propre
    with st.spinner("Sauvegarde de l'historique sur GitHub…"):
        try:
            _, sha_actuel = github_lire()
            github_ecrire(
                brut_complet,           # ← brutes uniquement, pas evol_df
                sha_actuel,
                f"historique: {PERIODE} — {NOM_ETAB}",
            )
            st.success("✅ Historique mis à jour sur GitHub.")
            recuperer_historique.clear()
        except Exception as e:
            st.error(f"❌ Erreur sauvegarde GitHub : {e}")

    # 3. Téléchargement PDF
    st.download_button(
        label="⬇️ Télécharger le rapport PDF",
        data=pdf_bytes,
        file_name=f"rapport_mensuel_{NOM_ETAB}_{PERIODE.replace(' → ', '_')}.pdf",
        mime="application/pdf",
    )
