
import streamlit as st
import core
import pandas as pd
import io
import json
import base64
import requests

st.set_page_config(layout="wide")
st.title("Générateur de rapport mensuel SSR")

#debug background
st.write("test bg:", str(Path(core.__file__).parent / core.CANVA_COVER_PATH))
from PIL import Image
try:
    img = Image.open(str(Path(core.__file__).parent / core.CANVA_COVER_PATH))
    st.write("✅ Image ouverte:", img.size)
except Exception as e:
    st.write("❌ Erreur:", e)
# ══════════════════════════════════════════════════════════════════════════════
#  GITHUB
# ══════════════════════════════════════════════════════════════════════════════


GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]
GITHUB_REPO  = st.secrets["GITHUB_REPO"]

GH_HEADERS = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json",
}


def gh_url(path):
    return f"https://api.github.com/repos/{GITHUB_REPO}/contents/{path}"


def github_lire_parquet(nom_etab):
    slug = nom_etab.lower().replace(" ", "_")
    r = requests.get(gh_url(f"data/historique_{slug}.parquet"), headers=GH_HEADERS)
    if r.status_code == 404:
        return None, None
    r.raise_for_status()
    meta = r.json()
    return pd.read_parquet(io.BytesIO(base64.b64decode(meta["content"]))), meta["sha"]


def github_ecrire_parquet(df, sha, nom_etab, message):
    slug = nom_etab.lower().replace(" ", "_")
    buf  = io.BytesIO()
    df.to_parquet(buf, index=False)
    payload = {"message": message, "content": base64.b64encode(buf.getvalue()).decode()}
    if sha:
        payload["sha"] = sha
    requests.put(gh_url(f"data/historique_{slug}.parquet"), headers=GH_HEADERS, json=payload).raise_for_status()


def github_lire_moy(nom_etab):
    slug = nom_etab.lower().replace(" ", "_")
    r = requests.get(gh_url(f"data/moy_annuelle_{slug}.json"), headers=GH_HEADERS)
    if r.status_code == 404:
        return None, None
    r.raise_for_status()
    meta = r.json()
    return json.loads(base64.b64decode(meta["content"])), meta["sha"]


def github_ecrire_moy(moy_dict, sha, nom_etab):
    slug    = nom_etab.lower().replace(" ", "_")
    payload = {
        "message": f"moy_annuelle: {nom_etab}",
        "content": base64.b64encode(json.dumps(moy_dict).encode()).decode(),
    }
    if sha:
        payload["sha"] = sha
    requests.put(gh_url(f"data/moy_annuelle_{slug}.json"), headers=GH_HEADERS, json=payload).raise_for_status()


def month_key(m):
    try:
        annee, num = m.split("_M")
        return (int(annee), int(num))
    except Exception:
        return (9999, 9999)


# ══════════════════════════════════════════════════════════════════════════════
#  NOM ÉTABLISSEMENT
# ══════════════════════════════════════════════════════════════════════════════

NOM_ETAB = st.text_input("🏥 Nom de l'établissement", placeholder="Attention à TOUJOURS bien mettre le même nom ! (Exemple : LB-Monchy)")
if not NOM_ETAB:
    st.warning("Veuillez saisir le nom de l'établissement.")
    st.stop()

NOM_ETAB = f"Centre Médical de {NOM_ETAB}"

# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT HISTORIQUE + MOY ANNUELLE DEPUIS GITHUB
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="Récupération de l'historique sur GitHub…", ttl=60)
def recuperer_historique(nom_etab):
    try:
        return github_lire_parquet(nom_etab)
    except Exception as e:
        st.warning(f"⚠️ Impossible de lire l'historique : {e}")
        return None, None


@st.cache_data(show_spinner="Récupération des moyennes annuelles…", ttl=60)
def recuperer_moy_annuelle(nom_etab):
    try:
        return github_lire_moy(nom_etab)
    except Exception as e:
        return None, None


hist_brut_df, hist_sha    = recuperer_historique(NOM_ETAB)
moy_annuelle, moy_sha     = recuperer_moy_annuelle(NOM_ETAB)

if hist_brut_df is not None:
    mois_connus = sorted(hist_brut_df["Mois"].unique(), key=month_key)
    st.info(f"📚 Historique **{NOM_ETAB}** — **{len(mois_connus)} mois** : {' · '.join(mois_connus)}")
else:
    st.info(f"📭 Aucun historique pour **{NOM_ETAB}** — premier chargement.")

if moy_annuelle is not None:
    st.info("📊 Moyenne année précédente chargée depuis GitHub.")

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION OPTIONNELLE — MOYENNES ANNÉE PRÉCÉDENTE
# ══════════════════════════════════════════════════════════════════════════════

with st.expander("📅 Charger les données de l'année précédente (facultatif)", expanded=moy_annuelle is None):
    st.caption(
        "Uploadez le ZIP contenant tous les dossiers mois de l'année passée. "
        "À faire une seule fois par établissement. "
        "Pas besoin du fichier CSV VisualValo."
    )
    uploaded_zip_annee = st.file_uploader(
        "📁 ZIP année précédente (tous les mois)",
        type=["zip"],
        key="zip_annee",
    )
    if uploaded_zip_annee is not None:
        if st.button("⚙️ Calculer et sauvegarder les moyennes"):
            with st.spinner("Calcul des moyennes…"):
                try:
                    nouvelles_moy = core.load_annee_precedente(io.BytesIO(uploaded_zip_annee.read()))
                    _, sha_actuel = github_lire_moy(NOM_ETAB)
                    github_ecrire_moy(nouvelles_moy, sha_actuel, NOM_ETAB)
                    moy_annuelle = nouvelles_moy
                    recuperer_moy_annuelle.clear()
                    st.success(
                        f"✅ Moyenne sauvegardée : "
                        f"Recette brute par séjour (2025) ={nouvelles_moy['recette_BR_moy_sej']:,.0f} € · "
                    )
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
    elif moy_annuelle is not None:
        st.success(
            f"✅ Moyenne déjà enregistrée : "
            f"recette_BR_moy_sej={moy_annuelle['recette_BR_moy_sej']:,.0f} € · "
        )

# ══════════════════════════════════════════════════════════════════════════════
#  UPLOADS MOIS COURANT
# ══════════════════════════════════════════════════════════════════════════════

st.subheader("📂 Données à intégrer")

uploaded_zip = st.file_uploader("📁 ZIP du nouveau mois à ajouter", type=["zip"])
uploaded_csv = st.file_uploader("📊 Fichier CSV VisualValoSejours", type=["csv"])

if not uploaded_zip or not uploaded_csv:
    st.warning("Veuillez uploader le fichier ZIP et le fichier CSV.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT + FUSION + RECALCUL
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="Chargement des nouvelles données…")
def charger_brut(zip_bytes, csv_bytes):
    return core.load_data_brut(io.BytesIO(zip_bytes), io.BytesIO(csv_bytes))


nouveau         = charger_brut(uploaded_zip.read(), uploaded_csv.read())
nouveau_brut_df = nouveau["brut_df"]

if hist_brut_df is not None:
    mois_nouveaux = set(nouveau_brut_df["Mois"].unique())
    mois_hist     = set(hist_brut_df["Mois"].unique())
    doublons      = mois_nouveaux & mois_hist
    if doublons:
        st.warning(f"⚠️ Mois ignorés (déjà présents) : {', '.join(sorted(doublons, key=month_key))}")
        nouveau_brut_df = nouveau_brut_df[~nouveau_brut_df["Mois"].isin(doublons)]
    brut_complet = (
        pd.concat([hist_brut_df, nouveau_brut_df], ignore_index=True)
        if not nouveau_brut_df.empty else hist_brut_df
    )
else:
    brut_complet = nouveau_brut_df

brut_complet = brut_complet.iloc[
    brut_complet["Mois"].map(month_key).argsort()
].reset_index(drop=True)

evol_df    = core.recalculer_derives(brut_complet)
mois_tries = sorted(evol_df["Mois"].unique(), key=month_key)
PERIODE    = f"{mois_tries[-1]}"

st.success(f"✅ Données prêtes — **{NOM_ETAB}** · {PERIODE}")
st.caption(f"Mois dans le rapport : {' · '.join(mois_tries)}")

# ══════════════════════════════════════════════════════════════════════════════
#  GRAPHES + COMMENTAIRES
# ══════════════════════════════════════════════════════════════════════════════

comments = {}
figures  = core.generate_all_figures(evol_df, moy_annuelle=moy_annuelle)

for theme, graphe_label, fig, plots in figures:
    st.subheader(f"{theme.strip()} — {graphe_label}")
    col1, col2 = st.columns([2, 1])
    with col1:
        st.pyplot(fig)
    with col2:
        for col, titre in plots:
            auto_comment = core.generate_comment(col, titre, evol_df)
            edited = st.text_area(titre, value=auto_comment, height=120, key=f"{theme}_{col}")
            comments[(theme, col)] = edited
    st.divider()

# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION PDF + SAUVEGARDE GITHUB
# ══════════════════════════════════════════════════════════════════════════════

st.subheader("📤 Export")

if st.button("📄 Générer le PDF et sauvegarder l'historique"):

    with st.spinner("Génération du PDF…"):
        pdf_bytes = core.generate_pdf(
            evol_df=evol_df,
            NOM_ETAB=NOM_ETAB,
            PERIODE=PERIODE,
            custom_comments=comments,
            moy_annuelle=moy_annuelle,
        )

    with st.spinner("Sauvegarde de l'historique sur GitHub…"):
        try:
            _, sha_actuel = github_lire_parquet(NOM_ETAB)
            github_ecrire_parquet(brut_complet, sha_actuel, NOM_ETAB, f"historique: {PERIODE} — {NOM_ETAB}")
            st.success(f"✅ Historique **{NOM_ETAB}** mis à jour sur GitHub.")
            recuperer_historique.clear()
        except Exception as e:
            st.error(f"❌ Erreur sauvegarde GitHub : {e}")

    st.download_button(
        label="⬇️ Télécharger le rapport PDF",
        data=pdf_bytes,
        file_name=f"rapport_mensuel_{NOM_ETAB}_{PERIODE.replace(' → ', '_')}.pdf",
        mime="application/pdf",
    )

# %%
