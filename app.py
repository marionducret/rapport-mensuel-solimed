import streamlit as st
import core

st.set_page_config(layout="wide")
st.title("Générateur de rapport")

# ── Uploads ──────────────────────────────────────────────────────────────────
uploaded_zip   = st.file_uploader("📦 Dossier compressé (.zip)", type=["zip"])
uploaded_excel = st.file_uploader("📊 Fichier Excel de valorisation (.xlsx)", type=["xlsx"])

if not uploaded_zip or not uploaded_excel:
    st.warning("Veuillez uploader le fichier zip ET le fichier Excel de valorisation.")
    st.stop()

# ── Chargement des données (mis en cache pour la session) ────────────────────
@st.cache_data(show_spinner="Chargement des données…")
def charger(zip_bytes, excel_bytes):
    import io
    return core.load_data(io.BytesIO(zip_bytes), io.BytesIO(excel_bytes))

result   = charger(uploaded_zip.read(), uploaded_excel.read())
evol_df  = result["evol_df"]
NOM_ETAB = result["NOM_ETAB"]
PERIODE  = result["PERIODE"]

st.success(f"✅ Données chargées — **{NOM_ETAB}** · {PERIODE}")

# ── Affichage des graphes + zones de commentaires ────────────────────────────
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
                f"{titre}",
                value=auto_comment,
                height=120,
                key=f"{theme}_{col}",
            )
            comments[(theme, col)] = edited

    st.divider()

# ── Génération du PDF ────────────────────────────────────────────────────────
if st.button("📄 Générer le PDF"):
    with st.spinner("Génération du PDF…"):
        pdf_bytes = core.generate_pdf(
            evol_df=evol_df,
            NOM_ETAB=NOM_ETAB,
            PERIODE=PERIODE,
            custom_comments=comments,
        )
    st.download_button(
        label="⬇️ Télécharger le rapport PDF",
        data=pdf_bytes,
        file_name="rapport_mensuel.pdf",
        mime="application/pdf",
    )
