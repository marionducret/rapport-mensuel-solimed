import streamlit as st
import core

st.set_page_config(layout="wide")

st.title("Générateur de rapport")

uploaded_zip = st.file_uploader("📦 Upload du dossier compressé (.zip)", type=["zip"])

if not uploaded_zip:
    st.warning("Veuillez uploader le fichier zip")
    st.stop()
    
comments = {}

figures = core.generate_all_figures(uploaded_files, uploaded_excel)

for theme, fig, plots in figures:
    st.header(theme)

    col1, col2 = st.columns([2, 1])  # ← ICI on crée les colonnes

    with col1:
        st.pyplot(fig)  # graphe à gauche

    with col2:
        for col, titre in plots:
            auto_comment = core.generate_comment(col, titre)

            edited = st.text_area(
                f"{titre}",
                value=auto_comment,
                height=120
            )

            comments[(theme, col)] = edited

    st.divider()

# bouton final
if st.button("Générer le PDF"):
    core.generate_pdf(custom_comments=comments)
    st.success("PDF généré !")