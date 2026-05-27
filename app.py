
import streamlit as st
import core
import pandas as pd
import io
import json
import base64
import requests
import re
import unicodedata
from pathlib import Path
from datetime import datetime, timedelta

st.set_page_config(layout="wide")
st.title("Générateur de rapport mensuel SSR")

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

def slug_etab(texte):
    texte = texte.strip().lower()
    texte = unicodedata.normalize("NFKD", texte).encode("ascii", "ignore").decode()
    texte = re.sub(r"[^a-z0-9]+", "_", texte)
    return texte.strip("_")

def nom_fichier_rapport(nom_etab, periode_code):
    mois = periode_code.split("_")[-1]
    nom = unicodedata.normalize("NFKD", nom_etab).encode("ascii", "ignore").decode()
    nom = re.sub(r"[^A-Za-z0-9]+", "_", nom).strip("_")
    return f"rapport_mensuel_{nom}_{mois}.pdf"

def extraire_nom_etab(etab_id):
    """
    Format attendu : 690000000_LB Monchy
    Retour PDF : LB Monchy
    """
    if "_" in etab_id:
        return etab_id.split("_", 1)[1].strip()
    return etab_id.strip()

def github_lire_parquet(etab_id):
    slug = slug_etab(etab_id)
    r = requests.get(gh_url(f"data/historique_{slug}.parquet"), headers=GH_HEADERS)
    if r.status_code == 404:
        return None, None
    r.raise_for_status()
    meta = r.json()
    return pd.read_parquet(io.BytesIO(base64.b64decode(meta["content"]))), meta["sha"]

def github_ecrire_parquet(df, sha, etab_id, message):
    slug = slug_etab(etab_id)
    buf  = io.BytesIO()
    df.to_parquet(buf, index=False)
    payload = {"message": message, "content": base64.b64encode(buf.getvalue()).decode()}
    if sha:
        payload["sha"] = sha
    requests.put(gh_url(f"data/historique_{slug}.parquet"), headers=GH_HEADERS, json=payload).raise_for_status()

def github_lire_moy(etab_id):
    slug = slug_etab(etab_id)
    r = requests.get(gh_url(f"data/moy_annuelle_{slug}.json"), headers=GH_HEADERS)
    if r.status_code == 404:
        return None, None
    r.raise_for_status()
    meta = r.json()
    return json.loads(base64.b64decode(meta["content"])), meta["sha"]

def github_ecrire_moy(moy_dict, sha, etab_id):
    slug = slug_etab(etab_id)
    payload = {
        "message": f"moy_annuelle: {etab_id}",
        "content": base64.b64encode(json.dumps(moy_dict).encode()).decode(),
    }
    if sha:
        payload["sha"] = sha
    requests.put(gh_url(f"data/moy_annuelle_{slug}.json"), headers=GH_HEADERS, json=payload).raise_for_status()

def github_lire_objectifs(etab_id):
    slug = slug_etab(etab_id)
    r = requests.get(gh_url(f"data/objectifs_{slug}.json"), headers=GH_HEADERS)
    if r.status_code == 404:
        return None, None
    r.raise_for_status()
    meta = r.json()
    return json.loads(base64.b64decode(meta["content"])), meta["sha"]

def github_ecrire_objectifs(obj_dict, sha, etab_id):
    slug = slug_etab(etab_id)
    payload = {
        "message": f"objectifs: {etab_id}",
        "content": base64.b64encode(json.dumps(obj_dict).encode()).decode(),
    }
    if sha:
        payload["sha"] = sha
    requests.put(gh_url(f"data/objectifs_{slug}.json"), headers=GH_HEADERS, json=payload).raise_for_status()

def github_supprimer_parquet(etab_id):
    slug = slug_etab(etab_id)
    r = requests.get(gh_url(f"data/historique_{slug}.parquet"), headers=GH_HEADERS)
    if r.status_code == 404:
        return False
    r.raise_for_status()
    sha = r.json()["sha"]
    payload = {
        "message": f"reset historique: {etab_id}",
        "sha": sha,
    }
    requests.delete(
        gh_url(f"data/historique_{slug}.parquet"),
        headers=GH_HEADERS,
        json=payload,
    ).raise_for_status()
    return True

def github_lire_etablissements():
    r = requests.get(gh_url("data/etablissements.json"), headers=GH_HEADERS)
    if r.status_code == 404:
        return [], None
    r.raise_for_status()
    meta = r.json()
    return json.loads(base64.b64decode(meta["content"])), meta["sha"]


def github_ecrire_etablissements(etabs, sha=None):
    payload = {
        "message": "maj liste etablissements",
        "content": base64.b64encode(
            json.dumps(etabs, ensure_ascii=False, indent=2).encode()
        ).decode(),
    }
    if sha:
        payload["sha"] = sha

    requests.put(
        gh_url("data/etablissements.json"),
        headers=GH_HEADERS,
        json=payload
    ).raise_for_status()

def sauvegarder_historique_github(brut_df, etab_id, nom_etab, nom_etab_simple, periode):
    _, sha_actuel = github_lire_parquet(etab_id)
    github_ecrire_parquet(
        brut_df,
        sha_actuel,
        etab_id,
        f"historique: {periode} — {nom_etab}"
    )

    etabs, sha_etabs = github_lire_etablissements()

    nouvel_etab = {
        "etab_id": etab_id,
        "nom_etab": nom_etab_simple,
        "slug": slug_etab(etab_id),
    }

    if not any(e["slug"] == nouvel_etab["slug"] for e in etabs):
        etabs.append(nouvel_etab)
        github_ecrire_etablissements(etabs, sha_etabs)
        lister_etabs_github.clear()

    recuperer_historique.clear()

@st.cache_data(show_spinner="Récupération des établissements enregistrés…", ttl=60)
def lister_etabs_github():
    etabs, _ = github_lire_etablissements()
    return etabs

def month_key(m):
    try:
        annee, num = m.split("_M")
        return (int(annee), int(num))
    except Exception:
        return (9999, 9999)

def excel_date_to_str(x):
    date = datetime(1899, 12, 30) + timedelta(days=int(x))
    return date.strftime("%d/%m/%y")

CALENDRIER_PERIODES = {
    "M1": ("29/12/25", "01/02/26"),
    "M2": ("29/12/25", "01/03/26"),
    "M3": ("29/12/25", "29/03/26"),
    "M4": ("29/12/25", "26/04/26"),
    "M5": ("29/12/25", "31/05/26"),
    "M6": ("29/12/25", "28/06/26"),
    "M7": ("29/12/25", "26/07/26"),
    "M8": ("29/12/25", "30/08/26"),
    "M9": ("29/12/25", "27/09/26"),
    "M10": ("29/12/25", "25/10/26"),
    "M11": ("29/12/25", "29/11/26"),
    "M12": ("29/12/25", "03/01/27"),
}

def libelle_periode_pmsi(periode_code):
    periode_simple = periode_code.split("_")[-1]  # 2026_M2 -> M2
    debut, fin = CALENDRIER_PERIODES.get(periode_simple, ("", ""))
    return f"{periode_simple} : du {debut} au {fin}"
# ══════════════════════════════════════════════════════════════════════════════
#  NOM ÉTABLISSEMENT
# ══════════════════════════════════════════════════════════════════════════════

etabs_connus = lister_etabs_github()

if etabs_connus:
    mode_etab = st.radio(
        "🏥 Établissement",
        ["Établissement déjà enregistré", "Nouvel établissement"],
        horizontal=True
    )
else:
    mode_etab = "Nouvel établissement"
    st.info("Aucun établissement enregistré pour le moment. Saisissez le premier établissement.")

if mode_etab == "Établissement déjà enregistré":
    etab_selection = st.selectbox(
        "Sélectionner un établissement",
        options=etabs_connus,
        format_func=lambda x: x["etab_id"]
    )

    ETAB_ID = etab_selection["etab_id"]
    NOM_ETAB_SIMPLE = etab_selection["nom_etab"]

else:
    ETAB_ID = st.text_input(
        "Nouvel établissement",
        placeholder="Format : Numéro Finess_Nom établissement"
    )

    if not ETAB_ID:
        st.warning("Veuillez saisir ou sélectionner un établissement.")
        st.stop()

    NOM_ETAB_SIMPLE = extraire_nom_etab(ETAB_ID)

NOM_ETAB_LAYOUT = f"Centre Médical de \n{NOM_ETAB_SIMPLE.upper()}"
NOM_ETAB = f"Centre Médical de {NOM_ETAB_SIMPLE}"

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

@st.cache_data(show_spinner="Récupération des objectifs…", ttl=60)
def recuperer_objectifs(nom_etab):
    try:
        return github_lire_objectifs(nom_etab)
    except Exception as e:
        return None, None

def message_moy_annuelle(moy_dict, prefixe="✅ Moyenne sauvegardée", annee="2025"):
    fragments = []
    suffixe_annee = f" ({annee})" if annee else ""
    if moy_dict.get("recette_BR_moy_jour") is not None:
        fragments.append(
            f"BR/jour HC{suffixe_annee} = {moy_dict['recette_BR_moy_jour']:,.0f} €"
        )
    if moy_dict.get("recette_BR_moy_jour_HTP") is not None:
        fragments.append(
            f"BR/jour HTP{suffixe_annee} = {moy_dict['recette_BR_moy_jour_HTP']:,.0f} €"
        )
    return f"{prefixe} : {' · '.join(fragments)}" if fragments else None

hist_brut_df, hist_sha    = recuperer_historique(ETAB_ID)
moy_annuelle, moy_sha     = recuperer_moy_annuelle(ETAB_ID)
objectifs, obj_sha        = recuperer_objectifs(ETAB_ID)

# Inconnu si pas d'historique : on affiche le champ HTP par défaut (sera ignoré
# au rendu KPI si pas d'activité HTP dans le mois courant).
hist_a_htp = True
if hist_brut_df is not None:
    hist_a_htp = bool(
        hist_brut_df[["effectif_transmis_HTP", "effectif_valorise_HTP"]]
        .fillna(0)
        .to_numpy()
        .sum()
        > 0
    )

if hist_brut_df is not None:
    mois_connus = sorted(hist_brut_df["Mois"].unique(), key=month_key)
    st.info(f"📚 Historique **{NOM_ETAB}** — **{len(mois_connus)} mois** : {' · '.join(mois_connus)}")
else:
    st.info(f"📭 Aucun historique pour **{NOM_ETAB}** — premier chargement.")

if moy_annuelle is not None:
    st.info("📊 Moyenne année précédente chargée depuis GitHub.")

if objectifs is not None:
    fragments_obj = []
    if objectifs.get("obj_BR_mois_HC"):
        fragments_obj.append(f"HC = {objectifs['obj_BR_mois_HC']:,.0f} €")
    if objectifs.get("obj_BR_mois_HTP"):
        fragments_obj.append(f"HTP = {objectifs['obj_BR_mois_HTP']:,.0f} €")
    if fragments_obj:
        st.info(f"🎯 Objectifs BR mensuels chargés : {' · '.join(fragments_obj)}")

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION OPTIONNELLE — MOYENNES ANNÉE PRÉCÉDENTE
# ══════════════════════════════════════════════════════════════════════════════

with st.expander("📅 Charger les données de l'année précédente (facultatif)", expanded=moy_annuelle is None):
    st.caption(
        "Uploadez le ZIP contenant tous les dossiers mois de l'année passée. "
        "À faire une seule fois par établissement. "
        "Le CSV VisualValo M12 est nécessaire pour calculer les jours valorisés HC."
    )
    uploaded_zip_annee = st.file_uploader(
        "📁 ZIP année précédente (M12 2025)",
        type=["zip"],
        key="zip_annee",
    )
    sans_valo_annee = st.checkbox("Je n'ai pas le VisualValo M12 année précédente")
    uploaded_csv_annee = None
    if not sans_valo_annee:
        uploaded_csv_annee = st.file_uploader(
            "📊 VisualValoSejours M12 année précédente",
            type=["csv"],
            key="csv_annee",
        )

    if uploaded_zip_annee is not None and (sans_valo_annee or uploaded_csv_annee is not None):
        if st.button("⚙️ Calculer et sauvegarder la moyenne"):
            with st.spinner("Calcul de la moyenne…"):
                try:
                    csv_annee_bytes = (
                        None
                        if sans_valo_annee
                        else io.BytesIO(uploaded_csv_annee.read())
                    )
                    nouvelles_moy = core.load_annee_precedente(
                        io.BytesIO(uploaded_zip_annee.read()),
                        csv_annee_bytes)
                    if not nouvelles_moy:
                        st.info("ℹ️ Aucune moyenne à sauvegarder : VisualValo absent et aucune activité HTP valorisée détectée.")
                    else:
                        _, sha_actuel = github_lire_moy(ETAB_ID)
                        github_ecrire_moy(nouvelles_moy, sha_actuel, ETAB_ID)
                        moy_annuelle = nouvelles_moy
                        recuperer_moy_annuelle.clear()
                        st.success(message_moy_annuelle(nouvelles_moy))
                except Exception as e:
                    st.error(f"❌ Erreur : {e}")
    elif moy_annuelle is not None:
        msg_moy = message_moy_annuelle(
            moy_annuelle,
            prefixe="✅ Moyenne déjà enregistrée",
            annee=None
        )
        if msg_moy:
            st.success(msg_moy)

# ══════════════════════════════════════════════════════════════════════════════
#  SECTION OPTIONNELLE — OBJECTIFS BR MENSUELS
# ══════════════════════════════════════════════════════════════════════════════

with st.expander("🎯 Objectifs BR mensuels (facultatif)", expanded=objectifs is None):
    st.caption(
        "Objectif de recette BR pour le mois supplémentaire, en €. "
        "Les badges « ✓ Objectif atteint » s'afficheront sur les KPI correspondants. "
        "Laisser à 0 pour ne pas afficher de badge."
    )

    obj_hc_init = float(objectifs.get("obj_BR_mois_HC", 0)) if objectifs else 0.0
    obj_htp_init = float(objectifs.get("obj_BR_mois_HTP", 0)) if objectifs else 0.0

    if hist_a_htp:
        col_obj_hc, col_obj_htp = st.columns(2)
        with col_obj_hc:
            obj_hc_input = st.number_input(
                "Objectif BR mensuel HC (€)",
                min_value=0.0,
                value=obj_hc_init,
                step=1000.0,
                format="%.0f",
                key="obj_BR_mois_HC_input",
            )
        with col_obj_htp:
            obj_htp_input = st.number_input(
                "Objectif BR mensuel HTP (€)",
                min_value=0.0,
                value=obj_htp_init,
                step=1000.0,
                format="%.0f",
                key="obj_BR_mois_HTP_input",
            )
    else:
        obj_hc_input = st.number_input(
            "Objectif BR mensuel HC (€)",
            min_value=0.0,
            value=obj_hc_init,
            step=1000.0,
            format="%.0f",
            key="obj_BR_mois_HC_input",
        )
        # Conserve la valeur existante côté HTP plutôt que de l'écraser à 0
        # si jamais un HTP avait été saisi par le passé.
        obj_htp_input = obj_htp_init

    if st.button("💾 Sauvegarder les objectifs"):
        nouveaux_obj = {
            "obj_BR_mois_HC": float(obj_hc_input),
            "obj_BR_mois_HTP": float(obj_htp_input),
        }
        try:
            _, sha_actuel = github_lire_objectifs(ETAB_ID)
            github_ecrire_objectifs(nouveaux_obj, sha_actuel, ETAB_ID)
            objectifs = nouveaux_obj
            recuperer_objectifs.clear()
            st.success(
                f"✅ Objectifs sauvegardés : HC = {nouveaux_obj['obj_BR_mois_HC']:,.0f} € · "
                f"HTP = {nouveaux_obj['obj_BR_mois_HTP']:,.0f} €"
            )
        except Exception as e:
            st.error(f"❌ Erreur sauvegarde objectifs : {e}")
    else:
        # Permet de prendre en compte les modifs dans la session courante
        # même sans sauvegarder sur GitHub.
        objectifs = {
            "obj_BR_mois_HC": float(obj_hc_input),
            "obj_BR_mois_HTP": float(obj_htp_input),
        }

# ══════════════════════════════════════════════════════════════════════════════
#  ZONE À RISQUE — RÉINITIALISATION HISTORIQUE
# ══════════════════════════════════════════════════════════════════════════════

if hist_brut_df is not None:
    with st.expander("⚠️ Réinitialiser l'historique (zone à risque)", expanded=False):
        st.warning(
            f"Cette action supprime **définitivement** le fichier d'historique de "
            f"**{NOM_ETAB}** sur GitHub. Tu devras réuploader tous les mois ensuite. "
            f"Les objectifs et la moyenne année précédente sont conservés."
        )
        confirm_reset = st.checkbox(
            "Je comprends que cette action est irréversible",
            key="confirm_reset_hist",
        )
        if st.button("🗑️ Réinitialiser l'historique", disabled=not confirm_reset):
            try:
                supprime = github_supprimer_parquet(ETAB_ID)
                recuperer_historique.clear()
                if supprime:
                    st.success(
                        f"✅ Historique de **{NOM_ETAB}** supprimé. "
                        "Recharge la page pour repartir à zéro."
                    )
                else:
                    st.info("ℹ️ Aucun historique à supprimer.")
                st.stop()
            except Exception as e:
                st.error(f"❌ Erreur lors de la suppression : {e}")

# ══════════════════════════════════════════════════════════════════════════════
#  UPLOADS MOIS COURANT
# ══════════════════════════════════════════════════════════════════════════════

st.subheader("📂 Données à intégrer")

uploaded_zip = st.file_uploader("📁 ZIP du nouveau mois à ajouter", type=["zip"])
sans_valo_periode = st.checkbox("Je n'ai pas le VisualValo de cette période")
uploaded_csv = None
if not sans_valo_periode:
    uploaded_csv = st.file_uploader("📊 Fichier CSV VisualValoSejours", type=["csv"])

if not uploaded_zip:
    st.warning("Veuillez uploader le fichier ZIP.")
    st.stop()

if not sans_valo_periode and not uploaded_csv:
    st.warning("Veuillez uploader le fichier CSV VisualValoSejours ou cocher l'option sans VisualValo.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  CHARGEMENT + FUSION + RECALCUL
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner="Chargement des nouvelles données…")
def charger_brut(zip_bytes, csv_bytes):
    csv_file = io.BytesIO(csv_bytes) if csv_bytes is not None else None
    return core.load_data_brut(io.BytesIO(zip_bytes), csv_file)

try:
    csv_bytes = None if sans_valo_periode else uploaded_csv.read()
    nouveau = charger_brut(uploaded_zip.read(), csv_bytes)
except Exception as e:
    st.error(str(e))
    st.stop()

nouveau_brut_df = nouveau["brut_df"]

dernier_mois_injecte = nouveau_brut_df["Mois"].iloc[0]
PERIODE = libelle_periode_pmsi(dernier_mois_injecte)

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

#détecter HTP
inclure_htp = (
    brut_complet[
        ["effectif_transmis_HTP", "effectif_valorise_HTP"]
    ]
    .fillna(0)
    .sum()
    .sum()
    > 0
)

if inclure_htp:
    st.info("✅ Activité HTP détectée : le rapport inclura les pages HC et HTP.")
else:
    st.info("ℹ️ Aucune activité HTP détectée : le rapport sera généré en HC uniquement.")

if sans_valo_periode:
    st.warning(
        "VisualValo absent : les indicateurs utilisant les jours valorisés HC "
        "seront indiqués comme non disponibles pour cette période."
    )

evol_df    = core.recalculer_derives(brut_complet)
mois_tries = sorted(evol_df["Mois"].unique(), key=month_key)

st.success(f"✅ Données prêtes — **{NOM_ETAB}** · {PERIODE}")
st.caption(f"Périodes dans le rapport : {' · '.join(mois_tries)}")

# ══════════════════════════════════════════════════════════════════════════════
#  SAUVEGARDE RAPIDE GITHUB
# ══════════════════════════════════════════════════════════════════════════════

st.subheader("📤 Sauvegarde")

if st.button("💾 Sauvegarder uniquement l'historique"):
    historique_sauvegarde = False
    with st.spinner("Sauvegarde de l'historique sur GitHub…"):
        try:
            sauvegarder_historique_github(
                brut_complet,
                ETAB_ID,
                NOM_ETAB,
                NOM_ETAB_SIMPLE,
                PERIODE
            )
            st.success(f"✅ Historique **{NOM_ETAB}** mis à jour sur GitHub.")
            historique_sauvegarde = True
        except Exception as e:
            st.error(f"❌ Erreur sauvegarde GitHub : {e}")
    if historique_sauvegarde:
        st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  GRAPHES + COMMENTAIRES
# ══════════════════════════════════════════════════════════════════════════════

comments = {}
figures  = core.generate_all_figures(evol_df, moy_annuelle=moy_annuelle, inclure_htp=inclure_htp)

for theme, graphe_label, fig, plots in figures:
    st.subheader(f"{theme.strip()} — {graphe_label}")
    col1, col2 = st.columns([2, 1])
    with col1:
        st.pyplot(fig)
    with col2:
        for col, titre in plots:
            auto_comment = core.generate_comment(col, titre, evol_df, moy_annuelle=moy_annuelle)
            edited = st.text_area(titre, value=auto_comment, height=120, key=f"{theme}_{col}")
            comments[(theme, col)] = edited
    st.divider()

# ══════════════════════════════════════════════════════════════════════════════
#  GÉNÉRATION PDF + SAUVEGARDE GITHUB
# ══════════════════════════════════════════════════════════════════════════════

st.subheader("📄 Export PDF")

if st.button("📄 Générer le PDF et sauvegarder l'historique"):

    with st.spinner("Génération du PDF…"):
        pdf_bytes = core.generate_pdf(
            evol_df=evol_df,
            NOM_ETAB=NOM_ETAB,
            NOM_ETAB_LAYOUT=NOM_ETAB_LAYOUT,
            PERIODE=PERIODE,
            custom_comments=comments,
            moy_annuelle=moy_annuelle,
            inclure_htp=inclure_htp,
            objectifs=objectifs,
        )

    with st.spinner("Sauvegarde de l'historique sur GitHub…"):
        try:
            sauvegarder_historique_github(
                brut_complet,
                ETAB_ID,
                NOM_ETAB,
                NOM_ETAB_SIMPLE,
                PERIODE
            )
            st.success(f"✅ Historique **{NOM_ETAB}** mis à jour sur GitHub.")

        except Exception as e:
            st.error(f"❌ Erreur sauvegarde GitHub : {e}")

    st.download_button(
        label="⬇️ Télécharger le rapport PDF",
        data=pdf_bytes,
        file_name=nom_fichier_rapport(NOM_ETAB_SIMPLE, dernier_mois_injecte),
        mime="application/pdf",
    )
