
import streamlit as st
from datetime import datetime
import tempfile
import os
import shutil
from rdv_generator import load_rdv_data, creer_rapport_rdv, sanitize_filename, normalize

COMMERCIAUX_CIBLES = ['Sandra', 'Ophélie', 'Arthur', 'Grégoire', 'Tania']

st.set_page_config(page_title="📅 Rapports de RDV", layout="centered")
st.title("📅 Générateur de rapports de rendez-vous commerciaux")

uploaded_rdv = st.file_uploader("📁 Importer le fichier Excel des RDV", type=["xlsx"])
uploaded_logo = st.file_uploader("🖼️ Logo (facultatif)", type=["png", "jpg", "jpeg"])

col1, col2 = st.columns(2)
with col1:
    mois = st.selectbox("📅 Mois", list(range(1, 13)), index=datetime.now().month - 1)
with col2:
    annee = st.selectbox("📆 Année", list(range(2022, 2030)), index=3)

col3, col4 = st.columns(2)
with col3:
    jour_debut = st.number_input("📍 Jour de début", min_value=1, max_value=31, value=1)
with col4:
    jour_fin = st.number_input("📍 Jour de fin", min_value=1, max_value=31, value=31)

if uploaded_rdv:
    if st.button("📋 Charger les données"):
        with st.spinner("🔍 Lecture et filtrage des données..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                rdv_path = os.path.join(temp_dir, "rdv.xlsx")
                with open(rdv_path, "wb") as f:
                    f.write(uploaded_rdv.read())

                rdv_data = load_rdv_data(rdv_path, jour_debut, jour_fin, mois, annee)

                def est_commercial_cible(nom):
                    nom_norm = normalize(nom)
                    return any(normalize(target) in nom_norm for target in COMMERCIAUX_CIBLES)

                rdv_data_cible = {k: v for k, v in rdv_data.items() if est_commercial_cible(k)}

                if rdv_data_cible:
                    selected_commerciaux = st.multiselect(
                        "👤 Choisir les commerciaux à inclure",
                        options=list(rdv_data_cible.keys()),
                        default=list(rdv_data_cible.keys())
                    )

                    if selected_commerciaux and st.button("🚀 Générer les rapports RDV"):
                        with st.spinner("📄 Génération des rapports en cours..."):
                            logo_path = None
                            if uploaded_logo:
                                logo_path = os.path.join(temp_dir, uploaded_logo.name)
                                with open(logo_path, "wb") as f:
                                    f.write(uploaded_logo.read())

                            output_dir = os.path.join(temp_dir, "rapports")
                            os.makedirs(output_dir, exist_ok=True)

                            for commercial in selected_commerciaux:
                                df = rdv_data_cible[commercial]
                                creer_rapport_rdv(df, commercial, jour_debut, jour_fin, mois, annee, output_dir, logo_path)

                            zip_path = shutil.make_archive(os.path.join(temp_dir, "Rapports_RDV"), 'zip', output_dir)
                            st.success("✅ Rapports générés avec succès.")
                            st.download_button("📥 Télécharger le fichier ZIP", open(zip_path, "rb"), file_name="Rapports_RDV.zip")
                else:
                    st.warning("Aucun commercial cible trouvé dans les RDV.")
