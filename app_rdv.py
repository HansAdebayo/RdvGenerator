import streamlit as st
from datetime import datetime
import tempfile
import os
import shutil
from rdv_generator import load_rdv_data, filter_rdv, create_word_report

st.set_page_config(page_title="📅 RDV - Générateur de rapports", layout="centered")

st.title("📅 Générateur de rapports de rendez-vous")

uploaded_rdv = st.file_uploader("📁 Fichier des rendez-vous", type=["xlsx"])
uploaded_logo = st.file_uploader("🖼️ Logo (facultatif)", type=["png", "jpg", "jpeg"])

col1, col2, col3 = st.columns(3)
with col1:
    jour = st.number_input("📆 Jour", min_value=1, max_value=31, value=datetime.now().day)
with col2:
    mois = st.selectbox("📅 Mois", list(range(1, 13)), index=datetime.now().month - 1)
with col3:
    annee = st.selectbox("📆 Année", list(range(2022, 2030)), index=2)

if uploaded_rdv and st.button("🚀 Générer les rapports RDV"):
    with st.spinner("📄 Génération des rapports en cours..."):
        report_date = datetime(annee, mois, jour)
        with tempfile.TemporaryDirectory() as temp_dir:
            rdv_path = os.path.join(temp_dir, "rdv.xlsx")
            with open(rdv_path, "wb") as f:
                f.write(uploaded_rdv.read())

            logo_path = None
            if uploaded_logo:
                logo_path = os.path.join(temp_dir, uploaded_logo.name)
                with open(logo_path, "wb") as f:
                    f.write(uploaded_logo.read())

            df = load_rdv_data(rdv_path)
            filtered = filter_rdv(df, report_date, report_date + timedelta(days=15))
            grouped = filtered.groupby('Commercial')

            output_dir = os.path.join(temp_dir, "rapports")
            for commercial, group in grouped:
                create_word_report(commercial, group, logo_path, report_date, output_dir)

            zip_path = shutil.make_archive(os.path.join(temp_dir, "Rapports_RDV"), 'zip', output_dir)
            st.success("✅ Rapports générés avec succès.")
            st.download_button("📥 Télécharger le ZIP", open(zip_path, "rb"), file_name="Rapports_RDV.zip")
