import streamlit as st
import tempfile
import os
from datetime import datetime
from rdv_generator import load_rdv_data, filter_rdv, generate_rdv_report

st.set_page_config(page_title="Rapport RDV", layout="centered")
st.title("📅 Générateur de rapports de rendez-vous commerciaux")

uploaded_file = st.file_uploader("📁 Importer le fichier des RDV", type=["xlsx"])
uploaded_logo = st.file_uploader("🖼️ Logo (optionnel)", type=["png", "jpg", "jpeg"])

col1, col2, col3 = st.columns(3)
with col1:
    annee = st.number_input("📆 Année", min_value=2020, max_value=2100, value=datetime.now().year)
with col2:
    mois = st.number_input("🗓 Mois", min_value=1, max_value=12, value=datetime.now().month)
with col3:
    jour_range = st.slider("📍 Intervalle de jours", 1, 31, (1, 31))

if uploaded_file:
    if st.button("🚀 Générer les rapports"):
        with st.spinner("Traitement en cours..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                file_path = os.path.join(tmpdir, "rdv.xlsx")
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.read())

                logo_path = None
                if uploaded_logo:
                    logo_path = os.path.join(tmpdir, uploaded_logo.name)
                    with open(logo_path, "wb") as f:
                        f.write(uploaded_logo.read())

                df, col_com, col_reason, col_address = load_rdv_data(file_path)
                filtered = filter_rdv(df, annee, mois, jour_range[0], jour_range[1])
                grouped = filtered.groupby(col_com)

                output_dir = os.path.join(tmpdir, "outputs")
                fichiers = []

                for commercial, group in grouped:
                    path = generate_rdv_report(commercial, group, logo_path, datetime.now(), col_reason, col_address, output_dir)
                    fichiers.append(path)

                for f in fichiers:
                    with open(f, "rb") as file:
                        st.download_button(f"📥 Télécharger {os.path.basename(f)}", file, file_name=os.path.basename(f))

        st.success("✅ Rapports générés avec succès.")
