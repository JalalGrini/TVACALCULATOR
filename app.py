import streamlit as st
import pandas as pd
from logic import calculate_ht_tva
from export_excel import export_to_excel

st.set_page_config(page_title="TVA Calculator", layout="centered")
st.title("🧾 TVA Calculator & Excel Export")

# Input entreprise name
enterprise_name = st.text_input("Nom de l'entreprise")

# ✅ Add this: input for date
date_str = st.text_input("Date (MM/YYYY)", value="07/2025")

# Initialize session state
if 'entries' not in st.session_state:
    st.session_state['entries'] = []

# Input form
with st.form("add_form"):
    col1, col2 = st.columns(2)
    with col1:
        role = st.selectbox("Type", ["Client", "Fournisseur"])
        service = st.text_input("Nom du service")
    with col2:
        ttc = st.number_input("Montant TTC", min_value=0.0, step=0.01)
        tva_rate = st.number_input("Taux de TVA %", min_value=0.0, max_value=100.0, value=20.0)

    submitted = st.form_submit_button("Ajouter à la liste")
    if submitted and service and ttc > 0:
        ht, tva = calculate_ht_tva(ttc, tva_rate)
        st.session_state['entries'].append({
            "Role": role,
            "Service": service,
            "TTC": ttc,
            "HT": ht,
            "TVA Rate": tva_rate,
            "TVA": tva
        })

# Display table
if st.session_state['entries']:
    df = pd.DataFrame(st.session_state['entries'])
    st.dataframe(df, use_container_width=True)

    # ✅ Export to Excel using the new function with date
    if st.button("📤 Exporter vers Excel"):
        export_path = export_to_excel(df, enterprise_name, date_str)
        st.success(f"✅ Fichier Excel généré : {export_path}")

# Button to reset entries
if st.button("🔁 Réinitialiser"):
    st.session_state['entries'] = []
