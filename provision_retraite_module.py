import streamlit as st
import pandas as pd
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta
from io import BytesIO

def parse_anciennete(value):
    if pd.isna(value):
        return 0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).lower()
    years = 0
    months = 0

    year_match = re.search(r'(\d+)\s*an', text)
    month_match = re.search(r'(\d+)\s*mois', text)

    if year_match:
        years = int(year_match.group(1))
    if month_match:
        months = int(month_match.group(1))

    return years + months / 12

def calculate_anciennete_from_date(date_entree, date_ref):
    if pd.isna(date_entree):
        return 0

    delta = relativedelta(date_ref, date_entree)
    return delta.years + delta.months / 12

def calculate_ir(sgmm, anc):
    if anc <= 5:
        T, K = 0.25, 0
    elif anc <= 10:
        T, K = 0.30, 0.25
    elif anc <= 20:
        T, K = 0.45, 1.75
    else:
        T, K = 0.50, 2.75

    return sgmm * (T * anc - K)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='RESULTAT')
    return output.getvalue()

def render_provision_retraite_app():
    st.header("📊 Calcul Provision Retraite")

    uploaded_file = st.file_uploader("Importer fichier Excel", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)

        df.columns = [col.upper().strip() for col in df.columns]

        date_ref = st.date_input("Date de clôture", value=datetime.today())

        if "DATE ENTREE" in df.columns:
            df["ANCIENNETE_CALC"] = df["DATE ENTREE"].apply(
                lambda x: calculate_anciennete_from_date(x, date_ref)
            )
        elif "ANCIENNETE" in df.columns:
            df["ANCIENNETE_CALC"] = df["ANCIENNETE"].apply(parse_anciennete)
        else:
            st.error("Aucune colonne DATE ENTREE ou ANCIENNETE trouvée")
            return

        if "SALAIRE MENSUEL" not in df.columns:
            if "SALAIRE ANNUEL" in df.columns:
                df["SALAIRE MENSUEL"] = df["SALAIRE ANNUEL"] / 12
            else:
                st.error("Aucun salaire trouvé")
                return

        df["PROVISION RECALCULEE"] = df.apply(
            lambda row: calculate_ir(row["SALAIRE MENSUEL"], row["ANCIENNETE_CALC"]),
            axis=1
        )

        if "PROVISION CLIENT" in df.columns:
            df["ECART"] = df["PROVISION RECALCULEE"] - df["PROVISION CLIENT"]
        else:
            df["ECART"] = None

        st.dataframe(df)

        excel_data = to_excel(df)

        st.download_button(
            label="📥 Télécharger le résultat",
            data=excel_data,
            file_name="provision_retraite.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )