import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from dateutil.relativedelta import relativedelta


def parse_anciennete(value):
    if pd.isna(value):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).lower().strip().replace(",", ".")

    try:
        return float(text)
    except ValueError:
        pass

    years = 0.0
    months = 0.0

    year_match = re.search(r"(\d+(?:\.\d+)?)\s*an", text)
    month_match = re.search(r"(\d+(?:\.\d+)?)\s*mois", text)

    if year_match:
        years = float(year_match.group(1))
    if month_match:
        months = float(month_match.group(1))

    return years + (months / 12)


def calculate_anciennete_from_date(date_entree, date_ref):
    if pd.isna(date_entree):
        return 0.0

    entry_date = pd.to_datetime(date_entree, errors="coerce", dayfirst=True)
    if pd.isna(entry_date):
        return 0.0

    if isinstance(date_ref, datetime):
        ref_date = date_ref.date()
    else:
        ref_date = date_ref

    delta = relativedelta(ref_date, entry_date.date())
    return delta.years + (delta.months / 12) + (delta.days / 365.25)


def calculate_ir(sgmm, anc):
    if pd.isna(sgmm) or pd.isna(anc):
        return None

    sgmm = float(sgmm)
    anc = float(anc)

    if sgmm < 0 or anc <= 0:
        return 0.0

    if anc <= 5:
        t, k = 0.25, 0
    elif anc <= 10:
        t, k = 0.30, 0.25
    elif anc <= 20:
        t, k = 0.45, 1.75
    else:
        t, k = 0.50, 2.75

    ir = sgmm * (t * anc - k)
    return max(ir, 0.0)


def format_number(x):
    if pd.isna(x) or x is None:
        return ""
    try:
        return f"{round(float(x)):,}".replace(",", " ")
    except Exception:
        return ""


def find_column(df, possible_names):
    columns_map = {str(col).upper().strip(): col for col in df.columns}
    for name in possible_names:
        normalized_name = name.upper().strip()
        if normalized_name in columns_map:
            return columns_map[normalized_name]
    return None


def to_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="RESULTAT")

        worksheet = writer.sheets["RESULTAT"]

        for idx, column in enumerate(df.columns, start=1):
            column_letter = worksheet.cell(row=1, column=idx).column_letter
            max_length = max(
                len(str(column)),
                df[column].astype(str).map(len).max() if not df.empty else 0,
            )
            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 30)

    return output.getvalue()


def render_provision_retraite_app():
    st.header("💰 Calcul Provision Retraite")

    uploaded_file = st.file_uploader(
        "Importer le fichier Excel",
        type=["xlsx", "xls"],
        key="provision_retraite_file",
    )

    if not uploaded_file:
        st.info("Ajoutez un fichier Excel pour lancer le calcul.")
        return

    try:
        df = pd.read_excel(uploaded_file)
    except Exception as exc:
        st.error(f"Impossible de lire le fichier : {exc}")
        return

    if df.empty:
        st.error("Le fichier importé est vide.")
        return

    df.columns = [str(col).upper().strip() for col in df.columns]

    matricule_col = find_column(df, ["MATRICULE", "MATRICULES"])
    nom_col = find_column(df, ["NOM"])
    prenom_col = find_column(df, ["PRENOM", "PRENOMS"])
    date_entree_col = find_column(df, ["DATE ENTREE", "DATE D'ENTREE", "DATE_ENTREE"])
    anciennete_col = find_column(df, ["ANCIENNETE", "ANCIENNETÉ"])
    salaire_annuel_col = find_column(df, ["SALAIRE ANNUEL"])
    salaire_mensuel_col = find_column(df, ["SALAIRE MENSUEL"])
    provision_client_col = find_column(df, ["PROVISION CLIENT"])

    if not date_entree_col and not anciennete_col:
        st.error("Le fichier doit contenir soit DATE ENTREE, soit ANCIENNETE.")
        return

    if not salaire_annuel_col and not salaire_mensuel_col:
        st.error("Le fichier doit contenir soit SALAIRE ANNUEL, soit SALAIRE MENSUEL.")
        return

    date_ref = st.date_input("Date de clôture", value=datetime.today())

    if date_entree_col:
        df[date_entree_col] = pd.to_datetime(df[date_entree_col], errors="coerce", dayfirst=True)
        df["ANCIENNETE_CALC"] = df[date_entree_col].apply(
            lambda x: calculate_anciennete_from_date(x, date_ref)
        )
        anciennete_output_col = date_entree_col
    else:
        df["ANCIENNETE_CALC"] = df[anciennete_col].apply(parse_anciennete)
        anciennete_output_col = anciennete_col

    if salaire_mensuel_col:
        df[salaire_mensuel_col] = pd.to_numeric(df[salaire_mensuel_col], errors="coerce")
        salary_calc_col = salaire_mensuel_col
        salaire_output_col = salaire_mensuel_col
    else:
        df[salaire_annuel_col] = pd.to_numeric(df[salaire_annuel_col], errors="coerce")
        df["SALAIRE MENSUEL CALC"] = df[salaire_annuel_col] / 12
        salary_calc_col = "SALAIRE MENSUEL CALC"
        salaire_output_col = salaire_annuel_col

    if provision_client_col:
        df[provision_client_col] = pd.to_numeric(df[provision_client_col], errors="coerce")
    else:
        df["PROVISION CLIENT"] = None
        provision_client_col = "PROVISION CLIENT"

    df["PROVISION RECALCULEE"] = df.apply(
        lambda row: calculate_ir(row[salary_calc_col], row["ANCIENNETE_CALC"]),
        axis=1,
    )

    df["ECART"] = df["PROVISION RECALCULEE"] - df[provision_client_col]

    result_columns = []

    if matricule_col:
        result_columns.append(matricule_col)
    if nom_col:
        result_columns.append(nom_col)
    if prenom_col:
        result_columns.append(prenom_col)

    result_columns.append(anciennete_output_col)
    result_columns.append(salaire_output_col)
    result_columns.append(provision_client_col)
    result_columns.append("PROVISION RECALCULEE")
    result_columns.append("ECART")

    result_df = df[result_columns].copy()

    if date_entree_col and anciennete_output_col in result_df.columns:
        result_df[anciennete_output_col] = pd.to_datetime(
            result_df[anciennete_output_col],
            errors="coerce"
        ).dt.strftime("%d/%m/%Y")

    numeric_cols = [
        salaire_output_col,
        provision_client_col,
        "PROVISION RECALCULEE",
        "ECART",
    ]

    for col in numeric_cols:
        if col in result_df.columns:
            result_df[col] = result_df[col].apply(format_number)

    st.success("Calcul terminé avec succès.")
    st.dataframe(result_df, width="stretch", hide_index=True)

    excel_data = to_excel(result_df)

    st.download_button(
        label="📥 Télécharger le résultat Excel",
        data=excel_data,
        file_name="provision_retraite_resultat.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )