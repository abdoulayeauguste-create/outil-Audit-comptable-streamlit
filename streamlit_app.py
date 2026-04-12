import io
from pathlib import Path

import pandas as pd
import streamlit as st

from amortissement_core import (
    build_example_frame,
    calculate_amortissements_frame,
    export_results_csv_bytes as export_amortissements_csv_bytes,
    export_results_excel_bytes as export_amortissements_excel_bytes,
    load_assets_frame,
)
from balance_core import (
    DISPLAY_COLUMNS as BALANCE_DISPLAY_COLUMNS,
    compare_balances,
    export_results_csv_bytes as export_balance_csv_bytes,
    export_results_excel_bytes as export_balance_excel_bytes,
    load_balance,
)

st.set_page_config(page_title="Outils comptables Streamlit", layout="wide")


def render_balance_module() -> None:
    st.subheader("Comparateur de balances N / N-1")
    st.write(
        "Chargez les balances N et N-1 pour obtenir automatiquement la variation en valeur absolue et en pourcentage."
    )

    col1, col2 = st.columns(2)
    with col1:
        balance_n_file = st.file_uploader(
            "Balance N",
            type=["csv", "txt", "xlsx", "xls"],
            key="balance_n",
        )
    with col2:
        balance_n1_file = st.file_uploader(
            "Balance N-1",
            type=["csv", "txt", "xlsx", "xls"],
            key="balance_n1",
        )

    if balance_n_file and balance_n1_file:
        try:
            balance_n = load_balance(
                io.BytesIO(balance_n_file.getvalue()),
                balance_n_file.name,
            )
            balance_n1 = load_balance(
                io.BytesIO(balance_n1_file.getvalue()),
                balance_n1_file.name,
            )
            results = compare_balances(balance_n, balance_n1)

            st.success(f"{len(results)} compte(s) compare(s) avec succes.")
            st.dataframe(results, width="stretch", hide_index=True)

            excel_bytes = export_balance_excel_bytes(results)
            csv_bytes = export_balance_csv_bytes(results)

            dl1, dl2 = st.columns(2)
            with dl1:
                st.download_button(
                    "Telecharger le resultat Excel",
                    data=excel_bytes,
                    file_name="variations_balances.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with dl2:
                st.download_button(
                    "Telecharger le resultat CSV",
                    data=csv_bytes,
                    file_name="variations_balances.csv",
                    mime="text/csv",
                )

            st.caption("Colonnes generees : " + " | ".join(BALANCE_DISPLAY_COLUMNS))
        except Exception as exc:
            st.error(f"Erreur lors du traitement : {exc}")
    else:
        st.info("Ajoutez les deux balances pour lancer le calcul.")


def render_amortissement_module() -> None:
    st.subheader("Amortissements des immobilisations")
    st.write(
        "Saisissez une liste d'immobilisations, choisissez l'annee de reference, puis laissez l'application calculer l'annuite complete, le prorata, l'amortissement cumule et la VNC."
    )

    settings_col1, settings_col2 = st.columns([1, 1])
    with settings_col1:
        reference_year = st.number_input(
            "Annee de reference",
            min_value=2000,
            max_value=2100,
            value=2025,
            step=1,
        )
    with settings_col2:
        prorata_label = st.selectbox(
            "Mode de prorata",
            options=[
                "Mensuel (mois entame compte)",
                "Journalier (exact au jour)",
            ],
            index=0,
        )
    prorata_mode = "daily" if prorata_label.startswith("Journalier") else "monthly"

    uploaded_file = st.file_uploader(
        "Importer une liste d'immobilisations",
        type=["csv", "xlsx", "xls"],
        key="amortissements_file",
        help="Colonnes attendues : REFERENCE, DESIGNATION, VALEUR ORIGINE, DATE ACQUISITION, DUREE (ANS).",
    )

    sample_path = Path(__file__).with_name("sample_immobilisations.csv")
    sample_bytes = sample_path.read_bytes()
    st.download_button(
        "Telecharger le modele CSV d'immobilisations",
        data=sample_bytes,
        file_name="sample_immobilisations.csv",
        mime="text/csv",
    )

    action_col1, action_col2 = st.columns(2)
    with action_col1:
        if st.button("Charger l'exemple des 5 cas", width="stretch"):
            st.session_state["amortissements_editor"] = build_example_frame()
    with action_col2:
        if st.button("Vider la liste", width="stretch"):
            st.session_state["amortissements_editor"] = pd.DataFrame(
                [
                    {
                        "REFERENCE": "",
                        "DESIGNATION": "",
                        "VALEUR ORIGINE": 0.0,
                        "DATE ACQUISITION": pd.NaT,
                        "DUREE (ANS)": 5,
                    }
                ]
            )

    if uploaded_file is not None:
        try:
            st.session_state["amortissements_editor"] = load_assets_frame(
                io.BytesIO(uploaded_file.getvalue()),
                uploaded_file.name,
            )
        except Exception as exc:
            st.error(f"Import impossible : {exc}")

    if "amortissements_editor" not in st.session_state:
        st.session_state["amortissements_editor"] = build_example_frame()

    df = st.session_state["amortissements_editor"].copy()

    expected_columns = [
        "REFERENCE",
        "DESIGNATION",
        "VALEUR ORIGINE",
        "DATE ACQUISITION",
        "DUREE (ANS)",
    ]

    for col in expected_columns:
        if col not in df.columns:
            if col in ["REFERENCE", "DESIGNATION"]:
                df[col] = ""
            elif col == "DATE ACQUISITION":
                df[col] = pd.NaT
            elif col == "DUREE (ANS)":
                df[col] = 5
            else:
                df[col] = 0.0

    df = df[expected_columns].copy()

    df["REFERENCE"] = df["REFERENCE"].fillna("").astype(str)
    df["DESIGNATION"] = df["DESIGNATION"].fillna("").astype(str)
    df["VALEUR ORIGINE"] = pd.to_numeric(df["VALEUR ORIGINE"], errors="coerce")
    df["DUREE (ANS)"] = pd.to_numeric(df["DUREE (ANS)"], errors="coerce")
    df["DATE ACQUISITION"] = pd.to_datetime(
        df["DATE ACQUISITION"],
        errors="coerce",
        dayfirst=True,
    )

    editor_frame = st.data_editor(
        df,
        width="stretch",
        num_rows="dynamic",
        hide_index=True,
        key="amortissements_editor_widget",
        column_config={
            "REFERENCE": st.column_config.TextColumn("REFERENCE"),
            "DESIGNATION": st.column_config.TextColumn("DESIGNATION"),
            "VALEUR ORIGINE": st.column_config.NumberColumn(
                "VALEUR ORIGINE",
                min_value=0.0,
            ),
            "DATE ACQUISITION": st.column_config.DateColumn(
                "DATE ACQUISITION",
                help="Formats acceptes : YYYY-MM-DD ou JJ/MM/AAAA.",
                format="DD/MM/YYYY",
            ),
            "DUREE (ANS)": st.column_config.NumberColumn(
                "DUREE (ANS)",
                min_value=1,
                step=1,
            ),
        },
    )

    st.session_state["amortissements_editor"] = editor_frame

    try:
        results_frame = calculate_amortissements_frame(
            editor_frame,
            reference_year=int(reference_year),
            prorata_mode=prorata_mode,
        )

        total_annuite = pd.to_numeric(
            results_frame[f"ANNUITE {int(reference_year)}"]
            .astype(str)
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False),
            errors="coerce",
        ).fillna(0)

        metric_col1, metric_col2, metric_col3 = st.columns(3)
        with metric_col1:
            st.metric("Immobilisations traitees", len(results_frame))
        with metric_col2:
            st.metric("Annee analysee", int(reference_year))
        with metric_col3:
            st.metric(
                "Total annuite",
                f"{total_annuite.sum():,.2f}".replace(",", " ").replace(".", ","),
            )

        st.dataframe(results_frame, width="stretch", hide_index=True)

        excel_bytes = export_amortissements_excel_bytes(results_frame)
        csv_bytes = export_amortissements_csv_bytes(results_frame)

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "Telecharger le plan d'amortissement Excel",
                data=excel_bytes,
                file_name=f"amortissements_{int(reference_year)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with dl2:
            st.download_button(
                "Telecharger le plan d'amortissement CSV",
                data=csv_bytes,
                file_name=f"amortissements_{int(reference_year)}.csv",
                mime="text/csv",
            )

        st.caption(
            "Cas couverts automatiquement : annuite complete, annuite incomplete en debut de vie, annuite incomplete en fin de vie, annuite nulle si l'immobilisation est deja totalement amortie."
        )
    except Exception as exc:
        st.warning(f"Le calcul ne peut pas encore etre affiche : {exc}")


st.title("Outils comptables Streamlit")
st.write(
    "Cette application rassemble plusieurs modules web accessibles a distance depuis un navigateur : comparateur de balances et calculateur d'amortissements."
)

tab1, tab2 = st.tabs(["Amortissements", "Balances"])

with tab1:
    render_amortissement_module()

with tab2:
    render_balance_module()