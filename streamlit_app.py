import io

import streamlit as st

from balance_core import (
    DISPLAY_COLUMNS,
    compare_balances,
    export_results_csv_bytes,
    export_results_excel_bytes,
    load_balance,
)


st.set_page_config(page_title="Comparateur de balances N / N-1", layout="wide")

st.title("Comparateur de balances N / N-1")
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
        balance_n = load_balance(io.BytesIO(balance_n_file.getvalue()), balance_n_file.name)
        balance_n1 = load_balance(io.BytesIO(balance_n1_file.getvalue()), balance_n1_file.name)
        results = compare_balances(balance_n, balance_n1)

        st.success(f"{len(results)} compte(s) compare(s) avec succes.")

        st.dataframe(results, use_container_width=True, hide_index=True)

        excel_bytes = export_results_excel_bytes(results)
        csv_bytes = export_results_csv_bytes(results)

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

        st.caption("Colonnes generees : " + " | ".join(DISPLAY_COLUMNS))
    except Exception as exc:
        st.error(f"Erreur lors du traitement : {exc}")
else:
    st.info("Ajoutez les deux balances pour lancer le calcul.")
