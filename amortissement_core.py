from __future__ import annotations

import calendar
import csv
import io
import math
from dataclasses import dataclass
from datetime import date, datetime


try:
    import pandas as pd  # type: ignore
except Exception:
    pd = None


BASE_DISPLAY_COLUMNS = (
    "REFERENCE",
    "DESIGNATION",
    "VALEUR ORIGINE",
    "DATE ACQUISITION",
    "DUREE (ANS)",
    "ANNUITE COMPLETE",
    "TAUX PRORATA",
    "ANNUITE",
    "AMORT. CUMULE",
    "VNC FIN",
    "STATUT",
)


EXAMPLE_ASSETS = [
    {
        "REFERENCE": "IMMO-001",
        "DESIGNATION": "Acquisition du 01/07/2025",
        "VALEUR ORIGINE": 1_000_000,
        "DATE ACQUISITION": "2025-07-01",
        "DUREE (ANS)": 5,
    },
    {
        "REFERENCE": "IMMO-002",
        "DESIGNATION": "Acquisition du 01/01/2025",
        "VALEUR ORIGINE": 1_000_000,
        "DATE ACQUISITION": "2025-01-01",
        "DUREE (ANS)": 5,
    },
    {
        "REFERENCE": "IMMO-003",
        "DESIGNATION": "Acquisition du 01/07/2020",
        "VALEUR ORIGINE": 1_000_000,
        "DATE ACQUISITION": "2020-07-01",
        "DUREE (ANS)": 5,
    },
    {
        "REFERENCE": "IMMO-004",
        "DESIGNATION": "Acquisition du 10/03/2022",
        "VALEUR ORIGINE": 1_000_000,
        "DATE ACQUISITION": "2022-03-10",
        "DUREE (ANS)": 5,
    },
    {
        "REFERENCE": "IMMO-005",
        "DESIGNATION": "Acquisition du 08/05/2018",
        "VALEUR ORIGINE": 1_000_000,
        "DATE ACQUISITION": "2018-05-08",
        "DUREE (ANS)": 5,
    },
]


@dataclass
class AssetRow:
    reference: str
    designation: str
    valeur_origine: float
    date_acquisition: date
    duree_ans: int


def load_assets_frame(source: io.BytesIO, filename: str):
    if pd is None:
        raise ValueError("L'import de fichier requiert pandas/openpyxl.")

    extension = filename.lower().rsplit(".", 1)[-1]
    if extension in {"xlsx", "xls"}:
        frame = pd.read_excel(source)
    else:
        text = source.getvalue().decode("utf-8-sig")
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=";,|\t")
        frame = pd.read_csv(io.StringIO(text), sep=dialect.delimiter)

    frame.columns = [str(column).strip().upper() for column in frame.columns]
    return frame


def calculate_amortissements_frame(frame, reference_year: int, prorata_mode: str = "monthly"):
    if pd is None:
        raise ValueError("Le calcul tabulaire requiert pandas.")

    assets = [_normalize_asset(row) for row in frame.to_dict(orient="records")]
    results = [
        calculate_asset_amortissement(asset, reference_year=reference_year, prorata_mode=prorata_mode)
        for asset in assets
    ]
    return pd.DataFrame(results, columns=get_display_columns(reference_year))


def calculate_asset_amortissement(
    asset: AssetRow, reference_year: int, prorata_mode: str = "monthly"
) -> dict[str, str]:
    period_start = date(reference_year, 1, 1)
    period_end_exclusive = date(reference_year + 1, 1, 1)
    depreciation_end = _add_years(asset.date_acquisition, asset.duree_ans)

    annual_amount = asset.valeur_origine / asset.duree_ans
    overlap_start = max(asset.date_acquisition, period_start)
    overlap_end = min(depreciation_end, period_end_exclusive)

    if overlap_start >= overlap_end:
        prorata_ratio = 0.0
        annuite_year = 0.0
        statut = "Aucune annuite"
    else:
        prorata_ratio = _prorata_ratio(overlap_start, overlap_end, reference_year, prorata_mode)
        annuite_year = annual_amount * prorata_ratio
        statut = "Annuite complete" if math.isclose(prorata_ratio, 1.0, abs_tol=1e-9) else "Annuite incomplete"

    amort_cumule = min(
        asset.valeur_origine,
        _accumulated_amount(asset, reference_year=reference_year, prorata_mode=prorata_mode),
    )
    vnc = max(0.0, asset.valeur_origine - amort_cumule)

    return {
        "REFERENCE": asset.reference,
        "DESIGNATION": asset.designation,
        "VALEUR ORIGINE": format_amount(asset.valeur_origine),
        "DATE ACQUISITION": asset.date_acquisition.strftime("%d/%m/%Y"),
        "DUREE (ANS)": str(asset.duree_ans),
        "ANNUITE COMPLETE": format_amount(annual_amount),
        "TAUX PRORATA": format_percent(prorata_ratio * 100),
        f"ANNUITE {reference_year}": format_amount(annuite_year),
        "AMORT. CUMULE": format_amount(amort_cumule),
        f"VNC FIN {reference_year}": format_amount(vnc),
        f"STATUT {reference_year}": statut,
    }


def export_results_csv_bytes(results_frame) -> bytes:
    if pd is None:
        raise ValueError("L'export CSV requiert pandas.")
    return results_frame.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")


def export_results_excel_bytes(results_frame) -> bytes:
    if pd is None:
        raise ValueError("L'export Excel requiert pandas/openpyxl.")
    buffer = io.BytesIO()
    results_frame.to_excel(buffer, index=False)
    return buffer.getvalue()


def build_example_frame():
    if pd is None:
        raise ValueError("L'exemple tabulaire requiert pandas.")
    return pd.DataFrame(EXAMPLE_ASSETS)


def get_display_columns(reference_year: int) -> tuple[str, ...]:
    return (
        "REFERENCE",
        "DESIGNATION",
        "VALEUR ORIGINE",
        "DATE ACQUISITION",
        "DUREE (ANS)",
        "ANNUITE COMPLETE",
        "TAUX PRORATA",
        f"ANNUITE {reference_year}",
        "AMORT. CUMULE",
        f"VNC FIN {reference_year}",
        f"STATUT {reference_year}",
    )


def _normalize_asset(row: dict[str, object]) -> AssetRow:
    normalized = {str(key).strip().upper(): value for key, value in row.items()}

    reference = str(normalized.get("REFERENCE", "") or "").strip()
    designation = str(normalized.get("DESIGNATION", "") or "").strip()
    valeur_origine = _parse_amount(normalized.get("VALEUR ORIGINE"))
    date_acquisition = _parse_date(normalized.get("DATE ACQUISITION"))
    duree_ans = int(float(normalized.get("DUREE (ANS)", 0) or 0))

    if not designation:
        designation = reference or "Immobilisation"
    if not reference:
        reference = designation
    if valeur_origine <= 0:
        raise ValueError(f"Valeur d'origine invalide pour {designation}.")
    if duree_ans <= 0:
        raise ValueError(f"Duree d'amortissement invalide pour {designation}.")

    return AssetRow(
        reference=reference,
        designation=designation,
        valeur_origine=valeur_origine,
        date_acquisition=date_acquisition,
        duree_ans=duree_ans,
    )


def _parse_amount(value: object) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(" ", "")
    if not text:
        return 0.0

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    else:
        text = text.replace(",", ".")

    return float(text)


def _parse_date(value: object) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    for date_format in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, date_format).date()
        except ValueError:
            continue
    raise ValueError(f"Date d'acquisition invalide : {value}")


def _add_years(input_date: date, years: int) -> date:
    try:
        return input_date.replace(year=input_date.year + years)
    except ValueError:
        return input_date.replace(month=2, day=28, year=input_date.year + years)


def _prorata_ratio(
    overlap_start: date, overlap_end_exclusive: date, reference_year: int, prorata_mode: str
) -> float:
    if prorata_mode == "daily":
        total_days = 366 if calendar.isleap(reference_year) else 365
        covered_days = (overlap_end_exclusive - overlap_start).days
        return covered_days / total_days

    covered_months = _covered_months_in_year(overlap_start, overlap_end_exclusive)
    return covered_months / 12


def _accumulated_amount(asset: AssetRow, reference_year: int, prorata_mode: str) -> float:
    annual_amount = asset.valeur_origine / asset.duree_ans
    depreciation_end = _add_years(asset.date_acquisition, asset.duree_ans)
    total = 0.0

    for year in range(asset.date_acquisition.year, reference_year + 1):
        period_start = date(year, 1, 1)
        period_end_exclusive = date(year + 1, 1, 1)
        overlap_start = max(asset.date_acquisition, period_start)
        overlap_end = min(depreciation_end, period_end_exclusive)
        if overlap_start >= overlap_end:
            continue
        total += annual_amount * _prorata_ratio(overlap_start, overlap_end, year, prorata_mode)

    return total


def _covered_months_in_year(start_date: date, end_date_exclusive: date) -> int:
    if start_date >= end_date_exclusive:
        return 0

    count = 0
    cursor = date(start_date.year, start_date.month, 1)
    while cursor < end_date_exclusive:
        month_end_exclusive = _first_day_next_month(cursor)
        if month_end_exclusive > start_date and cursor < end_date_exclusive:
            count += 1
        cursor = month_end_exclusive
    return count


def _first_day_next_month(current: date) -> date:
    if current.month == 12:
        return date(current.year + 1, 1, 1)
    return date(current.year, current.month + 1, 1)


def format_amount(value: float) -> str:
    if math.isclose(value, 0.0, abs_tol=1e-12):
        value = 0.0
    formatted = f"{round(value, 2):,.2f}"
    return formatted.replace(",", " ").replace(".", ",")


def format_percent(value: float) -> str:
    if math.isclose(value, 0.0, abs_tol=1e-12):
        value = 0.0
    formatted = f"{value:,.2f} %"
    return formatted.replace(",", " ").replace(".", ",")
