import csv
import io
import math
import unicodedata
from dataclasses import dataclass
from pathlib import Path


DISPLAY_COLUMNS = (
    "COMPTE",
    "LIBELLE",
    "SOLDE N",
    "SOLDE N-1",
    "VARIATION (ABS)",
    "VARIATION (%)",
)


try:
    import pandas as pd  # type: ignore
except Exception:
    pd = None


@dataclass
class BalanceRow:
    compte: str
    libelle: str
    solde: float


def load_balance(source: Path | io.BytesIO, filename: str | None = None) -> dict[str, BalanceRow]:
    source_name = filename or (source.name if hasattr(source, "name") else "fichier")
    extension = Path(source_name).suffix.lower()

    if extension in {".xlsx", ".xls"}:
        rows = _load_excel_rows(source)
    else:
        rows = _load_csv_rows(source)

    normalized_rows = [_normalize_row(row) for row in rows]
    data = {row.compte: row for row in normalized_rows if row.compte}
    if not data:
        raise ValueError(
            f"Aucune ligne exploitable n'a ete trouvee dans {source_name}. "
            "Le fichier doit contenir au minimum COMPTE, LIBELLE et SOLDE."
        )
    return data


def compare_balances(
    balance_n: dict[str, BalanceRow], balance_n1: dict[str, BalanceRow]
) -> list[dict[str, str]]:
    comptes = sorted(set(balance_n) | set(balance_n1))
    results: list[dict[str, str]] = []

    for compte in comptes:
        row_n = balance_n.get(compte)
        row_n1 = balance_n1.get(compte)
        libelle = (row_n.libelle if row_n and row_n.libelle else None) or (
            row_n1.libelle if row_n1 else ""
        )
        solde_n = row_n.solde if row_n else 0.0
        solde_n1 = row_n1.solde if row_n1 else 0.0
        variation_abs = solde_n - solde_n1

        if math.isclose(solde_n1, 0.0, abs_tol=1e-12):
            variation_pct = "" if math.isclose(solde_n, 0.0, abs_tol=1e-12) else "N/A"
        else:
            variation_pct = format_percent((variation_abs / solde_n1) * 100)

        results.append(
            {
                "COMPTE": compte,
                "LIBELLE": libelle or "",
                "SOLDE N": format_amount(solde_n),
                "SOLDE N-1": format_amount(solde_n1),
                "VARIATION (ABS)": format_amount(variation_abs),
                "VARIATION (%)": variation_pct,
            }
        )

    return results


def export_results_csv_bytes(results: list[dict[str, str]]) -> bytes:
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=DISPLAY_COLUMNS, delimiter=";")
    writer.writeheader()
    writer.writerows(results)
    return buffer.getvalue().encode("utf-8-sig")


def export_results_excel_bytes(results: list[dict[str, str]]) -> bytes:
    if pd is None:
        raise ValueError(
            "L'export Excel requiert pandas/openpyxl. "
            "Choisissez l'export CSV ou installez ces bibliotheques."
        )
    frame = pd.DataFrame(results, columns=DISPLAY_COLUMNS)
    buffer = io.BytesIO()
    frame.to_excel(buffer, index=False)
    return buffer.getvalue()


def export_results_csv_file(results: list[dict[str, str]], path: Path) -> None:
    path.write_bytes(export_results_csv_bytes(results))


def export_results_excel_file(results: list[dict[str, str]], path: Path) -> None:
    path.write_bytes(export_results_excel_bytes(results))


def _load_csv_rows(source: Path | io.BytesIO) -> list[dict[str, str]]:
    if isinstance(source, Path):
        with open(source, "r", encoding="utf-8-sig", newline="") as handle:
            return _read_csv_dicts(handle.read())

    content = source.getvalue().decode("utf-8-sig")
    return _read_csv_dicts(content)


def _read_csv_dicts(content: str) -> list[dict[str, str]]:
    sample = content[:4096]
    dialect = csv.Sniffer().sniff(sample, delimiters=";,|\t")
    reader = csv.DictReader(io.StringIO(content), dialect=dialect)
    return [dict(row) for row in reader if row]


def _load_excel_rows(source: Path | io.BytesIO) -> list[dict[str, str | float | int | None]]:
    if pd is None:
        raise ValueError(
            "Le support Excel requiert pandas/openpyxl. "
            "Convertissez le fichier en CSV ou installez ces bibliotheques."
        )
    frame = pd.read_excel(source)
    frame = frame.where(frame.notna(), None)
    return frame.to_dict(orient="records")


def _normalize_row(row: dict[str, str | float | int | None]) -> BalanceRow:
    lookup = {_normalize_header(str(key)): value for key, value in row.items() if key is not None}

    compte = _first_non_empty(
        lookup,
        "compte",
        "numero de compte",
        "numero compte",
        "n compte",
        "numcompte",
        "code compte",
    )
    libelle = _first_non_empty(
        lookup,
        "libelle",
        "intitule",
        "libelle compte",
        "compte libelle",
    )
    solde_raw = _first_non_empty(
        lookup,
        "solde",
        "solde final",
        "solde balance",
        "solde n",
        "solde debiteur",
        "solde crediteur",
        "balance",
        "montant",
        "net",
    )

    if not compte:
        raise ValueError("Colonne compte introuvable dans le fichier charge.")
    if solde_raw in (None, ""):
        raise ValueError(f"Solde introuvable pour le compte {compte}.")

    return BalanceRow(
        compte=str(compte).strip(),
        libelle=str(libelle or "").strip(),
        solde=_parse_amount(solde_raw),
    )


def _normalize_header(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value.strip().lower())
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    for source in ("_", "-", "\n"):
        normalized = normalized.replace(source, " ")
    return " ".join(normalized.split())


def _first_non_empty(mapping: dict[str, object], *keys: str) -> object | None:
    for key in keys:
        if key in mapping and mapping[key] not in (None, ""):
            return mapping[key]
    return None


def _parse_amount(value: str | float | int | None) -> float:
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


def format_amount(value: float) -> str:
    if math.isclose(value, 0.0, abs_tol=1e-12):
        value = 0.0
    formatted = f"{value:,.2f}"
    return formatted.replace(",", " ").replace(".", ",")


def format_percent(value: float) -> str:
    if math.isclose(value, 0.0, abs_tol=1e-12):
        value = 0.0
    formatted = f"{value:,.2f} %"
    return formatted.replace(",", " ").replace(".", ",")
