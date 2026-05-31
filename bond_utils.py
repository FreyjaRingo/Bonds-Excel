import re

import pandas as pd
from openpyxl.utils import get_column_letter


USD_PREFIXES = ("INDON", "INDOIS")
SYARIAH_PREFIXES = ("PBS", "SR", "ST", "INDOIS")


def normalize_column_name(col, idx):
    if pd.isna(col):
        return f"_unnamed_{idx}"

    name = str(col).replace("\n", "").replace("\r", "").strip()
    return name or f"_unnamed_{idx}"


def normalize_columns(columns):
    normalized = []
    seen = {}

    for idx, col in enumerate(columns, start=1):
        name = normalize_column_name(col, idx)
        seen[name] = seen.get(name, 0) + 1
        if seen[name] > 1:
            name = f"{name}_{seen[name]}"
        normalized.append(name)

    return normalized


def drop_repeated_header_rows(df):
    normalized_headers = [str(c).strip().lower() for c in df.columns]

    def looks_like_header(row):
        row_values = [
            str(v).replace("\n", "").replace("\r", "").strip().lower()
            for v in row.tolist()
        ]
        return row_values == normalized_headers

    return df[~df.apply(looks_like_header, axis=1)].reset_index(drop=True)


def clean_numeric(val):
    if pd.isna(val) or val == "":
        return val

    if not isinstance(val, str):
        try:
            return float(val)
        except (TypeError, ValueError):
            return val

    s = re.sub(r"\s+", "", val.strip())
    if s in ("", "-"):
        return val

    s = s.replace("%", "")

    if not all(c in "0123456789.,-" for c in s):
        return val

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        if s.count(",") > 1:
            s = s.replace(",", "")
        else:
            parts = s.split(",")
            if len(parts[-1]) == 3:
                s = s.replace(",", "")
            else:
                s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return val


def parse_date_series(series):
    text_values = series.astype("string").str.strip()
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    date_formats = (
        "%d-%b-%y",
        "%d-%b-%Y",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
    )
    for date_format in date_formats:
        missing = parsed.isna() & text_values.notna()
        if not missing.any():
            break
        parsed.loc[missing] = pd.to_datetime(
            text_values.loc[missing],
            format=date_format,
            errors="coerce",
        )

    return parsed


def parse_date_value(value):
    return parse_date_series(pd.Series([value])).iloc[0]


def to_percent_str(val):
    try:
        if pd.isna(val):
            return None
        return f"{round(float(val), 6)}%"
    except (TypeError, ValueError):
        return val


def parse_percent_rate(val):
    try:
        if pd.isna(val):
            return None
        return float(str(val).replace("%", "").replace(",", ".")) / 100.0
    except (TypeError, ValueError):
        return None


def parse_change_value(val):
    if pd.isna(val):
        return None

    text = str(val).strip()
    if text.upper() in ("", "-", "#N/A", "N/A", "NA"):
        return None

    text = text.replace("%", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def style_change_cell(val):
    change = parse_change_value(val)
    if change is None:
        return ""

    if change < 0:
        return "background-color: #fde2e2; color: #991b1b; font-weight: 700;"
    if change > 0:
        return "background-color: #dcfce7; color: #166534; font-weight: 700;"
    return "background-color: #ffedd5; color: #9a3412; font-weight: 700;"


def classify_currency(product_code):
    if not isinstance(product_code, str):
        return "IDR"

    code = product_code.strip().upper()
    return "USD" if code.startswith(USD_PREFIXES) else "IDR"


def currency_matches(series, currency):
    if currency in ("US", "USD"):
        return series.isin(["US", "USD"])
    return series == currency


def product_code_key(product_code):
    return str(product_code).strip().upper()


def product_options_for_currency(df, currency):
    if currency == "Semua" or "currency check" not in df.columns:
        source_df = df
    else:
        source_df = df[currency_matches(df["currency check"], currency)]

    return source_df["product_code"].dropna().unique().tolist()


def classify_bond_type(product_code):
    code = product_code_key(product_code)
    return "Syariah" if code.startswith(SYARIAH_PREFIXES) else "Konvensional"


def product_options_for_segment(df, currency, bond_type):
    products = product_options_for_currency(df, currency)
    return [product for product in products if classify_bond_type(product) == bond_type]


def auto_benchmark_products(df, benchmark_product_codes, currency, bond_type=None):
    return [
        product
        for product in product_options_for_currency(df, currency)
        if product_code_key(product) in benchmark_product_codes
        and (bond_type is None or classify_bond_type(product) == bond_type)
    ]


def excel_ref(df, column_name, row_number):
    col_idx = df.columns.get_loc(column_name) + 1
    return f"{get_column_letter(col_idx)}{row_number}"
