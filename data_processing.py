import re

import pandas as pd

from bond_utils import (
    classify_currency,
    clean_numeric,
    parse_date_series,
    to_percent_str,
)


NUMERIC_CANDIDATES = [
    "kupon",
    "mbi_beli",
    "mbi_jual",
    "yield_mbi_beli",
    "yield_mbi_jual",
    "Inventory",
]
DATE_CANDIDATES = ["Maturity", "Settlement date"]


def find_maturity_column(df):
    return next(
        (col for col in df.columns if pd.notna(col) and "maturity" in str(col).lower()),
        None,
    )


def extract_year(date_val):
    if pd.isna(date_val):
        return ""
    if isinstance(date_val, pd.Timestamp):
        return str(date_val.year)

    match = re.search(r"((?:19|20)\d{2})", str(date_val))
    return match.group(1) if match else str(date_val)


def clean_numeric_columns(df):
    df = df.copy()
    for col in NUMERIC_CANDIDATES:
        if col in df.columns:
            df[col] = df[col].apply(clean_numeric)
    return df


def normalize_date_columns(df):
    df = df.copy()
    for col in DATE_CANDIDATES:
        if col in df.columns:
            df[col] = parse_date_series(df[col]).dt.strftime("%d-%m-%Y")
    return df


def add_display_columns(df):
    df = df.copy()

    if "product_code" in df.columns:
        df["currency check"] = df["product_code"].apply(classify_currency)

    maturity_col = find_maturity_column(df)
    if maturity_col:
        df["year maturity"] = df[maturity_col].apply(extract_year)

    if "kupon" in df.columns:
        df["kupon %"] = df["kupon"].apply(to_percent_str)
        df.drop(columns=["kupon", "Settlement date", "_unnamed_4"], inplace=True, errors="ignore")
    else:
        df["kupon %"] = None

    if "yield_mbi_beli" in df.columns:
        df["y mbi beli"] = df["yield_mbi_beli"].apply(to_percent_str)
        df.drop(columns=["yield_mbi_beli"], inplace=True, errors="ignore")

    if "yield_mbi_jual" in df.columns:
        df["y mbi jual"] = df["yield_mbi_jual"].apply(to_percent_str)
        df.drop(columns=["yield_mbi_jual"], inplace=True, errors="ignore")
    else:
        df["y mbi jual"] = None

    return df, maturity_col


def prepare_bond_dataframe(raw_df):
    df = clean_numeric_columns(raw_df)
    df = normalize_date_columns(df)
    return add_display_columns(df)
