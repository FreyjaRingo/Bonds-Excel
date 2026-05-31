import re

import pandas as pd

from bond_utils import (
    drop_repeated_header_rows,
    normalize_columns,
    product_code_key,
)


BENCHMARK_FONT_FAMILY = "arialblack"
MAYBANK_COUPON_RE = re.compile(r"^-?\d+(?:\.\d+)?%$")
MAYBANK_DATE_RE = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{2,4}$")


def is_benchmark_font(fontname):
    font = str(fontname).replace("-", "").replace("_", "").lower()
    return BENCHMARK_FONT_FAMILY in font


def extract_bold_product_codes(pdf):
    bold_codes = set()

    for page in pdf.pages:
        try:
            words = page.extract_words(extra_attrs=["fontname"])
        except Exception:
            continue

        for word in words:
            text = str(word.get("text", "")).strip()
            fontname = str(word.get("fontname", ""))

            if not text or not is_benchmark_font(fontname):
                continue
            if text.lower() == "product_code" or word.get("top", 0) < 120:
                continue
            if word.get("x0", 999) > 100 or not re.search(r"[A-Za-z]", text):
                continue

            bold_codes.add(product_code_key(text))

    return bold_codes


def is_maybank_price_indication_pdf(pdf):
    sample_text = "\n".join((page.extract_text() or "") for page in pdf.pages[:2]).lower()
    has_price_header = "bond price indication" in sample_text
    has_table_header = (
        "prod_code" in sample_text
        and "mbi beli" in sample_text
        and "yield mbi jual" in sample_text
    )
    has_maybank_header = "maybank" in sample_text and has_table_header
    return (has_price_header and has_table_header) or has_maybank_header


def parse_maybank_price_line(line):
    tokens = line.split()
    if len(tokens) < 8:
        return None

    coupon_idx = next(
        (idx for idx, token in enumerate(tokens[1:], start=1) if MAYBANK_COUPON_RE.match(token)),
        None,
    )
    if coupon_idx is None:
        return None

    maturity_idx = coupon_idx + 1
    if maturity_idx >= len(tokens) or not MAYBANK_DATE_RE.match(tokens[maturity_idx]):
        return None

    value_start = maturity_idx + 1
    if value_start + 4 >= len(tokens):
        return None

    return {
        "product_code": tokens[0],
        "type": " ".join(tokens[1:coupon_idx]),
        "kupon": tokens[coupon_idx],
        "Maturity": tokens[maturity_idx],
        "mbi_beli": tokens[value_start],
        "yield_mbi_beli": tokens[value_start + 1],
        "mbi_jual": tokens[value_start + 2],
        "yield_mbi_jual": tokens[value_start + 3],
        "1D": tokens[value_start + 4],
    }


def parse_maybank_price_indication_pdf(pdf):
    rows = []
    benchmark_product_codes = set()
    current_section = ""

    for page in pdf.pages:
        for raw_line in (page.extract_text() or "").splitlines():
            line = " ".join(raw_line.split())
            if not line:
                continue

            lowered = line.lower()
            if lowered.startswith("benchmark ") or lowered.startswith("non benchmark "):
                current_section = line
                continue

            row = parse_maybank_price_line(line)
            if not row:
                continue

            rows.append(row)
            if current_section.lower().startswith("benchmark "):
                benchmark_product_codes.add(product_code_key(row["product_code"]))

    return pd.DataFrame(rows), benchmark_product_codes


def parse_table_pdf(pdf):
    all_rows = []
    benchmark_product_codes = extract_bold_product_codes(pdf)

    for page in pdf.pages:
        table = page.extract_table()
        if table:
            all_rows.extend(table)

    if not all_rows:
        return pd.DataFrame(), benchmark_product_codes

    df = pd.DataFrame(all_rows[1:], columns=all_rows[0])
    df.columns = normalize_columns(df.columns)
    df = drop_repeated_header_rows(df)
    return df, benchmark_product_codes


def extract_pdf_dataframe(pdf):
    if is_maybank_price_indication_pdf(pdf):
        return parse_maybank_price_indication_pdf(pdf)
    return parse_table_pdf(pdf)
