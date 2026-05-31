import io

import pandas as pd

from bond_utils import excel_ref, parse_percent_rate


def clean_for_excel(val):
    parsed = parse_percent_rate(val)
    return parsed if parsed is not None else val


def build_excel_buffer(excel_df, settlement_date, cut_rate_input, yield_hike_input, yield_cut_input):
    excel_df = excel_df.copy()

    for col in ["kupon %", "y mbi beli", "y mbi jual"]:
        if col in excel_df.columns:
            excel_df[col] = excel_df[col].apply(clean_for_excel)

    if {"Maturity", "kupon %", "y mbi jual"}.issubset(excel_df.columns):
        excel_df["MDURATION"] = [
            (
                f'=IFERROR(MDURATION(\'_Parameters\'!$B$1,{excel_ref(excel_df, "Maturity", i + 2)},'
                f'{excel_ref(excel_df, "kupon %", i + 2)},{excel_ref(excel_df, "y mbi jual", i + 2)},2,1),"")'
            )
            for i in range(len(excel_df))
        ]

    if "y mbi jual" in excel_df.columns:
        excel_df["Rate Hike"] = [
            f'=-{excel_ref(excel_df, "y mbi jual", i + 2)}*{cut_rate_input / 100.0}'
            for i in range(len(excel_df))
        ]
        excel_df["Rate Cut"] = [
            f'={excel_ref(excel_df, "y mbi jual", i + 2)}*{cut_rate_input / 100.0}'
            for i in range(len(excel_df))
        ]

    if {"mbi_jual", "Rate Hike", "Rate Cut"}.issubset(excel_df.columns):
        excel_df["Rate Hike Price"] = [
            (
                f'={excel_ref(excel_df, "mbi_jual", i + 2)}+'
                f'({excel_ref(excel_df, "Rate Hike", i + 2)}*{excel_ref(excel_df, "mbi_jual", i + 2)})'
            )
            for i in range(len(excel_df))
        ]
        excel_df["Rate Cut Price"] = [
            (
                f'={excel_ref(excel_df, "mbi_jual", i + 2)}+'
                f'({excel_ref(excel_df, "Rate Cut", i + 2)}*{excel_ref(excel_df, "mbi_jual", i + 2)})'
            )
            for i in range(len(excel_df))
        ]

    if {"y mbi jual", "Total Year to Maturity", "kupon %"}.issubset(excel_df.columns):
        excel_df["Price if Yield Hike"] = [
            (
                f'=-PV(({excel_ref(excel_df, "y mbi jual", i + 2)}+{yield_hike_input / 100.0}),'
                f'{excel_ref(excel_df, "Total Year to Maturity", i + 2)},'
                f'(100*{excel_ref(excel_df, "kupon %", i + 2)}),100,0)'
            )
            for i in range(len(excel_df))
        ]
        excel_df["Price if Yield Cut"] = [
            (
                f'=-PV(({excel_ref(excel_df, "y mbi jual", i + 2)}-{yield_cut_input / 100.0}),'
                f'{excel_ref(excel_df, "Total Year to Maturity", i + 2)},'
                f'(100*{excel_ref(excel_df, "kupon %", i + 2)}),100,0)'
            )
            for i in range(len(excel_df))
        ]

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        excel_df.to_excel(writer, index=False, sheet_name="Sheet1")
        params_sheet = writer.book.create_sheet("_Parameters")
        params_sheet["A1"] = "settlement_date"
        params_sheet["B1"] = pd.to_datetime(settlement_date).date()
        params_sheet["B1"].number_format = "m/d/yyyy"
        params_sheet.sheet_state = "hidden"

    return buffer.getvalue()
