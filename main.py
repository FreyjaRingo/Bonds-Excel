import pdfplumber
import pandas as pd
import streamlit as st

from calculations import add_simulation_columns
from data_processing import prepare_bond_dataframe
from excel_export import build_excel_buffer
from pdf_parsers import extract_pdf_dataframe
from ui_components import (
    render_copy_table,
    render_styled_table,
    render_yield_curve,
)


st.set_page_config(layout="wide", page_title="PDF Table Extractor & Editor")
st.title("PDF Table Extractor & Editor")


def render_simulation_inputs():
    st.write("---")
    st.subheader("Pengaturan Parameter Simulasi")

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        settlement_date = st.date_input("Tanggal Settlement ($Y$4)", value=pd.to_datetime("today"))
    with col2:
        cut_rate_input = st.number_input(
            "Cut / Hike Rate (%)",
            value=0.2500,
            step=0.0100,
            format="%.4f",
        )
    with col3:
        base_date_input = st.date_input("Tanggal Basis (Hari Ini)", value=pd.to_datetime("today"))
    with col4:
        yield_hike_input = st.number_input(
            "Yield Naik (%)",
            value=0.0,
            step=0.0100,
            format="%.4f",
        )
    with col5:
        yield_cut_input = st.number_input(
            "Yield Turun (%)",
            value=0.0,
            step=0.0100,
            format="%.4f",
        )

    return {
        "settlement_date": settlement_date,
        "cut_rate_input": cut_rate_input,
        "base_date_input": base_date_input,
        "yield_hike_input": yield_hike_input,
        "yield_cut_input": yield_cut_input,
    }


def render_download_button(edited_df, params):
    excel_data = build_excel_buffer(
        edited_df,
        params["settlement_date"],
        params["cut_rate_input"],
        params["yield_hike_input"],
        params["yield_cut_input"],
    )

    st.download_button(
        label="Unduh Data (Excel)",
        data=excel_data,
        file_name="hasil_edit.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def process_uploaded_pdf(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        raw_df, benchmark_product_codes = extract_pdf_dataframe(pdf)

    if raw_df.empty:
        st.warning("Tabel tidak terdeteksi pada PDF ini.")
        return

    df, maturity_col = prepare_bond_dataframe(raw_df)
    if "y mbi jual" in df.columns and df["y mbi jual"].isna().all():
        st.warning("Kolom yield_mbi_jual tidak ditemukan; simulasi yield dan grafik akan kosong.")

    params = render_simulation_inputs()
    df = add_simulation_columns(
        df,
        maturity_col,
        params["settlement_date"],
        params["cut_rate_input"],
        params["base_date_input"],
        params["yield_hike_input"],
        params["yield_cut_input"],
    )

    st.write("Edit data langsung pada tabel di bawah:")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    render_styled_table(edited_df)
    render_copy_table(edited_df)
    render_yield_curve(edited_df, benchmark_product_codes)
    render_download_button(edited_df, params)


uploaded_file = st.file_uploader("Unggah file PDF Excel", type="pdf")
if uploaded_file is not None:
    try:
        process_uploaded_pdf(uploaded_file)
    except Exception as exc:
        st.error(f"Gagal memproses file: {exc}")
