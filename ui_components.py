import html

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from bond_utils import (
    auto_benchmark_products,
    currency_matches,
    product_options_for_currency,
    product_options_for_segment,
    style_change_cell,
)


def render_styled_table(edited_df):
    st.write("Tabel dengan gradasi warna:")
    col1, col2 = st.columns([1, 3])
    with col1:
        price_threshold = st.number_input("Threshold Price4", value=100.0, step=1.0)
    with col2:
        usd_idr_filter = st.selectbox(
            "Filter Mata Uang untuk Bonds:",
            options=["Semua", "IDR", "USD"],
            index=0,
        )

    year_col = next((c for c in edited_df.columns if "year" in str(c).lower()), None)
    styled_df = edited_df.copy()

    if "currency check" in styled_df.columns:
        if usd_idr_filter == "USD":
            styled_df = styled_df[styled_df["currency check"] == "USD"]
        elif usd_idr_filter == "IDR":
            styled_df = styled_df[styled_df["currency check"] == "IDR"]

    if "mbi_jual" in styled_df.columns:
        styled_df["mbi_jual"] = pd.to_numeric(styled_df["mbi_jual"], errors="coerce")
        styled_df = styled_df[styled_df["mbi_jual"] <= price_threshold]

    styler = styled_df.style

    if year_col:
        styler = styler.background_gradient(
            subset=[year_col],
            cmap="Blues",
            gmap=pd.to_numeric(styled_df[year_col], errors="coerce"),
        )

    change_col = next(
        (
            c
            for c in styled_df.columns
            if str(c).strip().lower() in ("1d", "1d change", "1d_change")
        ),
        None,
    )
    if change_col:
        styler = styler.map(style_change_cell, subset=[change_col])

    st.dataframe(styler, use_container_width=True, hide_index=True)


def render_copy_table(edited_df):
    copy_cols = []
    if "product_code" in edited_df.columns and "Inventory" in edited_df.columns:
        start_idx = edited_df.columns.get_loc("product_code")
        end_idx = edited_df.columns.get_loc("Inventory")
        if start_idx <= end_idx:
            copy_cols = edited_df.columns[start_idx : end_idx + 1].tolist()

    if not copy_cols:
        st.info("Kolom product_code atau Inventory tidak ditemukan untuk fitur copy.")
        return

    copy_df = edited_df[copy_cols].copy()
    copy_text = copy_df.to_csv(index=False, sep="\t")
    escaped_copy_text = html.escape(copy_text)
    st.write("Copy tabel (kolom product_code sampai Inventory):")
    components.html(
        f"""
        <div style='display:flex; gap:8px; align-items:center;'>
          <button onclick='navigator.clipboard.writeText(document.getElementById("copy_area").value)'>Copy Table</button>
          <span id='copy_status' style='font-size:12px; color:#333;'></span>
        </div>
        <textarea id='copy_area' style='width:100%; height:160px; margin-top:8px;'>{escaped_copy_text}</textarea>
        <script>
          const btn = document.querySelector('button');
          const status = document.getElementById('copy_status');
          btn.addEventListener('click', () => {{
            status.textContent = 'Copied';
            setTimeout(() => status.textContent = '', 1500);
          }});
        </script>
        """,
        height=230,
    )


def build_benchmark_inputs(edited_df, benchmark_product_codes, selected_currency):
    benchmark_inputs = []
    if selected_currency in ("Semua", "IDR"):
        benchmark_inputs = [
            (
                "Benchmark IDR Konvensional",
                product_options_for_segment(edited_df, "IDR", "Konvensional"),
                auto_benchmark_products(edited_df, benchmark_product_codes, "IDR", "Konvensional"),
                "red",
            ),
            (
                "Benchmark IDR Syariah",
                product_options_for_segment(edited_df, "IDR", "Syariah"),
                auto_benchmark_products(edited_df, benchmark_product_codes, "IDR", "Syariah"),
                "green",
            ),
        ]

    if selected_currency in ("Semua", "USD", "US"):
        benchmark_inputs.append(
            (
                "Benchmark USD",
                product_options_for_currency(edited_df, "USD"),
                auto_benchmark_products(edited_df, benchmark_product_codes, "USD"),
                "orange",
            )
        )

    return benchmark_inputs


def render_benchmark_selectors(benchmark_inputs, selected_currency):
    benchmark_configs = []
    if not benchmark_inputs:
        return benchmark_configs

    bench_cols = st.columns(len(benchmark_inputs))
    for idx, (label, options, default, color) in enumerate(benchmark_inputs):
        with bench_cols[idx]:
            benchmark_selection = st.multiselect(
                f"Pilih Seri Obligasi {label}",
                options=options,
                default=default,
                key=f"benchmark_selection_{idx}_{selected_currency}",
            )
            benchmark_configs.append((label, benchmark_selection, color))

    return benchmark_configs


def parse_yield_chart(val):
    try:
        return float(str(val).replace("%", "").replace(",", "."))
    except Exception:
        return None


def render_yield_curve(edited_df, benchmark_product_codes):
    st.write("---")

    required_cols = {"product_code", "year maturity", "y mbi jual"}
    if not required_cols.issubset(edited_df.columns):
        return

    if "currency check" in edited_df.columns:
        available_currencies = ["Semua"] + edited_df["currency check"].dropna().unique().tolist()
        selected_currency = st.radio(
            "Filter Mata Uang Grafik:",
            options=available_currencies,
            horizontal=True,
        )
    else:
        selected_currency = "Semua"

    chart_title = (
        f"Bonds Chart {selected_currency if selected_currency != 'Semua' else ''} - Mark to Market"
    ).replace("  ", " ")
    st.subheader(chart_title)

    filtered_df = edited_df.copy()
    if selected_currency != "Semua" and "currency check" in filtered_df.columns:
        filtered_df = filtered_df[currency_matches(filtered_df["currency check"], selected_currency)]

    benchmark_inputs = build_benchmark_inputs(edited_df, benchmark_product_codes, selected_currency)
    benchmark_configs = render_benchmark_selectors(benchmark_inputs, selected_currency)

    chart_df = filtered_df.copy()
    chart_df["Year Numeric"] = pd.to_numeric(chart_df["year maturity"], errors="coerce")
    chart_df["Yield Numeric"] = chart_df["y mbi jual"].apply(parse_yield_chart)
    chart_df = chart_df.dropna(subset=["Year Numeric", "Yield Numeric"])

    if chart_df.empty:
        return

    import plotly.graph_objects as go

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=chart_df["Year Numeric"],
            y=chart_df["Yield Numeric"],
            mode="markers+text",
            name="Seri Bonds",
            text=chart_df["product_code"],
            textposition="top right",
            marker=dict(size=8, color="#5D9CEC"),
            textfont=dict(color="black", size=10),
        )
    )

    summary_tables = []
    for bench_name, selected_products, color in benchmark_configs:
        if not selected_products:
            continue

        bench_df = chart_df[chart_df["product_code"].isin(selected_products)].copy()
        bench_df = bench_df.sort_values(by="Year Numeric")

        fig.add_trace(
            go.Scatter(
                x=bench_df["Year Numeric"],
                y=bench_df["Yield Numeric"],
                mode="lines+markers",
                name=f"{bench_name} Curve",
                line=dict(color=color, width=3),
                marker=dict(size=10, color=color),
            )
        )

        summary_df = bench_df[["product_code", "year maturity", "y mbi jual"]].copy()
        summary_df.columns = [bench_name, "Year", "Yield"]
        summary_tables.append((bench_name, summary_df))

    if summary_tables:
        st.write("**Tabel Ringkasan Benchmark**")
        summary_cols = st.columns(len(summary_tables))
        for idx, (bench_name, summary_df) in enumerate(summary_tables):
            with summary_cols[idx]:
                st.write(f"**{bench_name}**")
                st.dataframe(summary_df, hide_index=True, use_container_width=True)

    fig.update_layout(
        xaxis_title="Year",
        yaxis_title="Yield (% MBI Jual)",
        hovermode="closest",
        yaxis=dict(tickformat=".2f", ticksuffix="%"),
        height=600,
        plot_bgcolor="white",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=20, r=20, t=40, b=20),
    )
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor="LightGray")
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor="LightGray")

    st.plotly_chart(fig, use_container_width=True)
