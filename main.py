import streamlit as st
import pdfplumber
import pandas as pd
import io
import numpy_financial as npf

# Konfigurasi agar halaman tampil full-width (lebar penuh)
st.set_page_config(layout="wide", page_title="PDF Table Extractor & Editor")

st.title("PDF Table Extractor & Editor")

uploaded_file = st.file_uploader("Unggah file PDF Excel", type="pdf")

if uploaded_file is not None:
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            all_rows = []
            
            # Ekstrak tabel dari semua halaman
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_rows.extend(table)
            
            if all_rows:
                # Jadikan baris pertama sebagai header
                header = all_rows[0]
                data = all_rows[1:]
                
                df = pd.DataFrame(data, columns=header)
                
                # Bersihkan nama kolom dari karakter 'enter' (newline/carriage return)
                df.columns = [str(col).replace('\n', '').replace('\r', '').strip() if pd.notna(col) else col for col in df.columns]
                
                # Tambah kolom kustom di sebelah kanan tabel
                import re
                
                # --- PEMBERSIHAN DATA DAN KONVERSI TIPE ---
                def clean_numeric(val):
                    if pd.isna(val) or val == '': return val
                    if not isinstance(val, str): return float(val)
                    s = val.strip()
                    # Buang spasi di antara angka (misal "1 000" atau "5 . 5" atau "4 75,000,000")
                    s = re.sub(r'\s+', '', s)
                    if s == '': return val
                    
                    # Deteksi kalau dia persentase
                    is_percent = '%' in s
                    s = s.replace('%', '')
                    
                    # Cek apakah berisi karakter angka (termasuk . , -)
                    if not all(c in '0123456789.,-' for c in s):
                        return val 
                        
                    # Handle format koma & titik
                    if ',' in s and '.' in s:
                        if s.rfind(',') > s.rfind('.'): # koma sebagai desimal
                            s = s.replace('.', '').replace(',', '.')
                        else: # koma sebagai ribuan
                            s = s.replace(',', '')
                    elif ',' in s:
                        # Bisa jadi desimal (5,5) atau ribuan (475,000)
                        if s.count(',') > 1:
                            s = s.replace(',', '')
                        else:
                            # Jika hanya ada 1 koma, cek jumlah digit di belakangnya
                            parts = s.split(',')
                            if len(parts[-1]) == 3:
                                s = s.replace(',', '') # Asumsi koma ribuan karena 3 digit tepat
                            else:
                                s = s.replace(',', '.') # Asumsi desimal
                        
                    try:
                        v = float(s)
                        if is_percent:
                            v = v / 100.0
                        return v
                    except:
                        return val

                # Konversi semua kolom numerik utama agar jadi float asli
                numeric_candidates = ['kupon', 'mbi_beli', 'mbi_jual', 'yield_mbi_beli', 'yield_mbi_jual', 'Inventory']
                for c in numeric_candidates:
                    if c in df.columns:
                        df[c] = df[c].apply(clean_numeric)
                
                # Konversi kolom tanggal ke Datetime
                date_candidates = ['Maturity', 'Settlement date']
                for c in date_candidates:
                    if c in df.columns:
                        df[c] = pd.to_datetime(df[c], errors='ignore')

                # 1. Currency Check
                if 'product_code' in df.columns:
                    df['currency check'] = df['product_code'].apply(
                        lambda x: 'US' if isinstance(x, str) and ('indon' in x.lower() or 'indois' in x.lower() or 'is' in x.lower() or 'in' in x.lower()) else 'IDR'
                    )
                
                # 2. Year Maturity
                maturity_col = next((col for col in df.columns if pd.notna(col) and 'maturity' in str(col).lower()), None)
                if maturity_col:
                    def extract_year(date_val):
                        if pd.isna(date_val): return ''
                        if isinstance(date_val, pd.Timestamp): return str(date_val.year)
                        match = re.search(r'((?:19|20)\d{2})', str(date_val))
                        return match.group(1) if match else str(date_val)
                    df['year maturity'] = df[maturity_col].apply(extract_year)
                
                # 3, 4, 5. Kolom Persentase (Kupon %, Y MBI Beli, Y MBI Jual)
                def to_percent_str(val):
                    try:
                        if pd.isna(val): return None
                        v = round(float(val), 6)
                        return f"{v}%"
                    except:
                        return val

                if 'kupon' in df.columns:
                    df['kupon %'] = df['kupon'].apply(to_percent_str)
                
                if 'yield_mbi_beli' in df.columns:
                    df['y mbi beli'] = df['yield_mbi_beli'].apply(to_percent_str)
                
                if 'yield_mbi_jual' in df.columns:
                    df['y mbi jual'] = df['yield_mbi_jual'].apply(to_percent_str)
                
                # 6. Kolom MDURATION 
                st.write("---")
                st.subheader("Pengaturan Parameter Simulasi")
                
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    settlement_date = st.date_input("Tanggal Settlement ($Y$4)", value=pd.to_datetime("today"))
                with col2:
                    cut_rate_input = st.number_input("Cut / Hike Rate (%)", value=0.2500, step=0.0100, format="%.4f")
                with col3:
                    base_date_input = st.date_input("Tanggal Basis (Hari Ini)", value=pd.to_datetime("today"))
                with col4:
                    yield_hike_input = st.number_input("Yield Naik (%)", value=0.0, step=0.0100, format="%.4f")
                with col5:
                    yield_cut_input = st.number_input("Yield Turun (%)", value=0.0, step=0.0100, format="%.4f")
                
                def hitung_mduration_ql(row):
                    try:
                        import QuantLib as ql
                        
                        mat_val = row.get('Maturity') if 'Maturity' in row else (row.get(maturity_col) if maturity_col else None)
                        kup_val = row.get('kupon %')
                        yld_val = row.get('y mbi jual')
                        
                        if pd.isna(mat_val) or pd.isna(kup_val) or pd.isna(yld_val):
                            return None
                        
                        maturity_date = pd.to_datetime(mat_val)
                        settlement_dt = pd.to_datetime(settlement_date)
                        
                        if pd.isna(maturity_date) or maturity_date <= settlement_dt:
                            return None
                            
                        def parse_to_rate(r):
                            return float(str(r).replace('%', '').replace(',', '.')) / 100.0
                            
                        c = parse_to_rate(kup_val)
                        y = parse_to_rate(yld_val)
                        
                        set_date_ql = ql.Date(settlement_dt.day, settlement_dt.month, settlement_dt.year)
                        mat_date_ql = ql.Date(maturity_date.day, maturity_date.month, maturity_date.year)
                        
                        ql.Settings.instance().evaluationDate = set_date_ql
                        
                        calendar = ql.NullCalendar()
                        day_count = ql.ActualActual(ql.ActualActual.ISMA) 
                        q_freq = ql.Semiannual
                        
                        schedule = ql.Schedule(set_date_ql, mat_date_ql, ql.Period(q_freq), calendar,
                                               ql.Unadjusted, ql.Unadjusted, ql.DateGeneration.Backward, False)
                        
                        bond = ql.FixedRateBond(0, 100.0, schedule, [c], day_count)
                        
                        interest_rate = ql.InterestRate(y, day_count, ql.Compounded, q_freq)
                        mod_dur = ql.BondFunctions.duration(bond, interest_rate, ql.Duration.Modified)
                        
                        return round(mod_dur, 4)
                    except:
                        return None

                df['MDURATION'] = df.apply(hitung_mduration_ql, axis=1)

                # 7 & 8. Kolom Rate Hike & Rate Cut
                def hitung_rate_impact(val, is_hike=False):
                    try:
                        if pd.isna(val): return None
                        v = float(str(val).replace('%', '').replace(',', '.')) / 100.0
                        
                        result = v * (cut_rate_input / 100.0) 
                        if is_hike:
                            result = -result 
                            
                        return round(result, 6)
                    except:
                        return None
                        
                df['Rate Hike'] = df['y mbi jual'].apply(lambda x: hitung_rate_impact(x, is_hike=True))
                df['Rate Cut'] = df['y mbi jual'].apply(lambda x: hitung_rate_impact(x, is_hike=False))

                if 'mbi_jual' in df.columns:
                    df['Rate Hike Price'] = df['mbi_jual'] + (df['Rate Hike'] * df['mbi_jual'])
                else:
                    df['Rate Hike Price'] = None

                if 'mbi_jual' in df.columns:
                    df['Rate Cut Price'] = df['mbi_jual'] + (df['Rate Cut'] * df['mbi_jual'])
                else:
                    df['Rate Cut Price'] = None
                
                base_date_pd = pd.to_datetime(base_date_input)
                
                # Kolom Total Year to Maturity
                def hitung_ytm_years(row):
                    mat = row.get('Maturity') if 'Maturity' in row else row.get(maturity_col)
                    if pd.isna(mat): return None
                    try:
                        diff = (pd.to_datetime(mat) - base_date_pd.normalize()).days / 365.25
                        return round(diff, 4)
                    except:
                        return None
                        
                df['Total Year to Maturity'] = df.apply(hitung_ytm_years, axis=1)
                
                # Kolom Price menggunakan formula PV
                def hitung_price_pv(row, is_hike=True):
                    try:
                        y_mbi = float(str(row.get('y mbi jual', '0')).replace('%', '').replace(',', '.')) / 100.0
                        kupon = float(str(row.get('kupon %', '0')).replace('%', '').replace(',', '.')) / 100.0
                        ytm_years = row.get('Total Year to Maturity')
                        
                        if pd.isna(y_mbi) or pd.isna(kupon) or pd.isna(ytm_years):
                            return None
                        
                        if is_hike:
                            rate = y_mbi + (yield_hike_input / 100.0)
                        else:
                            rate = y_mbi - (yield_cut_input / 100.0)
                            
                        nper = ytm_years
                        pmt = 100 * kupon
                        fv = 100
                        
                        price = -npf.pv(rate, nper, pmt, fv, when=0)
                        return round(price, 4)
                    except:
                        return None
                        
                df['Price if Yield Hike'] = df.apply(lambda r: hitung_price_pv(r, is_hike=True), axis=1)
                df['Price if Yield Cut'] = df.apply(lambda r: hitung_price_pv(r, is_hike=False), axis=1)

                st.write("Edit data langsung pada tabel di bawah:")
                
                # Memastikan tabel memanfaatkan seluruh lebar kontainer
                edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
                
                excel_df = edited_df.copy()

                # --- VISUALISASI YIELD CURVE BENCHMARK ---
                st.write("---")
                
                if 'product_code' in edited_df.columns and 'year maturity' in edited_df.columns and 'y mbi jual' in edited_df.columns:
                    
                    if 'currency check' in edited_df.columns:
                        available_currencies = ["Semua"] + edited_df['currency check'].dropna().unique().tolist()
                        selected_currency = st.radio("Filter Mata Uang Grafik:", options=available_currencies, horizontal=True)
                    else:
                        selected_currency = "Semua"

                    chart_title = f"Bonds Chart {selected_currency if selected_currency != 'Semua' else ''} - Mark to Market".replace("  ", " ")
                    st.subheader(chart_title)
                    
                    filtered_df_for_chart = edited_df.copy()
                    if selected_currency != "Semua" and 'currency check' in filtered_df_for_chart.columns:
                        filtered_df_for_chart = filtered_df_for_chart[filtered_df_for_chart['currency check'] == selected_currency]
                        
                    valid_products = filtered_df_for_chart['product_code'].dropna().unique().tolist()
                    
                    benchmark_selection = st.multiselect(
                        "Pilih Seri Obligasi Benchmark",
                        options=valid_products,
                        default=[]
                    )
                    
                    chart_df = filtered_df_for_chart.copy()
                    chart_df['Year Numeric'] = pd.to_numeric(chart_df['year maturity'], errors='coerce')
                    
                    def parse_yield_chart(val):
                        try:
                            return float(str(val).replace('%', '').replace(',', '.'))
                        except:
                            return None
                    chart_df['Yield Numeric'] = chart_df['y mbi jual'].apply(parse_yield_chart)
                    
                    chart_df = chart_df.dropna(subset=['Year Numeric', 'Yield Numeric'])
                    
                    if not chart_df.empty:
                        import plotly.graph_objects as go
                        
                        fig = go.Figure()
                        
                        # 1. Scatter Plot
                        fig.add_trace(go.Scatter(
                            x=chart_df['Year Numeric'],
                            y=chart_df['Yield Numeric'],
                            mode='markers+text',
                            name='Seri Bonds',
                            text=chart_df['product_code'],
                            textposition="top right",
                            marker=dict(size=8, color='#5D9CEC'),
                            textfont=dict(color='black', size=10) 
                        ))
                        
                        # 2. Line Plot (Benchmark Curve)
                        if benchmark_selection:
                            bench_df = chart_df[chart_df['product_code'].isin(benchmark_selection)].copy()
                            bench_df = bench_df.sort_values(by='Year Numeric')
                            
                            fig.add_trace(go.Scatter(
                                x=bench_df['Year Numeric'],
                                y=bench_df['Yield Numeric'],
                                mode='lines+markers',
                                name='Benchmark Curve',
                                line=dict(color='red', width=3),
                                marker=dict(size=10, color='red')
                            ))
                            
                            st.write("**Tabel Ringkasan Benchmark**")
                            summary_df = bench_df[['product_code', 'year maturity', 'y mbi jual']].copy()
                            summary_df.columns = ['Benchmark', 'Year', 'Yield']
                            
                            # Memastikan tabel rekap menggunakan seluruh lebar layar
                            st.dataframe(summary_df, hide_index=True, use_container_width=True)
                            
                        # Layout Diperbaiki agar rapi saat ditarik full layar
                        fig.update_layout(
                            xaxis_title="Year",
                            yaxis_title="Yield (% MBI Jual)",
                            hovermode="closest",
                            yaxis=dict(tickformat=".2f", ticksuffix="%"),
                            height=600,
                            plot_bgcolor='white',
                            showlegend=True,
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), # Legenda diletakkan di atas
                            margin=dict(l=20, r=20, t=40, b=20)
                        )
                        
                        fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
                        fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
                        
                        st.plotly_chart(fig, use_container_width=True)
                
                # --- PERSIAPAN DOWNLOAD EXCEL ---
                def clean_for_excel(val):
                    try:
                        return float(str(val).replace('%', '').replace(',', '.')) / 100.0
                    except:
                        return val
                        
                for c in ['kupon %', 'y mbi beli', 'y mbi jual']:
                    if c in excel_df.columns:
                        excel_df[c] = excel_df[c].apply(clean_for_excel)
                
                excel_df['MDURATION'] = [f"=MDURATION($Y$4, H{i+2}, M{i+2}, O{i+2}, 2, 1)" for i in range(len(excel_df))]
                excel_df['Rate Hike'] = [f"=-O{i+2} * {cut_rate_input/100.0}" for i in range(len(excel_df))]
                excel_df['Rate Cut']  = [f"=O{i+2} * {cut_rate_input/100.0}" for i in range(len(excel_df))]
                excel_df['Rate Hike Price'] = [f"=Q{i+2} * E{i+2}" for i in range(len(excel_df))]
                
                excel_df['Price if Yield Hike'] = [f"=-PV((O{i+2}+{yield_hike_input/100.0}), U{i+2}, (100*M{i+2}), 100, 0)" for i in range(len(excel_df))]
                excel_df['Price if Yield Cut']  = [f"=-PV((O{i+2}-{yield_cut_input/100.0}), U{i+2}, (100*M{i+2}), 100, 0)" for i in range(len(excel_df))]
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    excel_df.to_excel(writer, index=False, sheet_name='Sheet1')
                
                st.download_button(
                    label="Unduh Data (Excel)",
                    data=buffer.getvalue(),
                    file_name="hasil_edit.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Tabel tidak terdeteksi pada PDF ini.")
                
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")