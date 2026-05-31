# Maybank Bonds PDF Extractor

Aplikasi Streamlit untuk membaca PDF indikasi harga obligasi, membersihkan data, menghitung simulasi risiko/yield, menampilkan yield curve benchmark, dan mengekspor hasil ke Excel.

## Fitur

- Upload PDF obligasi langsung dari UI Streamlit.
- Deteksi otomatis dua format PDF:
  - Format tabel lama dari file pricing.
  - Format `BOND PRICE INDICATION` Maybank.
- Cleaning angka, persen, dan tanggal.
- Deteksi currency `IDR` / `USD` dari `product_code`.
- Deteksi benchmark otomatis:
  - Format lama: product code dengan font `Arial-Black`.
  - Format Maybank baru: section `Benchmark ...`.
- Grafik yield curve dengan benchmark:
  - `Benchmark IDR Konvensional`
  - `Benchmark IDR Syariah`
  - `Benchmark USD`
- Styling kolom `1D`:
  - Merah untuk minus.
  - Hijau untuk naik.
  - Orange untuk stagnan.
- Kalkulasi simulasi:
  - `MDURATION`
  - `Rate Hike`
  - `Rate Cut`
  - `Rate Hike Price`
  - `Rate Cut Price`
  - `Total Year to Maturity`
  - `Price if Yield Hike`
  - `Price if Yield Cut`
- Export Excel dengan formula dinamis.

## Cara Menjalankan

Install dependency:

```powershell
pip install -r requirements.txt
```

Jalankan aplikasi:

```powershell
streamlit run main.py
```

Buka URL yang ditampilkan oleh Streamlit, biasanya:

```text
http://localhost:8501
```

## Struktur Project

```text
main.py              Entry point aplikasi Streamlit.
pdf_parsers.py       Parser PDF format lama dan format Maybank baru.
data_processing.py   Cleaning, normalisasi kolom, tanggal, currency, dan persen.
calculations.py      Kalkulasi MDURATION, rate impact, PV, dan maturity.
excel_export.py      Export Excel dan formula-formula Excel.
ui_components.py     Komponen UI: tabel, copy table, dan chart benchmark.
bond_utils.py        Helper umum untuk parsing, currency, benchmark, dan Excel ref.
requirements.txt     Dependency Python.
```

## Format PDF Yang Didukung

### Format Lama

Format lama dibaca dari tabel PDF menggunakan `pdfplumber.extract_table()`.

Kolom yang umum dipakai:

```text
product_code
kupon
mbi_beli
mbi_jual
yield_mbi_beli
yield_mbi_jual
Maturity
Settlement date
Inventory
```

Benchmark pada format ini dideteksi dari `product_code` yang memakai font `Arial-Black`.

### Format Maybank Baru

Format ini dideteksi otomatis jika PDF memiliki header seperti:

```text
BOND PRICE INDICATION
PROD_CODE
MBI BELI
YIELD MBI JUAL
```

Parser membaca section seperti:

```text
Benchmark IDR
Benchmark USD
Benchmark PBS Series
Non Benchmark IDR
Non Benchmark USD
```

Baris yang berada di section `Benchmark ...` otomatis dijadikan benchmark.

## Catatan Benchmark

Pembagian benchmark:

- IDR Konvensional: contoh `FR`, `ORI`, dan selain prefix syariah.
- IDR Syariah: prefix `PBS`, `SR`, `ST`, `INDOIS`.
- USD: tidak dipecah syariah/konvensional, ditampilkan sebagai satu garis `Benchmark USD`.

## Catatan Development

Untuk cek syntax semua modul:

```powershell
python -m py_compile main.py bond_utils.py pdf_parsers.py data_processing.py calculations.py excel_export.py ui_components.py
```

Jika menambah format PDF baru, titik masuk terbaik adalah:

```text
pdf_parsers.py
```

Tambahkan fungsi deteksi format, parser baru, lalu hubungkan di `extract_pdf_dataframe()`.
