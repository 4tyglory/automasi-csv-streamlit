import streamlit as st
import pandas as pd
import re
import math
import os
from io import BytesIO

st.title("Automasi CSV dari Excel")

uploaded_excel = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])
uploaded_db = st.file_uploader("Upload file Database (.xlsx)", type=["xlsx"])

if uploaded_excel and uploaded_db:
    # Load file
    try:
        db = pd.read_excel(uploaded_db)
        xls = pd.ExcelFile(uploaded_excel)
    except Exception as e:
        st.error(f"Error membaca file: {e}")
        st.stop()

    sheet_names = xls.sheet_names
    sheet_selected = st.selectbox("Pilih sheet yang akan diproses", sheet_names)

    batch_size = 1000

    def normalize_text(s):
        if not isinstance(s, str):
            s = str(s)
        return re.sub(r'\s+', '', s.lower())

    def parse_sheet_name(name):
        kuota = None
        kuota_match = re.search(r'(\d+[\.,]?\d*)\s*gb', name, re.IGNORECASE)
        if kuota_match:
            kuota = kuota_match.group(1).replace(',', '.')

        hari = None
        hari_match = re.search(r'(\d+)\s*h', name, re.IGNORECASE)
        if hari_match:
            hari = hari_match.group(1)

        zona = None
        zona_match = re.search(r'(z\d+)', name, re.IGNORECASE)
        if zona_match:
            zona = zona_match.group(1).lower()

        return kuota, hari, zona

    def format_decimal_with_koma(s):
        s = s.lower()
        if '.' in s:
            s = s.replace('.', 'koma')
        return s

    if st.button("Proses"):
        with st.spinner('Memproses data...'):
            try:
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_selected, header=None)

                numbers_12digit = []
                for _, row in df.iterrows():
                    for val in row.astype(str):
                        found = re.findall(r'\b\d{12}\b', val)
                        if found:
                            numbers_12digit.extend(found)

                total_numbers = len(numbers_12digit)
                if total_numbers == 0:
                    st.warning(f"Tidak ada angka 12 digit ditemukan di sheet '{sheet_selected}'.")
                    st.stop()

                num_files = math.ceil(total_numbers / batch_size)

                db['Nama Barang Norm'] = db['Nama Barang'].apply(normalize_text)
                sheet_norm = normalize_text(sheet_selected)
                match_db = db[db['Nama Barang Norm'] == sheet_norm]

                if not match_db.empty:
                    harga = match_db.iloc[0]['Harga']
                    bulk_code = match_db.iloc[0]['bulk']
                    zona_db = match_db.iloc[0]['Zona']
                else:
                    harga = None
                    bulk_code = 'bulkUnknown'
                    zona_db = ''

                kuota, hari, zona_sheet = parse_sheet_name(sheet_selected)
                zona = zona_db if zona_db and str(zona_db).strip() != '-' and str(zona_db).strip() != '' else ''

                kuota_text_raw = f"{kuota}gb" if kuota else "unknowngb"
                kuota_text = format_decimal_with_koma(kuota_text_raw)

                bulk_text_raw = bulk_code.replace(' ', '')
                bulk_text = format_decimal_with_koma(bulk_text_raw)

                files_buffers = []

                for i in range(num_files):
                    batch_numbers = numbers_12digit[i*batch_size:(i+1)*batch_size]
                    qty = len(batch_numbers)
                    file_index = i + 1

                    hari_text = f"{hari}hari" if hari else "unknownhari"

                    if zona:
                        filename = f"{file_index} vcr fisik internet {hari_text} {kuota_text} {bulk_text} {zona} {qty}.csv"
                    else:
                        filename = f"{file_index} vcr fisik internet {hari_text} {kuota_text} {bulk_text} {qty}.csv"

                    # Buat file CSV di memory buffer
                    buffer = BytesIO()
                    buffer.write('\n'.join(batch_numbers).encode('utf-8'))
                    buffer.seek(0)

                    files_buffers.append((filename, buffer))

                st.success("Proses selesai! Unduh file CSV di bawah ini:")

                for filename, buffer in files_buffers:
                    st.download_button(
                        label=f"Download {filename}",
                        data=buffer,
                        file_name=filename,
                        mime='text/csv'
                    )

            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
