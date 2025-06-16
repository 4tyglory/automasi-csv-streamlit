import streamlit as st
import pandas as pd
import re
import math
import io
import zipfile

st.title("Automasi CSV Multi Sheet dengan Pilihan Download Otomatis")

uploaded_excel = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])
uploaded_db = st.file_uploader("Upload file Database (.xlsx)", type=["xlsx"])

def normalize_text(s):
    if not isinstance(s, str):
        s = str(s)
    s = s.lower().replace('.', ',').replace(' ', '')
    return s

def extract_package_name(sheet_name):
    pattern = r'(\d+[\.,]?\d*\s*gb)\s*(\d+\s*(h|hari))\s*(z\d+)'
    match = re.search(pattern, sheet_name, re.IGNORECASE)
    if match:
        kuota = match.group(1).replace(',', '.').lower()
        hari_num = re.findall(r'\d+', match.group(2))[0]
        hari = f"{hari_num}h"
        zona = match.group(4).lower()
        return f"{kuota} {hari} {zona}"
    else:
        return sheet_name.lower()

def format_decimal_with_koma(s):
    s = s.lower()
    if '.' in s:
        s = s.replace('.', 'koma')
    return s

batch_size = 1000

if uploaded_excel and uploaded_db:
    try:
        db = pd.read_excel(uploaded_db)
        xls = pd.ExcelFile(uploaded_excel)
    except Exception as e:
        st.error(f"Error membaca file: {e}")
        st.stop()

    sheet_names = xls.sheet_names

    if 'processed_sheets' not in st.session_state:
        st.session_state.processed_sheets = {}

    if st.button("Proses Semua Sheet"):
        with st.spinner('Memproses semua sheet...'):
            processed = {}
            for sheet_selected in sheet_names:
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_selected, header=None)
                numbers_12digit = []
                for _, row in df.iterrows():
                    for val in row.astype(str):
                        found = re.findall(r'\b\d{12}\b', val)
                        if found:
                            numbers_12digit.extend(found)
                total_numbers = len(numbers_12digit)
                if total_numbers == 0:
                    continue
                num_files = math.ceil(total_numbers / batch_size)

                db['Nama Barang Norm'] = db['Nama Barang'].apply(lambda x: normalize_text(extract_package_name(x)))
                sheet_norm = normalize_text(extract_package_name(sheet_selected))
                match_db = db[db['Nama Barang Norm'] == sheet_norm]

                if not match_db.empty:
                    harga = match_db.iloc[0]['Harga']
                    bulk_code = match_db.iloc[0]['bulk']
                    zona_db = match_db.iloc[0]['Zona']
                else:
                    harga = None
                    bulk_code = 'bulkUnknown'
                    zona_db = ''

                kuota, hari, zona_sheet = None, None, None
                pattern_kuota = r'(\d+[\.,]?\d*)\s*gb'
                kuota_match = re.search(pattern_kuota, sheet_selected, re.IGNORECASE)
                if kuota_match:
                    kuota = kuota_match.group(1).replace(',', '.')
                hari_match = re.search(r'(\d+)\s*h', sheet_selected, re.IGNORECASE)
                if hari_match:
                    hari = hari_match.group(1)
                zona_match = re.search(r'(z\d+)', sheet_selected, re.IGNORECASE)
                if zona_match:
                    zona_sheet = zona_match.group(1).lower()

                zona = zona_db if zona_db and str(zona_db).strip() != '-' and str(zona_db).strip() != '' else (zona_sheet if zona_sheet else '')

                kuota_text_raw = f"{kuota}gb" if kuota else "unknowngb"
                kuota_text = format_decimal_with_koma(kuota_text_raw)

                bulk_text_raw = bulk_code.replace(' ', '')
                bulk_text = format_decimal_with_koma(bulk_text_raw)

                files_buffers = []

                for i in range(num_files):
                    batch_numbers = numbers_12digit[i*batch_size:(i+1)*batch_size]
                    qty = len(batch_numbers)
                    file_index = i + 1

                    hari_raw = f"{hari}h" if hari else "unknownhari"
                    hari_text = re.sub(r'h$', 'hari', hari_raw)

                    if zona:
                        filename = f"{file_index} vcr fisik internet {hari_text} {kuota_text} {bulk_text} {zona} {qty}.csv"
                    else:
                        filename = f"{file_index} vcr fisik internet {hari_text} {kuota_text} {bulk_text} {qty}.csv"

                    buffer = io.BytesIO()
                    buffer.write('\n'.join(batch_numbers).encode('utf-8'))
                    buffer.seek(0)

                    files_buffers.append((filename, buffer))

                processed[sheet_selected] = files_buffers

            st.session_state.processed_sheets = processed
            st.success("Semua sheet selesai diproses.")

    if st.session_state.processed_sheets:
        st.subheader("Pilih sheet untuk download:")
        selected_sheets = st.multiselect(
            "Centang sheet yang ingin di-download:",
            options=list(st.session_state.processed_sheets.keys())
        )

        if st.button("Download ZIP dari Sheet Terpilih"):
            if not selected_sheets:
                st.warning("Silakan pilih minimal 1 sheet.")
            else:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for sheet_name in selected_sheets:
                        for filename, buffer in st.session_state.processed_sheets[sheet_name]:
                            zip_path = f"{sheet_name}/{filename}"
                            zip_file.writestr(zip_path, buffer.getvalue())
                zip_buffer.seek(0)

                zip_filename = f"hasil_csv_{'_'.join([s.replace(' ', '_') for s in selected_sheets])}.zip"

                st.download_button(
                    label="Download ZIP",
                    data=zip_buffer,
                    file_name=zip_filename,
                    mime="application/zip"
                )

else:
    st.info("Silakan upload file Excel dan Database terlebih dahulu.")
