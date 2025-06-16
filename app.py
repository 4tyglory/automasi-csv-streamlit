import streamlit as st
import pandas as pd
import re
import math
import io
import zipfile

st.set_page_config(page_title="Automasi CSV Multi Sheet", layout="wide")

st.title("🚀 Automasi CSV Multi Sheet dengan Pilihan Download")

uploaded_excel = st.file_uploader("📥 Upload file Excel (.xlsx)", type=["xlsx"])
uploaded_db = st.file_uploader("📥 Upload file Database (.xlsx)", type=["xlsx"])

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

# CSS untuk horizontal scroll checkbox
st.markdown("""
<style>
.checkbox-container {
    display: flex;
    overflow-x: auto;
    white-space: nowrap;
    padding: 10px 0;
    border: 1px solid #ddd;
    border-radius: 8px;
}
.checkbox-item {
    flex: 0 0 auto;
    margin-right: 25px;
    padding: 5px 10px;
    background-color: #f5f5f5;
    border-radius: 6px;
    user-select: none;
}
.footer {
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background-color: #f1f1f1;
    color: #555;
    text-align: center;
    padding: 5px 0;
    font-size: 14px;
    font-family: Arial, sans-serif;
    border-top: 1px solid #ddd;
    z-index: 1000;
}
.footer a {
    color: #0366d6;
    text-decoration: none;
    font-weight: bold;
}
.footer a:hover {
    text-decoration: underline;
}
</style>
""", unsafe_allow_html=True)

if uploaded_excel and uploaded_db:
    try:
        db = pd.read_excel(uploaded_db)
        xls = pd.ExcelFile(uploaded_excel)
    except Exception as e:
        st.error(f"⚠️ Error membaca file: {e}")
        st.stop()

    sheet_names = xls.sheet_names

    if 'processed_sheets' not in st.session_state:
        st.session_state.processed_sheets = {}

    if st.button("▶️ Proses Semua Sheet"):
        with st.spinner('⚙️ Memproses semua sheet...'):
            progress_bar = st.progress(0)
            processed = {}
            for idx, sheet_selected in enumerate(sheet_names):
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_selected, header=None)
                numbers_12digit = []
                for _, row in df.iterrows():
                    for val in row.astype(str):
                        found = re.findall(r'\b\d{12}\b', val)
                        if found:
                            numbers_12digit.extend(found)
                total_numbers = len(numbers_12digit)
                if total_numbers == 0:
                    st.warning(f"⚠️ Tidak ada angka 12 digit ditemukan di sheet **{sheet_selected}**.")
                    progress_bar.progress((idx+1)/len(sheet_names))
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
                progress_bar.progress((idx+1)/len(sheet_names))

            st.session_state.processed_sheets = processed
            st.success("🎉 Semua sheet selesai diproses.")

    if st.session_state.processed_sheets:
        st.subheader("📋 Sheet yang sudah diproses:")
        sheet_list = list(st.session_state.processed_sheets.keys())

        # Container dengan class checkbox-container
        st.markdown('<div class="checkbox-container">', unsafe_allow_html=True)
        selected_sheets = []
        for sheet in sheet_list:
            key = f"cb_{sheet}"
            checked = st.checkbox(sheet, key=key)
            if checked:
                selected_sheets.append(sheet)
            st.markdown('<div class="checkbox-item"></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button("📦 Download ZIP dari Sheet Terpilih"):
            if not selected_sheets:
                st.warning("⚠️ Silakan pilih minimal 1 sheet.")
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
                    label="📥 Download ZIP",
                    data=zip_buffer,
                    file_name=zip_filename,
                    mime="application/zip"
                )

# Footer dengan info author dan GitHub
st.markdown("""
<style>
.footer {
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background-color: #f1f1f1;
    color: #555;
    text-align: center;
    padding: 5px 0;
    font-size: 14px;
    font-family: Arial, sans-serif;
    border-top: 1px solid #ddd;
    z-index: 1000;
}
.footer a {
    color: #0366d6;
    text-decoration: none;
    font-weight: bold;
}
.footer a:hover {
    text-decoration: underline;
}
</style>
<div class="footer">
    Dibuat oleh: Muhammad Aldi Yusuf | Github: <a href="https://github.com/4tyglory" target="_blank">4tyglory</a>
</div>
""", unsafe_allow_html=True)
