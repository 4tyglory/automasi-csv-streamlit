import streamlit as st
import pandas as pd
import re
import math
import io
import zipfile
import os

# ------- Universal Key Extraction --------
def extract_key(s):
    """
    Ambil universal key: '[ANGKA]GB [ANGKA]H Z[ANGKA]'
    Support: titik/koma, "h" atau "hari", ignore kata tambahan
    """
    if not isinstance(s, str):
        s = str(s)
    s = s.lower().replace(',', '.')
    m = re.search(r'(\d+\.?\d*)\s*gb\s*(\d+)\s*(h|hari)?\s*z(\d+)', s)
    if m:
        gb = m.group(1)
        h = m.group(2)
        z = m.group(4)
        return f"{gb}gb {h}h z{z}"
    return s.strip()

# ---------- Load database dari repo ----------
def load_bulk_database(path='database.xlsx'):
    db = pd.read_excel(path)
    if 'Nama Barang' not in db.columns or 'bulk' not in db.columns:
        st.error("Kolom 'Nama Barang' dan/atau 'bulk' tidak ditemukan di database.")
        return None
    db['key'] = db['Nama Barang'].apply(extract_key)
    return db

def cari_bulk(sheet_name, db):
    key = extract_key(sheet_name)
    row = db[db['key'] == key]
    if not row.empty:
        return row.iloc[0]['bulk']
    return "bulkUnknown"

def extract_12digit_numbers(df):
    numbers = []
    for _, row in df.iterrows():
        for val in row.astype(str):
            found = re.findall(r'\b\d{12}\b', val)
            if found:
                numbers.extend(found)
    return numbers

def find_duplicates(numbers):
    seen = set()
    duplicates = set()
    for num in numbers:
        if num in seen:
            duplicates.add(num)
        else:
            seen.add(num)
    return duplicates

def format_decimal_with_koma(s: str) -> str:
    return s.lower().replace('.', 'koma')

def parse_sheet_name(sheet_name: str):
    gb, hari, zona = None, None, None
    m = re.search(r'(\d+[\.,]?\d*)\s*gb', sheet_name, re.IGNORECASE)
    if m:
        gb = m.group(1).replace(',', '.').replace(' ', '')
    m2 = re.search(r'(\d+)\s*(h|hari)', sheet_name, re.IGNORECASE)
    if m2:
        hari = m2.group(1)
    m3 = re.search(r'(z\d+)', sheet_name, re.IGNORECASE)
    if m3:
        zona = m3.group(1).lower()
    return gb, hari, zona

def buat_nama_file(file_index, sheet_name, qty, bulk_text):
    gb, hari, zona = parse_sheet_name(sheet_name)
    if gb:
        kuota_text = format_decimal_with_koma(f"{gb}gb")
    else:
        kuota_text = "unknowngb"
    if bulk_text:
        bulk_text = bulk_text.replace(" ", "")
        bulk_text = format_decimal_with_koma(bulk_text)
        if not bulk_text.startswith("bulk"):
            bulk_text = "bulk" + bulk_text
    else:
        bulk_text = "bulkUnknown"
    hari_text = f"{hari}hari" if hari else "unknownhari"
    if zona:
        filename = f"{file_index} vcr fisik internet {hari_text} {kuota_text} {bulk_text} {zona} {qty}.csv"
    else:
        filename = f"{file_index} vcr fisik internet {hari_text} {kuota_text} {bulk_text} {qty}.csv"
    return filename

# ------- UI STREAMLIT --------
st.title("Automasi CSV Multi Sheet Voucher (Smart Matching Bulk)")

if not os.path.exists('database.xlsx'):
    st.error("File database.xlsx tidak ditemukan di direktori repo.")
    st.stop()
bulk_db = load_bulk_database('database.xlsx')
if bulk_db is None:
    st.stop()

uploaded_excel = st.file_uploader("Upload file Excel Voucher (.xlsx)", type=["xlsx"])
batch_size = st.number_input("Batch Size", min_value=1, max_value=10000, value=1000)

if uploaded_excel:
    xls = pd.ExcelFile(uploaded_excel)
    sheet_names = xls.sheet_names

    if 'processed_sheets' not in st.session_state:
        st.session_state.processed_sheets = {}

    if st.button("▶️ Proses Semua Sheet"):
        with st.spinner('⚙️ Memproses semua sheet...'):
            progress_bar = st.progress(0)
            processed = {}
            for idx, sheet_selected in enumerate(sheet_names):
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_selected, header=None)
                numbers_12digit = extract_12digit_numbers(df)
                total_numbers = len(numbers_12digit)
                if total_numbers == 0:
                    st.warning(f"⚠️ Tidak ada angka 12 digit ditemukan di sheet **{sheet_selected}**.")
                    progress_bar.progress((idx+1)/len(sheet_names))
                    continue

                duplicates = find_duplicates(numbers_12digit)
                if duplicates:
                    st.warning(f"⚠️ Sheet **{sheet_selected}** ditemukan angka duplikat: {', '.join(sorted(duplicates))}")
                else:
                    st.success(f"✅ Sheet **{sheet_selected}** tidak mengandung duplikat angka.")

                num_files = math.ceil(total_numbers / batch_size)
                # ------------- SMART MATCHING BULK -------------
                bulk_text = cari_bulk(sheet_selected, bulk_db)
                # Debug
                # st.write(f"Sheet: {sheet_selected} | Key: {extract_key(sheet_selected)} | Bulk: {bulk_text}")

                files_buffers = []
                for i in range(num_files):
                    batch_numbers = numbers_12digit[i*batch_size:(i+1)*batch_size]
                    qty = len(batch_numbers)
                    file_index = i + 1

                    filename = buat_nama_file(file_index, sheet_selected, qty, bulk_text)
                    buffer = io.BytesIO()
                    buffer.write('\n'.join(batch_numbers).encode('utf-8'))
                    buffer.seek(0)
                    files_buffers.append((filename, buffer))

                processed[sheet_selected] = files_buffers
                progress_bar.progress((idx+1)/len(sheet_names))

            st.session_state.processed_sheets = processed
            st.success("🎉 Semua sheet selesai diproses.")

    if 'processed_sheets' in st.session_state and st.session_state.processed_sheets:
        st.subheader("📋 Sheet yang sudah diproses:")
        sheet_list = list(st.session_state.processed_sheets.keys())
        st.markdown('<div style="display:flex; overflow-x:auto; white-space:nowrap; padding:10px 0; border:1px solid #ddd; border-radius:8px;">', unsafe_allow_html=True)
        selected_sheets = []
        for sheet in sheet_list:
            key = f"cb_{sheet}"
            checked = st.checkbox(sheet, key=key)
            if checked:
                selected_sheets.append(sheet)
            st.markdown('<div style="flex:0 0 auto; margin-right:25px; padding:5px 10px; background:#f5f5f5; border-radius:6px;"></div>', unsafe_allow_html=True)
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

                zip_filename = (
                    "all_sheet.zip"
                    if len(selected_sheets) == len(st.session_state.processed_sheets)
                    else f"hasil_csv_{'_'.join([s.replace(' ', '_') for s in selected_sheets])}.zip"
                )

                st.download_button(
                    label="📥 Download ZIP",
                    data=zip_buffer,
                    file_name=zip_filename,
                    mime="application/zip"
                )
else:
    st.info("Silakan upload file Excel voucher (xlsx) terlebih dahulu.")

# Footer
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
