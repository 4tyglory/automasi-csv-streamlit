import streamlit as st
import pandas as pd
import re
import math
import io
import zipfile

def format_decimal_with_koma(s: str) -> str:
    return s.lower().replace('.', 'koma')

def parse_sheet_name(sheet_name: str):
    kuota = None
    hari = None
    zona = None

    kuota_match = re.search(r'(\d+[\.,]?\d*)\s*gb', sheet_name, re.IGNORECASE)
    if kuota_match:
        kuota = kuota_match.group(1).replace(',', '.').replace(' ', '')

    hari_match = re.search(r'(\d+)\s*h', sheet_name, re.IGNORECASE)
    if hari_match:
        hari = hari_match.group(1)

    zona_match = re.search(r'(z\d+)', sheet_name, re.IGNORECASE)
    if zona_match:
        zona = zona_match.group(1).lower()

    return kuota, hari, zona

def buat_nama_file(file_index, sheet_name, qty, bulk_text):
    kuota, hari, zona = parse_sheet_name(sheet_name)
    if kuota:
        kuota_text = format_decimal_with_koma(f"{kuota}gb")
    else:
        kuota_text = "unknowngb"

    if bulk_text:
        bulk_text = bulk_text.replace(" ", "")
        bulk_text = format_decimal_with_koma(bulk_text)
    else:
        bulk_text = "bulkUnknown"

    hari_text = f"{hari}hari" if hari else "unknownhari"

    if zona:
        filename = f"{file_index} vcr fisik internet {hari_text} {kuota_text} {bulk_text} {zona} {qty}.csv"
    else:
        filename = f"{file_index} vcr fisik internet {hari_text} {kuota_text} {bulk_text} {qty}.csv"

    return filename

st.title("Automasi CSV Multi Sheet dengan Penamaan File Khusus")

uploaded_excel = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_excel:
    xls = pd.ExcelFile(uploaded_excel)
    sheet_names = xls.sheet_names

    batch_size = 1000

    if 'processed_sheets' not in st.session_state:
        st.session_state.processed_sheets = {}

    # Simulasi database bulk per sheet, bisa diganti sesuai database asli
    contoh_database_bulk = {
        "2.5 GB 5H Z3": "bulk10.9K",
        "3 GB 5H Z3": "bulk12K",
        "4 GB 7H Z3": "bulk18K",
        # tambah sesuai kebutuhan
    }

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

                bulk_text = contoh_database_bulk.get(sheet_selected, "bulkUnknown")

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

                if len(selected_sheets) == len(st.session_state.processed_sheets):
                    zip_filename = "all sheet.zip"
                else:
                    zip_filename = f"hasil_csv_{'_'.join([s.replace(' ', '_') for s in selected_sheets])}.zip"

                st.download_button(
                    label="📥 Download ZIP",
                    data=zip_buffer,
                    file_name=zip_filename,
                    mime="application/zip"
                )
else:
    st.info("Silakan upload file Excel terlebih dahulu.")
