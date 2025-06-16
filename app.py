import streamlit as st
import pandas as pd
import re
import math
import io
import zipfile

def extract_12digit_numbers(df):
    numbers = []
    for _, row in df.iterrows():
        for val in row.astype(str):
            found = re.findall(r'\b\d{12}\b', val)
            if found:
                numbers.extend(found)
    return numbers

def check_duplicates(lst):
    seen = set()
    duplicates = set()
    for x in lst:
        if x in seen:
            duplicates.add(x)
        else:
            seen.add(x)
    return duplicates

def normalize_name(name: str) -> str:
    name = name.lower()
    name = name.replace(" ", "")
    name = name.replace(".", "koma")
    name = name.replace("-", "")
    name = name.replace("_", "")
    # Tambah penggantian lain jika perlu
    return name

st.set_page_config(page_title="Automation & Validation CSV", layout="wide")
st.sidebar.title("Menu")
menu = st.sidebar.radio("Pilih menu:", ["Automasi CSV", "Validasi CSV"])

if menu == "Automasi CSV":
    st.title("Automasi CSV")
    uploaded_excel = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])
    if uploaded_excel:
        xls = pd.ExcelFile(uploaded_excel)
        sheet_names = xls.sheet_names
        st.write(f"File Excel berisi sheet: {sheet_names}")

        batch_size = 1000

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

                    files_buffers = []
                    for i in range(num_files):
                        batch_numbers = numbers_12digit[i*batch_size:(i+1)*batch_size]
                        qty = len(batch_numbers)
                        file_index = i + 1

                        filename = f"{file_index}_{sheet_selected}_{qty}.csv"

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

elif menu == "Validasi CSV":
    st.title("Validasi CSV")

    uploaded_excel = st.file_uploader("Upload file Excel asli (.xlsx) untuk validasi", type=["xlsx"])
    uploaded_csv = st.file_uploader("Upload file CSV hasil proses untuk validasi", type=["csv"], accept_multiple_files=True)

    if uploaded_excel and uploaded_csv:
        try:
            xls = pd.ExcelFile(uploaded_excel)
            excel_sheets = xls.sheet_names
        except Exception as e:
            st.error(f"Error baca file Excel: {e}")
            st.stop()

        validation_results = {}

        for csv_file in uploaded_csv:
            filename = csv_file.name.lower()
            csv_name_norm = normalize_name(filename)

            matched_sheet = None
            for sheet in excel_sheets:
                sheet_norm = normalize_name(sheet)
                # Debug output nama normalisasi
                st.write(f"Debug: normalisasi CSV '{csv_name_norm}' vs Sheet '{sheet_norm}'")
                if sheet_norm in csv_name_norm:
                    matched_sheet = sheet
                    break

            if not matched_sheet:
                validation_results[filename] = {"valid": False, "reason": "Sheet terkait tidak ditemukan di Excel"}
                continue

            try:
                df_csv = pd.read_csv(csv_file, header=None)
                df_excel = pd.read_excel(uploaded_excel, sheet_name=matched_sheet, header=None)
            except Exception as e:
                validation_results[filename] = {"valid": False, "reason": f"Gagal baca file: {e}"}
                continue

            excel_numbers = set(extract_12digit_numbers(df_excel))
            csv_numbers = df_csv[0].astype(str).tolist()

            not_in_excel = [num for num in csv_numbers if num not in excel_numbers]
            duplicates = check_duplicates(csv_numbers)

            angka_valid = (len(not_in_excel) == 0)
            duplikat_valid = (len(duplicates) == 0)

            reason = []
            if not angka_valid:
                reason.append(f"Ada angka di CSV yang tidak ada di Excel: {not_in_excel}")
            if not duplikat_valid:
                reason.append(f"Ada duplikat angka di CSV: {list(duplicates)}")

            validation_results[matched_sheet] = {
                "valid": angka_valid and duplikat_valid,
                "reason": "; ".join(reason) if reason else "Valid"
            }

        st.write("### Hasil Validasi")
        for sheet_name, res in validation_results.items():
            if res["valid"]:
                st.success(f"✅ {sheet_name}: {res['reason']}")
            else:
                st.error(f"❌ {sheet_name}: {res['reason']}")
    else:
        st.info("Silakan upload file Excel asli dan minimal satu file CSV hasil proses.")

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
