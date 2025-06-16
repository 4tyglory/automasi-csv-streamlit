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
    uploaded_csv = st.file_uploader(
        "Upload file CSV hasil proses untuk validasi",
        type=["csv"],
        accept_multiple_files=True
    )

    if uploaded_csv:
        validation_results = {}

        for csv_file in uploaded_csv:
            filename = csv_file.name

            try:
                df_csv = pd.read_csv(csv_file, header=None)
                csv_numbers = df_csv[0].astype(str).tolist()
            except Exception as e:
                validation_results[filename] = {
                    "valid": False,
                    "reason": f"Gagal baca file CSV: {e}"
                }
                continue

            invalid_numbers = [num for num in csv_numbers if not re.fullmatch(r'\d{12}', num)]

            seen = set()
            duplicates = set()
            for num in csv_numbers:
                if num in seen:
                    duplicates.add(num)
                else:
                    seen.add(num)

            if invalid_numbers:
                reason = f"Ada angka tidak valid (bukan 12 digit): {', '.join(invalid_numbers)}"
                valid = False
            elif duplicates:
                reason = f"Ada angka duplikat: {', '.join(duplicates)}"
                valid = False
            else:
                reason = "Valid: Semua angka 12 digit dan tanpa duplikat"
                valid = True

            validation_results[filename] = {"valid": valid, "reason": reason}

        st.write("### Hasil Validasi CSV")
        for fname, res in validation_results.items():
            if res["valid"]:
                st.success(f"✅ {fname}: {res['reason']}")
            else:
                st.error(f"⚠️ {fname}: {res['reason']}")
    else:
        st.info("Silakan upload minimal satu file CSV untuk validasi.")

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
