def normalize_name(name: str) -> str:
    name = name.lower()
    name = name.replace(" ", "")
    name = name.replace(".", "koma")
    # Tambahkan replace lain jika perlu
    return name

# Dalam blok validasi CSV:

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
