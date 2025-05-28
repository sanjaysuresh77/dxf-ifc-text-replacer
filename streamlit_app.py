import os
import zipfile
import tempfile
import streamlit as st
import openpyxl
import xlrd

def replace_text_advanced(line, replacements):
    if line.strip() in replacements:
        return replacements[line.strip()]
    for key, val in replacements.items():
        if key in line:
            line = line.replace(key, val)
    return line

def load_replacements_from_excel(uploaded_file):
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    replacements = {}

    if ext == ".xlsx":
        wb = openpyxl.load_workbook(uploaded_file)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:
                replacements[str(row[0]).strip()] = str(row[1]).strip()

    elif ext == ".xls":
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name
        wb = xlrd.open_workbook(tmp_path)
        sheet = wb.sheet_by_index(0)
        for i in range(1, sheet.nrows):
            original, replace = sheet.row_values(i)[:2]
            if original and replace:
                replacements[str(original).strip()] = str(replace).strip()

    return replacements

def main():
    st.title("IFC DESIGN TOOL")
    st.write("Upload an Excel file and a ZIP file of DXF drawings.")

    excel_file = st.file_uploader("Upload Excel File (.xls or .xlsx)", type=["xls", "xlsx"])
    zip_file = st.file_uploader("Upload stack of typical Files", type=["zip"])

    if st.button("Generate IFC"):
        if not excel_file or not zip_file:
            st.warning("Please upload both Excel and ZIP files.")
            return

        with st.spinner("Processing..."):
            try:
                # Extract ZIP
                temp_dir = tempfile.mkdtemp()
                zip_input_path = os.path.join(temp_dir, "input.zip")
                with open(zip_input_path, "wb") as f:
                    f.write(zip_file.read())

                with zipfile.ZipFile(zip_input_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)

                # Load Excel Replacements
                replacements = load_replacements_from_excel(excel_file)

                updated_files = []
                for fname in os.listdir(temp_dir):
                    if fname.lower().endswith(".dxf"):
                        in_path = os.path.join(temp_dir, fname)
                        out_path = os.path.join(temp_dir, f"updated_{fname}")

                        with open(in_path, "r", encoding="latin1") as f:
                            lines = f.read().splitlines()

                        new_lines = [replace_text_advanced(line, replacements) for line in lines]

                        with open(out_path, "w", encoding="latin1", newline="") as f:
                            f.write("\r\n".join(new_lines) + "\r\n")

                        updated_files.append(out_path)

                # Create output ZIP
                output_zip_path = os.path.join(temp_dir, "updated_typical.zip")
                with zipfile.ZipFile(output_zip_path, "w") as zipf:
                    for file in updated_files:
                        zipf.write(file, arcname=os.path.basename(file))

                with open(output_zip_path, "rb") as f:
                    st.success("IFC successfully generated!")
                    st.download_button("Download Circuits", f, file_name="Updated IFC.zip")

            except Exception as e:
                st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
