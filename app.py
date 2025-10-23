# app.py
import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
import io
import zipfile

def generate_and_zip_pdfs(pdf_file, excel_file, sheet_name, filename_col):
    """
    Reads a multi-page PDF and an Excel sheet, then generates a zip file
    containing individually named, single-page PDFs.
    """
    try:
        # --- 1. Read the Excel File ---
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        if filename_col not in df.columns:
            st.error(f"Error: Column '{filename_col}' not found in the sheet '{sheet_name}'.")
            st.error(f"Available columns are: {df.columns.tolist()}")
            return None
        
        filenames = df[filename_col].dropna().tolist()
        
        # --- 2. Read the PDF File ---
        pdf_reader = PdfReader(pdf_file)
        num_pdf_pages = len(pdf_reader.pages)
        
        if num_pdf_pages < len(filenames):
            st.warning(f"Warning: The Excel file lists {len(filenames)} certificates, but the PDF only has {num_pdf_pages} pages. Some certificates will not be generated.")
        
        st.info(f"Preparing to generate {min(len(filenames), num_pdf_pages)} certificates...")
        
        # --- 3. Create a Zip File in Memory ---
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f:
            progress_bar = st.progress(0)
            for i, filename in enumerate(filenames):
                page_num = i + 1
                if page_num > num_pdf_pages:
                    break
                
                if not filename.lower().endswith('.pdf'):
                    filename += '.pdf'
                
                st.write(f"Processing page {page_num}: {filename}")

                pdf_writer = PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[i])
                
                pdf_output_buffer = io.BytesIO()
                pdf_writer.write(pdf_output_buffer)
                pdf_output_buffer.seek(0)
                
                zip_f.writestr(filename, pdf_output_buffer.getvalue())
                
                progress_bar.progress((i + 1) / len(filenames))

        zip_buffer.seek(0)
        return zip_buffer

    except ValueError:
        st.error(f"Error: Sheet '{sheet_name}' not found in the Excel file. This should not happen with the dropdown selection.")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        return None

# --- STREAMLIT USER INTERFACE ---
st.set_page_config(layout="wide")
st.title("Automated Certificate Generator")
st.markdown("""
This tool generates individual PDF certificates from a multi-page master PDF and an Excel data sheet.
**Workflow:**
1.  Upload your multi-page certificate master file (saved as a **PDF**).
2.  Upload your Excel data file.
3.  Use the dropdowns to select the correct sheet and filename column.
4.  Click "Generate Certificates" and download the resulting `.zip` file.
""")

st.header("1. Upload Files")

col1, col2 = st.columns(2)
with col1:
    pdf_template_file = st.file_uploader("Upload Multi-Page Certificate PDF", type=['pdf'])
with col2:
    excel_data_file = st.file_uploader("Upload Excel Data Source", type=['xlsx', 'xls'])

st.header("2. Configure Settings")

col3, col4 = st.columns(2)

# --- START OF THE NEW DYNAMIC DROPDOWN LOGIC ---
if excel_data_file is not None:
    try:
        xls = pd.ExcelFile(excel_data_file)
        sheet_names = xls.sheet_names
        
        with col3:
            sheet_name_input = st.selectbox("Select the Excel Sheet", sheet_names, help="The app automatically reads the sheet names from your file.")
        
        if sheet_name_input:
            df_sheet = pd.read_excel(excel_data_file, sheet_name=sheet_name_input)
            column_names = df_sheet.columns.tolist()
            with col4:
                filename_column_input = st.selectbox("Select the Filename Column", column_names, help="The app automatically reads the column headers from your selected sheet.")
        else:
            with col4:
                st.selectbox("Select the Filename Column", ["Select a sheet first"], disabled=True)
                filename_column_input = None
                
    except Exception as e:
        st.error(f"Could not read the Excel file. Please ensure it is a valid .xlsx or .xls file. Error: {e}")
        sheet_name_input = None
        filename_column_input = None
else:
    # Display disabled dropdowns if no file is uploaded
    with col3:
        st.selectbox("Select the Excel Sheet", ["Upload an Excel file first"], disabled=True)
        sheet_name_input = None
    with col4:
        st.selectbox("Select the Filename Column", ["Upload an Excel file first"], disabled=True)
        filename_column_input = None
# --- END OF THE NEW DYNAMIC DROPDOWN LOGIC ---

st.header("3. Generate and Download")

if st.button("Generate Certificates"):
    if not all([pdf_template_file, excel_data_file, sheet_name_input, filename_column_input]):
        st.warning("Please fill in all the fields and upload both files before generating.")
    else:
        with st.expander("Processing Log", expanded=True):
            zip_file_buffer = generate_and_zip_pdfs(
                pdf_template_file,
                excel_data_file,
                sheet_name_input,
                filename_column_input
            )
        
        if zip_file_buffer:
            st.success("Certificate generation successful!")
            st.download_button(
                label="Download Certificates (.zip)",
                data=zip_file_buffer,
                file_name="generated_certificates.zip",
                mime="application/zip"
            )
