import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
from io import BytesIO
import pytesseract
from PIL import Image

# --------------------------------------------------
# Page Config
# --------------------------------------------------
st.set_page_config(
    page_title="PDF to Excel Extractor (Row-wise)",
    layout="wide"
)

# --------------------------------------------------
# Helper Functions
# --------------------------------------------------
def extract_text_from_pdf(pdf_bytes):
    """Extract text from PDF using PyMuPDF"""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text
    except Exception as e:
        st.error(f"Text extraction error: {e}")
        return None


def extract_text_with_ocr(pdf_bytes):
    """OCR fallback using Tesseract"""
    try:
        # Check Tesseract availability
        pytesseract.get_tesseract_version()

        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = ""

        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            text += pytesseract.image_to_string(img)

        doc.close()
        return text
    except Exception as e:
        st.error(f"OCR error: {e}")
        return None


def normalize(text):
    return re.sub(r'[^a-z0-9]', '', text.lower())


def parse_customer_data(text):
    """Parse extracted text into key-value pairs"""
    fields = [
        "Type of Customer", "Name of Customer", "Company Code", "Customer Group",
        "Sales Group", "Region", "Zone", "Sub Zone", "State", "Sales Office",
        "SAP Dealer code to be mapped", "Search Term",
        "Search Term 1- Old customer code", "Search Term 2 - District",
        "Mobile Number", "E-Mail ID", "Lattitude", "Longitude",
        "Address 1", "Address 2", "Address 3", "Address 4",
        "PIN", "City", "District", "Whatsapp No.",
        "Date of Birth", "Date of Anniversary",
        "Is GST Present", "GSTIN", "Trade Name", "Legal Name",
        "PAN", "PAN Holder Name", "PAN Status",
        "IFSC Number", "Account Number", "Bank Name", "Bank Branch",
        "Aadhaar Number", "Gender", "DOB",
        "Logistics Transportation Zone", "Transportation Zone Code",
        "Date of Appointment", "Delivering Plant", "Plant Name",
        "Credit Limit (In Rs.)", "Sales Officer to be mapped",
        "Initiator Name", "Initiator Email ID",
        "Final Result"
    ]

    data = {}
    lines = text.split("\n")

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        for field in fields:
            if normalize(line).startswith(normalize(field)):
                value = line[len(field):].strip(" :-")
                if not value and i + 1 < len(lines):
                    value = lines[i + 1].strip()
                data[field] = value
                break

    return data


def create_excel_rowwise(dataframes_dict):
    """Create Excel with ONLY row-wise format"""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for filename, df in dataframes_dict.items():
            rows = []

            for col in df.columns:
                value = df[col].iloc[0]
                rows.append([col, "" if pd.isna(value) else value])

            df_row = pd.DataFrame(rows, columns=["Field Name", "Value"])
            sheet_name = filename.replace(".pdf", "")[:31]
            df_row.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.sheets[sheet_name]
            ws.column_dimensions["A"].width = 45
            ws.column_dimensions["B"].width = 55

    output.seek(0)
    return output


def display_rowwise(data_dict):
    return pd.DataFrame(
        [{"Field Name": k, "Value": v} for k, v in data_dict.items()]
    )

# --------------------------------------------------
# UI
# --------------------------------------------------
st.title("📄 PDF to Excel Extractor (Row-wise)")
st.write("Extract customer creation form data and download it in **row-wise Excel format**.")

with st.sidebar:
    st.header("⚙️ Options")
    use_ocr = st.checkbox("Enable OCR (for scanned PDFs)", value=False)

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# --------------------------------------------------
# Processing
# --------------------------------------------------
if uploaded_files:
    st.success(f"{len(uploaded_files)} file(s) uploaded")

    if st.button("🚀 Extract Data", type="primary"):
        dataframes_dict = {}
        preview_data = {}

        progress = st.progress(0)

        for idx, file in enumerate(uploaded_files):
            pdf_bytes = file.read()

            text = extract_text_from_pdf(pdf_bytes)

            if (not text or len(text.strip()) < 100) and use_ocr:
                text = extract_text_with_ocr(pdf_bytes)

            if text:
                parsed = parse_customer_data(text)
                if parsed:
                    df = pd.DataFrame([parsed])
                    dataframes_dict[file.name] = df
                    preview_data[file.name] = display_rowwise(parsed)
                else:
                    st.warning(f"No data found in {file.name}")
            else:
                st.error(f"Failed to read {file.name}")

            progress.progress((idx + 1) / len(uploaded_files))

        # --------------------------------------------------
        # Display Preview
        # --------------------------------------------------
        if dataframes_dict:
            st.success("✅ Extraction completed")

            for fname, df_preview in preview_data.items():
                with st.expander(f"📄 {fname}", expanded=len(preview_data) == 1):
                    st.dataframe(df_preview, use_container_width=True, hide_index=True)

            excel_file = create_excel_rowwise(dataframes_dict)

            st.download_button(
                "📥 Download Excel (Row-wise)",
                data=excel_file,
                file_name="customer_data_rowwise.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        else:
            st.error("No data could be extracted.")

# --------------------------------------------------
# Footer
# --------------------------------------------------
st.markdown("---")
st.markdown("Built with Streamlit | PyMuPDF | Tesseract OCR")
