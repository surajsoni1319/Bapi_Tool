import streamlit as st
import fitz  # PyMuPDF
import re
from io import BytesIO
import pytesseract
from PIL import Image
import pandas as pd

# --------------------------------------------------
# Page Config
# --------------------------------------------------
st.set_page_config(
    page_title="PDF to Excel Extractor (Row-wise)",
    layout="wide"
)

# --------------------------------------------------
# PDF Text Extraction
# --------------------------------------------------
def extract_text_from_pdf(pdf_bytes):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text
    except Exception as e:
        st.error(f"Text extraction failed: {e}")
        return None


def extract_text_with_ocr(pdf_bytes):
    try:
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
        st.error(f"OCR failed: {e}")
        return None


# --------------------------------------------------
# Utility
# --------------------------------------------------
def normalize(txt):
    return re.sub(r"[^a-z0-9]", "", txt.lower())


def parse_customer_data(text):
    fields = [
        "Type of Customer", "Name of Customer", "Company Code",
        "Customer Group", "Sales Group", "Region", "Zone", "Sub Zone",
        "State", "Sales Office", "Mobile Number", "Whatsapp No.",
        "E-Mail ID", "Address 1", "Address 2", "Address 3", "Address 4",
        "City", "District", "PIN", "GSTIN", "PAN",
        "PAN Holder Name", "PAN Status",
        "IFSC Number", "Account Number", "Bank Name", "Bank Branch",
        "Aadhaar Number", "Gender", "DOB",
        "Credit Limit (In Rs.)", "Sales Officer to be mapped",
        "Initiator Name", "Initiator Email ID", "Final Result"
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


# --------------------------------------------------
# Excel Creation (STRICT ROW-WISE)
# --------------------------------------------------
def create_excel_rowwise_only(all_data: dict):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for filename, record in all_data.items():

            # 🔑 THIS is the key line (dict → rows)
            df = pd.DataFrame(
                list(record.items()),
                columns=["Field Name", "Value"]
            )

            sheet_name = filename.replace(".pdf", "")[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.sheets[sheet_name]
            ws.column_dimensions["A"].width = 45
            ws.column_dimensions["B"].width = 60

    output.seek(0)
    return output


# --------------------------------------------------
# UI
# --------------------------------------------------
st.title("📄 PDF to Excel Extractor (Row-wise Only)")
st.write("Extract customer data from PDFs and download **row-wise Excel output**.")

with st.sidebar:
    st.header("⚙️ Options")
    use_ocr = st.checkbox("Enable OCR for scanned PDFs", value=False)

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
        extracted_data = {}
        progress = st.progress(0)

        for idx, file in enumerate(uploaded_files):
            pdf_bytes = file.read()

            text = extract_text_from_pdf(pdf_bytes)

            if (not text or len(text.strip()) < 100) and use_ocr:
                text = extract_text_with_ocr(pdf_bytes)

            if text:
                parsed = parse_customer_data(text)
                if parsed:
                    extracted_data[file.name] = parsed
                else:
                    st.warning(f"No data found in {file.name}")
            else:
                st.error(f"Failed to read {file.name}")

            progress.progress((idx + 1) / len(uploaded_files))

        # --------------------------------------------------
        # Preview
        # --------------------------------------------------
        if extracted_data:
            st.success("✅ Extraction completed")

            for fname, record in extracted_data.items():
                with st.expander(f"📄 {fname}", expanded=len(extracted_data) == 1):
                    st.dataframe(
                        pd.DataFrame(list(record.items()), columns=["Field Name", "Value"]),
                        use_container_width=True,
                        hide_index=True
                    )

            excel_file = create_excel_rowwise_only(extracted_data)

            st.download_button(
                label="📥 Download Excel (Row-wise)",
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
