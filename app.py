import streamlit as st
import fitz
import pandas as pd
from io import BytesIO

# -----------------------------------------
# PAGE CONFIG
# -----------------------------------------
st.set_page_config(
    page_title="ZSUB PDF Data Extractor",
    page_icon="📄",
    layout="wide"
)

st.title("📄 ZSUB Customer PDF Extractor (Fixed)")
st.caption("Template-driven | Line-based | No OCR")

st.markdown("---")

# -----------------------------------------
# FIELD LABELS (EXACT PDF LABELS)
# -----------------------------------------
FIELDS = [
    "Type of Customer",
    "Name of Customer",
    "Company Code",
    "Customer Group",
    "Sales Group",
    "Region",
    "Zone",
    "Sub Zone",
    "State",
    "Sales Office",
    "SAP Dealer code to be mapped Search Term 2",
    "Search Term 1- Old customer code",
    "Search Term 2 - District",
    "Mobile Number",
    "E-Mail ID",
    "Lattitude",
    "Longitude",
    "Name of the Customers (Trade Name or Legal Name)",
    "Address 1",
    "Address 2",
    "Address 3",
    "Address 4",
    "PIN",
    "City",
    "District",
    "Whatsapp No.",
    "Date of Birth",
    "Date of Anniversary",
    "Counter Potential - Maximum",
    "Counter Potential - Minimum",
    "Is GST Present",
    "GSTIN",
    "Trade Name",
    "Legal Name",
    "PAN",
    "PAN Holder Name",
    "PAN Status",
    "IFSC Number",
    "Account Number",
    "Name of Account Holder",
    "Bank Name",
    "Bank Branch",
    "Is Aadhaar Linked with Mobile?",
    "Aadhaar Number",
    "Gender",
    "DOB",
    "Logistics Transportation Zone",
    "Transportation Zone Description",
    "Transportation Zone Code",
    "Date of Appointment",
    "Plant Name",
    "Plant Code",
    "Incoterns",
    "Incoterns Code",
    "Initiator Name",
    "Initiator Email ID",
    "Initiator Mobile Number",
    "Final Result"
]

# -----------------------------------------
# FUNCTIONS
# -----------------------------------------

def extract_text_lines(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    return lines


def extract_fields_from_lines(lines):
    data = {field: "" for field in FIELDS}

    for i, line in enumerate(lines):
        for field in FIELDS:
            if line == field:
                # Take next non-empty line as value
                if i + 1 < len(lines):
                    data[field] = lines[i + 1]
    return data


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# -----------------------------------------
# UI
# -----------------------------------------
uploaded_files = st.file_uploader(
    "📤 Upload ZSUB PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} PDF(s) uploaded")

    if st.button("🚀 Extract Data", use_container_width=True):
        rows = []

        with st.spinner("Extracting data..."):
            for file in uploaded_files:
                lines = extract_text_lines(file)
                record = extract_fields_from_lines(lines)
                record["Source File"] = file.name
                rows.append(record)

        df = pd.DataFrame(rows)

        st.subheader("📊 Extracted Data Preview")
        st.dataframe(df, use_container_width=True)

        st.download_button(
            "⬇️ Download Excel",
            data=to_excel(df),
            file_name="zsub_extracted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload a ZSUB PDF to begin.")
