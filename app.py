import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(
    page_title="ZSUB PDF Data Extractor",
    page_icon="📄",
    layout="wide"
)

st.title("📄 ZSUB Customer PDF Data Extractor")
st.caption("Template-Driven | No OCR | High Accuracy")

st.markdown("---")

# --------------------------------------------------
# FIELD → REGEX MAPPING
# --------------------------------------------------

FIELD_PATTERNS = {
    "Type of Customer": r"Type of Customer\s+(.*)",
    "Name of Customer": r"Name of Customer\s+(.*)",
    "Company Code": r"Company Code\s+(.*)",
    "Customer Group": r"Customer Group\s+(.*)",
    "Sales Group": r"Sales Group\s+(.*)",
    "Region": r"Region\s+(.*)",
    "Zone": r"Zone\s+(.*)",
    "Sub Zone": r"Sub Zone\s+(.*)",
    "State": r"State\s+(.*)",
    "Sales Office": r"Sales Office\s+(.*)",

    "SAP Dealer Code": r"SAP Dealer code.*?\s+(\d+)",
    "Old Customer Code": r"Search Term 1- Old customer code\s*(.*)",
    "Search Term District": r"Search Term 2 - District\s*(.*)",

    "Mobile Number": r"Mobile Number\s+(\d{10})",
    "Email": r"E-Mail ID\s*([\w\.-]+@[\w\.-]+)?",

    "Latitude": r"Lattitude\s+([\d\.]+)",
    "Longitude": r"Longitude\s+([\d\.]+)",

    "Trade / Legal Name": r"Name of the Customers.*?\s+(.*)",

    "Address 1": r"Address 1\s+(.*)",
    "Address 2": r"Address 2\s+(.*)",
    "Address 3": r"Address 3\s+(.*)",
    "Address 4": r"Address 4\s+(.*)",

    "PIN": r"PIN\s+(\d{6})",
    "City": r"City\s+(.*)",
    "District": r"District\s+(.*)",
    "Whatsapp No": r"Whatsapp No\.\s*(\d{10})?",

    "Date of Birth": r"Date of Birth\s+([\d\-\/]+)",
    "Date of Anniversary": r"Date of Anniversary\s+([\d\-\/]+)?",

    "Is GST Present": r"Is GST Present\s+(Yes|No)",
    "GSTIN": r"GSTIN\s+([A-Z0-9]{15})?",

    "PAN": r"PAN\s+([A-Z0-9]{10})",
    "PAN Holder Name": r"PAN Holder Name\s+(.*)",
    "PAN Status": r"PAN Status\s+(.*)",

    "IFSC Number": r"IFSC Number\s+([A-Z0-9]+)",
    "Account Number": r"Account Number\s+(\d+)",
    "Account Holder": r"Name of Account Holder\s+(.*)",
    "Bank Name": r"Bank Name\s+(.*)",
    "Bank Branch": r"Bank Branch\s+(.*)",

    "Aadhaar Linked with Mobile": r"Is Aadhaar Linked with Mobile\?\s+(Yes|No)",
    "Aadhaar Number": r"Aadhaar Number\s+([X\d\s]+)",

    "Gender": r"Gender\s+(.*)",

    "Logistics Transportation Zone": r"Logistics Transportation Zone\s+(.*)",
    "Transportation Zone Code": r"Transportation Zone Code\s+(\d+)",

    "Plant Name": r"Plant Name\s+(.*)",
    "Plant Code": r"Plant Code\s+(.*)",

    "Incoterms": r"Incoterns\s+(.*)",
    "Incoterms Code": r"Incoterns Code\s+(.*)",

    "Security Deposit": r"Dealer\s+(\d+)",
    "Credit Limit": r"Credit Limit \(In Rs\.?\)\s*(\d+)?",

    "Initiator Name": r"Initiator Name\s+(.*)",
    "Initiator Email": r"Initiator Email ID\s+([\w\.-]+@[\w\.-]+)",
    "Initiator Mobile": r"Initiator Mobile Number\s+(\d{10})",

    "Final Result": r"Final Result\s+(True|False)"
}

# --------------------------------------------------
# FUNCTIONS
# --------------------------------------------------

def extract_text_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    return re.sub(r"\s+", " ", text)


def extract_fields(text):
    record = {}
    for field, pattern in FIELD_PATTERNS.items():
        match = re.search(pattern, text, re.IGNORECASE)
        record[field] = match.group(1).strip() if match else ""
    return record


def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --------------------------------------------------
# UI
# --------------------------------------------------

uploaded_files = st.file_uploader(
    "📤 Upload ZSUB PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded")

    if st.button("🚀 Extract Data", use_container_width=True):
        rows = []

        with st.spinner("Extracting data from PDFs..."):
            for file in uploaded_files:
                text = extract_text_from_pdf(file)
                data = extract_fields(text)
                data["Source File"] = file.name
                rows.append(data)

        df = pd.DataFrame(rows)

        st.subheader("📊 Extracted Data Preview")
        st.dataframe(df, use_container_width=True)

        excel_data = convert_df_to_excel(df)

        st.download_button(
            label="⬇️ Download Excel",
            data=excel_data,
            file_name="zsub_extracted_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload one or more ZSUB PDF files to begin.")
