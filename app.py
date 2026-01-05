import streamlit as st
import fitz
import pandas as pd
import re
from io import BytesIO

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(page_title="ZSUB PDF Extractor", layout="wide")
st.title("📄 ZSUB PDF Data Extractor (Final – Section Aware)")
st.caption("No OCR | SAP-safe | Production Ready")

# -------------------------------------------------
# SECTION → FIELD MAPPING (ORDER MATTERS)
# -------------------------------------------------
SECTIONS = {
    "Customer Creation": [
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
        "Longitude"
    ],

    "Customer Address": [
        "Name of the Customers (Trade Name or Legal Name)",
        "Mobile Number",
        "E-mail",
        "Address 1",
        "Address 2",
        "Address 3",
        "Address 4",
        "PIN",
        "City",
        "District",
        "State",
        "Whatsapp No.",
        "Date of Birth",
        "Date of Anniversary",
        "Counter Potential - Maximum",
        "Counter Potential - Minimum"
    ],

    "GSTIN Details": [
        "Is GST Present",
        "GSTIN",
        "Trade Name",
        "Legal Name",
        "Reg Date",
        "City",
        "Type",
        "Building No.",
        "District Code",
        "State Code",
        "Street",
        "PIN Code"
    ],

    "PAN Details": [
        "PAN",
        "PAN Holder Name",
        "PAN Status",
        "PAN - Aadhaar Linking Status"
    ],

    "Bank Details": [
        "IFSC Number",
        "Account Number",
        "Name of Account Holder",
        "Bank Name",
        "Bank Branch"
    ],

    "Aadhaar Details": [
        "Is Aadhaar Linked with Mobile?",
        "Aadhaar Number",
        "Name",
        "Gender",
        "DOB",
        "Address",
        "PIN",
        "City",
        "State"
    ],

    "Zone Details": [
        "Logistics Transportation Zone",
        "Transportation Zone Description",
        "Transportation Zone Code",
        "Postal Code",
        "Logistics team to vet the T zone selected by Sales Officer",
        "Selection of Available T Zones from T Zone Master list, if found.",
        "If NEW T Zone need to be created, details to be provided by Logistics team"
    ],

    "Other Details": [
        "Date of Appointment",
        "Delivering Plant",
        "Plant Name",
        "Plant Code",
        "Incoterns",
        "Incoterns Code",
        "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
        "Credit Limit (In Rs.)"
    ],

    # SALES HIERARCHY (critical missing section earlier)
    "Regional Head to be mapped": [
        "Regional Head to be mapped",
        "Zonal Head to be mapped",
        "Sub-Zonal Head (RSM) to be mapped",
        "Area Sales Manager to be mapped",
        "Sales Officer to be mapped",
        "Sales Promoter to be mapped",
        "Sales Promoter Number",
        "Internal control code",
        "SAP CODE"
    ],

    "Initiator Name": [
        "Initiator Name",
        "Initiator Email ID",
        "Initiator Mobile Number",
        "Created By Customer UserID",
        "Sales Head Name",
        "Sales Head Email",
        "Sales Head Mobile Number"
    ],

    "Duplicity Check": [
        "Extra2",
        "PAN Result",
        "Mobile Number Result",
        "Email Result",
        "GST Result",
        "Final Result"
    ]
}

# -------------------------------------------------
# HELPERS
# -------------------------------------------------
def normalize_text(text: str) -> str:
    text = re.sub(r"\s*\n\s*", " ", text)  # remove line breaks safely
    text = re.sub(r"\s+", " ", text)       # normalize spaces
    return text.strip()


def extract_text(pdf_file) -> str:
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    return normalize_text(text)


def slice_sections(text: str) -> dict:
    section_text = {}
    section_names = list(SECTIONS.keys())

    for i, section in enumerate(section_names):
        start = text.find(section)
        if start == -1:
            section_text[section] = ""
            continue

        start += len(section)

        if i + 1 < len(section_names):
            end = text.find(section_names[i + 1], start)
            section_text[section] = text[start:end] if end != -1 else text[start:]
        else:
            section_text[section] = text[start:]

    return section_text


def extract_fields_from_section(section_body: str, fields: list) -> dict:
    data = {}

    for i, field in enumerate(fields):
        # tolerate line breaks and multiple spaces in labels
        label = re.escape(field).replace("\\ ", "\\s+")

        if i + 1 < len(fields):
            next_label = re.escape(fields[i + 1]).replace("\\ ", "\\s+")
            pattern = rf"{label}\s*(.*?)\s*(?={next_label})"
        else:
            pattern = rf"{label}\s*(.*)"

        match = re.search(pattern, section_body, re.IGNORECASE | re.DOTALL)
        data[field] = match.group(1).strip(" :-") if match else ""

    return data


def extract_all_fields(text: str) -> dict:
    all_data = {}
    sections = slice_sections(text)

    for section, fields in SECTIONS.items():
        section_data = extract_fields_from_section(sections.get(section, ""), fields)
        all_data.update(section_data)

    return all_data


def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# -------------------------------------------------
# UI
# -------------------------------------------------
uploaded_files = st.file_uploader(
    "📤 Upload ZSUB PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files and st.button("🚀 Extract Data"):
    rows = []

    with st.spinner("Extracting data from PDFs..."):
        for file in uploaded_files:
            text = extract_text(file)
            record = extract_all_fields(text)
            record["Source File"] = file.name
            rows.append(record)

    df = pd.DataFrame(rows)

    st.subheader("📊 Extracted Data Preview")
    st.dataframe(df, use_container_width=True)

    st.download_button(
        "⬇️ Download Excel",
        data=to_excel(df),
        file_name="zsub_final_extracted_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
