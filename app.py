import streamlit as st
import fitz
import pandas as pd
import re
from io import BytesIO

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(page_title="ZSUB PDF Extractor", layout="wide")
st.title("📄 ZSUB PDF Data Extractor (Final Stable)")
st.caption("Section-aware • No OCR • SAP-ready")

# -------------------------------------------------
# GLOBAL CONTROLS
# -------------------------------------------------

STOP_WORDS = [
    "Supporting Document",
    "Zone Details",
    "Other Details",
    "PAN Details",
    "Bank Details",
    "Aadhaar Details"
]

EMPTY_ALLOWED_FIELDS = {
    "If NEW T Zone need to be created, details to be provided by Logistics team"
}

# ⚠️ ONLY ID / NUMBER FIELDS HERE
SINGLE_TOKEN_FIELDS = {
    "SAP Dealer code to be mapped Search Term 2",
    "Regional Head to be mapped",
    "Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped",
    "Area Sales Manager to be mapped",
    "Sales Officer to be mapped",
    "Sales Promoter Number",
    "Transportation Zone Code",
    "Plant Code",
    "IFSC Number",
    "Account Number",
    "Initiator Mobile Number",
    "Sales Head Mobile Number",
    "Mobile Number"
}

# -------------------------------------------------
# SECTION → FIELD MAPPING
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
        "Incoterns (1)",
        "Incoterns (2)",
        "Incoterns Code",
        "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
        "Credit Limit (In Rs.)"
    ],

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
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def clean_value(val: str) -> str:
    for w in STOP_WORDS:
        if w in val:
            val = val.split(w)[0]
    return val.strip()

def extract_text(pdf_file) -> str:
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    return normalize_text(text)

def slice_sections(text: str) -> dict:
    section_text = {}
    names = list(SECTIONS.keys())

    for i, section in enumerate(names):
        start = text.find(section)
        if start == -1:
            section_text[section] = ""
            continue

        start += len(section)

        if i + 1 < len(names):
            pattern = re.escape(names[i + 1]).replace("\\ ", "\\s+")
            m = re.search(pattern, text[start:], re.IGNORECASE)
            end = start + m.start() if m else -1
            section_text[section] = text[start:end] if end != -1 else text[start:]
        else:
            section_text[section] = text[start:]

    return section_text

def extract_fields_from_section(section_body: str, fields: list) -> dict:
    data = {}

    for i, field in enumerate(fields):
        value = ""

        base_field = field.replace("(1)", "").replace("(2)", "")
        label = re.escape(base_field).replace("\\ ", "\\s+")

        if i + 1 < len(fields):
            next_base = fields[i + 1].replace("(1)", "").replace("(2)", "")
            next_label = re.escape(next_base).replace("\\ ", "\\s+")
            pattern = rf"{label}\s*(.*?)\s*(?={next_label})"
        else:
            pattern = rf"{label}\s*(.*)"

        match = re.search(pattern, section_body, re.IGNORECASE | re.DOTALL)
        if match:
            value = match.group(1).strip(" :-")

        # Explicit empty override
        if field in EMPTY_ALLOWED_FIELDS:
            value = ""

        # Incoterms split
        if field.startswith("Incoterns"):
            parts = value.split("Incoterns")
            if field == "Incoterns (1)":
                value = parts[0].strip()
            elif field == "Incoterns (2)" and len(parts) > 1:
                value = parts[1].strip()

        # Trim ONLY numeric/code fields
        if field in SINGLE_TOKEN_FIELDS and value:
            value = value.split()[0]

        value = clean_value(value)
        data[field] = value

    return data

def extract_all_fields(text: str) -> dict:
    all_data = {}
    sections = slice_sections(text)

    for section, fields in SECTIONS.items():
        extracted = extract_fields_from_section(sections.get(section, ""), fields)
        all_data.update(extracted)

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
    st.dataframe(df, use_container_width=True)

    st.download_button(
        "⬇️ Download Excel",
        data=to_excel(df),
        file_name="zsub_final_stable_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
