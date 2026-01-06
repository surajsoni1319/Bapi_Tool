import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# =============================
# PAGE CONFIG
# =============================
st.set_page_config(
    page_title="SAP PDF → Final Excel",
    page_icon="📄",
    layout="wide"
)

st.title("📄 SAP Customer PDF → Final Excel")
st.write("Raw extraction ➜ cleaning ➜ final SAP-ready Excel")
st.markdown("---")

uploaded_files = st.file_uploader(
    "Upload SAP Customer Creation PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# =============================
# CONFIG
# =============================

HEADERS_TO_DROP = {
    "Customer Creation",
    "Customer Address",
    "GSTIN Details",
    "PAN Details",
    "Bank Details",
    "Aadhaar Details",
    "Supporting Document",
    "Zone Details",
    "Other Details",
    "Duplicity Check"
}

FINAL_FIELDS = [
    "Type of Customer","Name of Customer","Company Code","Customer Group",
    "Sales Group","Region","Zone","Sub Zone","State","Sales Office",
    "SAP Dealer code to be mapped Search Term 2",
    "Search Term 1- Old customer code","Search Term 2 - District",
    "Mobile Number","E-Mail ID","Lattitude","Longitude",
    "Name of the Customers (Trade Name or Legal Name)",
    "Mobile Number","E-mail",
    "Address 1","Address 2","Address 3","Address 4",
    "PIN","City","District","State","Whatsapp No.",
    "Date of Birth","Date of Anniversary",
    "Counter Potential - Maximum","Counter Potential - Minimum",
    "Is GST Present","GSTIN","Trade Name","Legal Name","Reg Date",
    "City","Type","Building No.","District Code","State Code","Street","PIN Code",
    "PAN","PAN Holder Name","PAN Status","PAN - Aadhaar Linking Status",
    "IFSC Number","Account Number","Name of Account Holder","Bank Name","Bank Branch",
    "Is Aadhaar Linked with Mobile?","Aadhaar Number","Name","Gender","DOB",
    "Address","PIN","City","State",
    "Logistics Transportation Zone","Transportation Zone Description",
    "Transportation Zone Code","Postal Code",
    "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment","Delivering Plant","Plant Name","Plant Code",
    "Incoterns","Incoterns","Incoterns Code",
    "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)","Regional Head to be mapped","Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped","Area Sales Manager to be mapped",
    "Sales Officer to be mapped","Sales Promoter to be mapped","Sales Promoter Number",
    "Internal control code","SAP CODE","Initiator Name","Initiator Email ID",
    "Initiator Mobile Number","Created By Customer UserID",
    "Sales Head Name","Sales Head Email","Sales Head Mobile Number",
    "Extra2","PAN Result","Mobile Number Result","Email Result","GST Result","Final Result"
]

# =============================
# STEP 1: RAW PDF EXTRACTION
# =============================

def pdf_to_raw_df(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue
            for line_no, line in enumerate(text.split("\n"), start=1):
                rows.append({
                    "Source File": pdf_file.name,
                    "Page": page_no,
                    "Line No": line_no,
                    "Text": line.strip()
                })
    return pd.DataFrame(rows)

# =============================
# STEP 2: CLEAN RAW LINES
# =============================

def clean_raw_lines(df):
    lines = df["Text"].tolist()
    cleaned = []
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        # Drop headers
        if line in HEADERS_TO_DROP:
            i += 1
            continue

        # Merge broken lines
        while i + 1 < len(lines) and (
            lines[i + 1].endswith(")") or
            lines[i + 1].startswith("Sales Officer") or
            lines[i + 1].startswith("Master list") or
            lines[i + 1].startswith("be provided") or
            re.fullmatch(r"\d+", lines[i + 1])
        ):
            line = f"{line} {lines[i + 1].strip()}"
            i += 1

        cleaned.append(line)
        i += 1

    return cleaned

# =============================
# STEP 3: MAP TO FINAL FIELDS
# =============================

def map_to_final_fields(cleaned_lines):
    result = {field: "" for field in FINAL_FIELDS}

    for line in cleaned_lines:
        for field in FINAL_FIELDS:
            if line.startswith(field):
                result[field] = line.replace(field, "").strip()

    return result

# =============================
# MAIN FLOW
# =============================

if uploaded_files:
    final_rows = []

    for pdf in uploaded_files:
        raw_df = pdf_to_raw_df(pdf)
        cleaned_lines = clean_raw_lines(raw_df)
        final_data = map_to_final_fields(cleaned_lines)
        final_data["Source File"] = pdf.name
        final_rows.append(final_data)

    final_df = pd.DataFrame(final_rows)

    st.subheader("✅ Final Extracted Excel Preview")
    st.dataframe(final_df, use_container_width=True)

    # Export FINAL Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="FINAL_SAP_DATA")

    st.download_button(
        "⬇️ Download FINAL Excel",
        data=output.getvalue(),
        file_name="FINAL_SAP_CUSTOMER.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
