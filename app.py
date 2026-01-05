import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(
    page_title="SAP Customer PDF → Excel",
    page_icon="📄",
    layout="wide"
)

st.title("📄 SAP Customer Creation PDF → Excel")
st.write("Converts SAP Customer Creation PDFs into a fully structured Excel (1 row per customer).")
st.markdown("---")

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# ==============================
# EXACT FIELD LIST (AS PROVIDED)
# ==============================
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
    "Mobile Number (Address)",
    "E-mail",
    "Address 1",
    "Address 2",
    "Address 3",
    "Address 4",
    "PIN",
    "City",
    "District",
    "State (Address)",
    "Whatsapp No.",
    "Date of Birth",
    "Date of Anniversary",
    "Counter Potential - Maximum",
    "Counter Potential - Minimum",
    "Is GST Present",
    "GSTIN",
    "Trade Name",
    "Legal Name",
    "Reg Date",
    "City (GST)",
    "Type",
    "Building No.",
    "District Code",
    "State Code",
    "Street",
    "PIN Code",
    "PAN",
    "PAN Holder Name",
    "PAN Status",
    "PAN - Aadhaar Linking Status",
    "IFSC Number",
    "Account Number",
    "Name of Account Holder",
    "Bank Name",
    "Bank Branch",
    "Is Aadhaar Linked with Mobile?",
    "Aadhaar Number",
    "Name",
    "Gender",
    "DOB",
    "Address",
    "PIN (Aadhaar)",
    "City (Aadhaar)",
    "State (Aadhaar)",
    "Logistics Transportation Zone",
    "Transportation Zone Description",
    "Transportation Zone Code",
    "Postal Code",
    "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment",
    "Delivering Plant",
    "Plant Name",
    "Plant Code",
    "Incoterns",
    "Incoterns (Duplicate)",
    "Incoterns Code",
    "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)",
    "Regional Head to be mapped",
    "Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped",
    "Area Sales Manager to be mapped",
    "Sales Officer to be mapped",
    "Sales Promoter to be mapped",
    "Sales Promoter Number",
    "Internal control code",
    "SAP CODE",
    "Initiator Name",
    "Initiator Email ID",
    "Initiator Mobile Number",
    "Created By Customer UserID",
    "Sales Head Name",
    "Sales Head Email",
    "Sales Head Mobile Number",
    "Extra2",
    "PAN Result",
    "Mobile Number Result",
    "Email Result",
    "GST Result",
    "Final Result"
]

def extract_data(pdf_file):
    row = {field: "" for field in FIELDS}

    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    for field in FIELDS:
        pattern = rf"{re.escape(field)}\s+(.+)"
        match = re.search(pattern, full_text)
        if match:
            row[field] = match.group(1).strip()

    return row

# ==============================
# PROCESS FILES
# ==============================
if uploaded_files:
    rows = []

    for pdf in uploaded_files:
        row = extract_data(pdf)
        row["Source File"] = pdf.name
        rows.append(row)

    df = pd.DataFrame(rows)

    st.subheader("📊 Extracted Data Preview")
    st.dataframe(df, use_container_width=True)

    # Export Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Customer_Data")

    st.download_button(
        "⬇️ Download Excel",
        data=output.getvalue(),
        file_name="SAP_Customer_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
