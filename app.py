import streamlit as st
import pdfplumber
import pandas as pd
import io

# -----------------------------
# Page config
# -----------------------------
st.set_page_config(
    page_title="PDF → Excel (Raw + Clean)",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Excel Converter (RAW → CLEAN)")
st.write("Step 1: Extract raw PDF text")
st.write("Step 2: Clean broken lines into Field | Value")
st.markdown("---")

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# -----------------------------
# Known field starters
# -----------------------------
KNOWN_FIELDS = [
    "Type of Customer","Name of Customer","Company Code","Customer Group",
    "Sales Group","Region","Zone","Sub Zone","State","Sales Office",
    "SAP Dealer code to be mapped Search Term",
    "Search Term 1- Old customer code","Search Term 2 - District",
    "Mobile Number","E-Mail ID","Lattitude","Longitude",
    "Name of the Customers","Legal Name","E-mail",
    "Address 1","Address 2","Address 3","Address 4",
    "PIN","City","District","Whatsapp No.",
    "Date of Birth","Date of Anniversary",
    "Counter Potential - Maximum","Counter Potential - Minimum",
    "Is GST Present","GSTIN","Trade Name","Reg Date",
    "PAN","PAN Holder Name","PAN Status","PAN - Aadhaar Linking Status",
    "IFSC Number","Account Number","Name of Account Holder",
    "Bank Name","Bank Branch",
    "Is Aadhaar Linked with Mobile?","Aadhaar Number",
    "Logistics Transportation Zone","Transportation Zone Description",
    "Transportation Zone Code",
    "Date of Appointment","Delivering Plant","Plant Name","Plant Code",
    "Incoterns","Incoterns Code",
    "Security Deposit Amount details to filled up",
    "Credit Limit (In Rs.)",
    "Regional Head to be mapped","Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped",
    "Area Sales Manager to be mapped","Sales Officer to be mapped",
    "Sales Promoter to be mapped","Sales Promoter Number",
    "Internal control code","SAP CODE",
    "Initiator Name","Initiator Email ID","Initiator Mobile Number",
    "Created By Customer UserID",
    "Sales Head Name","Sales Head Email","Sales Head Mobile Number",
    "PAN Result","Mobile Number Result","Email Result","GST Result","Final Result"
]

# -----------------------------
# Step 1: PDF → RAW
# -----------------------------
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

# -----------------------------
# Step 2: RAW → CLEAN
# -----------------------------
def raw_to_clean_df(raw_df):
    records = []
    current_field = None
    current_value = ""

    for text in raw_df["Text"]:
        matched = False

        for field in KNOWN_FIELDS:
            if text.startswith(field):
                if current_field:
                    records.append([current_field, current_value.strip()])
                current_field = field
                current_value = text.replace(field, "").strip()
                matched = True
                break

        if not matched and current_field:
            current_value += " " + text

    if current_field:
        records.append([current_field, current_value.strip()])

    return pd.DataFrame(records, columns=["Field", "Value"])

# -----------------------------
# Main flow
# -----------------------------
if uploaded_files:
    all_raw = []
    all_clean = []

    for pdf in uploaded_files:
        raw_df = pdf_to_raw_df(pdf)
        raw_df["Source File"] = pdf.name
        all_raw.append(raw_df)

        clean_df = raw_to_clean_df(raw_df)
        clean_df["Source File"] = pdf.name
        all_clean.append(clean_df)

    final_raw = pd.concat(all_raw, ignore_index=True)
    final_clean = pd.concat(all_clean, ignore_index=True)

    st.subheader("🔹 RAW PDF DATA")
    st.dataframe(final_raw, use_container_width=True)

    st.subheader("🔹 CLEAN FIELD | VALUE DATA")
    st.dataframe(final_clean, use_container_width=True)

    # Export Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_raw.to_excel(writer, index=False, sheet_name="RAW_TEXT")
        final_clean.to_excel(writer, index=False, sheet_name="CLEAN_DATA")

    st.download_button(
        "⬇️ Download Excel (RAW + CLEAN)",
        data=output.getvalue(),
        file_name="PDF_RAW_AND_CLEAN.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
