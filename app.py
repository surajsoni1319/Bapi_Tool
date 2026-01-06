import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(
    page_title="SAP PDF → Final Excel",
    page_icon="📄",
    layout="wide"
)

st.title("📄 SAP Customer PDF → Final Excel")
st.write("PDF → Raw row data → Final SAP-ready Excel (single flow)")
st.markdown("---")

uploaded_files = st.file_uploader(
    "Upload SAP Customer Creation PDF(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# =====================================================
# REQUIRED OUTPUT FIELDS (EXACT)
# =====================================================
FIELDS = [
    "Type of Customer","Name of Customer","Company Code","Customer Group",
    "Sales Group","Region","Zone","Sub Zone","State","Sales Office",
    "SAP Dealer code to be mapped Search Term 2",
    "Search Term 1- Old customer code","Search Term 2 - District",
    "Mobile Number","E-Mail ID","Lattitude","Longitude",
    "Name of the Customers (Trade Name or Legal Name)",
    "Mobile Number (Address)","E-mail",
    "Address 1","Address 2","Address 3","Address 4",
    "PIN","City","District","State (Address)",
    "Whatsapp No.","Date of Birth","Date of Anniversary",
    "Counter Potential - Maximum","Counter Potential - Minimum",
    "Is GST Present","GSTIN","Trade Name","Legal Name","Reg Date",
    "City (GST)","Type","Building No.","District Code","State Code",
    "Street","PIN Code",
    "PAN","PAN Holder Name","PAN Status","PAN - Aadhaar Linking Status",
    "IFSC Number","Account Number","Name of Account Holder","Bank Name","Bank Branch",
    "Is Aadhaar Linked with Mobile?","Aadhaar Number","Name","Gender","DOB",
    "Address (Aadhaar)","PIN (Aadhaar)","City (Aadhaar)","State (Aadhaar)",
    "Logistics Transportation Zone","Transportation Zone Description",
    "Transportation Zone Code","Postal Code",
    "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment","Delivering Plant","Plant Name","Plant Code",
    "Incoterns","Incoterns (2)","Incoterns Code",
    "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)","Regional Head to be mapped","Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped","Area Sales Manager to be mapped",
    "Sales Officer to be mapped","Sales Promoter to be mapped",
    "Sales Promoter Number","Internal control code","SAP CODE",
    "Initiator Name","Initiator Email ID","Initiator Mobile Number",
    "Created By Customer UserID","Sales Head Name","Sales Head Email",
    "Sales Head Mobile Number","Extra2","PAN Result","Mobile Number Result",
    "Email Result","GST Result","Final Result"
]

# =====================================================
# STEP 1: PDF → RAW ROW DATA
# =====================================================
def pdf_to_raw_rows(pdf_file):
    rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue

            for line_no, line in enumerate(text.split("\n"), start=1):
                rows.append({
                    "Page": page_no,
                    "Line No": line_no,
                    "Text": line.strip()
                })

    return pd.DataFrame(rows)

# =====================================================
# HELPER FUNCTIONS ON RAW DATA
# =====================================================
def value_same_line(df, label):
    match = df[df["Text"].str.startswith(label + " ")]
    if not match.empty:
        return match.iloc[0]["Text"].replace(label, "").strip()
    return ""

def value_next_line(df, label):
    idx = df.index[df["Text"] == label]
    if not idx.empty and idx[0] + 1 in df.index:
        return df.loc[idx[0] + 1, "Text"]
    return ""

# =====================================================
# STEP 2: RAW ROWS → FINAL SAP ROW
# =====================================================
def build_final_row(df):
    row = {f: "" for f in FIELDS}

    # Header
    row["Type of Customer"] = value_same_line(df, "Type of Customer")
    row["Name of Customer"] = value_same_line(df, "Name of Customer")
    row["Company Code"] = value_same_line(df, "Company Code")
    row["Customer Group"] = value_same_line(df, "Customer Group")
    row["Sales Group"] = value_same_line(df, "Sales Group")
    row["Region"] = value_same_line(df, "Region")
    row["Zone"] = value_same_line(df, "Zone")
    row["Sub Zone"] = value_same_line(df, "Sub Zone")
    row["State"] = value_same_line(df, "State")
    row["Sales Office"] = value_same_line(df, "Sales Office")

    # Search Terms
    row["SAP Dealer code to be mapped Search Term 2"] = value_same_line(
        df, "SAP Dealer code to be mapped Search Term"
    )

    # Contact & Geo
    row["Mobile Number"] = value_same_line(df, "Mobile Number")
    row["E-Mail ID"] = value_same_line(df, "E-Mail ID")
    row["Lattitude"] = value_same_line(df, "Lattitude")
    row["Longitude"] = value_same_line(df, "Longitude")

    # Address
    row["Address 1"] = value_same_line(df, "Address 1")
    row["Address 2"] = value_same_line(df, "Address 2")
    row["Address 3"] = value_next_line(df, "Address 3")
    row["Address 4"] = value_next_line(df, "Address 4")
    row["PIN"] = value_same_line(df, "PIN")
    row["City"] = value_same_line(df, "City")
    row["District"] = value_same_line(df, "District")
    row["Whatsapp No."] = value_same_line(df, "Whatsapp No.")

    # Dates & Counters
    row["Date of Birth"] = value_same_line(df, "Date of Birth")
    row["Date of Anniversary"] = value_same_line(df, "Date of Anniversary")
    row["Counter Potential - Maximum"] = value_same_line(df, "Counter Potential - Maximum")
    row["Counter Potential - Minimum"] = value_same_line(df, "Counter Potential - Minimum")

    # GST / PAN
    row["Is GST Present"] = value_same_line(df, "Is GST Present")
    row["PAN"] = value_same_line(df, "PAN")
    row["PAN Holder Name"] = value_same_line(df, "PAN Holder Name")
    row["PAN Status"] = value_same_line(df, "PAN Status")
    row["PAN - Aadhaar Linking Status"] = value_same_line(df, "PAN - Aadhaar Linking Status")

    # Bank
    row["IFSC Number"] = value_same_line(df, "IFSC Number")
    row["Account Number"] = value_same_line(df, "Account Number")
    row["Name of Account Holder"] = value_same_line(df, "Name of Account Holder")
    row["Bank Name"] = value_same_line(df, "Bank Name")
    row["Bank Branch"] = value_same_line(df, "Bank Branch")

    # Aadhaar
    row["Is Aadhaar Linked with Mobile?"] = value_same_line(df, "Is Aadhaar Linked with Mobile?")
    row["Aadhaar Number"] = value_same_line(df, "Aadhaar Number")
    row["Name"] = value_same_line(df, "Name")
    row["Gender"] = value_same_line(df, "Gender")
    row["DOB"] = value_same_line(df, "DOB")

    # Logistics
    row["Logistics Transportation Zone"] = value_same_line(df, "Logistics Transportation Zone")
    row["Transportation Zone Description"] = value_same_line(df, "Transportation Zone Description")
    row["Transportation Zone Code"] = value_same_line(df, "Transportation Zone Code")

    # Results
    row["PAN Result"] = value_same_line(df, "PAN Result")
    row["Mobile Number Result"] = value_same_line(df, "Mobile Number Result")
    row["Email Result"] = value_same_line(df, "Email Result")
    row["GST Result"] = value_same_line(df, "GST Result")
    row["Final Result"] = value_same_line(df, "Final Result")

    return row

# =====================================================
# RUN PIPELINE
# =====================================================
if uploaded_files:
    final_rows = []

    for pdf in uploaded_files:
        raw_df = pdf_to_raw_rows(pdf)
        final_row = build_final_row(raw_df)
        final_row["Source File"] = pdf.name
        final_rows.append(final_row)

    final_df = pd.DataFrame(final_rows)

    st.subheader("✅ Final SAP Output")
    st.dataframe(final_df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="SAP_Data")

    st.download_button(
        "⬇️ Download Final Excel",
        data=output.getvalue(),
        file_name="Final_SAP_Customer_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
