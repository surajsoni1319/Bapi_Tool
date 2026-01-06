import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="RAW PDF → SAP Excel",
    page_icon="📄",
    layout="wide"
)

st.title("RAW Extracted PDF → Final SAP Excel")
st.write("Uses raw extracted data. No PDF parsing. No guessing.")
st.markdown("---")

uploaded_file = st.file_uploader(
    "Upload RAW extracted Excel (Source File | Page | Line No | Text)",
    type=["xlsx"]
)

# =====================================================
# REQUIRED FIELDS (EXACT ORDER)
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

def extract_value(df, label):
    """Extract value from same-line text"""
    match = df[df["Text"].str.startswith(label + " ")]
    if not match.empty:
        return match.iloc[0]["Text"].replace(label, "").strip()
    return ""

def extract_next_line(df, label):
    """Extract value from next line"""
    idx = df.index[df["Text"] == label]
    if not idx.empty and idx[0] + 1 in df.index:
        return df.loc[idx[0] + 1, "Text"]
    return ""

def build_final_row(df):
    row = {f: "" for f in FIELDS}

    # Header
    row["Type of Customer"] = extract_value(df, "Type of Customer")
    row["Name of Customer"] = extract_value(df, "Name of Customer")
    row["Company Code"] = extract_value(df, "Company Code")
    row["Customer Group"] = extract_value(df, "Customer Group")
    row["Sales Group"] = extract_value(df, "Sales Group")
    row["Region"] = extract_value(df, "Region")
    row["Zone"] = extract_value(df, "Zone")
    row["Sub Zone"] = extract_value(df, "Sub Zone")
    row["State"] = extract_value(df, "State")
    row["Sales Office"] = extract_value(df, "Sales Office")

    # Search Terms
    row["SAP Dealer code to be mapped Search Term 2"] = extract_value(
        df, "SAP Dealer code to be mapped Search Term"
    )

    # Contact
    row["Mobile Number"] = extract_value(df, "Mobile Number")
    row["E-Mail ID"] = extract_value(df, "E-Mail ID")
    row["Lattitude"] = extract_value(df, "Lattitude")
    row["Longitude"] = extract_value(df, "Longitude")

    # Address
    row["Address 1"] = extract_value(df, "Address 1")
    row["Address 2"] = extract_value(df, "Address 2")
    row["Address 3"] = extract_next_line(df, "Address 3")
    row["Address 4"] = extract_next_line(df, "Address 4")
    row["PIN"] = extract_value(df, "PIN")
    row["City"] = extract_value(df, "City")
    row["District"] = extract_value(df, "District")
    row["Whatsapp No."] = extract_value(df, "Whatsapp No.")

    # PAN
    row["PAN"] = extract_value(df, "PAN")
    row["PAN Holder Name"] = extract_value(df, "PAN Holder Name")
    row["PAN Status"] = extract_value(df, "PAN Status")
    row["PAN - Aadhaar Linking Status"] = extract_value(df, "PAN - Aadhaar Linking Status")

    # Bank
    row["IFSC Number"] = extract_value(df, "IFSC Number")
    row["Account Number"] = extract_value(df, "Account Number")
    row["Name of Account Holder"] = extract_value(df, "Name of Account Holder")
    row["Bank Name"] = extract_value(df, "Bank Name")
    row["Bank Branch"] = extract_value(df, "Bank Branch")

    # Aadhaar
    row["Is Aadhaar Linked with Mobile?"] = extract_value(df, "Is Aadhaar Linked with Mobile?")
    row["Aadhaar Number"] = extract_value(df, "Aadhaar Number")
    row["Name"] = extract_value(df, "Name")
    row["Gender"] = extract_value(df, "Gender")
    row["DOB"] = extract_value(df, "DOB")

    # Logistics
    row["Logistics Transportation Zone"] = extract_value(df, "Logistics Transportation Zone")
    row["Transportation Zone Description"] = extract_value(df, "Transportation Zone Description")
    row["Transportation Zone Code"] = extract_value(df, "Transportation Zone Code")

    # Results
    row["PAN Result"] = extract_value(df, "PAN Result")
    row["Mobile Number Result"] = extract_value(df, "Mobile Number Result")
    row["Email Result"] = extract_value(df, "Email Result")
    row["GST Result"] = extract_value(df, "GST Result")
    row["Final Result"] = extract_value(df, "Final Result")

    return row

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file)
    final_row = build_final_row(raw_df)
    final_df = pd.DataFrame([final_row])

    st.subheader("✅ Final SAP-Ready Output")
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
