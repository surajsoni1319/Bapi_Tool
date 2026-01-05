import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(
    page_title="SAP Customer PDF → Excel (Corrected)",
    page_icon="📄",
    layout="wide"
)

st.title("📄 SAP Customer Creation PDF → Excel (Correct Logic)")
st.markdown("Uses section-based extraction. Safe for SAP forms.")
st.markdown("---")

uploaded_files = st.file_uploader(
    "Upload SAP Customer Creation PDF(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# Exact output columns
COLUMNS = [
    "Type of Customer","Name of Customer","Company Code","Customer Group",
    "Sales Group","Region","Zone","Sub Zone","State","Sales Office",
    "SAP Dealer code to be mapped Search Term 2",
    "Search Term 1- Old customer code","Search Term 2 - District",
    "Mobile Number","E-Mail ID","Lattitude","Longitude",
    "Address 1","Address 2","Address 3","Address 4",
    "PIN","City","District","Whatsapp No.",
    "Date of Birth","Date of Anniversary",
    "Counter Potential - Maximum","Counter Potential - Minimum",
    "Is GST Present","GSTIN","PAN",
    "PAN Holder Name","PAN Status","PAN - Aadhaar Linking Status",
    "IFSC Number","Account Number","Name of Account Holder",
    "Bank Name","Bank Branch",
    "Is Aadhaar Linked with Mobile?","Aadhaar Number",
    "Logistics Transportation Zone","Transportation Zone Description",
    "Transportation Zone Code",
    "Date of Appointment","Delivering Plant","Plant Name","Plant Code",
    "Incoterms","Incoterms Code",
    "Security Deposit Amount","Credit Limit (In Rs.)",
    "Regional Head to be mapped","Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped",
    "Area Sales Manager to be mapped","Sales Officer to be mapped",
    "Internal control code","SAP CODE",
    "Initiator Name","Initiator Email ID","Initiator Mobile Number",
    "Sales Head Name","Sales Head Email","Sales Head Mobile Number",
    "PAN Result","Mobile Number Result","Email Result","GST Result","Final Result"
]

def extract_pdf(pdf):
    data = {col: "" for col in COLUMNS}

    with pdfplumber.open(pdf) as p:
        lines = []
        for page in p.pages:
            text = page.extract_text()
            if text:
                lines.extend([l.strip() for l in text.split("\n") if l.strip()])

    def next_line_value(label):
        for i, line in enumerate(lines):
            if line == label and i + 1 < len(lines):
                return lines[i + 1]
        return ""

    # Header section
    data["Type of Customer"] = next_line_value("Type of Customer")
    data["Name of Customer"] = next_line_value("Name of Customer")
    data["Company Code"] = next_line_value("Company Code")
    data["Customer Group"] = next_line_value("Customer Group")
    data["Sales Group"] = next_line_value("Sales Group")
    data["Region"] = next_line_value("Region")
    data["Zone"] = next_line_value("Zone")
    data["Sub Zone"] = next_line_value("Sub Zone")
    data["State"] = next_line_value("State")
    data["Sales Office"] = next_line_value("Sales Office")

    # Search terms
    data["SAP Dealer code to be mapped Search Term 2"] = next_line_value(
        "SAP Dealer code to be mapped Search Term 2"
    )

    # Contact
    data["Mobile Number"] = next_line_value("Mobile Number")
    data["E-Mail ID"] = next_line_value("E-Mail ID")
    data["Whatsapp No."] = next_line_value("Whatsapp No.")

    # Location
    data["Lattitude"] = next_line_value("Lattitude")
    data["Longitude"] = next_line_value("Longitude")

    # Address block
    data["Address 1"] = next_line_value("Address 1")
    data["Address 2"] = next_line_value("Address 2")
    data["Address 3"] = next_line_value("Address 3")
    data["Address 4"] = next_line_value("Address 4")
    data["PIN"] = next_line_value("PIN")
    data["City"] = next_line_value("City")
    data["District"] = next_line_value("District")

    # GST / PAN
    data["Is GST Present"] = next_line_value("Is GST Present")
    data["GSTIN"] = next_line_value("GSTIN")
    data["PAN"] = next_line_value("PAN")
    data["PAN Holder Name"] = next_line_value("PAN Holder Name")
    data["PAN Status"] = next_line_value("PAN Status")
    data["PAN - Aadhaar Linking Status"] = next_line_value("PAN - Aadhaar Linking Status")

    # Bank
    data["IFSC Number"] = next_line_value("IFSC Number")
    data["Account Number"] = next_line_value("Account Number")
    data["Name of Account Holder"] = next_line_value("Name of Account Holder")
    data["Bank Name"] = next_line_value("Bank Name")
    data["Bank Branch"] = next_line_value("Bank Branch")

    # Aadhaar
    data["Is Aadhaar Linked with Mobile?"] = next_line_value("Is Aadhaar Linked with Mobile?")
    data["Aadhaar Number"] = next_line_value("Aadhaar Number")

    # Logistics
    data["Logistics Transportation Zone"] = next_line_value("Logistics Transportation Zone")
    data["Transportation Zone Description"] = next_line_value("Transportation Zone Description")
    data["Transportation Zone Code"] = next_line_value("Transportation Zone Code")

    # Other
    data["Date of Appointment"] = next_line_value("Date of Appointment")
    data["Delivering Plant"] = next_line_value("Delivering Plant")
    data["Plant Name"] = next_line_value("Plant Name")
    data["Plant Code"] = next_line_value("Plant Code")

    # Results
    data["PAN Result"] = next_line_value("PAN Result")
    data["Mobile Number Result"] = next_line_value("Mobile Number Result")
    data["Email Result"] = next_line_value("Email Result")
    data["GST Result"] = next_line_value("GST Result")
    data["Final Result"] = next_line_value("Final Result")

    return data

if uploaded_files:
    rows = [extract_pdf(pdf) for pdf in uploaded_files]
    df = pd.DataFrame(rows)

    st.subheader("✅ Corrected Extraction Preview")
    st.dataframe(df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="SAP_Data")

    st.download_button(
        "⬇️ Download Correct Excel",
        data=output.getvalue(),
        file_name="SAP_Customer_Data_Correct.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
