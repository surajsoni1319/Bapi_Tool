import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# Page config
st.set_page_config(page_title="PDF Customer Data Extractor", page_icon="📄", layout="wide")

# Title
st.title("📄 PDF Customer Data Extractor")
st.markdown("---")

# Output column names
OUTPUT_COLUMNS = [
    "Type of Customer", "Name of Customer", "Company Code", "Customer Group",
    "Sales Group", "Region", "Zone", "Sub Zone", "State", "Sales Office",
    "SAP Dealer code to be mapped Search Term 2", "Search Term 1- Old customer code",
    "Search Term 2 - District", "Mobile Number", "E-Mail ID", "Lattitude", "Longitude",
    "Name of the Customers (Trade Name or Legal Name)", "E-mail",
    "Address 1", "Address 2", "Address 3", "Address 4", "PIN", "City", "District",
    "Whatsapp No.", "Date of Birth", "Date of Anniversary",
    "Counter Potential - Maximum", "Counter Potential - Minimum",
    "Is GST Present", "GSTIN", "Trade Name", "Legal Name", "Reg Date",
    "Type", "Building No.", "District Code", "State Code", "Street", "PIN Code",
    "PAN", "PAN Holder Name", "PAN Status", "PAN - Aadhaar Linking Status",
    "IFSC Number", "Account Number", "Name of Account Holder", "Bank Name", "Bank Branch",
    "Is Aadhaar Linked with Mobile?", "Aadhaar Number", "Name", "Gender", "DOB", "Address",
    "Logistics Transportation Zone", "Transportation Zone Description",
    "Transportation Zone Code", "Postal Code",
    "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment", "Delivering Plant", "Plant Name", "Plant Code",
    "Incoterns", "Incoterns Code",
    "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)", "Regional Head to be mapped", "Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped", "Area Sales Manager to be mapped",
    "Sales Officer to be mapped", "Sales Promoter to be mapped", "Sales Promoter Number",
    "Internal control code", "SAP CODE", "Initiator Name", "Initiator Email ID",
    "Initiator Mobile Number", "Created By Customer UserID", "Sales Head Name",
    "Sales Head Email", "Sales Head Mobile Number", "Extra2",
    "PAN Result", "Mobile Number Result", "Email Result", "GST Result", "Final Result"
]

EXCLUDE_HEADERS = [
    'Customer Creation', 'Customer Address', 'GSTIN Details', 
    'PAN Details', 'Bank Details', 'Aadhaar Details', 
    'Supporting Document', 'Zone Details', 'Other Details', 
    'Duplicity Check'
]

def extract_value_after_field(text, field_name):
    """Extract value that comes after a field name"""
    # Try pattern: Field Name   Value (on same line)
    pattern = re.escape(field_name) + r'\s+(.+?)(?=\n|$)'
    match = re.search(pattern, text, re.MULTILINE)
    if match:
        value = match.group(1).strip()
        # Make sure it's not another field name
        if not any(value.startswith(h) for h in EXCLUDE_HEADERS):
            return value
    return ""

def extract_data_from_pdf(pdf_file):
    """Extract key-value pairs from PDF file"""
    
    # Read PDF
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Get all text
    full_text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        full_text += page.get_text()
    
    pdf_document.close()
    
    # Dictionary to store extracted data
    data = {}
    
    # Extract each field
    data["Type of Customer"] = extract_value_after_field(full_text, "Type of Customer")
    data["Name of Customer"] = extract_value_after_field(full_text, "Name of Customer")
    data["Company Code"] = extract_value_after_field(full_text, "Company Code")
    data["Customer Group"] = extract_value_after_field(full_text, "Customer Group")
    data["Sales Group"] = extract_value_after_field(full_text, "Sales Group")
    data["Region"] = extract_value_after_field(full_text, "Region")
    
    # Zone - careful with Sub Zone and Transportation Zone
    zone_match = re.search(r'^Zone\s+([A-Z0-9]+)\s*$', full_text, re.MULTILINE)
    if zone_match:
        data["Zone"] = zone_match.group(1)
    
    data["Sub Zone"] = extract_value_after_field(full_text, "Sub Zone")
    
    # State - first occurrence
    state_match = re.search(r'^State\s+([A-Z]+)\s*$', full_text, re.MULTILINE)
    if state_match:
        data["State"] = state_match.group(1)
    
    data["Sales Office"] = extract_value_after_field(full_text, "Sales Office")
    
    # SAP Dealer code - complex multi-line
    sap_pattern = r'SAP Dealer code to be mapped Search Term\s*\n\s*2\s*\n\s*(\d+)'
    sap_match = re.search(sap_pattern, full_text)
    if sap_match:
        data["SAP Dealer code to be mapped Search Term 2"] = sap_match.group(1)
    
    # Search Terms
    data["Search Term 1- Old customer code"] = extract_value_after_field(full_text, "Search Term 1- Old customer code")
    data["Search Term 2 - District"] = extract_value_after_field(full_text, "Search Term 2 - District")
    
    # Mobile Number - first occurrence
    mobile_match = re.search(r'^Mobile Number\s+(\d+)', full_text, re.MULTILINE)
    if mobile_match:
        data["Mobile Number"] = mobile_match.group(1)
    
    data["E-Mail ID"] = extract_value_after_field(full_text, "E-Mail ID")
    data["Lattitude"] = extract_value_after_field(full_text, "Lattitude")
    data["Longitude"] = extract_value_after_field(full_text, "Longitude")
    
    # Name of Customers - multi-line
    name_pattern = r'Name of the Customers \(Trade Name or\s*\n\s*Legal Name\)\s+(.+?)(?=\n|$)'
    name_match = re.search(name_pattern, full_text)
    if name_match:
        data["Name of the Customers (Trade Name or Legal Name)"] = name_match.group(1).strip()
    
    data["E-mail"] = extract_value_after_field(full_text, "E-mail")
    data["Address 1"] = extract_value_after_field(full_text, "Address 1")
    data["Address 2"] = extract_value_after_field(full_text, "Address 2")
    data["Address 3"] = extract_value_after_field(full_text, "Address 3")
    data["Address 4"] = extract_value_after_field(full_text, "Address 4")
    
    # PIN - first occurrence
    pin_match = re.search(r'^PIN\s+(\d+)', full_text, re.MULTILINE)
    if pin_match:
        data["PIN"] = pin_match.group(1)
    
    # City - first occurrence
    city_match = re.search(r'^City\s+(\w+)', full_text, re.MULTILINE)
    if city_match:
        data["City"] = city_match.group(1)
    
    # District - first occurrence
    district_match = re.search(r'^District\s+(\w+)', full_text, re.MULTILINE)
    if district_match:
        data["District"] = district_match.group(1)
    
    data["Whatsapp No."] = extract_value_after_field(full_text, "Whatsapp No.")
    data["Date of Birth"] = extract_value_after_field(full_text, "Date of Birth")
    data["Date of Anniversary"] = extract_value_after_field(full_text, "Date of Anniversary")
    data["Counter Potential - Maximum"] = extract_value_after_field(full_text, "Counter Potential - Maximum")
    data["Counter Potential - Minimum"] = extract_value_after_field(full_text, "Counter Potential - Minimum")
    data["Is GST Present"] = extract_value_after_field(full_text, "Is GST Present")
    data["GSTIN"] = extract_value_after_field(full_text, "GSTIN")
    data["Trade Name"] = extract_value_after_field(full_text, "Trade Name")
    data["Legal Name"] = extract_value_after_field(full_text, "Legal Name")
    data["Reg Date"] = extract_value_after_field(full_text, "Reg Date")
    data["Type"] = extract_value_after_field(full_text, "Type")
    data["Building No."] = extract_value_after_field(full_text, "Building No.")
    data["District Code"] = extract_value_after_field(full_text, "District Code")
    data["State Code"] = extract_value_after_field(full_text, "State Code")
    data["Street"] = extract_value_after_field(full_text, "Street")
    data["PIN Code"] = extract_value_after_field(full_text, "PIN Code")
    data["PAN"] = extract_value_after_field(full_text, "PAN")
    data["PAN Holder Name"] = extract_value_after_field(full_text, "PAN Holder Name")
    data["PAN Status"] = extract_value_after_field(full_text, "PAN Status")
    data["PAN - Aadhaar Linking Status"] = extract_value_after_field(full_text, "PAN - Aadhaar Linking Status")
    data["IFSC Number"] = extract_value_after_field(full_text, "IFSC Number")
    data["Account Number"] = extract_value_after_field(full_text, "Account Number")
    data["Name of Account Holder"] = extract_value_after_field(full_text, "Name of Account Holder")
    data["Bank Name"] = extract_value_after_field(full_text, "Bank Name")
    data["Bank Branch"] = extract_value_after_field(full_text, "Bank Branch")
    data["Is Aadhaar Linked with Mobile?"] = extract_value_after_field(full_text, "Is Aadhaar Linked with Mobile?")
    data["Aadhaar Number"] = extract_value_after_field(full_text, "Aadhaar Number")
    
    # Name in Aadhaar section
    name_aadhaar_pattern = r'Aadhaar Number\s+.+?\n\s*Name\s+(.+?)(?=\n|$)'
    name_aadhaar_match = re.search(name_aadhaar_pattern, full_text, re.DOTALL)
    if name_aadhaar_match:
        data["Name"] = name_aadhaar_match.group(1).strip()
    
    data["Gender"] = extract_value_after_field(full_text, "Gender")
    
    # DOB in Aadhaar section
    dob_match = re.search(r'^DOB\s+([\d/]+)', full_text, re.MULTILINE)
    if dob_match:
        data["DOB"] = dob_match.group(1)
    
    # Address in Aadhaar section - multi-line
    address_pattern = r'DOB\s+[\d/]+\s*\n\s*Address\s+(.+?)(?=\n\s*PIN|$)'
    address_match = re.search(address_pattern, full_text, re.DOTALL)
    if address_match:
        data["Address"] = address_match.group(1).strip()
    
    data["Logistics Transportation Zone"] = extract_value_after_field(full_text, "Logistics Transportation Zone")
    data["Transportation Zone Description"] = extract_value_after_field(full_text, "Transportation Zone Description")
    data["Transportation Zone Code"] = extract_value_after_field(full_text, "Transportation Zone Code")
    data["Postal Code"] = extract_value_after_field(full_text, "Postal Code")
    
    # Multi-line logistics fields
    logistics_pattern = r'Logistics team to vet the T zone selected by\s*\n\s*Sales Officer\s*\n?\s*(\w+)'
    logistics_match = re.search(logistics_pattern, full_text)
    if logistics_match:
        data["Logistics team to vet the T zone selected by Sales Officer"] = logistics_match.group(1)
    
    tzone_pattern = r'Selection of Available T Zones from T Zone\s*\n\s*Master list, if found\.\s*\n?\s*(.+?)(?=\n|$)'
    tzone_match = re.search(tzone_pattern, full_text)
    if tzone_match:
        data["Selection of Available T Zones from T Zone Master list, if found."] = tzone_match.group(1).strip()
    
    new_tzone_pattern = r'If NEW T Zone need to be created, details to\s*\n\s*be provided by Logistics team\s*\n?\s*(.+?)(?=\n|$)'
    new_tzone_match = re.search(new_tzone_pattern, full_text)
    if new_tzone_match:
        data["If NEW T Zone need to be created, details to be provided by Logistics team"] = new_tzone_match.group(1).strip()
    
    data["Date of Appointment"] = extract_value_after_field(full_text, "Date of Appointment")
    data["Delivering Plant"] = extract_value_after_field(full_text, "Delivering Plant")
    data["Plant Name"] = extract_value_after_field(full_text, "Plant Name")
    data["Plant Code"] = extract_value_after_field(full_text, "Plant Code")
    
    # Incoterns - first occurrence
    inco_match = re.search(r'^Incoterns\s+(.+?)(?=\n|$)', full_text, re.MULTILINE)
    if inco_match:
        data["Incoterns"] = inco_match.group(1).strip()
    
    data["Incoterns Code"] = extract_value_after_field(full_text, "Incoterns Code")
    
    # Security Deposit - multi-line
    security_pattern = r'Security Deposit Amount details to filled up,\s*\n\s*as per checque received by Customer / Dealer\s*\n?\s*(\d+)'
    security_match = re.search(security_pattern, full_text)
    if security_match:
        data["Security Deposit Amount details to filled up, as per checque received by Customer / Dealer"] = security_match.group(1)
    
    data["Credit Limit (In Rs.)"] = extract_value_after_field(full_text, "Credit Limit (In Rs.)")
    data["Regional Head to be mapped"] = extract_value_after_field(full_text, "Regional Head to be mapped")
    data["Zonal Head to be mapped"] = extract_value_after_field(full_text, "Zonal Head to be mapped")
    data["Sub-Zonal Head (RSM) to be mapped"] = extract_value_after_field(full_text, "Sub-Zonal Head (RSM) to be mapped")
    data["Area Sales Manager to be mapped"] = extract_value_after_field(full_text, "Area Sales Manager to be mapped")
    data["Sales Officer to be mapped"] = extract_value_after_field(full_text, "Sales Officer to be mapped")
    data["Sales Promoter to be mapped"] = extract_value_after_field(full_text, "Sales Promoter to be mapped")
    data["Sales Promoter Number"] = extract_value_after_field(full_text, "Sales Promoter Number")
    data["Internal control code"] = extract_value_after_field(full_text, "Internal control code")
    data["SAP CODE"] = extract_value_after_field(full_text, "SAP CODE")
    data["Initiator Name"] = extract_value_after_field(full_text, "Initiator Name")
    data["Initiator Email ID"] = extract_value_after_field(full_text, "Initiator Email ID")
    data["Initiator Mobile Number"] = extract_value_after_field(full_text, "Initiator Mobile Number")
    data["Created By Customer UserID"] = extract_value_after_field(full_text, "Created By Customer UserID")
    data["Sales Head Name"] = extract_value_after_field(full_text, "Sales Head Name")
    data["Sales Head Email"] = extract_value_after_field(full_text, "Sales Head Email")
    data["Sales Head Mobile Number"] = extract_value_after_field(full_text, "Sales Head Mobile Number")
    data["Extra2"] = extract_value_after_field(full_text, "Extra2")
    data["PAN Result"] = extract_value_after_field(full_text, "PAN Result")
    data["Mobile Number Result"] = extract_value_after_field(full_text, "Mobile Number Result")
    data["Email Result"] = extract_value_after_field(full_text, "Email Result")
    data["GST Result"] = extract_value_after_field(full_text, "GST Result")
    data["Final Result"] = extract_value_after_field(full_text, "Final Result")
    
    return data

def convert_to_dataframe(data_dict):
    """Convert dictionary to pandas DataFrame with all columns"""
    df_data = []
    for col_name in OUTPUT_COLUMNS:
        value = data_dict.get(col_name, "")
        df_data.append([col_name, value])
    df = pd.DataFrame(df_data, columns=['Field', 'Value'])
    return df

def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Customer Data')
    return output.getvalue()

# Sidebar
with st.sidebar:
    st.header("⚙️ Options")
    show_empty = st.checkbox("Show empty fields", value=True)
    show_stats = st.checkbox("Show statistics", value=True)
    st.markdown("---")
    st.markdown("### 📊 Info")
    st.info(f"Total fields: {len(OUTPUT_COLUMNS)}")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type=['pdf'])

if uploaded_file is not None:
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    if st.button("🚀 Extract Data", type="primary", use_container_width=True):
        with st.spinner("Extracting data from PDF..."):
            try:
                uploaded_file.seek(0)
                extracted_data = extract_data_from_pdf(uploaded_file)
                df = convert_to_dataframe(extracted_data)
                
                if not show_empty:
                    df = df[df['Value'] != ""]
                
                filled_count = len(df[df['Value'] != ""])
                st.success(f"✅ Successfully extracted data!")
                
                if show_stats:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Fields", len(OUTPUT_COLUMNS))
                    with col2:
                        st.metric("Fields Filled", filled_count)
                    with col3:
                        completion = (filled_count / len(OUTPUT_COLUMNS)) * 100
                        st.metric("Completion", f"{completion:.1f}%")
                
                st.subheader("📋 Extracted Data")
                st.dataframe(df, use_container_width=True, height=500)
                
                st.subheader("📥 Download Options")
                col1, col2 = st.columns(2)
                
                with col1:
                    csv_data = convert_df_to_csv(df)
                    st.download_button(
                        "Download as CSV",
                        csv_data,
                        f"extracted_customer_data_{uploaded_file.name.replace('.pdf', '')}.csv",
                        "text/csv"
                    )
                
                with col2:
                    excel_data = convert_df_to_excel(df)
                    st.download_button(
                        "Download as Excel",
                        excel_data,
                        f"extracted_customer_data_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                with st.expander("Error Details"):
                    st.code(str(e))

else:
    st.info("👆 Upload a PDF file to begin extraction")
    
    with st.expander("📖 Instructions"):
        st.markdown("""
        ### How to use:
        1. Upload your customer creation PDF file
        2. Click "Extract Data"
        3. Review the extracted data in the table
        4. Download as CSV or Excel
        
        ### Features:
        - ✅ Extracts all 94 customer fields
        - ✅ Handles empty fields correctly
        - ✅ Shows completion statistics
        - ✅ Export to CSV or Excel
        """)

st.markdown("---")
st.markdown("Built with ❤️ using Streamlit")
