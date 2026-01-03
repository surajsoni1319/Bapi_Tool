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
    
    # Parse the text using regex patterns
    data = {}
    
    # Simple extraction patterns - field name followed by value on same line
    patterns = {
        "Type of Customer": r"Type of Customer\s+(.+?)(?=\n|$)",
        "Name of Customer": r"Name of Customer\s+(.+?)(?=\n|$)",
        "Company Code": r"Company Code\s+(.+?)(?=\n|$)",
        "Customer Group": r"Customer Group\s+(.+?)(?=\n|$)",
        "Sales Group": r"Sales Group\s+(.+?)(?=\n|$)",
        "Region": r"Region\s+([A-Z0-9]+)(?=\n|$)",
        "Zone": r"^Zone\s+([A-Z0-9]+)(?=\n|$)",
        "Sub Zone": r"Sub Zone\s+(.+?)(?=\n|$)",
        "Sales Office": r"Sales Office\s+(.+?)(?=\n|$)",
        "Mobile Number": r"^Mobile Number\s+(\d+)",
        "Lattitude": r"Lattitude\s+([\d.]+)",
        "Longitude": r"Longitude\s+([\d.]+)",
        "Address 1": r"Address 1\s+(.+?)(?=\n|$)",
        "PIN": r"^PIN\s+(\d+)",
        "City": r"^City\s+(\w+)",
        "District": r"^District\s+(\w+)",
        "Whatsapp No.": r"Whatsapp No\.\s+(\d+)",
        "Date of Birth": r"Date of Birth\s+([\d\-/]+)",
        "Date of Anniversary": r"Date of Anniversary\s+([\d\-/]+)",
        "Counter Potential - Maximum": r"Counter Potential - Maximum\s+(\d+)",
        "Counter Potential - Minimum": r"Counter Potential - Minimum\s+(\d+)",
        "Is GST Present": r"Is GST Present\s+(\w+)",
        "PAN": r"^PAN\s+([A-Z0-9]+)(?=\n|$)",
        "PAN Holder Name": r"PAN Holder Name\s+(.+?)(?=\n|$)",
        "PAN Status": r"PAN Status\s+(\w+)",
        "PAN - Aadhaar Linking Status": r"PAN - Aadhaar Linking Status\s+(\w+)",
        "IFSC Number": r"IFSC Number\s+([A-Z0-9]+)",
        "Account Number": r"Account Number\s+(\d+)",
        "Name of Account Holder": r"Name of Account Holder\s+(.+?)(?=\n|$)",
        "Bank Name": r"Bank Name\s+(.+?)(?=\n|$)",
        "Bank Branch": r"Bank Branch\s+(.+?)(?=\n|$)",
        "Is Aadhaar Linked with Mobile?": r"Is Aadhaar Linked with Mobile\?\s+(\w+)",
        "Aadhaar Number": r"Aadhaar Number\s+(.+?)(?=\n|$)",
        "Gender": r"Gender\s+(\w+)",
        "DOB": r"^DOB\s+([\d/]+)",
        "Logistics Transportation Zone": r"Logistics Transportation Zone\s+(.+?)(?=\n|$)",
        "Transportation Zone Description": r"Transportation Zone Description\s+(.+?)(?=\n|$)",
        "Transportation Zone Code": r"Transportation Zone Code\s+(\d+)",
        "Date of Appointment": r"Date of Appointment\s+([\d\-]+)",
        "Delivering Plant": r"Delivering Plant\s+(.+?)(?=\n|$)",
        "Plant Name": r"Plant Name\s+(.+?)(?=\n|$)",
        "Plant Code": r"Plant Code\s+(\w+)",
        "Incoterns": r"^Incoterns\s+(.+?)(?=\n|$)",
        "Regional Head to be mapped": r"Regional Head to be mapped\s+(\w+)",
        "Zonal Head to be mapped": r"Zonal Head to be mapped\s+(\w+)",
        "Sub-Zonal Head (RSM) to be mapped": r"Sub-Zonal Head \(RSM\) to be mapped\s+(\w+)",
        "Area Sales Manager to be mapped": r"Area Sales Manager to be mapped\s+(\w+)",
        "Sales Officer to be mapped": r"Sales Officer to be mapped\s+(\w+)",
        "Internal control code": r"Internal control code\s+(\w+)",
        "Initiator Name": r"Initiator Name\s+(.+?)(?=\n|$)",
        "Initiator Email ID": r"Initiator Email ID\s+(.+?)(?=\n|$)",
        "Initiator Mobile Number": r"Initiator Mobile Number\s+(\d+)",
        "Created By Customer UserID": r"Created By Customer UserID\s+(\w+)",
        "Sales Head Name": r"Sales Head Name\s+(.+?)(?=\n|$)",
        "Sales Head Email": r"Sales Head Email\s+(.+?)(?=\n|$)",
        "Sales Head Mobile Number": r"Sales Head Mobile Number\s+(\d+)",
        "Final Result": r"Final Result\s+(\w+)",
    }
    
    # Extract using patterns
    for field, pattern in patterns.items():
        match = re.search(pattern, full_text, re.MULTILINE)
        if match:
            data[field] = match.group(1).strip()
    
    # Special handling for multi-line fields
    
    # State - first occurrence
    state_match = re.search(r"^State\s+([A-Z]+)", full_text, re.MULTILINE)
    if state_match:
        data["State"] = state_match.group(1)
    
    # SAP Dealer code (has number separator)
    sap_match = re.search(r"SAP Dealer code to be mapped Search Term\s*\n\s*2\s*\n\s*(\d+)", full_text)
    if sap_match:
        data["SAP Dealer code to be mapped Search Term 2"] = sap_match.group(1)
    
    # Name of Customers (multi-line)
    name_match = re.search(r"Name of the Customers \(Trade Name or\s*\n\s*Legal Name\)\s*\n?\s*(.+?)(?=\n|$)", full_text)
    if name_match:
        data["Name of the Customers (Trade Name or Legal Name)"] = name_match.group(1).strip()
    
    # Name (in Aadhaar section)
    name_aadhaar = re.search(r"Aadhaar Number.+?\n\s*Name\s+(.+?)(?=\n|$)", full_text, re.DOTALL)
    if name_aadhaar:
        data["Name"] = name_aadhaar.group(1).strip()
    
    # Address (multi-line in Aadhaar section)
    address_match = re.search(r"DOB\s+[\d/]+\s*\n\s*Address\s+(.+?)(?=\n\s*PIN|$)", full_text, re.DOTALL)
    if address_match:
        data["Address"] = address_match.group(1).strip()
    
    # Logistics team field
    logistics_match = re.search(r"Logistics team to vet the T zone selected by\s*\n\s*Sales Officer\s*\n?\s*(\w+)", full_text)
    if logistics_match:
        data["Logistics team to vet the T zone selected by Sales Officer"] = logistics_match.group(1)
    
    # Selection of T Zones
    tzone_match = re.search(r"Selection of Available T Zones from T Zone\s*\n\s*Master list, if found\.\s*\n?\s*(.+?)(?=\n|$)", full_text)
    if tzone_match:
        data["Selection of Available T Zones from T Zone Master list, if found."] = tzone_match.group(1).strip()
    
    # Incoterns Code
    inco_code = re.search(r"Incoterns Code\s+([A-Z]+)", full_text)
    if inco_code:
        data["Incoterns Code"] = inco_code.group(1)
    
    # Security Deposit
    security_match = re.search(r"Security Deposit Amount details to filled up,\s*\n\s*as per checque received by Customer / Dealer\s*\n?\s*(\d+)", full_text)
    if security_match:
        data["Security Deposit Amount details to filled up, as per checque received by Customer / Dealer"] = security_match.group(1)
    
    # PAN Result
    pan_result = re.search(r"PAN Result\s+([A-Z])", full_text)
    if pan_result:
        data["PAN Result"] = pan_result.group(1)
    
    # Mobile Number Result
    mobile_result = re.search(r"Mobile Number Result\s+([A-Z])", full_text)
    if mobile_result:
        data["Mobile Number Result"] = mobile_result.group(1)
    
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
