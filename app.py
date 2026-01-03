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

# Field mapping - exact field names as they appear in PDF and their output names
FIELD_MAPPING = {
    "Type of Customer": "Type of Customer",
    "Name of Customer": "Name of Customer",
    "Company Code": "Company Code",
    "Customer Group": "Customer Group",
    "Sales Group": "Sales Group",
    "Region": "Region",
    "Zone": "Zone",
    "Sub Zone": "Sub Zone",
    "State": "State",
    "Sales Office": "Sales Office",
    "SAP Dealer code to be mapped Search Term": "SAP Dealer code to be mapped Search Term 2",
    "Search Term 1- Old customer code": "Search Term 1- Old customer code",
    "Search Term 2 - District": "Search Term 2 - District",
    "Mobile Number": "Mobile Number",
    "E-Mail ID": "E-Mail ID",
    "Lattitude": "Lattitude",
    "Longitude": "Longitude",
    "Name of the Customers (Trade Name or": "Name of the Customers (Trade Name or Legal Name)",
    "E-mail": "E-mail",
    "Address 1": "Address 1",
    "Address 2": "Address 2",
    "Address 3": "Address 3",
    "Address 4": "Address 4",
    "PIN": "PIN",
    "City": "City",
    "District": "District",
    "Whatsapp No.": "Whatsapp No.",
    "Date of Birth": "Date of Birth",
    "Date of Anniversary": "Date of Anniversary",
    "Counter Potential - Maximum": "Counter Potential - Maximum",
    "Counter Potential - Minimum": "Counter Potential - Minimum",
    "Is GST Present": "Is GST Present",
    "GSTIN": "GSTIN",
    "Trade Name": "Trade Name",
    "Legal Name": "Legal Name",
    "Reg Date": "Reg Date",
    "Type": "Type",
    "Building No.": "Building No.",
    "District Code": "District Code",
    "State Code": "State Code",
    "Street": "Street",
    "PIN Code": "PIN Code",
    "PAN": "PAN",
    "PAN Holder Name": "PAN Holder Name",
    "PAN Status": "PAN Status",
    "PAN - Aadhaar Linking Status": "PAN - Aadhaar Linking Status",
    "IFSC Number": "IFSC Number",
    "Account Number": "Account Number",
    "Name of Account Holder": "Name of Account Holder",
    "Bank Name": "Bank Name",
    "Bank Branch": "Bank Branch",
    "Is Aadhaar Linked with Mobile?": "Is Aadhaar Linked with Mobile?",
    "Aadhaar Number": "Aadhaar Number",
    "Name": "Name",
    "Gender": "Gender",
    "DOB": "DOB",
    "Address": "Address",
    "Logistics Transportation Zone": "Logistics Transportation Zone",
    "Transportation Zone Description": "Transportation Zone Description",
    "Transportation Zone Code": "Transportation Zone Code",
    "Postal Code": "Postal Code",
    "Logistics team to vet the T zone selected by": "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone": "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to": "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment": "Date of Appointment",
    "Delivering Plant": "Delivering Plant",
    "Plant Name": "Plant Name",
    "Plant Code": "Plant Code",
    "Incoterns": "Incoterns",
    "Incoterns Code": "Incoterns Code",
    "Security Deposit Amount details to filled up,": "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)": "Credit Limit (In Rs.)",
    "Regional Head to be mapped": "Regional Head to be mapped",
    "Zonal Head to be mapped": "Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped": "Sub-Zonal Head (RSM) to be mapped",
    "Area Sales Manager to be mapped": "Area Sales Manager to be mapped",
    "Sales Officer to be mapped": "Sales Officer to be mapped",
    "Sales Promoter to be mapped": "Sales Promoter to be mapped",
    "Sales Promoter Number": "Sales Promoter Number",
    "Internal control code": "Internal control code",
    "SAP CODE": "SAP CODE",
    "Initiator Name": "Initiator Name",
    "Initiator Email ID": "Initiator Email ID",
    "Initiator Mobile Number": "Initiator Mobile Number",
    "Created By Customer UserID": "Created By Customer UserID",
    "Sales Head Name": "Sales Head Name",
    "Sales Head Email": "Sales Head Email",
    "Sales Head Mobile Number": "Sales Head Mobile Number",
    "Extra2": "Extra2",
    "PAN Result": "PAN Result",
    "Mobile Number Result": "Mobile Number Result",
    "Email Result": "Email Result",
    "GST Result": "GST Result",
    "Final Result": "Final Result"
}

# All output column names in order
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
    data = {}
    
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    full_text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        full_text += page.get_text()
    
    pdf_document.close()
    
    lines = full_text.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        # Skip empty lines and headers
        if not line or line in EXCLUDE_HEADERS:
            i += 1
            continue
        
        # Try to match with each field in our mapping
        matched = False
        for field_key, field_output in FIELD_MAPPING.items():
            if line.startswith(field_key):
                # Extract value after the field name
                value = line[len(field_key):].strip()
                
                # Handle multi-line fields
                if field_key == "Name of the Customers (Trade Name or":
                    # Next line has "Legal Name)" and then the value
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if next_line.startswith("Legal Name)"):
                            value = next_line[len("Legal Name)"):].strip()
                            i += 1
                
                elif field_key == "SAP Dealer code to be mapped Search Term":
                    # Next line might have "2" and then the code
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if next_line.isdigit():
                            if i + 2 < len(lines):
                                value = lines[i + 2].strip()
                                i += 2
                            else:
                                i += 1
                
                elif field_key in ["Logistics team to vet the T zone selected by",
                                  "Selection of Available T Zones from T Zone",
                                  "If NEW T Zone need to be created, details to",
                                  "Security Deposit Amount details to filled up,"]:
                    # Multi-line field names - get value from next line
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        # Skip the rest of field name if present
                        if not any(next_line.startswith(k) for k in FIELD_MAPPING.keys()):
                            if i + 2 < len(lines):
                                value = lines[i + 2].strip()
                                i += 2
                            else:
                                i += 1
                
                # Store the value
                if field_output not in data or not data[field_output]:
                    data[field_output] = value
                
                matched = True
                break
        
        if not matched:
            # Try exact pattern matching for fields with values on same line
            for field_key, field_output in FIELD_MAPPING.items():
                pattern = re.escape(field_key) + r'\s+(.*?)$'
                match = re.match(pattern, line)
                if match:
                    value = match.group(1).strip()
                    if field_output not in data or not data[field_output]:
                        data[field_output] = value
                    break
        
        i += 1
    
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
        - ✅ Handles multi-line field names
        - ✅ Shows completion statistics
        - ✅ Export to CSV or Excel
        """)

st.markdown("---")
st.markdown("Built with ❤️ using Streamlit")
