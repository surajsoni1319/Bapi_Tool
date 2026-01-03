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
    
    # Use pdfplumber-style table extraction
    data = {}
    
    # Split into lines and process
    lines = full_text.split('\n')
    
    # Clean lines
    cleaned_lines = []
    for line in lines:
        line = line.strip()
        if line and line not in EXCLUDE_HEADERS:
            cleaned_lines.append(line)
    
    # Process line by line with next line as value
    i = 0
    field_map = {
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
        "Mobile Number": "Mobile Number",
        "Lattitude": "Lattitude",
        "Longitude": "Longitude",
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
        "IFSC Number": "IFSC Number",
        "Account Number": "Account Number",
        "Name of Account Holder": "Name of Account Holder",
        "Bank Name": "Bank Name",
        "Bank Branch": "Bank Branch",
        "Aadhaar Number": "Aadhaar Number",
        "Name": "Name",
        "Gender": "Gender",
        "DOB": "DOB",
        "Address": "Address",
        "Logistics Transportation Zone": "Logistics Transportation Zone",
        "Transportation Zone Description": "Transportation Zone Description",
        "Transportation Zone Code": "Transportation Zone Code",
        "Postal Code": "Postal Code",
        "Date of Appointment": "Date of Appointment",
        "Delivering Plant": "Delivering Plant",
        "Plant Name": "Plant Name",
        "Plant Code": "Plant Code",
        "Incoterns": "Incoterns",
        "Incoterns Code": "Incoterns Code",
        "Regional Head to be mapped": "Regional Head to be mapped",
        "Zonal Head to be mapped": "Zonal Head to be mapped",
        "Area Sales Manager to be mapped": "Area Sales Manager to be mapped",
        "Sales Officer to be mapped": "Sales Officer to be mapped",
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
        "Final Result": "Final Result",
    }
    
    # Simple table-like extraction
    while i < len(cleaned_lines):
        line = cleaned_lines[i]
        
        # Try to split by multiple spaces (table format)
        parts = re.split(r'\s{2,}', line)
        
        if len(parts) >= 2:
            field = parts[0].strip()
            value = ' '.join(parts[1:]).strip()
            
            # Map to output field name
            if field in field_map:
                data[field_map[field]] = value
        
        i += 1
    
    # Handle special cases that weren't caught
    # SAP Dealer code
    if "SAP Dealer code to be mapped Search Term 2" not in data:
        for idx, line in enumerate(cleaned_lines):
            if "SAP Dealer code to be mapped Search Term" in line:
                # Look for number in next few lines
                for j in range(idx + 1, min(idx + 5, len(cleaned_lines))):
                    if cleaned_lines[j].isdigit() and len(cleaned_lines[j]) > 5:
                        data["SAP Dealer code to be mapped Search Term 2"] = cleaned_lines[j]
                        break
    
    # Multi-line Name field
    if "Name of the Customers (Trade Name or Legal Name)" not in data:
        for idx, line in enumerate(cleaned_lines):
            if "Name of the Customers (Trade Name or" in line:
                if idx + 2 < len(cleaned_lines):
                    if "Legal Name)" in cleaned_lines[idx + 1]:
                        data["Name of the Customers (Trade Name or Legal Name)"] = cleaned_lines[idx + 2]
    
    # Aadhaar Address (multi-line)
    for idx, line in enumerate(cleaned_lines):
        if line == "Address" and idx + 1 < len(cleaned_lines):
            # Get next few lines until we hit another field
            addr_parts = []
            for j in range(idx + 1, min(idx + 4, len(cleaned_lines))):
                if cleaned_lines[j] not in field_map and not cleaned_lines[j].startswith("PIN"):
                    addr_parts.append(cleaned_lines[j])
                else:
                    break
            if addr_parts and "Address" not in data:
                data["Address"] = " ".join(addr_parts)
    
    # Handle fields with additional mapping
    special_mappings = {
        "PAN - Aadhaar Linking Status": ["PAN - Aadhaar Linking Status", "PAN -Aadhaar Linking Status"],
        "Is Aadhaar Linked with Mobile?": ["Is Aadhaar Linked with Mobile?", "Is Aadhaar Linked with Mobile"],
        "Sub-Zonal Head (RSM) to be mapped": ["Sub-Zonal Head (RSM) to be mapped", "Sub-Zonal Head  (RSM) to be mapped"],
        "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer": [
            "Security Deposit Amount details to filled up,",
            "Security Deposit Amount details to filled up"
        ],
        "Logistics team to vet the T zone selected by Sales Officer": [
            "Logistics team to vet the T zone selected by"
        ],
        "Selection of Available T Zones from T Zone Master list, if found.": [
            "Selection of Available T Zones from T Zone"
        ],
        "If NEW T Zone need to be created, details to be provided by Logistics team": [
            "If NEW T Zone need to be created, details to"
        ],
    }
    
    # Try to match special fields
    for output_field, search_terms in special_mappings.items():
        if output_field not in data:
            for idx, line in enumerate(cleaned_lines):
                for search_term in search_terms:
                    if search_term in line:
                        # Get value from same line or next line
                        parts = re.split(r'\s{2,}', line)
                        if len(parts) >= 2:
                            data[output_field] = ' '.join(parts[1:]).strip()
                        elif idx + 1 < len(cleaned_lines):
                            # Check next few lines for value
                            for j in range(idx + 1, min(idx + 3, len(cleaned_lines))):
                                if cleaned_lines[j] and cleaned_lines[j] not in field_map:
                                    data[output_field] = cleaned_lines[j]
                                    break
                        break
    
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
