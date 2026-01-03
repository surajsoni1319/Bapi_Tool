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

# Fixed column names
COLUMN_NAMES = [
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
    "E-mail",
    "Address 1",
    "Address 2",
    "Address 3",
    "Address 4",
    "PIN",
    "City",
    "District",
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

# Headers to exclude
EXCLUDE_HEADERS = [
    'Customer Creation', 'Customer Address', 'GSTIN Details', 
    'PAN Details', 'Bank Details', 'Aadhaar Details', 
    'Supporting Document', 'Zone Details', 'Other Details', 
    'Duplicity Check'
]

def extract_data_from_pdf(pdf_file):
    """Extract key-value pairs from PDF file using PyMuPDF"""
    data = {}
    
    # Read the PDF file
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Get all text from PDF
    full_text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        full_text += page.get_text()
    
    pdf_document.close()
    
    # Split into lines
    lines = full_text.split('\n')
    
    # Process each line
    for i, line in enumerate(lines):
        line = line.strip()
        
        # Skip empty lines and header lines
        if not line or line in EXCLUDE_HEADERS:
            continue
        
        # Check if this line matches any of our column names
        for col_name in COLUMN_NAMES:
            # Try exact match first
            if line.startswith(col_name):
                # Get the value after the column name
                value = line[len(col_name):].strip()
                
                # If value is empty, check next line
                if not value and i + 1 < len(lines):
                    value = lines[i + 1].strip()
                
                # Store the data
                if col_name not in data:  # Avoid duplicates
                    data[col_name] = value
                break
            
            # Try split by multiple spaces or tabs
            parts = re.split(r'\s{2,}|\t+', line, maxsplit=1)
            if len(parts) == 2:
                potential_key = parts[0].strip()
                if potential_key == col_name:
                    value = parts[1].strip()
                    if col_name not in data:
                        data[col_name] = value
                    break
    
    return data

def convert_to_dataframe(data_dict):
    """Convert dictionary to pandas DataFrame with all columns"""
    # Create DataFrame with all column names
    df_data = []
    for col_name in COLUMN_NAMES:
        value = data_dict.get(col_name, "")
        df_data.append([col_name, value])
    
    df = pd.DataFrame(df_data, columns=['Field', 'Value'])
    return df

def convert_df_to_csv(df):
    """Convert dataframe to CSV for download"""
    return df.to_csv(index=False).encode('utf-8')

def convert_df_to_excel(df):
    """Convert dataframe to Excel for download"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Customer Data')
    return output.getvalue()

# Sidebar options
with st.sidebar:
    st.header("⚙️ Options")
    show_empty = st.checkbox("Show empty fields", value=True)
    st.markdown("---")
    st.markdown("### 📊 Total Fields")
    st.info(f"{len(COLUMN_NAMES)} fields")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type=['pdf'])

if uploaded_file is not None:
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    # Add a button to process
    if st.button("Extract Data", type="primary"):
        with st.spinner("Extracting data from PDF..."):
            try:
                # Reset file pointer
                uploaded_file.seek(0)
                
                # Extract data
                extracted_data = extract_data_from_pdf(uploaded_file)
                
                # Convert to DataFrame
                df = convert_to_dataframe(extracted_data)
                
                # Filter empty fields if needed
                if not show_empty:
                    df = df[df['Value'] != ""]
                
                # Display success message
                filled_count = len(df[df['Value'] != ""])
                st.success(f"✅ Successfully extracted data! {filled_count} fields filled out of {len(COLUMN_NAMES)}")
                
                # Display the data
                st.subheader("Extracted Data")
                st.dataframe(df, use_container_width=True, height=600)
                
                # Download options
                st.subheader("📥 Download Options")
                col1, col2 = st.columns(2)
                
                with col1:
                    csv_data = convert_df_to_csv(df)
                    st.download_button(
                        label="Download as CSV",
                        data=csv_data,
                        file_name=f"extracted_customer_data_{uploaded_file.name.replace('.pdf', '')}.csv",
                        mime="text/csv",
                    )
                
                with col2:
                    excel_data = convert_df_to_excel(df)
                    st.download_button(
                        label="Download as Excel",
                        data=excel_data,
                        file_name=f"extracted_customer_data_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                
                # Show statistics
                st.subheader("📊 Statistics")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Total Fields", len(COLUMN_NAMES))
                
                with col2:
                    st.metric("Fields Filled", filled_count)
                
                with col3:
                    completion_rate = (filled_count / len(COLUMN_NAMES)) * 100
                    st.metric("Completion Rate", f"{completion_rate:.1f}%")
                
            except Exception as e:
                st.error(f"❌ Error processing PDF: {str(e)}")
                st.info("Please make sure the PDF format matches the expected structure.")
                with st.expander("Show error details"):
                    st.code(str(e))

else:
    # Instructions
    st.info("👆 Please upload a PDF file to begin extraction")
    
    with st.expander("📖 Instructions"):
        st.markdown("""
        ### How to use this tool:
        
        1. **Upload** a PDF file containing customer creation data
        2. Click the **Extract Data** button
        3. Review the extracted data in the table
        4. **Download** the results as CSV or Excel
        
        ### Features:
        - ✅ Extracts all 94 predefined fields
        - ✅ Shows completion statistics
        - ✅ Option to hide/show empty fields
        - ✅ Exports to CSV or Excel format
        - ✅ Handles multi-page PDFs
        
        ### Supported Fields:
        The tool extracts data for all standard customer creation fields including:
        - Customer information
        - Contact details
        - Address information
        - PAN, GST, Aadhaar details
        - Bank information
        - Sales and zone mapping
        - And more...
        """)
    
    with st.expander("📋 View All Field Names"):
        st.write(pd.DataFrame(COLUMN_NAMES, columns=['Field Name']))

# Footer
st.markdown("---")
st.markdown("Built with ❤️ using Streamlit | Powered by PyMuPDF")
