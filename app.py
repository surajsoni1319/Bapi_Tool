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

# All expected fields in order as they appear in PDF
EXPECTED_FIELDS = [
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
    "Mobile Number",  # Second occurrence
    "E-mail",
    "Address 1",
    "Address 2",
    "Address 3",
    "Address 4",
    "PIN",
    "City",
    "District",
    "State",  # Second occurrence
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
    "City",  # Second occurrence
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
    "PIN",  # Third occurrence
    "City",  # Third occurrence
    "State",  # Third occurrence
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
    "Incoterns",  # Second occurrence
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

SECTION_HEADERS = [
    'Customer Creation', 'Customer Address', 'GSTIN Details', 
    'PAN Details', 'Bank Details', 'Aadhaar Details', 
    'Supporting Document', 'Zone Details', 'Other Details', 
    'Duplicity Check'
]

def extract_data_from_pdf(pdf_file):
    """Extract data using simple line matching"""
    
    # Read PDF
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Get all text
    full_text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        full_text += page.get_text()
    
    pdf_document.close()
    
    # Split by tabs and multiple spaces to get key-value pairs
    lines = full_text.split('\n')
    
    # Create raw table
    raw_data = []
    for line in lines:
        line = line.strip()
        if line and line not in SECTION_HEADERS:
            # Split by tab or 2+ spaces
            if '\t' in line:
                parts = line.split('\t')
            else:
                parts = re.split(r'\s{2,}', line)
            
            if len(parts) >= 2:
                field = parts[0].strip()
                value = '\t'.join(parts[1:]).strip() if '\t' in line else ' '.join(parts[1:]).strip()
                raw_data.append({'Field': field, 'Value': value})
            elif len(parts) == 1:
                raw_data.append({'Field': parts[0].strip(), 'Value': ''})
    
    # Create DataFrame from raw data
    df = pd.DataFrame(raw_data)
    
    # Now map to expected output format
    output_data = []
    field_counts = {}  # Track occurrences of duplicate fields
    
    for expected_field in EXPECTED_FIELDS:
        # Handle duplicate fields
        base_field = expected_field
        
        # For duplicate tracking
        if expected_field in field_counts:
            field_counts[expected_field] += 1
            occurrence = field_counts[expected_field]
        else:
            field_counts[expected_field] = 1
            occurrence = 1
        
        # Find matching row in extracted data
        value = ""
        
        # Special handling for specific fields
        if expected_field == "SAP Dealer code to be mapped Search Term 2":
            # Look for SAP Dealer code pattern
            matches = df[df['Field'].str.contains('SAP Dealer code', na=False, case=False)]
            if not matches.empty:
                value = matches.iloc[0]['Value']
            # Also check for standalone number after "2"
            if not value:
                for i, row in df.iterrows():
                    if row['Field'] == '2' and i + 1 < len(df):
                        value = df.iloc[i + 1]['Field']
                        break
        
        elif expected_field == "Name of the Customers (Trade Name or Legal Name)":
            # Multi-line field
            for i, row in df.iterrows():
                if 'Name of the Customers' in row['Field']:
                    # Check next few rows
                    if i + 1 < len(df) and 'Legal Name' in df.iloc[i + 1]['Field']:
                        value = df.iloc[i + 1]['Value']
                        if not value and i + 2 < len(df):
                            value = df.iloc[i + 2]['Field']
                    break
        
        elif expected_field in ["Logistics team to vet the T zone selected by Sales Officer",
                                "Selection of Available T Zones from T Zone Master list, if found.",
                                "If NEW T Zone need to be created, details to be provided by Logistics team",
                                "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer"]:
            # Multi-line field names
            search_terms = {
                "Logistics team to vet the T zone selected by Sales Officer": "Logistics team to vet",
                "Selection of Available T Zones from T Zone Master list, if found.": "Selection of Available T Zones",
                "If NEW T Zone need to be created, details to be provided by Logistics team": "If NEW T Zone",
                "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer": "Security Deposit Amount"
            }
            search_term = search_terms.get(expected_field, expected_field)
            matches = df[df['Field'].str.contains(search_term, na=False, case=False)]
            if not matches.empty:
                value = matches.iloc[0]['Value']
        
        else:
            # Standard field matching - handle duplicates
            matches = df[df['Field'] == base_field]
            if not matches.empty:
                if occurrence <= len(matches):
                    value = matches.iloc[occurrence - 1]['Value']
                else:
                    value = matches.iloc[0]['Value']
        
        output_data.append({'Field': expected_field, 'Value': value})
    
    return pd.DataFrame(output_data), df

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
    show_raw_extraction = st.checkbox("Show raw extraction", value=False)
    show_empty = st.checkbox("Show empty fields", value=True)
    show_stats = st.checkbox("Show statistics", value=True)
    st.markdown("---")
    st.markdown("### 📊 Info")
    st.info(f"Total fields: {len(EXPECTED_FIELDS)}")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type=['pdf'])

if uploaded_file is not None:
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    if st.button("🚀 Extract Data", type="primary", use_container_width=True):
        with st.spinner("Extracting data from PDF..."):
            try:
                uploaded_file.seek(0)
                
                # Extract data
                final_df, raw_df = extract_data_from_pdf(uploaded_file)
                
                # Show raw extraction if requested
                if show_raw_extraction:
                    st.subheader("🔍 Raw Extraction (Debug)")
                    st.dataframe(raw_df, use_container_width=True, height=400)
                    st.markdown("---")
                
                # Filter empty if needed
                display_df = final_df.copy()
                if not show_empty:
                    display_df = display_df[display_df['Value'] != ""]
                
                filled_count = len(final_df[final_df['Value'] != ""])
                total_count = len(final_df)
                
                st.success(f"✅ Successfully extracted data!")
                
                if show_stats:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Fields", total_count)
                    with col2:
                        st.metric("Fields Filled", filled_count)
                    with col3:
                        completion = (filled_count / total_count) * 100 if total_count > 0 else 0
                        st.metric("Completion", f"{completion:.1f}%")
                
                st.subheader("📋 Extracted Data")
                st.dataframe(display_df, use_container_width=True, height=500)
                
                st.subheader("📥 Download Options")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    csv_data = convert_df_to_csv(final_df)
                    st.download_button(
                        "Download Final CSV",
                        csv_data,
                        f"extracted_{uploaded_file.name.replace('.pdf', '')}.csv",
                        "text/csv"
                    )
                
                with col2:
                    excel_data = convert_df_to_excel(final_df)
                    st.download_button(
                        "Download Final Excel",
                        excel_data,
                        f"extracted_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col3:
                    raw_csv = convert_df_to_csv(raw_df)
                    st.download_button(
                        "Download Raw Extraction (Debug)",
                        raw_csv,
                        f"raw_{uploaded_file.name.replace('.pdf', '')}.csv",
                        "text/csv"
                    )
                
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                with st.expander("Error Details"):
                    st.code(str(e))
                    import traceback
                    st.code(traceback.format_exc())

else:
    st.info("👆 Upload a PDF file to begin extraction")
    
    with st.expander("📖 Instructions"):
        st.markdown("""
        ### How to use:
        1. Upload your customer creation PDF file
        2. Click "Extract Data"
        3. Enable "Show raw extraction" to see what was extracted
        4. Review the final data in the table
        5. Download as CSV or Excel
        
        ### Features:
        - ✅ Extracts all 101 fields (including duplicates)
        - ✅ Handles duplicate field names (Mobile Number, State, City, PIN, etc.)
        - ✅ Shows raw extraction for debugging
        - ✅ Exports to CSV or Excel
        
        ### Debug Tips:
        - Enable "Show raw extraction" to see field-value pairs as extracted
        - Compare with your PDF to identify issues
        - Download raw extraction to analyze in Excel
        """)

st.markdown("---")
st.markdown("Built with ❤️ using Streamlit")
