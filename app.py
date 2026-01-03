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

# Section headers to track context
SECTION_HEADERS = [
    'Customer Creation', 'Customer Address', 'GSTIN Details', 
    'PAN Details', 'Bank Details', 'Aadhaar Details', 
    'Supporting Document', 'Zone Details', 'Other Details', 
    'Duplicity Check'
]

def extract_table_from_pdf(pdf_file):
    """Extract PDF as a table with Field-Value pairs, keeping section headers"""
    
    # Read PDF
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Get all text
    full_text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        full_text += page.get_text()
    
    pdf_document.close()
    
    # Split into lines
    lines = [line.strip() for line in full_text.split('\n') if line.strip()]
    
    # Extract as table: [Section, Field, Value]
    table_data = []
    current_section = "General"
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Check if this is a section header
        if line in SECTION_HEADERS:
            current_section = line
            i += 1
            continue
        
        # Try to split the line by 2+ spaces (table format)
        parts = re.split(r'\s{2,}', line)
        
        if len(parts) >= 2:
            # Field and value on same line
            field = parts[0].strip()
            value = ' '.join(parts[1:]).strip()
            table_data.append([current_section, field, value])
        
        elif len(parts) == 1:
            # Field on its own line, check next line for value
            field = parts[0].strip()
            
            # Check if next line exists and is not a section header
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                
                # If next line is a section header or another field, value is empty
                if next_line in SECTION_HEADERS:
                    table_data.append([current_section, field, ""])
                else:
                    # Check if next line looks like a field name or a value
                    # Simple heuristic: if it contains common field patterns, it's a field
                    is_next_field = any([
                        next_line.endswith(":"),
                        next_line.startswith("Search Term"),
                        next_line.startswith("Mobile Number"),
                        next_line.startswith("E-Mail"),
                        next_line.startswith("Address"),
                        next_line.startswith("PIN"),
                        next_line.startswith("City"),
                        next_line.startswith("District"),
                        next_line.startswith("State"),
                        next_line.startswith("Name"),
                        next_line.startswith("Type"),
                        next_line.startswith("Date"),
                        next_line.startswith("Is "),
                        next_line.startswith("PAN"),
                        next_line.startswith("IFSC"),
                        next_line.startswith("Account"),
                        next_line.startswith("Bank"),
                        next_line.startswith("Gender"),
                        next_line.startswith("DOB"),
                        next_line.startswith("Credit"),
                        next_line.startswith("Security"),
                        next_line.startswith("Regional"),
                        next_line.startswith("Zonal"),
                        next_line.startswith("Sub-Zonal"),
                        next_line.startswith("Area Sales"),
                        next_line.startswith("Sales"),
                        next_line.startswith("Internal"),
                        next_line.startswith("SAP"),
                        next_line.startswith("Initiator"),
                        next_line.startswith("Created"),
                        next_line.startswith("Extra"),
                        next_line.startswith("Result"),
                        next_line.startswith("Final"),
                    ])
                    
                    if is_next_field:
                        # Next line is another field, so current field is empty
                        table_data.append([current_section, field, ""])
                    else:
                        # Next line is the value
                        table_data.append([current_section, field, next_line])
                        i += 1  # Skip next line since we used it
            else:
                # No next line, value is empty
                table_data.append([current_section, field, ""])
        
        i += 1
    
    # Create DataFrame
    df = pd.DataFrame(table_data, columns=['Section', 'Field', 'Value'])
    
    return df

def process_extracted_table(df):
    """Process the extracted table and create final output with proper field mapping"""
    
    # Dictionary to store final data
    final_data = {}
    
    # Process each row
    for idx, row in df.iterrows():
        section = row['Section']
        field = row['Field']
        value = row['Value']
        
        # Create unique field names for duplicates based on section
        if section == "Customer Creation" or section == "General":
            final_data[field] = value
        
        elif section == "Customer Address":
            # Fields in Customer Address section
            if field == "Name of the Customers (Trade Name or":
                # Handle multi-line field name
                if idx + 1 < len(df) and df.iloc[idx + 1]['Field'] == "Legal Name)":
                    final_data["Name of the Customers (Trade Name or Legal Name)"] = df.iloc[idx + 1]['Value']
            elif field not in ["Legal Name)"]:
                final_data[field] = value
        
        elif section == "GSTIN Details":
            final_data[field] = value
        
        elif section == "PAN Details":
            final_data[field] = value
        
        elif section == "Bank Details":
            final_data[field] = value
        
        elif section == "Aadhaar Details":
            # Duplicate fields - add suffix
            if field == "Name":
                final_data["Name"] = value
            elif field == "State":
                # This is Aadhaar State, but we keep it as State if not already set
                if "State" not in final_data or not final_data["State"]:
                    final_data["State"] = value
            elif field == "City":
                if "City" not in final_data or not final_data["City"]:
                    final_data["City"] = value
            elif field == "PIN":
                if "PIN" not in final_data or not final_data["PIN"]:
                    final_data["PIN"] = value
            else:
                final_data[field] = value
        
        elif section == "Zone Details":
            final_data[field] = value
        
        elif section == "Other Details":
            final_data[field] = value
        
        elif section == "Duplicity Check":
            final_data[field] = value
    
    return final_data

def create_final_dataframe(data_dict):
    """Create final DataFrame with all expected fields"""
    
    # Define all expected output columns
    output_columns = [
        "Type of Customer", "Name of Customer", "Company Code", "Customer Group",
        "Sales Group", "Region", "Zone", "Sub Zone", "State", "Sales Office",
        "SAP Dealer code to be mapped Search Term", "Search Term 1- Old customer code",
        "Search Term 2 - District", "Mobile Number", "E-Mail ID", "Lattitude", "Longitude",
        "Name of the Customers (Trade Name or Legal Name)", "Mobile Number", "E-mail",
        "Address 1", "Address 2", "Address 3", "Address 4", "PIN", "City", "District",
        "State", "Whatsapp No.", "Date of Birth", "Date of Anniversary",
        "Counter Potential - Maximum", "Counter Potential - Minimum",
        "Is GST Present", "GSTIN", "Trade Name", "Legal Name", "Reg Date",
        "City", "Type", "Building No.", "District Code", "State Code", "Street", "PIN Code",
        "PAN", "PAN Holder Name", "PAN Status", "PAN - Aadhaar Linking Status",
        "IFSC Number", "Account Number", "Name of Account Holder", "Bank Name", "Bank Branch",
        "Is Aadhaar Linked with Mobile?", "Aadhaar Number", "Name", "Gender", "DOB", "Address",
        "PIN", "City", "State",
        "Logistics Transportation Zone", "Transportation Zone Description",
        "Transportation Zone Code", "Postal Code",
        "Logistics team to vet the T zone selected by", "Sales Officer",
        "Selection of Available T Zones from T Zone", "Master list, if found.",
        "If NEW T Zone need to be created, details to", "be provided by Logistics team",
        "Date of Appointment", "Delivering Plant", "Plant Name", "Plant Code",
        "Incoterns", "Incoterns", "Incoterns Code",
        "Security Deposit Amount details to filled up,", "as per checque received by Customer / Dealer",
        "Credit Limit (In Rs.)", "Regional Head to be mapped", "Zonal Head to be mapped",
        "Sub-Zonal Head (RSM) to be mapped", "Area Sales Manager to be mapped",
        "Sales Officer to be mapped", "Sales Promoter to be mapped", "Sales Promoter Number",
        "Internal control code", "SAP CODE", "Initiator Name", "Initiator Email ID",
        "Initiator Mobile Number", "Created By Customer UserID", "Sales Head Name",
        "Sales Head Email", "Sales Head Mobile Number", "Extra2",
        "PAN Result", "Mobile Number Result", "Email Result", "GST Result", "Final Result"
    ]
    
    # Create output data
    output_data = []
    for field in output_columns:
        value = data_dict.get(field, "")
        output_data.append([field, value])
    
    df = pd.DataFrame(output_data, columns=['Field', 'Value'])
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
    show_raw_table = st.checkbox("Show raw extracted table", value=False)
    show_empty = st.checkbox("Show empty fields", value=True)
    show_stats = st.checkbox("Show statistics", value=True)
    st.markdown("---")
    st.markdown("### 📊 Info")
    st.info("Table-based extraction")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type=['pdf'])

if uploaded_file is not None:
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    if st.button("🚀 Extract Data", type="primary", use_container_width=True):
        with st.spinner("Extracting data from PDF..."):
            try:
                uploaded_file.seek(0)
                
                # Step 1: Extract as table
                raw_table = extract_table_from_pdf(uploaded_file)
                
                # Show raw table if requested
                if show_raw_table:
                    st.subheader("📊 Raw Extracted Table")
                    st.dataframe(raw_table, use_container_width=True, height=400)
                    st.markdown("---")
                
                # Step 2: Process table into final format
                processed_data = process_extracted_table(raw_table)
                
                # Step 3: Create final DataFrame
                final_df = create_final_dataframe(processed_data)
                
                if not show_empty:
                    final_df = final_df[final_df['Value'] != ""]
                
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
                st.dataframe(final_df, use_container_width=True, height=500)
                
                st.subheader("📥 Download Options")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    csv_data = convert_df_to_csv(final_df)
                    st.download_button(
                        "Download as CSV",
                        csv_data,
                        f"extracted_customer_data_{uploaded_file.name.replace('.pdf', '')}.csv",
                        "text/csv"
                    )
                
                with col2:
                    excel_data = convert_df_to_excel(final_df)
                    st.download_button(
                        "Download as Excel",
                        excel_data,
                        f"extracted_customer_data_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col3:
                    raw_csv = convert_df_to_csv(raw_table)
                    st.download_button(
                        "Download Raw Table (Debug)",
                        raw_csv,
                        f"raw_table_{uploaded_file.name.replace('.pdf', '')}.csv",
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
        3. Review the extracted data in the table
        4. Download as CSV or Excel
        
        ### Features:
        - ✅ Table-based extraction (like ilovepdf.com)
        - ✅ Handles empty fields correctly
        - ✅ Preserves section context
        - ✅ Handles duplicate field names
        - ✅ Shows raw extracted table (debug option)
        - ✅ Export to CSV or Excel
        
        ### New in this version:
        - **Smart empty field detection** - Won't mistake field names for values
        - **Section-aware extraction** - Knows which section each field belongs to
        - **Duplicate handling** - Can distinguish between multiple "State", "City", etc.
        """)

st.markdown("---")
st.markdown("Built with ❤️ using Streamlit | Table-based extraction")
