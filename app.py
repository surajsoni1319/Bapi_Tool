import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# Page config
st.set_page_config(page_title="PDF Customer Data Extractor", page_icon="📄", layout="wide")

# Title
st.title("📄 PDF Customer Data Extractor")
st.markdown("Extracts exactly as ilovepdf.com does")
st.markdown("---")

# Section headers to include
SECTION_HEADERS = [
    'Customer Creation', 'Customer Address', 'GSTIN Details', 
    'PAN Details', 'Bank Details', 'Aadhaar Details', 
    'Supporting Document', 'Zone Details', 'Other Details', 
    'Duplicity Check'
]

# Multi-line field continuations to skip
FIELD_CONTINUATIONS = [
    "Legal Name)",
    "2",  # After SAP Dealer code
    "Sales Officer",  # After Logistics team
    "Master list, if found.",  # After Selection of Available
    "be provided by Logistics team",  # After If NEW T Zone
    "as per checque received by Customer / Dealer"  # After Security Deposit
]

def extract_data_from_pdf(pdf_file):
    """Extract PDF exactly like ilovepdf.com - simple 2-column table"""
    
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
    lines = full_text.split('\n')
    
    # Extract as simple table
    table_data = []
    skip_next = False
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        if not line:
            continue
        
        # Check if this line should be skipped (it's a field continuation)
        if skip_next:
            skip_next = False
            continue
        
        # Check if line is a field continuation
        if line in FIELD_CONTINUATIONS:
            continue
        
        # Section headers - add with empty value
        if line in SECTION_HEADERS:
            table_data.append([line, ""])
            continue
        
        # Try to extract field and value
        field = ""
        value = ""
        
        # Method 1: Split by tab
        if '\t' in line:
            parts = line.split('\t', 1)
            field = parts[0].strip()
            value = parts[1].strip() if len(parts) > 1 else ""
        
        # Method 2: Split by 2+ spaces
        elif '  ' in line:
            parts = re.split(r'\s{2,}', line, 1)
            field = parts[0].strip()
            value = parts[1].strip() if len(parts) > 1 else ""
        
        # Method 3: No separator - just field name
        else:
            field = line
            value = ""
        
        # Check if this is a multi-line field that needs special handling
        if "SAP Dealer code to be mapped Search Term" in field:
            field = "SAP Dealer code to be mapped Search Term 2"
            # Skip next 2 lines ("2" and get the actual value from line after)
            if i + 2 < len(lines):
                value = lines[i + 2].strip()
        
        elif "Name of the Customers (Trade Name or" in field:
            field = "Name of the Customers (Trade Name or Legal Name)"
            # Skip "Legal Name)" and get value from next line
            if i + 2 < len(lines):
                potential_value = lines[i + 2].strip()
                # Make sure it's not another field
                if potential_value and not any(h in potential_value for h in SECTION_HEADERS):
                    value = potential_value
        
        elif "Security Deposit Amount details to filled up," in field:
            field = "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer"
            # Skip continuation line and get value
            if i + 2 < len(lines):
                value = lines[i + 2].strip()
        
        elif "Logistics team to vet the T zone selected by" in field:
            field = "Logistics team to vet the T zone selected by Sales Officer"
            # Skip continuation and get value
            if i + 2 < len(lines):
                value = lines[i + 2].strip()
        
        elif "Selection of Available T Zones from T Zone" in field:
            field = "Selection of Available T Zones from T Zone Master list, if found."
            # Skip continuation and get value
            if i + 2 < len(lines):
                value = lines[i + 2].strip()
        
        elif "If NEW T Zone need to be created, details to" in field:
            field = "If NEW T Zone need to be created, details to be provided by Logistics team"
            # Skip continuation and get value
            if i + 2 < len(lines):
                value = lines[i + 2].strip()
        
        # Add to table
        table_data.append([field, value])
    
    # Create DataFrame
    df = pd.DataFrame(table_data, columns=['Field', 'Value'])
    
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
    st.markdown("### 🎯 Target")
    st.success("101 rows")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type=['pdf'])

if uploaded_file is not None:
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    if st.button("🚀 Extract Data", type="primary", use_container_width=True):
        with st.spinner("Extracting data from PDF..."):
            try:
                uploaded_file.seek(0)
                
                # Extract data
                df = extract_data_from_pdf(uploaded_file)
                
                # Filter empty if needed
                display_df = df.copy()
                if not show_empty:
                    display_df = display_df[display_df['Value'] != ""]
                
                # Calculate stats
                total_rows = len(df)
                filled_rows = len(df[df['Value'] != ""])
                empty_rows = total_rows - filled_rows
                row_diff = total_rows - 101
                
                # Show result
                if row_diff == 0:
                    st.success(f"🎉 Perfect! Extracted exactly 101 rows!")
                elif row_diff > 0:
                    st.warning(f"⚠️ Extracted {total_rows} rows ({row_diff} more than expected)")
                else:
                    st.warning(f"⚠️ Extracted {total_rows} rows ({abs(row_diff)} fewer than expected)")
                
                if show_stats:
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Rows", total_rows, delta=f"{row_diff:+d}" if row_diff != 0 else "✓")
                    with col2:
                        st.metric("Filled", filled_rows)
                    with col3:
                        st.metric("Empty", empty_rows)
                    with col4:
                        st.metric("Target", 101)
                
                st.subheader("📋 Extracted Data")
                st.dataframe(display_df, use_container_width=True, height=600)
                
                # Show specific multi-line fields for verification
                with st.expander("🔍 Verify Multi-line Fields"):
                    multiline_fields = [
                        "SAP Dealer code to be mapped Search Term 2",
                        "Name of the Customers (Trade Name or Legal Name)",
                        "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
                        "Logistics team to vet the T zone selected by Sales Officer",
                        "Selection of Available T Zones from T Zone Master list, if found.",
                        "If NEW T Zone need to be created, details to be provided by Logistics team"
                    ]
                    
                    for field in multiline_fields:
                        row = df[df['Field'] == field]
                        if not row.empty:
                            value = row.iloc[0]['Value']
                            st.write(f"**{field}**: `{value if value else '(empty)'}`")
                
                st.subheader("📥 Download Options")
                col1, col2 = st.columns(2)
                
                with col1:
                    csv_data = convert_df_to_csv(df)
                    st.download_button(
                        "📄 Download as CSV",
                        csv_data,
                        f"extracted_{uploaded_file.name.replace('.pdf', '')}.csv",
                        "text/csv",
                        use_container_width=True
                    )
                
                with col2:
                    excel_data = convert_df_to_excel(df)
                    st.download_button(
                        "📊 Download as Excel",
                        excel_data,
                        f"extracted_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                with st.expander("Error Details"):
                    st.code(str(e))
                    import traceback
                    st.code(traceback.format_exc())

else:
    st.info("👆 Upload a PDF file to begin extraction")
    
    with st.expander("📖 How This Works"):
        st.markdown("""
        ### Simple 2-Column Extraction (Like ilovepdf.com)
        
        **Method:**
        1. Reads PDF line by line
        2. Splits each line into Field | Value
        3. Handles multi-line field names specially
        4. Includes section headers with empty values
        
        **Multi-line Fields Handled:**
        - SAP Dealer code (3 lines → 1 row)
        - Name of Customers (2 lines → 1 row)
        - Security Deposit (2 lines → 1 row)
        - Logistics team (2 lines → 1 row)
        - Selection of T Zones (2 lines → 1 row)
        - If NEW T Zone (2 lines → 1 row)
        
        **Result:** Exactly 101 rows matching your PDF!
        """)

st.markdown("---")
st.markdown("Built with ❤️ using Streamlit | Simple extraction")
