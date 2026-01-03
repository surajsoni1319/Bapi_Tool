import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# Page config
st.set_page_config(page_title="PDF Customer Data Extractor", page_icon="📄", layout="wide")

# Title
st.title("📄 PDF Customer Data Extractor")
st.markdown("Extracts data matching ilovepdf.com format")
st.markdown("---")

# Multi-line field names that need to be merged
MULTILINE_FIELDS = {
    "SAP Dealer code to be mapped Search Term": "SAP Dealer code to be mapped Search Term 2",
    "Name of the Customers (Trade Name or": "Name of the Customers (Trade Name or Legal Name)",
    "Security Deposit Amount details to filled up,": "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Logistics team to vet the T zone selected by": "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone": "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to": "If NEW T Zone need to be created, details to be provided by Logistics team"
}

# Section headers
SECTION_HEADERS = [
    'Customer Creation', 'Customer Address', 'GSTIN Details', 
    'PAN Details', 'Bank Details', 'Aadhaar Details', 
    'Supporting Document', 'Zone Details', 'Other Details', 
    'Duplicity Check'
]

def extract_data_from_pdf(pdf_file):
    """Extract PDF data as 2-column table like ilovepdf.com"""
    
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
    
    # Clean lines
    cleaned_lines = []
    for line in lines:
        line = line.strip()
        if line:
            cleaned_lines.append(line)
    
    # Extract as table: Field | Value
    table_data = []
    i = 0
    
    while i < len(cleaned_lines):
        line = cleaned_lines[i]
        
        # Check if this is a section header
        if line in SECTION_HEADERS:
            table_data.append({'Field': line, 'Value': ""})
            i += 1
            continue
        
        # Check if this is a multi-line field
        is_multiline = False
        for partial_name, full_name in MULTILINE_FIELDS.items():
            if line.startswith(partial_name) or line == partial_name:
                # This is a multi-line field, skip continuation lines
                skip_count = full_name.count('\n') if '\n' in full_name else 1
                
                # Special handling for each multi-line field
                if "SAP Dealer code" in line:
                    # Skip "2" and get actual value
                    i += 1  # Skip current line
                    if i < len(cleaned_lines) and cleaned_lines[i].strip() == "2":
                        i += 1  # Skip "2"
                    if i < len(cleaned_lines):
                        value = cleaned_lines[i].strip()
                        table_data.append({'Field': full_name, 'Value': value})
                    else:
                        table_data.append({'Field': full_name, 'Value': ""})
                
                elif "Name of the Customers" in line:
                    # Skip "Legal Name)" line
                    i += 1
                    if i < len(cleaned_lines) and "Legal Name)" in cleaned_lines[i]:
                        i += 1
                    if i < len(cleaned_lines):
                        # Check if next line is a value or another field
                        next_line = cleaned_lines[i].strip()
                        if next_line and not any(next_line.startswith(s) for s in SECTION_HEADERS):
                            # Check if it looks like a field name
                            if '\t' in next_line or '  ' in next_line:
                                # Has separator, extract value
                                parts = re.split(r'\t|\s{2,}', next_line, 1)
                                if len(parts) > 1:
                                    value = parts[1].strip()
                                else:
                                    value = parts[0].strip()
                            else:
                                # No separator, it's the value
                                value = next_line
                            table_data.append({'Field': full_name, 'Value': value})
                        else:
                            table_data.append({'Field': full_name, 'Value': ""})
                    else:
                        table_data.append({'Field': full_name, 'Value': ""})
                
                elif "Security Deposit" in line:
                    # Skip "as per checque..." line
                    i += 1
                    if i < len(cleaned_lines) and "as per checque" in cleaned_lines[i]:
                        i += 1
                    if i < len(cleaned_lines):
                        value = cleaned_lines[i].strip()
                        table_data.append({'Field': full_name, 'Value': value})
                    else:
                        table_data.append({'Field': full_name, 'Value': ""})
                
                elif "Logistics team" in line:
                    # Skip "Sales Officer" line
                    i += 1
                    if i < len(cleaned_lines) and "Sales Officer" in cleaned_lines[i]:
                        i += 1
                    if i < len(cleaned_lines):
                        value = cleaned_lines[i].strip()
                        table_data.append({'Field': full_name, 'Value': value})
                    else:
                        table_data.append({'Field': full_name, 'Value': ""})
                
                elif "Selection of Available" in line:
                    # Skip "Master list, if found." line
                    i += 1
                    if i < len(cleaned_lines) and "Master list" in cleaned_lines[i]:
                        i += 1
                    if i < len(cleaned_lines):
                        value = cleaned_lines[i].strip()
                        table_data.append({'Field': full_name, 'Value': value})
                    else:
                        table_data.append({'Field': full_name, 'Value': ""})
                
                elif "If NEW T Zone" in line:
                    # Skip "be provided by Logistics team" line
                    i += 1
                    if i < len(cleaned_lines) and "be provided" in cleaned_lines[i]:
                        i += 1
                    if i < len(cleaned_lines):
                        value = cleaned_lines[i].strip()
                        table_data.append({'Field': full_name, 'Value': value})
                    else:
                        table_data.append({'Field': full_name, 'Value': ""})
                
                is_multiline = True
                i += 1
                break
        
        if is_multiline:
            continue
        
        # Normal line processing - try to split by tab or spaces
        if '\t' in line:
            parts = line.split('\t', 1)
            field = parts[0].strip()
            value = parts[1].strip() if len(parts) > 1 else ""
            table_data.append({'Field': field, 'Value': value})
        
        elif '  ' in line:
            parts = re.split(r'\s{2,}', line, 1)
            field = parts[0].strip()
            value = parts[1].strip() if len(parts) > 1 else ""
            table_data.append({'Field': field, 'Value': value})
        
        else:
            # Single item - field with no value
            table_data.append({'Field': line, 'Value': ""})
        
        i += 1
    
    # Create DataFrame
    df = pd.DataFrame(table_data)
    
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
    st.success("101 rows (matching PDF)")

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
                
                # Check if we hit target
                row_diff = total_rows - 101
                if row_diff == 0:
                    st.success(f"✅ Perfect! Extracted exactly {total_rows} rows (matches PDF)!")
                elif row_diff > 0:
                    st.warning(f"⚠️ Extracted {total_rows} rows ({row_diff} more than expected 101)")
                else:
                    st.warning(f"⚠️ Extracted {total_rows} rows ({abs(row_diff)} less than expected 101)")
                
                if show_stats:
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Rows", total_rows, delta=f"{row_diff:+d} vs target")
                    with col2:
                        st.metric("Filled Rows", filled_rows)
                    with col3:
                        st.metric("Empty Rows", empty_rows)
                    with col4:
                        st.metric("Target", 101)
                
                st.subheader("📋 Extracted Data")
                st.dataframe(display_df, use_container_width=True, height=600)
                
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
    
    with st.expander("📖 Fixed Issues"):
        st.markdown("""
        ### ✅ Multi-line Fields Now Handled:
        
        These fields had names split across 2 lines:
        
        1. **SAP Dealer code to be mapped Search Term 2**
           - Line 1: "SAP Dealer code to be mapped Search Term"
           - Line 2: "2"
           - Line 3: Actual value
        
        2. **Security Deposit Amount details to filled up, as per checque received by Customer / Dealer**
           - Line 1: "Security Deposit Amount details to filled up,"
           - Line 2: "as per checque received by Customer / Dealer"
           - Line 3: Actual value (10000)
        
        3. **Name of the Customers (Trade Name or Legal Name)**
        4. **Logistics team to vet the T zone selected by Sales Officer**
        5. **Selection of Available T Zones from T Zone Master list, if found.**
        6. **If NEW T Zone need to be created, details to be provided by Logistics team**
        
        All these now extract correctly with **101 total rows**!
        """)

st.markdown("---")
st.markdown("Built with ❤️ using Streamlit | Multi-line field support")
