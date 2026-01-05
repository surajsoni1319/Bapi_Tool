import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# Page config
st.set_page_config(page_title="PDF Customer Data Extractor", page_icon="📄", layout="wide")

st.title("📄 PDF Customer Data Extractor")
st.markdown("Extracts Field | Value format (2 columns)")
st.markdown("---")

def extract_pdf_as_table(pdf_file):
    """Extract PDF into Field | Value table format"""
    
    # Read PDF
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Get text from all pages
    full_text = ""
    for page in pdf_document:
        full_text += page.get_text()
    pdf_document.close()
    
    # Split into lines
    lines = [line.strip() for line in full_text.split('\n') if line.strip()]
    
    # Section headers (include with empty values)
    sections = [
        'Customer Creation', 'Customer Address', 'GSTIN Details', 
        'PAN Details', 'Bank Details', 'Aadhaar Details', 
        'Supporting Document', 'Zone Details', 'Other Details', 
        'Duplicity Check'
    ]
    
    # Process each line
    rows = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # Section headers
        if line in sections:
            rows.append({'Field': line, 'Value': ''})
            i += 1
            continue
        
        # Split line into field and value
        field = ''
        value = ''
        
        # Try tab separator
        if '\t' in line:
            parts = line.split('\t', 1)
            field = parts[0].strip()
            value = parts[1].strip() if len(parts) > 1 else ''
        
        # Try 2+ space separator
        elif '  ' in line:
            match = re.match(r'^(.+?)\s{2,}(.*)$', line)
            if match:
                field = match.group(1).strip()
                value = match.group(2).strip()
            else:
                field = line
                value = ''
        
        # No separator - field only
        else:
            field = line
            value = ''
        
        # Handle special multi-line fields
        handled = False
        
        # SAP Dealer code
        if 'SAP Dealer code to be mapped Search Term' in field:
            field = 'SAP Dealer code to be mapped Search Term 2'
            # Next lines: "2", then actual value
            if i + 2 < len(lines) and lines[i + 1].strip() == '2':
                value = lines[i + 2].strip()
                i += 2
            handled = True
        
        # Name of Customers
        elif 'Name of the Customers (Trade Name or' in field:
            field = 'Name of the Customers (Trade Name or Legal Name)'
            # Next line: "Legal Name)", then value
            if i + 2 < len(lines) and 'Legal Name)' in lines[i + 1]:
                value = lines[i + 2].strip() if not value else value
                i += 1
            handled = True
        
        # Security Deposit
        elif 'Security Deposit Amount details to filled up,' in field:
            field = 'Security Deposit Amount details to filled up, as per checque received by Customer / Dealer'
            # Next line: continuation, then value
            if i + 2 < len(lines) and 'as per checque' in lines[i + 1]:
                value = lines[i + 2].strip()
                i += 1
            handled = True
        
        # Logistics team
        elif 'Logistics team to vet the T zone selected by' in field:
            field = 'Logistics team to vet the T zone selected by Sales Officer'
            if i + 2 < len(lines) and 'Sales Officer' in lines[i + 1]:
                value = lines[i + 2].strip()
                i += 1
            handled = True
        
        # Selection of T Zones
        elif 'Selection of Available T Zones from T Zone' in field:
            field = 'Selection of Available T Zones from T Zone Master list, if found.'
            if i + 2 < len(lines) and 'Master list' in lines[i + 1]:
                value = lines[i + 2].strip()
                i += 1
            handled = True
        
        # If NEW T Zone
        elif 'If NEW T Zone need to be created, details to' in field:
            field = 'If NEW T Zone need to be created, details to be provided by Logistics team'
            if i + 2 < len(lines) and 'be provided' in lines[i + 1]:
                value = lines[i + 2].strip()
                i += 1
            handled = True
        
        # Add row
        rows.append({'Field': field, 'Value': value})
        i += 1
    
    # Create DataFrame
    df = pd.DataFrame(rows)
    return df

def to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

# Sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    show_empty = st.checkbox("Show empty rows", value=True)
    st.markdown("---")
    st.info("**Target:** 101 rows\n\n**Format:** Field | Value")

# Main
uploaded = st.file_uploader("Upload PDF", type=['pdf'])

if uploaded:
    st.success(f"Uploaded: {uploaded.name}")
    
    if st.button("Extract", type="primary", use_container_width=True):
        with st.spinner("Extracting..."):
            try:
                uploaded.seek(0)
                df = extract_pdf_as_table(uploaded)
                
                # Stats
                total = len(df)
                filled = len(df[df['Value'] != ''])
                diff = total - 101
                
                if diff == 0:
                    st.success(f"✅ Perfect! {total} rows extracted")
                else:
                    st.warning(f"⚠️ {total} rows ({diff:+d} vs target 101)")
                
                # Show stats
                col1, col2, col3 = st.columns(3)
                col1.metric("Total", total, delta=f"{diff:+d}" if diff != 0 else "✓")
                col2.metric("Filled", filled)
                col3.metric("Empty", total - filled)
                
                # Display
                st.subheader("Extracted Data")
                display = df if show_empty else df[df['Value'] != '']
                st.dataframe(display, use_container_width=True, height=600)
                
                # Sample
                with st.expander("📋 Sample (first 10 rows)"):
                    st.dataframe(df.head(10), use_container_width=True)
                
                # Download
                st.subheader("Download")
                col1, col2 = st.columns(2)
                
                with col1:
                    csv = to_csv(df)
                    st.download_button(
                        "📄 CSV",
                        csv,
                        f"{uploaded.name.replace('.pdf', '')}.csv",
                        "text/csv",
                        use_container_width=True
                    )
                
                with col2:
                    excel = to_excel(df)
                    st.download_button(
                        "📊 Excel",
                        excel,
                        f"{uploaded.name.replace('.pdf', '')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
            except Exception as e:
                st.error(f"Error: {e}")
                import traceback
                st.code(traceback.format_exc())

else:
    st.info("👆 Upload a PDF to start")
    
    with st.expander("ℹ️ Expected Output"):
        st.markdown("""
        **Format:** 2 columns
        
        | Field | Value |
        |-------|-------|
        | Customer Creation | |
        | Type of Customer | ZSUB - RSSD Sub-Dealer |
        | Name of Customer | GULZAR HARDWARE |
        | ... | ... |
        
        **Total:** ~101 rows
        """)

st.markdown("---")
st.caption("PDF Data Extractor • Field | Value format")
