import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# Page config
st.set_page_config(page_title="PDF Customer Data Extractor", page_icon="📄", layout="wide")

# Title
st.title("📄 PDF Customer Data Extractor")
st.markdown("Extracts data in ilovepdf.com format")
st.markdown("---")

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
    
    # Extract as table: Field | Value
    table_data = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Try to split by tab first (most reliable)
        if '\t' in line:
            parts = line.split('\t', 1)  # Split only on first tab
            field = parts[0].strip()
            value = parts[1].strip() if len(parts) > 1 else ""
            table_data.append({'Field': field, 'Value': value})
        
        # Try to split by 2+ spaces
        elif '  ' in line:
            parts = re.split(r'\s{2,}', line, 1)  # Split only once
            field = parts[0].strip()
            value = parts[1].strip() if len(parts) > 1 else ""
            table_data.append({'Field': field, 'Value': value})
        
        # Single item - it's a field with no value
        else:
            table_data.append({'Field': line, 'Value': ""})
    
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
    st.markdown("### ℹ️ About")
    st.info("Extracts PDF in same format as ilovepdf.com:\n\n2 columns: Field | Value")

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
                
                st.success(f"✅ Successfully extracted {total_rows} rows!")
                
                if show_stats:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Rows", total_rows)
                    with col2:
                        st.metric("Filled Rows", filled_rows)
                    with col3:
                        st.metric("Empty Rows", empty_rows)
                
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
                
                # Show comparison with ilovepdf format
                with st.expander("🔍 Sample Comparison"):
                    st.markdown("""
                    ### Your extraction should look like:
                    
                    | Field | Value |
                    |-------|-------|
                    | Customer Creation | |
                    | Type of Customer | ZSUB - RSSD Sub-Dealer |
                    | Name of Customer | GULZAR HARDWARE |
                    | Company Code | 1010 - SCL-STAR CEMENT LTD. |
                    | ... | ... |
                    
                    **Section headers** (like "Customer Creation") are included with empty values.
                    """)
                
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                with st.expander("Error Details"):
                    st.code(str(e))
                    import traceback
                    st.code(traceback.format_exc())

else:
    st.info("👆 Upload a PDF file to begin extraction")
    
    with st.expander("📖 How it works"):
        st.markdown("""
        ### This tool extracts PDFs exactly like ilovepdf.com:
        
        **Output Format:**
        - 2 columns: **Field** | **Value**
        - Section headers included (Customer Creation, PAN Details, etc.)
        - Empty fields show as blank values
        - All rows preserved in order
        
        **What you get:**
        - Same structure as ilovepdf.com
        - Easy to compare with your reference
        - Ready to use in Excel/CSV
        
        **Steps:**
        1. Upload your PDF
        2. Click "Extract Data"
        3. Compare with your ilovepdf output
        4. Download as CSV or Excel
        """)
    
    with st.expander("📋 Expected Format"):
        st.code("""
Field                          | Value
-------------------------------|---------------------------
Customer Creation              | 
Type of Customer               | ZSUB - RSSD Sub-Dealer
Name of Customer               | GULZAR HARDWARE
Company Code                   | 1010 - SCL-STAR CEMENT LTD.
Search Term 1- Old customer... | 
Search Term 2 - District       | 
Mobile Number                  | 8638595914
E-Mail ID                      | 
        """, language="text")

st.markdown("---")
st.markdown("Built with ❤️ using Streamlit | Mimics ilovepdf.com extraction")
