import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io
import re

def extract_pdf_data(pdf_file):
    """Extract key-value pairs from PDF and return as DataFrame"""
    
    # Headers to skip
    skip_headers = [
        'Customer Creation', 'Customer Address', 'GSTIN Details', 
        'PAN Details', 'Bank Details', 'Aadhaar Details', 
        'Supporting Document', 'Zone Details', 'Other Details', 
        'Duplicity Check'
    ]
    
    data = []
    
    # Read PDF bytes
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        text = page.get_text()
        
        if text:
            lines = text.split('\n')
            
            for line in lines:
                # Skip empty lines and header lines
                line = line.strip()
                if not line or line in skip_headers:
                    continue
                
                # Try to split by multiple spaces or tabs
                parts = re.split(r'\s{2,}|\t+', line, maxsplit=1)
                
                if len(parts) == 2:
                    key = parts[0].strip()
                    value = parts[1].strip()
                    
                    # Skip if key is in skip headers
                    if key not in skip_headers:
                        data.append({
                            'Field': key,
                            'Value': value
                        })
    
    pdf_document.close()
    return pd.DataFrame(data)

def main():
    st.title("PDF Data Extractor")
    st.markdown("Extract customer data from PDF files in tabular format")
    
    # File uploader
    uploaded_file = st.file_uploader("Upload PDF file", type=['pdf'])
    
    if uploaded_file is not None:
        try:
            # Extract data
            with st.spinner('Extracting data from PDF...'):
                df = extract_pdf_data(uploaded_file)
            
            if not df.empty:
                st.success(f"Successfully extracted {len(df)} fields!")
                
                # Display the dataframe
                st.subheader("Extracted Data")
                st.dataframe(df, use_container_width=True, height=400)
                
                # Download options
                st.subheader("Download Options")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Download as CSV
                    csv = df.to_csv(index=False)
                    st.download_button(
                        label="Download as CSV",
                        data=csv,
                        file_name="extracted_data.csv",
                        mime="text/csv"
                    )
                
                with col2:
                    # Download as Excel
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Data')
                    
                    st.download_button(
                        label="Download as Excel",
                        data=buffer.getvalue(),
                        file_name="extracted_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Show statistics
                st.subheader("Statistics")
                st.info(f"Total fields extracted: {len(df)}")
                
            else:
                st.warning("No data could be extracted from the PDF.")
                
        except Exception as e:
            st.error(f"Error processing PDF: {str(e)}")
            st.exception(e)
    
    # Instructions
    with st.expander("ℹ️ Instructions"):
        st.markdown("""
        ### How to use:
        1. Upload your PDF file using the file uploader
        2. The tool will automatically extract field-value pairs
        3. Review the extracted data in the table
        4. Download the data as CSV or Excel format
        
        ### Note:
        - The following section headers are excluded from extraction:
          - Customer Creation, Customer Address, GSTIN Details
          - PAN Details, Bank Details, Aadhaar Details
          - Supporting Document, Zone Details, Other Details
          - Duplicity Check
        """)

if __name__ == "__main__":
    main()
