import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# Page config
st.set_page_config(page_title="PDF Customer Data Extractor", page_icon="📄", layout="wide")

# Title
st.title("📄 PDF Customer Data Extractor")
st.markdown("---")

# Headers to exclude
EXCLUDE_HEADERS = [
    'Customer Creation', 'Customer Address', 'GSTIN Details', 
    'PAN Details', 'Bank Details', 'Aadhaar Details', 
    'Supporting Document', 'Zone Details', 'Other Details', 
    'Duplicity Check'
]

def extract_data_from_pdf(pdf_file):
    """Extract key-value pairs from PDF file"""
    data = {}
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            
            if text:
                lines = text.split('\n')
                
                for line in lines:
                    # Skip empty lines and header lines
                    line = line.strip()
                    if not line or line in EXCLUDE_HEADERS:
                        continue
                    
                    # Try to split by multiple spaces or tabs
                    # Pattern: Field name followed by value
                    parts = re.split(r'\s{2,}|\t+', line, maxsplit=1)
                    
                    if len(parts) == 2:
                        key = parts[0].strip()
                        value = parts[1].strip()
                        
                        # Skip if key is a header
                        if key not in EXCLUDE_HEADERS and key and value:
                            data[key] = value
    
    return data

def convert_to_dataframe(data_dict):
    """Convert dictionary to pandas DataFrame"""
    df = pd.DataFrame(list(data_dict.items()), columns=['Field', 'Value'])
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

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type=['pdf'])

if uploaded_file is not None:
    st.success(f"File uploaded: {uploaded_file.name}")
    
    # Add a button to process
    if st.button("Extract Data", type="primary"):
        with st.spinner("Extracting data from PDF..."):
            try:
                # Extract data
                extracted_data = extract_data_from_pdf(uploaded_file)
                
                if extracted_data:
                    # Convert to DataFrame
                    df = convert_to_dataframe(extracted_data)
                    
                    # Display success message
                    st.success(f"✅ Successfully extracted {len(df)} fields!")
                    
                    # Display the data
                    st.subheader("Extracted Data")
                    st.dataframe(df, use_container_width=True, height=600)
                    
                    # Download options
                    st.subheader("Download Options")
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        csv_data = convert_df_to_csv(df)
                        st.download_button(
                            label="📥 Download as CSV",
                            data=csv_data,
                            file_name="extracted_customer_data.csv",
                            mime="text/csv",
                        )
                    
                    with col2:
                        excel_data = convert_df_to_excel(df)
                        st.download_button(
                            label="📥 Download as Excel",
                            data=excel_data,
                            file_name="extracted_customer_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    
                    # Show some statistics
                    st.subheader("Statistics")
                    st.metric("Total Fields Extracted", len(df))
                    
                else:
                    st.warning("No data could be extracted from the PDF.")
                    
            except Exception as e:
                st.error(f"Error processing PDF: {str(e)}")
                st.info("Please make sure the PDF format matches the expected structure.")

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
        - Automatically extracts field-value pairs from PDF
        - Excludes header sections (Customer Creation, PAN Details, etc.)
        - Exports data to CSV or Excel format
        - Handles multiple pages in PDF
        
        ### Supported Format:
        The tool works best with PDFs that have:
        - Field names on the left
        - Values on the right
        - Clear separation between field and value
        """)

# Footer
st.markdown("---")
st.markdown("Built with ❤️ using Streamlit")
