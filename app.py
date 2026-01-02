import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
from io import BytesIO
import pytesseract
from PIL import Image
import io

# Configure page
st.set_page_config(page_title="PDF to Excel Extractor", layout="wide")

def extract_text_from_pdf(pdf_file):
    """Extract text from PDF using PyMuPDF"""
    try:
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text
    except Exception as e:
        st.error(f"Error extracting text: {str(e)}")
        return None

def extract_text_with_ocr(pdf_file):
    """Extract text using OCR (Tesseract) as fallback"""
    try:
        pdf_file.seek(0)
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        text = ""
        for page_num in range(len(doc)):
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Higher resolution
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            text += pytesseract.image_to_string(img)
        doc.close()
        return text
    except Exception as e:
        st.error(f"Error with OCR: {str(e)}")
        return None

def parse_customer_data(text):
    """Parse the extracted text and create a dictionary of field-value pairs"""
    data = {}
    
    # Define all possible fields to extract (as they appear in PDF)
    fields = [
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
        "SAP Dealer code to be mapped",
        "Search Term",
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
        "Incoterns FOR - FREE ON ROAD",
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
    
    # Clean the text
    lines = text.split('\n')
    
    # Create a more robust parsing approach
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
            
        # Try to match field patterns
        for field in fields:
            # Check if line starts with the field name
            if line.startswith(field):
                # Extract value after the field name
                value = line[len(field):].strip()
                
                # Remove common separators
                if value.startswith(':'):
                    value = value[1:].strip()
                elif value.startswith('-'):
                    value = value[1:].strip()
                
                # If value is empty, check next line
                if not value and i + 1 < len(lines):
                    value = lines[i + 1].strip()
                
                data[field] = value
                break
    
    # Alternative parsing using regex for key-value pairs
    if len(data) < 10:  # If we didn't capture much, try regex approach
        pattern = r'([A-Za-z\s\-\(\)\.\/]+?)\s*[:−-]\s*(.+?)(?=\n[A-Za-z\s\-\(\)\.\/]+?\s*[:−-]|\Z)'
        matches = re.findall(pattern, text, re.MULTILINE | re.DOTALL)
        
        for field_name, value in matches:
            field_name = field_name.strip()
            value = value.strip().replace('\n', ' ')
            
            if field_name in fields:
                data[field_name] = value
    
    return data

def create_excel_file(dataframes_dict):
    """Create Excel file with data from multiple PDFs"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Combine all dataframes
        if dataframes_dict:
            combined_df = pd.concat(dataframes_dict.values(), ignore_index=True)
            combined_df.to_excel(writer, sheet_name='All Customers', index=False)
            
            # Individual sheets for each customer
            for filename, df in dataframes_dict.items():
                sheet_name = filename[:30]  # Excel sheet name limit
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    output.seek(0)
    return output

# Streamlit UI
st.title("📄 PDF to Excel Data Extractor")
st.markdown("### Customer Creation Form Extractor")
st.write("Upload PDF files containing customer creation forms to extract data into Excel format.")

# Sidebar for options
with st.sidebar:
    st.header("⚙️ Options")
    use_ocr = st.checkbox("Use OCR (if text extraction fails)", value=False)
    st.info("💡 **Tip:** Enable OCR for scanned PDFs or images")
    
    st.markdown("---")
    st.markdown("### About")
    st.markdown("""
    This tool extracts customer data from PDF forms and converts them to Excel format.
    
    **Features:**
    - Extract text from PDF
    - OCR support for scanned documents
    - Multiple PDF processing
    - Excel export with multiple sheets
    """)

# File uploader
uploaded_files = st.file_uploader(
    "Choose PDF file(s)", 
    type=['pdf'], 
    accept_multiple_files=True,
    help="Upload one or more PDF files"
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")
    
    # Process button
    if st.button("🚀 Extract Data", type="primary"):
        dataframes_dict = {}
        
        # Progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Processing: {uploaded_file.name}...")
            
            # Extract text
            text = extract_text_from_pdf(uploaded_file)
            
            # If text extraction failed or OCR is enabled
            if (not text or len(text.strip()) < 100) and use_ocr:
                status_text.text(f"Running OCR on: {uploaded_file.name}...")
                text = extract_text_with_ocr(uploaded_file)
            
            if text:
                # Parse data
                customer_data = parse_customer_data(text)
                
                if customer_data:
                    # Convert to DataFrame
                    df = pd.DataFrame([customer_data])
                    dataframes_dict[uploaded_file.name] = df
                    
                    # Show preview
                    with st.expander(f"📋 Preview: {uploaded_file.name}"):
                        st.dataframe(df.T, use_container_width=True)
                else:
                    st.warning(f"⚠️ No data extracted from {uploaded_file.name}")
            else:
                st.error(f"❌ Failed to extract text from {uploaded_file.name}")
            
            # Update progress
            progress_bar.progress((idx + 1) / len(uploaded_files))
        
        status_text.text("✅ Processing complete!")
        
        # Create and download Excel
        if dataframes_dict:
            st.success(f"🎉 Successfully extracted data from {len(dataframes_dict)} file(s)!")
            
            # Show combined preview
            st.subheader("📊 Combined Data Preview")
            combined_df = pd.concat(dataframes_dict.values(), ignore_index=True)
            st.dataframe(combined_df, use_container_width=True)
            
            # Statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Records", len(combined_df))
            with col2:
                st.metric("Total Fields", len(combined_df.columns))
            with col3:
                st.metric("Files Processed", len(dataframes_dict))
            
            # Download button
            excel_file = create_excel_file(dataframes_dict)
            st.download_button(
                label="📥 Download Excel File",
                data=excel_file,
                file_name="customer_data_extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            st.error("❌ No data could be extracted from the uploaded files.")

# Instructions
with st.expander("📖 How to Use"):
    st.markdown("""
    1. **Upload PDF Files**: Click on the file uploader and select one or more PDF files
    2. **Enable OCR (Optional)**: If your PDFs are scanned images, enable OCR in the sidebar
    3. **Extract Data**: Click the "Extract Data" button to process the files
    4. **Preview**: Review the extracted data in the preview section
    5. **Download**: Click "Download Excel File" to save the extracted data
    
    **Supported Fields:**
    - Customer information (Name, Type, Code, etc.)
    - Contact details (Mobile, Email, Address)
    - Location data (Latitude, Longitude, Zone, etc.)
    - Financial details (PAN, Bank, GST)
    - Sales team mapping
    - And many more...
    """)

# Footer
st.markdown("---")
st.markdown("Made with ❤️ using Streamlit | Powered by PyMuPDF & Tesseract OCR")
