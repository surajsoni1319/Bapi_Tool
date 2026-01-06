import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(
    page_title="PDF → Excel (Structured)",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Structured Excel Converter")
st.write("Converts PDF to clean field-value Excel format")

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# Headers to exclude
HEADERS_TO_EXCLUDE = {
    "Customer Creation", "Customer Address", "GSTIN Details",
    "PAN Details", "Bank Details", "Aadhaar Details",
    "Supporting Document", "Zone Details", "Other Details",
    "Duplicity Check"
}

# Define all expected fields in order
FIELD_DEFINITIONS = [
    "Type of Customer", "Name of Customer", "Company Code", "Customer Group",
    "Sales Group", "Region", "Zone", "Sub Zone", "State", "Sales Office",
    "SAP Dealer code to be mapped Search Term 2", "Search Term 1- Old customer code",
    "Search Term 2 - District", "Mobile Number", "E-Mail ID", "Lattitude", "Longitude",
    "Name of the Customers (Trade Name or Legal Name)", "Mobile Number",
    "E-mail", "Address 1", "Address 2", "Address 3", "Address 4",
    "PIN", "City", "District", "State", "Whatsapp No.", "Date of Birth",
    "Date of Anniversary", "Counter Potential - Maximum", "Counter Potential - Minimum",
    "Is GST Present", "GSTIN", "Trade Name", "Legal Name", "Reg Date",
    "City", "Type", "Building No.", "District Code", "State Code",
    "Street", "PIN Code", "PAN", "PAN Holder Name", "PAN Status",
    "PAN - Aadhaar Linking Status", "IFSC Number", "Account Number",
    "Name of Account Holder", "Bank Name", "Bank Branch",
    "Is Aadhaar Linked with Mobile?", "Aadhaar Number", "Name",
    "Gender", "DOB", "Address", "PIN", "City", "State",
    "Logistics Transportation Zone", "Transportation Zone Description",
    "Transportation Zone Code", "Postal Code",
    "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment", "Delivering Plant", "Plant Name", "Plant Code",
    "Incoterns", "Incoterns", "Incoterns Code",
    "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)", "Regional Head to be mapped",
    "Zonal Head to be mapped", "Sub-Zonal Head (RSM) to be mapped",
    "Area Sales Manager to be mapped", "Sales Officer to be mapped",
    "Sales Promoter to be mapped", "Sales Promoter Number",
    "Internal control code", "SAP CODE", "Initiator Name",
    "Initiator Email ID", "Initiator Mobile Number",
    "Created By Customer UserID", "Sales Head Name", "Sales Head Email",
    "Sales Head Mobile Number", "Extra2", "PAN Result",
    "Mobile Number Result", "Email Result", "GST Result", "Final Result"
]

def extract_raw_lines(pdf_file):
    """Extract all lines from PDF"""
    lines = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    cleaned = line.strip()
                    if cleaned and cleaned not in HEADERS_TO_EXCLUDE:
                        lines.append(cleaned)
    return lines

def smart_parse_field_value(lines):
    """Parse lines into field-value pairs with intelligent merging"""
    data = {}
    i = 0
    
    while i < len(lines):
        current_line = lines[i]
        next_line = lines[i + 1] if i + 1 < len(lines) else ""
        
        # Check if current line matches start of any known field
        matched_field = None
        for field in FIELD_DEFINITIONS:
            # Check exact match
            if current_line == field:
                matched_field = field
                break
            # Check if line starts with field name
            elif current_line.startswith(field):
                matched_field = field
                break
        
        if matched_field:
            # Extract value from same line if present
            remainder = current_line[len(matched_field):].strip()
            
            # Check if next line completes the field name
            combined_field = current_line + " " + next_line
            full_field_match = None
            for field in FIELD_DEFINITIONS:
                if combined_field.startswith(field):
                    full_field_match = field
                    break
            
            if full_field_match and full_field_match != matched_field:
                # Field name spans two lines
                matched_field = full_field_match
                remainder = combined_field[len(matched_field):].strip()
                i += 1  # Skip next line as it's part of field name
            
            # Get value (might be empty or on next line)
            if remainder:
                value = remainder
                # Check if value continues on next line
                if i + 1 < len(lines):
                    peek = lines[i + 1]
                    # If next line doesn't look like a field, it's continuation
                    is_field = any(peek == f or peek.startswith(f[:20]) for f in FIELD_DEFINITIONS)
                    if not is_field:
                        value += " " + peek
                        i += 1
            else:
                # Value on next line
                if i + 1 < len(lines):
                    value = lines[i + 1]
                    i += 1
                    # Check for multi-line values
                    if i + 1 < len(lines):
                        peek = lines[i + 1]
                        is_field = any(peek == f or peek.startswith(f[:20]) for f in FIELD_DEFINITIONS)
                        if not is_field:
                            value += " " + peek
                            i += 1
                else:
                    value = ""
            
            data[matched_field] = value.strip()
        
        i += 1
    
    return data

def process_pdf(pdf_file):
    """Process single PDF file"""
    lines = extract_raw_lines(pdf_file)
    data = smart_parse_field_value(lines)
    
    # Create row with all fields
    row = {"Source File": pdf_file.name}
    for field in FIELD_DEFINITIONS:
        row[field] = data.get(field, "")
    
    return row

if uploaded_files:
    all_rows = []
    
    with st.spinner("Processing PDFs..."):
        for pdf in uploaded_files:
            row = process_pdf(pdf)
            all_rows.append(row)
    
    final_df = pd.DataFrame(all_rows)
    
    st.subheader("📊 Preview (Structured Data)")
    st.dataframe(final_df, use_container_width=True)
    
    st.info(f"✅ Processed {len(all_rows)} PDF(s) with {len(FIELD_DEFINITIONS)} fields each")
    
    # Download button
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Structured_Data")
    
    st.download_button(
        "⬇️ Download Structured Excel",
        data=output.getvalue(),
        file_name="PDF_Structured_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👆 Upload PDF files to begin processing")
