import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(
    page_title="PDF → Excel (Smart Parser)",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Smart PDF to Excel Converter")
st.write("Intelligently parses PDF forms and maps to predefined fields")

# Define the fixed field list
FIELD_LIST = [
    "Type of Customer", "Name of Customer", "Company Code", "Customer Group",
    "Sales Group", "Region", "Zone", "Sub Zone", "State", "Sales Office",
    "SAP Dealer code to be mapped Search Term 2",
    "Search Term 1- Old customer code", "Search Term 2 - District",
    "Mobile Number", "E-Mail ID", "Lattitude", "Longitude",
    "Name of the Customers (Trade Name or Legal Name)",
    "E-mail", "Address 1", "Address 2", "Address 3", "Address 4",
    "PIN", "City", "District", "Whatsapp No.", "Date of Birth",
    "Date of Anniversary", "Counter Potential - Maximum",
    "Counter Potential - Minimum", "Is GST Present", "GSTIN",
    "Trade Name", "Legal Name", "Reg Date", "Type", "Building No.",
    "District Code", "State Code", "Street", "PIN Code", "PAN",
    "PAN Holder Name", "PAN Status", "PAN - Aadhaar Linking Status",
    "IFSC Number", "Account Number", "Name of Account Holder",
    "Bank Name", "Bank Branch", "Is Aadhaar Linked with Mobile?",
    "Aadhaar Number", "Name", "Gender", "DOB", "Address",
    "Logistics Transportation Zone", "Transportation Zone Description",
    "Transportation Zone Code", "Postal Code",
    "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment", "Delivering Plant", "Plant Name", "Plant Code",
    "Incoterns", "Incoterns Code",
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

HEADERS_TO_SKIP = [
    "Customer Creation", "Customer Address", "GSTIN Details",
    "PAN Details", "Bank Details", "Aadhaar Details",
    "Supporting Document", "Zone Details", "Other Details", "Duplicity Check"
]

def extract_raw_lines(pdf_file):
    """Extract all lines from PDF"""
    lines = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    line = line.strip()
                    if line and line not in HEADERS_TO_SKIP:
                        lines.append(line)
    return lines

def fuzzy_match_field(text, field_list):
    """Find field that matches start of text, handling partial matches"""
    text_lower = text.lower().strip()
    
    # Sort fields by length (longest first) for better matching
    sorted_fields = sorted(field_list, key=len, reverse=True)
    
    for field in sorted_fields:
        field_lower = field.lower().strip()
        
        # Direct start match
        if text_lower.startswith(field_lower):
            value = text[len(field):].strip()
            return field, value
        
        # Handle partial field names (for split cases)
        # Check if text starts with significant portion of field
        words = field_lower.split()
        if len(words) > 2:
            # Check if at least first 60% of words match
            partial = ' '.join(words[:max(2, int(len(words) * 0.6))])
            if text_lower.startswith(partial):
                return field, None  # Value might be on next line
    
    return None, None

def smart_parse_pdf(pdf_file, field_list):
    """Parse PDF with intelligent field-value separation"""
    
    lines = extract_raw_lines(pdf_file)
    parsed_data = {}
    i = 0
    
    while i < len(lines):
        current_line = lines[i]
        
        # Try to match with a known field
        field, value = fuzzy_match_field(current_line, field_list)
        
        if field:
            # If no value found, check next lines
            if not value:
                j = i + 1
                accumulated = []
                
                # Look ahead to collect value lines
                while j < len(lines):
                    next_line = lines[j]
                    next_field, _ = fuzzy_match_field(next_line, field_list)
                    
                    # If next line is a new field, stop
                    if next_field:
                        break
                    
                    # Add to accumulated value
                    accumulated.append(next_line)
                    j += 1
                    
                    # Check if we've completed the field name
                    combined = current_line + " " + " ".join(accumulated)
                    field_check, val_check = fuzzy_match_field(combined, field_list)
                    
                    if field_check and val_check:
                        value = val_check
                        i = j - 1  # Update position
                        break
                
                # If still no clear value, use accumulated
                if not value and accumulated:
                    value = " ".join(accumulated)
                    i = j - 1
            
            # Store parsed data
            if field in parsed_data:
                # Handle duplicate fields (like Mobile Number appears twice)
                if parsed_data[field] and value:
                    parsed_data[field] += f" | {value}"
                elif value:
                    parsed_data[field] = value
            else:
                parsed_data[field] = value if value else ""
        
        i += 1
    
    # Create result dataframe
    result = []
    for field in field_list:
        value = parsed_data.get(field, "")
        result.append({"Field": field, "Value": value})
    
    return pd.DataFrame(result)

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    for pdf in uploaded_files:
        st.subheader(f"📄 {pdf.name}")
        
        # Parse the PDF
        df = smart_parse_pdf(pdf, FIELD_LIST)
        
        # Show preview
        st.dataframe(df, use_container_width=True, height=400)
        
        # Download button
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Parsed_Data")
        
        st.download_button(
            f"⬇️ Download {pdf.name}.xlsx",
            data=output.getvalue(),
            file_name=f"{pdf.name.replace('.pdf', '')}_Parsed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_{pdf.name}"
        )
