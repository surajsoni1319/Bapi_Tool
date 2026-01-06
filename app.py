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

HEADERS_TO_EXCLUDE = {
    "Customer Creation", "Customer Address", "GSTIN Details",
    "PAN Details", "Bank Details", "Aadhaar Details",
    "Supporting Document", "Zone Details", "Other Details",
    "Duplicity Check"
}

# Field definitions with patterns for better matching
FIELD_PATTERNS = [
    ("Type of Customer", r"^Type of Customer"),
    ("Name of Customer", r"^Name of Customer"),
    ("Company Code", r"^Company Code"),
    ("Customer Group", r"^Customer Group"),
    ("Sales Group", r"^Sales Group"),
    ("Region", r"^Region"),
    ("Zone", r"^Zone"),
    ("Sub Zone", r"^Sub Zone"),
    ("State", r"^State"),
    ("Sales Office", r"^Sales Office"),
    ("SAP Dealer code to be mapped Search Term 2", r"^SAP Dealer code to be mapped Search Term"),
    ("Search Term 1- Old customer code", r"^Search Term 1"),
    ("Search Term 2 - District", r"^Search Term 2"),
    ("Mobile Number", r"^Mobile Number"),
    ("E-Mail ID", r"^E-Mail ID"),
    ("Lattitude", r"^Lattitude"),
    ("Longitude", r"^Longitude"),
    ("Name of the Customers (Trade Name or Legal Name)", r"^Name of the Customers"),
    ("E-mail", r"^E-mail"),
    ("Address 1", r"^Address 1"),
    ("Address 2", r"^Address 2"),
    ("Address 3", r"^Address 3"),
    ("Address 4", r"^Address 4"),
    ("PIN", r"^PIN"),
    ("City", r"^City"),
    ("District", r"^District"),
    ("Whatsapp No.", r"^Whatsapp No"),
    ("Date of Birth", r"^Date of Birth"),
    ("Date of Anniversary", r"^Date of Anniversary"),
    ("Counter Potential - Maximum", r"^Counter Potential - Maximum"),
    ("Counter Potential - Minimum", r"^Counter Potential - Minimum"),
    ("Is GST Present", r"^Is GST Present"),
    ("GSTIN", r"^GSTIN"),
    ("Trade Name", r"^Trade Name"),
    ("Legal Name", r"^Legal Name"),
    ("Reg Date", r"^Reg Date"),
    ("Type", r"^Type"),
    ("Building No.", r"^Building No"),
    ("District Code", r"^District Code"),
    ("State Code", r"^State Code"),
    ("Street", r"^Street"),
    ("PIN Code", r"^PIN Code"),
    ("PAN", r"^PAN"),
    ("PAN Holder Name", r"^PAN Holder Name"),
    ("PAN Status", r"^PAN Status"),
    ("PAN - Aadhaar Linking Status", r"^PAN - Aadhaar Linking Status"),
    ("IFSC Number", r"^IFSC Number"),
    ("Account Number", r"^Account Number"),
    ("Name of Account Holder", r"^Name of Account Holder"),
    ("Bank Name", r"^Bank Name"),
    ("Bank Branch", r"^Bank Branch"),
    ("Is Aadhaar Linked with Mobile?", r"^Is Aadhaar Linked with Mobile"),
    ("Aadhaar Number", r"^Aadhaar Number"),
    ("Name", r"^Name"),
    ("Gender", r"^Gender"),
    ("DOB", r"^DOB"),
    ("Address", r"^Address"),
    ("Logistics Transportation Zone", r"^Logistics Transportation Zone"),
    ("Transportation Zone Description", r"^Transportation Zone Description"),
    ("Transportation Zone Code", r"^Transportation Zone Code"),
    ("Postal Code", r"^Postal Code"),
    ("Logistics team to vet the T zone selected by Sales Officer", r"^Logistics team to vet the T zone"),
    ("Selection of Available T Zones from T Zone Master list, if found.", r"^Selection of Available T Zones"),
    ("If NEW T Zone need to be created, details to be provided by Logistics team", r"^If NEW T Zone need to be created"),
    ("Date of Appointment", r"^Date of Appointment"),
    ("Delivering Plant", r"^Delivering Plant"),
    ("Plant Name", r"^Plant Name"),
    ("Plant Code", r"^Plant Code"),
    ("Incoterns", r"^Incoterns"),
    ("Incoterns Code", r"^Incoterns Code"),
    ("Security Deposit Amount details to filled up, as per checque received by Customer / Dealer", r"^Security Deposit Amount"),
    ("Credit Limit (In Rs.)", r"^Credit Limit"),
    ("Regional Head to be mapped", r"^Regional Head to be mapped"),
    ("Zonal Head to be mapped", r"^Zonal Head to be mapped"),
    ("Sub-Zonal Head (RSM) to be mapped", r"^Sub-Zonal Head"),
    ("Area Sales Manager to be mapped", r"^Area Sales Manager to be mapped"),
    ("Sales Officer to be mapped", r"^Sales Officer to be mapped"),
    ("Sales Promoter to be mapped", r"^Sales Promoter to be mapped"),
    ("Sales Promoter Number", r"^Sales Promoter Number"),
    ("Internal control code", r"^Internal control code"),
    ("SAP CODE", r"^SAP CODE"),
    ("Initiator Name", r"^Initiator Name"),
    ("Initiator Email ID", r"^Initiator Email ID"),
    ("Initiator Mobile Number", r"^Initiator Mobile Number"),
    ("Created By Customer UserID", r"^Created By Customer UserID"),
    ("Sales Head Name", r"^Sales Head Name"),
    ("Sales Head Email", r"^Sales Head Email"),
    ("Sales Head Mobile Number", r"^Sales Head Mobile Number"),
    ("Extra2", r"^Extra2"),
    ("PAN Result", r"^PAN Result"),
    ("Mobile Number Result", r"^Mobile Number Result"),
    ("Email Result", r"^Email Result"),
    ("GST Result", r"^GST Result"),
    ("Final Result", r"^Final Result")
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

def find_field_match(line):
    """Find which field this line starts with"""
    for field_name, pattern in FIELD_PATTERNS:
        if re.match(pattern, line, re.IGNORECASE):
            return field_name
    return None

def extract_value_from_line(line, field_name):
    """Extract value after field name from the same line"""
    # Remove field name from beginning
    value = line[len(field_name):].strip()
    # Remove common separators
    value = re.sub(r'^[:\-\s]+', '', value)
    return value

def parse_pdf_lines(lines):
    """Parse lines into field-value dictionary"""
    data = {}
    i = 0
    
    while i < len(lines):
        current_line = lines[i]
        field_name = find_field_match(current_line)
        
        if field_name:
            # Extract value from current line
            value = extract_value_from_line(current_line, field_name)
            
            # Check if field name continues on next line
            if not value and i + 1 < len(lines):
                next_line = lines[i + 1]
                combined = current_line + " " + next_line
                
                # Check if combined line forms a complete field
                combined_field = find_field_match(combined)
                if combined_field and len(combined_field) > len(field_name):
                    field_name = combined_field
                    value = extract_value_from_line(combined, field_name)
                    i += 1  # Skip next line
                else:
                    # Next line is the value
                    value = next_line
                    i += 1
            
            # Check if value continues on next line (for addresses, etc.)
            if value and i + 1 < len(lines):
                peek_line = lines[i + 1]
                peek_field = find_field_match(peek_line)
                
                # If next line is not a field, it might be value continuation
                if not peek_field:
                    # Special handling for long values
                    if field_name in ["Address", "Address 1"] or "Address" in field_name:
                        value += " " + peek_line
                        i += 1
            
            data[field_name] = value.strip()
        
        i += 1
    
    return data

def create_output_row(pdf_file, data):
    """Create output row with all fields"""
    row = {"Source File": pdf_file.name}
    
    for field_name, _ in FIELD_PATTERNS:
        row[field_name] = data.get(field_name, "")
    
    return row

if uploaded_files:
    all_rows = []
    
    with st.spinner("Processing PDFs..."):
        for pdf in uploaded_files:
            lines = extract_raw_lines(pdf)
            
            # Show raw lines for debugging
            with st.expander(f"📝 Raw lines from {pdf.name}"):
                for idx, line in enumerate(lines[:50], 1):  # Show first 50 lines
                    st.text(f"{idx:3d}. {line}")
            
            data = parse_pdf_lines(lines)
            row = create_output_row(pdf, data)
            all_rows.append(row)
    
    final_df = pd.DataFrame(all_rows)
    
    st.subheader("📊 Preview (Structured Data)")
    st.dataframe(final_df, use_container_width=True)
    
    st.success(f"✅ Processed {len(all_rows)} PDF(s) with {len(FIELD_PATTERNS)} fields")
    
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
