import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# Page config
st.set_page_config(page_title="PDF Customer Data Extractor", page_icon="📄", layout="wide")

# Title
st.title("📄 PDF Customer Data Extractor")
st.markdown("---")

# All field names that can appear in the PDF
FIELD_NAMES = [
    "Type of Customer", "Name of Customer", "Company Code", "Customer Group",
    "Sales Group", "Region", "Zone", "Sub Zone", "State", "Sales Office",
    "SAP Dealer code to be mapped Search Term", "Search Term 1- Old customer code",
    "Search Term 2 - District", "Mobile Number", "E-Mail ID", "Lattitude", "Longitude",
    "Name of the Customers (Trade Name or", "Legal Name)", "E-mail",
    "Address 1", "Address 2", "Address 3", "Address 4", "PIN", "City", "District",
    "Whatsapp No.", "Date of Birth", "Date of Anniversary",
    "Counter Potential - Maximum", "Counter Potential - Minimum",
    "Is GST Present", "GSTIN", "Trade Name", "Legal Name", "Reg Date",
    "Type", "Building No.", "District Code", "State Code", "Street", "PIN Code",
    "PAN", "PAN Holder Name", "PAN Status", "PAN - Aadhaar Linking Status",
    "IFSC Number", "Account Number", "Name of Account Holder", "Bank Name", "Bank Branch",
    "Is Aadhaar Linked with Mobile?", "Aadhaar Number", "Name", "Gender", "DOB", "Address",
    "Logistics Transportation Zone", "Transportation Zone Description",
    "Transportation Zone Code", "Postal Code",
    "Logistics team to vet the T zone selected by", "Sales Officer",
    "Selection of Available T Zones from T Zone", "Master list, if found.",
    "If NEW T Zone need to be created, details to", "be provided by Logistics team",
    "Date of Appointment", "Delivering Plant", "Plant Name", "Plant Code",
    "Incoterns", "Incoterns Code",
    "Security Deposit Amount details to filled up,", "as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)", "Regional Head to be mapped", "Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped", "Area Sales Manager to be mapped",
    "Sales Officer to be mapped", "Sales Promoter to be mapped", "Sales Promoter Number",
    "Internal control code", "SAP CODE", "Initiator Name", "Initiator Email ID",
    "Initiator Mobile Number", "Created By Customer UserID", "Sales Head Name",
    "Sales Head Email", "Sales Head Mobile Number", "Extra2", "Duplicity Check",
    "PAN Result", "Mobile Number Result", "Email Result", "GST Result", "Final Result"
]

# Output column names (standardized)
OUTPUT_COLUMNS = [
    "Type of Customer", "Name of Customer", "Company Code", "Customer Group",
    "Sales Group", "Region", "Zone", "Sub Zone", "State", "Sales Office",
    "SAP Dealer code to be mapped Search Term 2", "Search Term 1- Old customer code",
    "Search Term 2 - District", "Mobile Number", "E-Mail ID", "Lattitude", "Longitude",
    "Name of the Customers (Trade Name or Legal Name)", "E-mail",
    "Address 1", "Address 2", "Address 3", "Address 4", "PIN", "City", "District",
    "Whatsapp No.", "Date of Birth", "Date of Anniversary",
    "Counter Potential - Maximum", "Counter Potential - Minimum",
    "Is GST Present", "GSTIN", "Trade Name", "Legal Name", "Reg Date",
    "Type", "Building No.", "District Code", "State Code", "Street", "PIN Code",
    "PAN", "PAN Holder Name", "PAN Status", "PAN - Aadhaar Linking Status",
    "IFSC Number", "Account Number", "Name of Account Holder", "Bank Name", "Bank Branch",
    "Is Aadhaar Linked with Mobile?", "Aadhaar Number", "Name", "Gender", "DOB", "Address",
    "Logistics Transportation Zone", "Transportation Zone Description",
    "Transportation Zone Code", "Postal Code",
    "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment", "Delivering Plant", "Plant Name", "Plant Code",
    "Incoterns", "Incoterns Code",
    "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)", "Regional Head to be mapped", "Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped", "Area Sales Manager to be mapped",
    "Sales Officer to be mapped", "Sales Promoter to be mapped", "Sales Promoter Number",
    "Internal control code", "SAP CODE", "Initiator Name", "Initiator Email ID",
    "Initiator Mobile Number", "Created By Customer UserID", "Sales Head Name",
    "Sales Head Email", "Sales Head Mobile Number", "Extra2",
    "PAN Result", "Mobile Number Result", "Email Result", "GST Result", "Final Result"
]

EXCLUDE_HEADERS = [
    'Customer Creation', 'Customer Address', 'GSTIN Details', 
    'PAN Details', 'Bank Details', 'Aadhaar Details', 
    'Supporting Document', 'Zone Details', 'Other Details', 
    'Duplicity Check'
]

def is_field_name(text):
    """Check if text is likely a field name"""
    text = text.strip()
    for field in FIELD_NAMES:
        if text.startswith(field):
            return True
    return False

def extract_data_from_pdf(pdf_file):
    """Extract key-value pairs from PDF file"""
    data = {}
    
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    full_text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        full_text += page.get_text()
    
    pdf_document.close()
    
    lines = [line.strip() for line in full_text.split('\n') if line.strip()]
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Skip headers
        if line in EXCLUDE_HEADERS:
            i += 1
            continue
        
        # Check for specific field patterns
        
        # Type of Customer
        if line.startswith("Type of Customer"):
            value = line.replace("Type of Customer", "").strip()
            data["Type of Customer"] = value
        
        # Name of Customer
        elif line.startswith("Name of Customer"):
            value = line.replace("Name of Customer", "").strip()
            data["Name of Customer"] = value
        
        # Company Code
        elif line.startswith("Company Code"):
            value = line.replace("Company Code", "").strip()
            data["Company Code"] = value
        
        # Customer Group
        elif line.startswith("Customer Group"):
            value = line.replace("Customer Group", "").strip()
            data["Customer Group"] = value
        
        # Sales Group
        elif line.startswith("Sales Group"):
            value = line.replace("Sales Group", "").strip()
            data["Sales Group"] = value
        
        # Region
        elif line.startswith("Region") and not line.startswith("Regional"):
            value = line.replace("Region", "").strip()
            data["Region"] = value
        
        # Zone (but not Sub Zone or Transportation Zone)
        elif line == "Zone" or (line.startswith("Zone") and not line.startswith("Sub Zone") and not line.startswith("Transportation Zone")):
            value = line.replace("Zone", "").strip()
            if not value and i + 1 < len(lines):
                next_line = lines[i + 1]
                if not is_field_name(next_line):
                    value = next_line
                    i += 1
            data["Zone"] = value
        
        # Sub Zone
        elif line.startswith("Sub Zone"):
            value = line.replace("Sub Zone", "").strip()
            data["Sub Zone"] = value
        
        # State (multiple occurrences - handle first one for customer state)
        elif line.startswith("State") and "State" not in data:
            value = line.replace("State", "").strip()
            if not value and i + 1 < len(lines):
                next_line = lines[i + 1]
                if not is_field_name(next_line):
                    value = next_line
                    i += 1
            data["State"] = value
        
        # Sales Office
        elif line.startswith("Sales Office"):
            value = line.replace("Sales Office", "").strip()
            data["Sales Office"] = value
        
        # SAP Dealer code
        elif line.startswith("SAP Dealer code to be mapped Search Term"):
            # Skip next lines until we find the actual code
            i += 1
            while i < len(lines):
                if lines[i].strip() and not is_field_name(lines[i]) and lines[i].strip() not in ["2", "Search Term"]:
                    data["SAP Dealer code to be mapped Search Term 2"] = lines[i].strip()
                    break
                i += 1
        
        # Search Term 1
        elif line.startswith("Search Term 1- Old customer code"):
            value = line.replace("Search Term 1- Old customer code", "").strip()
            if value and not is_field_name(value):
                data["Search Term 1- Old customer code"] = value
            else:
                data["Search Term 1- Old customer code"] = ""
        
        # Search Term 2
        elif line.startswith("Search Term 2 - District"):
            value = line.replace("Search Term 2 - District", "").strip()
            if value and not is_field_name(value):
                data["Search Term 2 - District"] = value
            else:
                data["Search Term 2 - District"] = ""
        
        # Mobile Number (first occurrence)
        elif line.startswith("Mobile Number") and "Mobile Number" not in data:
            value = line.replace("Mobile Number", "").strip()
            data["Mobile Number"] = value
        
        # E-Mail ID
        elif line.startswith("E-Mail ID"):
            value = line.replace("E-Mail ID", "").strip()
            if value and not is_field_name(value):
                data["E-Mail ID"] = value
            else:
                data["E-Mail ID"] = ""
        
        # Lattitude
        elif line.startswith("Lattitude"):
            value = line.replace("Lattitude", "").strip()
            data["Lattitude"] = value
        
        # Longitude
        elif line.startswith("Longitude"):
            value = line.replace("Longitude", "").strip()
            data["Longitude"] = value
        
        # Name of the Customers (multi-line)
        elif line.startswith("Name of the Customers (Trade Name or"):
            i += 1
            if i < len(lines) and lines[i].startswith("Legal Name)"):
                value = lines[i].replace("Legal Name)", "").strip()
                data["Name of the Customers (Trade Name or Legal Name)"] = value
        
        # E-mail (different from E-Mail ID)
        elif line == "E-mail":
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                if not is_field_name(next_line):
                    data["E-mail"] = next_line
                    i += 1
                else:
                    data["E-mail"] = ""
        
        # Address fields
        elif line.startswith("Address 1"):
            value = line.replace("Address 1", "").strip()
            data["Address 1"] = value
        
        elif line.startswith("Address 2"):
            value = line.replace("Address 2", "").strip()
            if value and not is_field_name(value):
                data["Address 2"] = value
            else:
                data["Address 2"] = ""
        
        elif line.startswith("Address 3"):
            value = line.replace("Address 3", "").strip()
            if value and not is_field_name(value):
                data["Address 3"] = value
            else:
                data["Address 3"] = ""
        
        elif line.startswith("Address 4"):
            value = line.replace("Address 4", "").strip()
            if value and not is_field_name(value):
                data["Address 4"] = value
            else:
                data["Address 4"] = ""
        
        # PIN
        elif line == "PIN" or (line.startswith("PIN") and not line.startswith("PIN Code")):
            value = line.replace("PIN", "").strip()
            if not value and i + 1 < len(lines):
                next_line = lines[i + 1]
                if not is_field_name(next_line):
                    value = next_line
                    i += 1
            data["PIN"] = value
        
        # City
        elif line.startswith("City") and "City" not in data:
            value = line.replace("City", "").strip()
            data["City"] = value
        
        # District
        elif line.startswith("District") and "District" not in data:
            value = line.replace("District", "").strip()
            data["District"] = value
        
        # Continue with other fields...
        elif line.startswith("Whatsapp No."):
            value = line.replace("Whatsapp No.", "").strip()
            data["Whatsapp No."] = value
        
        elif line.startswith("Date of Birth"):
            value = line.replace("Date of Birth", "").strip()
            data["Date of Birth"] = value
        
        elif line.startswith("Date of Anniversary"):
            value = line.replace("Date of Anniversary", "").strip()
            data["Date of Anniversary"] = value
        
        elif line.startswith("Counter Potential - Maximum"):
            value = line.replace("Counter Potential - Maximum", "").strip()
            data["Counter Potential - Maximum"] = value
        
        elif line.startswith("Counter Potential - Minimum"):
            value = line.replace("Counter Potential - Minimum", "").strip()
            data["Counter Potential - Minimum"] = value
        
        elif line.startswith("Is GST Present"):
            value = line.replace("Is GST Present", "").strip()
            data["Is GST Present"] = value
        
        elif line == "GSTIN":
            if i + 1 < len(lines) and not is_field_name(lines[i + 1]):
                data["GSTIN"] = lines[i + 1]
                i += 1
            else:
                data["GSTIN"] = ""
        
        elif line == "Trade Name" and "Trade Name" not in data:
            if i + 1 < len(lines) and not is_field_name(lines[i + 1]):
                data["Trade Name"] = lines[i + 1]
                i += 1
            else:
                data["Trade Name"] = ""
        
        elif line == "Legal Name" and "Legal Name" not in data:
            value = ""
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                if not is_field_name(next_line) and next_line != ")":
                    value = next_line
                    i += 1
            data["Legal Name"] = value
        
        elif line.startswith("Reg Date"):
            value = line.replace("Reg Date", "").strip()
            if value and not is_field_name(value):
                data["Reg Date"] = value
            else:
                data["Reg Date"] = ""
        
        elif line == "Type" and "Type" not in data:
            if i + 1 < len(lines) and not is_field_name(lines[i + 1]):
                data["Type"] = lines[i + 1]
                i += 1
            else:
                data["Type"] = ""
        
        elif line.startswith("Building No."):
            value = line.replace("Building No.", "").strip()
            if value and not is_field_name(value):
                data["Building No."] = value
            else:
                data["Building No."] = ""
        
        elif line.startswith("District Code"):
            value = line.replace("District Code", "").strip()
            if value and not is_field_name(value):
                data["District Code"] = value
            else:
                data["District Code"] = ""
        
        elif line.startswith("State Code"):
            value = line.replace("State Code", "").strip()
            if value and not is_field_name(value):
                data["State Code"] = value
            else:
                data["State Code"] = ""
        
        elif line.startswith("Street"):
            value = line.replace("Street", "").strip()
            if value and not is_field_name(value):
                data["Street"] = value
            else:
                data["Street"] = ""
        
        elif line.startswith("PIN Code"):
            value = line.replace("PIN Code", "").strip()
            if value and not is_field_name(value):
                data["PIN Code"] = value
            else:
                data["PIN Code"] = ""
        
        elif line.startswith("PAN") and not line.startswith("PAN Holder") and not line.startswith("PAN Status") and not line.startswith("PAN -"):
            value = line.replace("PAN", "").strip()
            data["PAN"] = value
        
        elif line.startswith("PAN Holder Name"):
            value = line.replace("PAN Holder Name", "").strip()
            data["PAN Holder Name"] = value
        
        elif line.startswith("PAN Status"):
            value = line.replace("PAN Status", "").strip()
            data["PAN Status"] = value
        
        elif line.startswith("PAN - Aadhaar Linking Status"):
            value = line.replace("PAN - Aadhaar Linking Status", "").strip()
            data["PAN - Aadhaar Linking Status"] = value
        
        elif line.startswith("IFSC Number"):
            value = line.replace("IFSC Number", "").strip()
            data["IFSC Number"] = value
        
        elif line.startswith("Account Number"):
            value = line.replace("Account Number", "").strip()
            data["Account Number"] = value
        
        elif line.startswith("Name of Account Holder"):
            value = line.replace("Name of Account Holder", "").strip()
            data["Name of Account Holder"] = value
        
        elif line.startswith("Bank Name"):
            value = line.replace("Bank Name", "").strip()
            data["Bank Name"] = value
        
        elif line.startswith("Bank Branch"):
            value = line.replace("Bank Branch", "").strip()
            data["Bank Branch"] = value
        
        elif line.startswith("Is Aadhaar Linked with Mobile?"):
            value = line.replace("Is Aadhaar Linked with Mobile?", "").strip()
            data["Is Aadhaar Linked with Mobile?"] = value
        
        elif line.startswith("Aadhaar Number"):
            value = line.replace("Aadhaar Number", "").strip()
            data["Aadhaar Number"] = value
        
        elif line == "Name" and "Name" not in data:
            if i + 1 < len(lines) and not is_field_name(lines[i + 1]):
                data["Name"] = lines[i + 1]
                i += 1
            else:
                data["Name"] = ""
        
        elif line.startswith("Gender"):
            value = line.replace("Gender", "").strip()
            data["Gender"] = value
        
        elif line.startswith("DOB"):
            value = line.replace("DOB", "").strip()
            data["DOB"] = value
        
        elif line == "Address" and "Address" not in data:
            if i + 1 < len(lines) and not is_field_name(lines[i + 1]):
                # Collect multiple lines for address if needed
                address_parts = []
                i += 1
                while i < len(lines) and not is_field_name(lines[i]):
                    address_parts.append(lines[i])
                    i += 1
                    if len(address_parts) >= 2:  # Limit address lines
                        break
                data["Address"] = " ".join(address_parts)
                i -= 1  # Step back one since we'll increment at the end
            else:
                data["Address"] = ""
        
        elif line.startswith("Logistics Transportation Zone"):
            value = line.replace("Logistics Transportation Zone", "").strip()
            data["Logistics Transportation Zone"] = value
        
        elif line.startswith("Transportation Zone Description"):
            value = line.replace("Transportation Zone Description", "").strip()
            data["Transportation Zone Description"] = value
        
        elif line.startswith("Transportation Zone Code"):
            value = line.replace("Transportation Zone Code", "").strip()
            data["Transportation Zone Code"] = value
        
        elif line.startswith("Postal Code"):
            value = line.replace("Postal Code", "").strip()
            if value and not is_field_name(value):
                data["Postal Code"] = value
            else:
                data["Postal Code"] = ""
        
        elif line.startswith("Logistics team to vet the T zone selected by"):
            i += 1
            if i < len(lines) and lines[i] == "Sales Officer":
                i += 1
                if i < len(lines) and not is_field_name(lines[i]):
                    data["Logistics team to vet the T zone selected by Sales Officer"] = lines[i]
                else:
                    data["Logistics team to vet the T zone selected by Sales Officer"] = ""
        
        elif line.startswith("Selection of Available T Zones from T Zone"):
            i += 1
            if i < len(lines) and lines[i].startswith("Master list"):
                i += 1
                if i < len(lines) and not is_field_name(lines[i]):
                    data["Selection of Available T Zones from T Zone Master list, if found."] = lines[i]
                else:
                    data["Selection of Available T Zones from T Zone Master list, if found."] = ""
        
        elif line.startswith("If NEW T Zone need to be created, details to"):
            i += 1
            if i < len(lines) and lines[i].startswith("be provided"):
                i += 1
                if i < len(lines) and not is_field_name(lines[i]):
                    data["If NEW T Zone need to be created, details to be provided by Logistics team"] = lines[i]
                else:
                    data["If NEW T Zone need to be created, details to be provided by Logistics team"] = ""
        
        elif line.startswith("Date of Appointment"):
            value = line.replace("Date of Appointment", "").strip()
            data["Date of Appointment"] = value
        
        elif line.startswith("Delivering Plant"):
            value = line.replace("Delivering Plant", "").strip()
            data["Delivering Plant"] = value
        
        elif line.startswith("Plant Name"):
            value = line.replace("Plant Name", "").strip()
            data["Plant Name"] = value
        
        elif line.startswith("Plant Code"):
            value = line.replace("Plant Code", "").strip()
            data["Plant Code"] = value
        
        elif line.startswith("Incoterns") and not line.startswith("Incoterns Code"):
            value = line.replace("Incoterns", "").strip()
            data["Incoterns"] = value
        
        elif line.startswith("Incoterns Code"):
            value = line.replace("Incoterns Code", "").strip()
            if value and not is_field_name(value):
                data["Incoterns Code"] = value
            else:
                data["Incoterns Code"] = ""
        
        elif line.startswith("Security Deposit Amount details to filled up,"):
            i += 1
            if i < len(lines) and lines[i].startswith("as per"):
                i += 1
                if i < len(lines) and not is_field_name(lines[i]):
                    data["Security Deposit Amount details to filled up, as per checque received by Customer / Dealer"] = lines[i]
                else:
                    data["Security Deposit Amount details to filled up, as per checque received by Customer / Dealer"] = ""
        
        elif line.startswith("Credit Limit (In Rs.)"):
            value = line.replace("Credit Limit (In Rs.)", "").strip()
            if value and not is_field_name(value):
                data["Credit Limit (In Rs.)"] = value
            else:
                data["Credit Limit (In Rs.)"] = ""
        
        elif line.startswith("Regional Head to be mapped"):
            value = line.replace("Regional Head to be mapped", "").strip()
            if value and not is_field_name(value):
                data["Regional Head to be mapped"] = value
            else:
                data["Regional Head to be mapped"] = ""
        
        elif line.startswith("Zonal Head to be mapped"):
            value = line.replace("Zonal Head to be mapped", "").strip()
            data["Zonal Head to be mapped"] = value
        
        elif line.startswith("Sub-Zonal Head (RSM) to be mapped"):
            value = line.replace("Sub-Zonal Head (RSM) to be mapped", "").strip()
            data["Sub-Zonal Head (RSM) to be mapped"] = value
        
        elif line.startswith("Area Sales Manager to be mapped"):
            value = line.replace("Area Sales Manager to be mapped", "").strip()
            data["Area Sales Manager to be mapped"] = value
        
        elif line.startswith("Sales Officer to be mapped"):
            value = line.replace("Sales Officer to be mapped", "").strip()
            if value and not is_field_name(value):
                data["Sales Officer to be mapped"] = value
            else:
                data["Sales Officer to be mapped"] = ""
        
        elif line.startswith("Sales Promoter to be mapped"):
            value = line.replace("Sales Promoter to be mapped", "").strip()
            if value and not is_field_name(value):
                data["Sales Promoter to be mapped"] = value
            else:
                data["Sales Promoter to be mapped"] = ""
        
        elif line.startswith("Sales Promoter Number"):
            value = line.replace("Sales Promoter Number", "").strip()
            if value and not is_field_name(value):
                data["Sales Promoter Number"] = value
            else:
                data["Sales Promoter Number"] = ""
        
        elif line.startswith("Internal control code"):
            value = line.replace("Internal control code", "").strip()
            data["Internal control code"] = value
        
        elif line.startswith("SAP CODE"):
            value = line.replace("SAP CODE", "").strip()
            if value and not is_field_name(value):
                data["SAP CODE"] = value
            else:
                data["SAP CODE"] = ""
        
        elif line.startswith("Initiator Name"):
            value = line.replace("Initiator Name", "").strip()
            data["Initiator Name"] = value
        
        elif line.startswith("Initiator Email ID"):
            value = line.replace("Initiator Email ID", "").strip()
            data["Initiator Email ID"] = value
        
        elif line.startswith("Initiator Mobile Number"):
            value = line.replace("Initiator Mobile Number", "").strip()
            data["Initiator Mobile Number"] = value
        
        elif line.startswith("Created By Customer UserID"):
            value = line.replace("Created By Customer UserID", "").strip()
            data["Created By Customer UserID"] = value
        
        elif line.startswith("Sales Head Name"):
            value = line.replace("Sales Head Name", "").strip()
            data["Sales Head Name"] = value
        
        elif line.startswith("Sales Head Email"):
            value = line.replace("Sales Head Email", "").strip()
            data["Sales Head Email"] = value
        
        elif line.startswith("Sales Head Mobile Number"):
            value = line.replace("Sales Head Mobile Number", "").strip()
            data["Sales Head Mobile Number"] = value
        
        elif line.startswith("Extra2"):
            value = line.replace("Extra2", "").strip()
            if value and not is_field_name(value):
                data["Extra2"] = value
            else:
                data["Extra2"] = ""
        
        elif line.startswith("PAN Result"):
            value = line.replace("PAN Result", "").strip()
            if value and not is_field_name(value):
                data["PAN Result"] = value
            else:
                data["PAN Result"] = ""
        
        elif line.startswith("Mobile Number Result"):
            value = line.replace("Mobile Number Result", "").strip()
            if value and not is_field_name(value):
                data["Mobile Number Result"] = value
            else:
                data["Mobile Number Result"] = ""
        
        elif line.startswith("Email Result"):
            value = line.replace("Email Result", "").strip()
            if value and not is_field_name(value):
                data["Email Result"] = value
            else:
                data["Email Result"] = ""
        
        elif line.startswith("GST Result"):
            value = line.replace("GST Result", "").strip()
            if value and not is_field_name(value):
                data["GST Result"] = value
            else:
                data["GST Result"] = ""
        
        elif line.startswith("Final Result"):
            value = line.replace("Final Result", "").strip()
            data["Final Result"] = value
        
        i += 1
    
    return data

def convert_to_dataframe(data_dict):
    """Convert dictionary to pandas DataFrame with all columns"""
    df_data = []
    for col_name in OUTPUT_COLUMNS:
        value = data_dict.get(col_name, "")
        df_data.append([col_name, value])
    df = pd.DataFrame(df_data, columns=['Field', 'Value'])
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
    st.markdown("### 📊 Info")
    st.info(f"Total fields: {len(OUTPUT_COLUMNS)}")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type=['pdf'])

if uploaded_file is not None:
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    if st.button("🚀 Extract Data", type="primary", use_container_width=True):
        with st.spinner("Extracting data from PDF..."):
            try:
                uploaded_file.seek(0)
                extracted_data = extract_data_from_pdf(uploaded_file)
                df = convert_to_dataframe(extracted_data)
                
                if not show_empty:
                    df = df[df['Value'] != ""]
                
                filled_count = len(df[df['Value'] != ""])
                st.success(f"✅ Successfully extracted data!")
                
                if show_stats:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Fields", len(OUTPUT_COLUMNS))
                    with col2:
                        st.metric("Fields Filled", filled_count)
                    with col3:
                        completion = (filled_count / len(OUTPUT_COLUMNS)) * 100
                        st.metric("Completion", f"{completion:.1f}%")
                
                st.subheader("📋 Extracted Data")
                st.dataframe(df, use_container_width=True, height=500)
                
                st.subheader("📥 Download Options")
                col1, col2 = st.columns(2)
                
                with col1:
                    csv_data = convert_df_to_csv(df)
                    st.download_button(
                        "Download as CSV",
                        csv_data,
                        f"extracted_customer_data_{uploaded_file.name.replace('.pdf', '')}.csv",
                        "text/csv"
                    )
                
                with col2:
                    excel_data = convert_df_to_excel(df)
                    st.download_button(
                        "Download as Excel",
                        excel_data,
                        f"extracted_customer_data_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                with st.expander("Error Details"):
                    st.code(str(e))

else:
    st.info("👆 Upload a PDF file to begin extraction")
    
    with st.expander("📖 Instructions"):
        st.markdown("""
        ### How to use:
        1. Upload your customer creation PDF file
        2. Click "Extract Data"
        3. Review the extracted data in the table
        4. Download as CSV or Excel
        
        ### Features:
        - ✅ Extracts all 94 customer fields
        - ✅ Handles empty fields correctly
        - ✅ Shows completion statistics
        - ✅ Export to CSV or Excel
        """)

st.markdown("---")
st.markdown("Built with ❤️ using Streamlit")
