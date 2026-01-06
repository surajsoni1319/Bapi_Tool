import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(
    page_title="PDF → Excel (Structured Conversion)",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Structured Excel Converter")
st.write("Converts PDF to clean field-value pairs in Excel format.")

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# Headers to exclude
HEADERS_TO_EXCLUDE = {
    "Customer Creation",
    "Customer Address",
    "GSTIN Details",
    "PAN Details",
    "Bank Details",
    "Aadhaar Details",
    "Supporting Document",
    "Zone Details",
    "Other Details",
    "Duplicity Check"
}

# Known field patterns (helps identify where fields end)
KNOWN_FIELD_PATTERNS = [
    r"^SAP Dealer code to be mapped",
    r"^Name of the Customers",
    r"^Address\s",
    r"^Logistics team to vet",
    r"^Selection of Available T Zones",
    r"^If NEW T Zone need to be created",
    r"^Security Deposit Amount details",
    r"^[A-Z][a-zA-Z\s]+(?:to|from|for|of|the|by|as|per)\s"
]

def extract_raw_lines(pdf_file):
    """Extract all lines from PDF with metadata"""
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue
            lines = text.split("\n")
            for line_no, line in enumerate(lines, start=1):
                cleaned_line = line.strip()
                if not cleaned_line or cleaned_line in HEADERS_TO_EXCLUDE:
                    continue
                rows.append({
                    "source_file": pdf_file.name,
                    "page": page_no,
                    "line_no": line_no,
                    "text": cleaned_line
                })
    return rows

def looks_like_continuation(text):
    """Check if text looks like a continuation line"""
    # Starts with lowercase or common continuation words
    continuation_patterns = [
        r"^[a-z]",  # Starts with lowercase
        r"^(Legal Name|Master list|Sales Officer|be provided|as per)",  # Common continuations
        r"^[0-9]",  # Starts with number (could be continuation)
        r"^PO\s",  # Address continuation
    ]
    for pattern in continuation_patterns:
        if re.match(pattern, text):
            return True
    return False

def split_field_value(text):
    """Try to split a line into field and value"""
    # Pattern 1: Field ending with common words, then value
    # Look for capital letter sequences ending with specific words, followed by uppercase or numbers
    match = re.match(r'^(.+?(?:to|from|for|of|the|by|as per|Master list|Search Term|or))\s+([A-Z0-9].+)$', text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    
    # Pattern 2: "Address" followed by address text
    match = re.match(r'^(Address)\s+(.+)$', text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    
    # Pattern 3: Long field name with embedded value (uppercase followed by UPPERCASE word)
    match = re.match(r'^(.+?)\s+([A-Z][A-Z\s]+(?:[A-Z]+|[0-9]+).*)$', text)
    if match and len(match.group(2)) > 3:
        return match.group(1).strip(), match.group(2).strip()
    
    # Pattern 4: Field ending with specific words before a value
    match = re.match(r'^(.+?(?:Officer|Dealer|team|list|found|Name))\s+(.+)$', text)
    if match:
        field = match.group(1).strip()
        value = match.group(2).strip()
        # Only split if value looks like actual data
        if re.match(r'^[A-Z0-9]', value) or value.lower() in ['yes', 'no']:
            return field, value
    
    return None, text

def parse_lines_to_fields(lines):
    """Parse lines into structured field-value pairs"""
    structured_data = []
    i = 0
    
    while i < len(lines):
        current = lines[i]
        current_text = current["text"]
        
        # Try to split current line into field and value
        field, value = split_field_value(current_text)
        
        # Look ahead for continuation lines
        continuation_lines = []
        j = i + 1
        while j < len(lines):
            next_line = lines[j]
            # Check if next line is from same page or consecutive page
            if next_line["page"] - current["page"] > 1:
                break
            
            next_text = next_line["text"]
            
            # If next line looks like continuation
            if looks_like_continuation(next_text):
                continuation_lines.append(next_line)
                j += 1
            else:
                # Check if it might be part of the field name
                # (doesn't start with uppercase word that could be a new field)
                if not re.match(r'^[A-Z][a-z]+\s+[A-Z]', next_text):
                    continuation_lines.append(next_line)
                    j += 1
                else:
                    break
        
        # Reconstruct field and value
        if field:
            # We already have a field, append continuations to value
            all_value_parts = [value] + [c["text"] for c in continuation_lines]
            final_value = " ".join(all_value_parts)
        else:
            # No clear field-value split
            if continuation_lines:
                # First line is field, rest is value
                final_field = current_text
                final_value = " ".join([c["text"] for c in continuation_lines])
                
                # Unless the continuation completes the field name
                if looks_like_continuation(continuation_lines[0]["text"]):
                    # Could be field continuation
                    combined = current_text + " " + continuation_lines[0]["text"]
                    field_part, value_part = split_field_value(combined)
                    if field_part:
                        final_field = field_part
                        final_value = value_part
                        if len(continuation_lines) > 1:
                            final_value += " " + " ".join([c["text"] for c in continuation_lines[1:]])
                
                field = final_field
                value = final_value
            else:
                # Single line with no clear split - treat as field with empty value
                field = current_text
                value = ""
        
        # Clean up field and value
        field = field.strip()
        value = value.strip()
        
        # Handle empty values
        if not value or value == "":
            value = "Blank - no value"
        
        structured_data.append({
            "Source File": current["source_file"],
            "Page": current["page"],
            "Field Name": field,
            "Value": value
        })
        
        # Move index forward
        i = j if continuation_lines else i + 1
    
    return structured_data

if uploaded_files:
    all_structured = []
    
    for pdf in uploaded_files:
        # Extract raw lines
        raw_lines = extract_raw_lines(pdf)
        
        # Parse into structured data
        structured = parse_lines_to_fields(raw_lines)
        all_structured.extend(structured)
    
    # Create DataFrame
    final_df = pd.DataFrame(all_structured)
    
    # Display preview
    st.subheader("📊 Structured Data Preview")
    st.dataframe(final_df, use_container_width=True)
    
    # Download button
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Structured_Data")
    
    st.download_button(
        "⬇️ Download Structured Excel",
        data=output.getvalue(),
        file_name="PDF_STRUCTURED_DATA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # Show statistics
    st.subheader("📈 Statistics")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Fields", len(final_df))
    with col2:
        st.metric("Unique Fields", final_df["Field Name"].nunique())
    with col3:
        st.metric("Empty Values", (final_df["Value"] == "Blank - no value").sum())
