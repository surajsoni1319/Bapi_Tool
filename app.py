import streamlit as st
import pdfplumber
import pandas as pd
import re

st.set_page_config(page_title="PDF to Table Extractor", layout="wide")

st.title("📄 SAP PDF → Tabular Data Extractor")
st.write("Upload SAP-style PDFs and extract data in clean tabular format.")

# Section headers to ignore
IGNORE_HEADERS = [
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
]

def extract_text_from_pdf(pdf_file):
    full_text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    return full_text


def clean_and_extract_key_values(text):
    data = {}

    lines = text.split("\n")

    for line in lines:
        line = line.strip()

        # Skip empty lines
        if not line:
            continue

        # Skip section headers
        if any(header.lower() in line.lower() for header in IGNORE_HEADERS):
            continue

        # Remove multiple spaces
        line = re.sub(r"\s{2,}", " ", line)

        # Split key-value using first space gap
        match = re.match(r"^(.*?)(\s{1,})(.+)$", line)
        if match:
            key = match.group(1).strip()
            value = match.group(3).strip()

            # Avoid numeric-only junk lines
            if len(key) > 2:
                data[key] = value

    return data


uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("Extracting data..."):
        raw_text = extract_text_from_pdf(uploaded_file)
        extracted_data = clean_and_extract_key_values(raw_text)

        if extracted_data:
            df = pd.DataFrame(
                extracted_data.items(),
                columns=["Field Name", "Value"]
            )

            st.success("✅ Data extracted successfully")
            st.dataframe(df, use_container_width=True)

            col1, col2 = st.columns(2)

            with col1:
                st.download_button(
                    "⬇ Download CSV",
                    df.to_csv(index=False),
                    file_name="extracted_data.csv",
                    mime="text/csv"
                )

            with col2:
                st.download_button(
                    "⬇ Download Excel",
                    df.to_excel(index=False),
                    file_name="extracted_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("No valid data found in the PDF.")
