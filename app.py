import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(
    page_title="PDF → Excel (Raw Conversion)",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Excel Converter (RAW)")
st.write("This converts PDF text to Excel exactly as-is. No parsing, no logic.")

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# Define headers to exclude
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

def pdf_to_raw_excel(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue
            lines = text.split("\n")
            for line_no, line in enumerate(lines, start=1):
                cleaned_line = line.strip()
                
                # Skip if line is empty or matches a header
                if not cleaned_line or cleaned_line in HEADERS_TO_EXCLUDE:
                    continue
                
                rows.append({
                    "Source File": pdf_file.name,
                    "Page": page_no,
                    "Line No": line_no,
                    "Text": cleaned_line
                })
    return pd.DataFrame(rows)

if uploaded_files:
    all_dfs = []
    for pdf in uploaded_files:
        df = pdf_to_raw_excel(pdf)
        all_dfs.append(df)
    
    final_df = pd.concat(all_dfs, ignore_index=True)
    
    st.subheader("📊 Preview (Raw PDF Content)")
    st.dataframe(final_df, use_container_width=True)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="RAW_PDF_TEXT")
    
    st.download_button(
        "⬇️ Download Excel",
        data=output.getvalue(),
        file_name="PDF_RAW_TO_EXCEL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
