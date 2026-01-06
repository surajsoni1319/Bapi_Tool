import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# =====================
# STREAMLIT CONFIG
# =====================
st.set_page_config(
    page_title="PDF → Clean Excel (In-Memory)",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF → Clean Excel Converter")
st.write("Raw extraction ➜ rule-based cleaning ➜ final Excel (all in memory)")
st.markdown("---")

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# =====================
# STEP 1: RAW EXTRACTION
# =====================
def extract_raw_lines(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue
            for line_no, line in enumerate(text.split("\n"), start=1):
                rows.append({
                    "Source File": pdf_file.name,
                    "Page": page_no,
                    "Line No": line_no,
                    "Text": line.strip()
                })
    return rows

# =====================
# STEP 2: CLEANING LOGIC
# =====================
HEADERS_TO_DROP = {
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

def clean_lines(lines):
    clean_rows = []
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        # 1. Drop headers
        if line in HEADERS_TO_DROP:
            i += 1
            continue

        # 2. SAP Search Term2
        if line.startswith("SAP Dealer code to be mapped Search Term"):
            value = re.findall(r"\d+", line)
            clean_rows.append({
                "Field": "SAP Dealer code to be mapped Search Term2",
                "Value": value[0] if value else ""
            })
            i += 2
            continue

        # 3. Broken Trade Name field
        if line.startswith("Name of the Customers (Trade Name or"):
            value = line.replace(
                "Name of the Customers (Trade Name or", ""
            ).strip()
            clean_rows.append({
                "Field": "Name of the Customers (Trade Name or Legal Name)",
                "Value": value
            })
            i += 2
            continue

        # 4. Multi-line Address
        if line.startswith("Address "):
            value = line.replace("Address", "").strip()
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if not re.match(r"^[A-Z][a-zA-Z ]+ \d+", next_line):
                    value = value + " " + next_line
                    i += 1
            clean_rows.append({"Field": "Address", "Value": value})
            i += 1
            continue

        # 5. Logistics vet field
        if line.startswith("Logistics team to vet the T zone selected by"):
            value = "No" if "No" in line else ""
            clean_rows.append({
                "Field": "Logistics team to vet the T zone selected by Sales Officer",
                "Value": value
            })
            i += 2
            continue

        # 6. Split long label (no value)
        if line.startswith("If NEW T Zone need to be created"):
            field = line
            if i + 1 < len(lines):
                field = field + " " + lines[i + 1].strip()
                i += 1
            clean_rows.append({"Field": field, "Value": ""})
            i += 1
            continue

        # 7. Security Deposit
        if line.startswith("Security Deposit Amount details"):
            amount = re.findall(r"\d+", line)
            field = line.replace(amount[0], "").strip() if amount else line
            if i + 1 < len(lines):
                field = field + " " + lines[i + 1].strip()
                i += 1
            clean_rows.append({
                "Field": field,
                "Value": amount[0] if amount else ""
            })
            i += 1
            continue

        # Default: split on last space
        parts = line.rsplit(" ", 1)
        if len(parts) == 2:
            clean_rows.append({"Field": parts[0], "Value": parts[1]})
        else:
            clean_rows.append({"Field": line, "Value": ""})

        i += 1

    return pd.DataFrame(clean_rows)

# =====================
# PIPELINE EXECUTION
# =====================
if uploaded_files:
    raw_rows = []

    for pdf in uploaded_files:
        raw_rows.extend(extract_raw_lines(pdf))

    raw_df = pd.DataFrame(raw_rows)

    st.subheader("🔍 Raw Extracted Data")
    st.dataframe(raw_df, use_container_width=True)

    clean_df = clean_lines(raw_df["Text"].tolist())

    st.subheader("✅ Cleaned Final Output (Field | Value)")
    st.dataframe(clean_df, use_container_width=True)

    # =====================
    # DOWNLOAD EXCEL (IN MEMORY)
    # =====================
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        clean_df.to_excel(writer, index=False, sheet_name="FINAL_CLEAN_DATA")

    st.download_button(
        "⬇️ Download Final Excel",
        data=output.getvalue(),
        file_name="FINAL_CLEAN_EXCEL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
