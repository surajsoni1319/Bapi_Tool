import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

# --------------------------------------------------
# Page config
# --------------------------------------------------
st.set_page_config(
    page_title="PDF Customer Data Extractor",
    page_icon="📄",
    layout="wide"
)

# --------------------------------------------------
# Title
# --------------------------------------------------
st.title("📄 PDF Customer Data Extractor")
st.markdown("---")

# --------------------------------------------------
# Expected Fields (UNCHANGED)
# --------------------------------------------------
EXPECTED_FIELDS = [
    "Type of Customer","Name of Customer","Company Code","Customer Group",
    "Sales Group","Region","Zone","Sub Zone","State","Sales Office",
    "SAP Dealer code to be mapped Search Term 2","Search Term 1- Old customer code",
    "Search Term 2 - District","Mobile Number","E-Mail ID","Lattitude","Longitude",
    "Name of the Customers (Trade Name or Legal Name)","Mobile Number","E-mail",
    "Address 1","Address 2","Address 3","Address 4","PIN","City","District","State",
    "Whatsapp No.","Date of Birth","Date of Anniversary",
    "Counter Potential - Maximum","Counter Potential - Minimum",
    "Is GST Present","GSTIN","Trade Name","Legal Name","Reg Date","City","Type",
    "Building No.","District Code","State Code","Street","PIN Code",
    "PAN","PAN Holder Name","PAN Status","PAN - Aadhaar Linking Status",
    "IFSC Number","Account Number","Name of Account Holder","Bank Name","Bank Branch",
    "Is Aadhaar Linked with Mobile?","Aadhaar Number","Name","Gender","DOB","Address",
    "PIN","City","State","Logistics Transportation Zone",
    "Transportation Zone Description","Transportation Zone Code","Postal Code",
    "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment","Delivering Plant","Plant Name","Plant Code",
    "Incoterns","Incoterns","Incoterns Code",
    "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)","Regional Head to be mapped","Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped","Area Sales Manager to be mapped",
    "Sales Officer to be mapped","Sales Promoter to be mapped",
    "Sales Promoter Number","Internal control code","SAP CODE","Initiator Name",
    "Initiator Email ID","Initiator Mobile Number","Created By Customer UserID",
    "Sales Head Name","Sales Head Email","Sales Head Mobile Number","Extra2",
    "PAN Result","Mobile Number Result","Email Result","GST Result","Final Result"
]

# --------------------------------------------------
# Helpers
# --------------------------------------------------
def normalize_text(txt):
    return re.sub(r"\s+", " ", txt.strip())


def extract_layout_pairs(pdf_bytes):
    """
    Layout-based extraction using X/Y coordinates
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pairs = []

    for page in doc:
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if "lines" not in block:
                continue

            for line in block["lines"]:
                spans = []
                for span in line["spans"]:
                    text = normalize_text(span["text"])
                    if text:
                        spans.append({
                            "text": text,
                            "x": span["bbox"][0]
                        })

                if len(spans) >= 2:
                    spans = sorted(spans, key=lambda s: s["x"])
                    field = spans[0]["text"]
                    value = " ".join(s["text"] for s in spans[1:])
                    pairs.append((field, value))
                elif len(spans) == 1:
                    pairs.append((spans[0]["text"], ""))

    doc.close()
    return pairs


# --------------------------------------------------
# MAIN EXTRACTION FUNCTION (UPDATED)
# --------------------------------------------------
def extract_data_from_pdf(pdf_file):
    pdf_bytes = pdf_file.read()

    raw_pairs = extract_layout_pairs(pdf_bytes)

    raw_df = pd.DataFrame(raw_pairs, columns=["Field", "Value"])
    raw_df["Field"] = raw_df["Field"].str.strip()
    raw_df["Value"] = raw_df["Value"].str.strip()
    raw_df = raw_df[raw_df["Field"] != ""]

    output_data = []
    field_counts = {}

    for expected_field in EXPECTED_FIELDS:
        field_counts[expected_field] = field_counts.get(expected_field, 0) + 1
        occurrence = field_counts[expected_field]

        value = ""

        matches = raw_df[
            raw_df["Field"].str.fullmatch(
                re.escape(expected_field), case=False, na=False
            )
        ]

        if not matches.empty:
            if occurrence <= len(matches):
                value = matches.iloc[occurrence - 1]["Value"]
            else:
                value = matches.iloc[0]["Value"]

        output_data.append({"Field": expected_field, "Value": value})

    final_df = pd.DataFrame(output_data)
    return final_df, raw_df


# --------------------------------------------------
# Sidebar
# --------------------------------------------------
with st.sidebar:
    st.header("⚙️ Options")
    show_raw = st.checkbox("Show raw extraction", False)
    hide_empty = st.checkbox("Hide empty fields", False)
    show_stats = st.checkbox("Show statistics", True)
    st.markdown("---")
    st.info(f"Total Expected Fields: {len(EXPECTED_FIELDS)}")

# --------------------------------------------------
# File uploader
# --------------------------------------------------
uploaded_file = st.file_uploader("Upload Customer Creation PDF", type=["pdf"])

if uploaded_file:
    st.success(f"Uploaded: {uploaded_file.name}")

    if st.button("🚀 Extract Data", type="primary", use_container_width=True):
        with st.spinner("Extracting using layout detection..."):
            final_df, raw_df = extract_data_from_pdf(uploaded_file)

        if show_raw:
            st.subheader("🔍 Raw Layout Extraction (Debug)")
            st.dataframe(raw_df, use_container_width=True, height=350)

        display_df = final_df.copy()
        if hide_empty:
            display_df = display_df[display_df["Value"] != ""]

        if show_stats:
            filled = (final_df["Value"] != "").sum()
            total = len(final_df)
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Fields", total)
            c2.metric("Fields Filled", filled)
            c3.metric("Completion", f"{(filled/total)*100:.1f}%")

        st.subheader("📋 Extracted Data (Row-wise)")
        st.dataframe(display_df, use_container_width=True, height=500)

        # Downloads
        csv_data = final_df.to_csv(index=False).encode("utf-8")
        excel_output = BytesIO()
        with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Customer Data")

        st.markdown("### 📥 Download")
        col1, col2 = st.columns(2)
        col1.download_button("Download CSV", csv_data, "customer_data.csv", "text/csv")
        col2.download_button(
            "Download Excel",
            excel_output.getvalue(),
            "customer_data.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("👆 Upload a PDF to start extraction")

st.markdown("---")
st.markdown("Built with Streamlit • PyMuPDF • Layout-based Extraction")
