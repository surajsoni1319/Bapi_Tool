import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(
    page_title="SAP PDF → Final Excel",
    page_icon="📄",
    layout="wide"
)

st.title("📄 SAP Customer Creation PDF → Final Excel")
st.write("PDF → Raw rows → Section-aware extraction → Final SAP-ready Excel")
st.markdown("---")

uploaded_files = st.file_uploader(
    "Upload SAP Customer Creation PDF(s)",
    type=["pdf"],
    accept_multiple_files=True
)

# =====================================================
# REQUIRED FINAL FIELDS (EXACT)
# =====================================================
FIELDS = [
    "Type of Customer","Name of Customer","Company Code","Customer Group",
    "Sales Group","Region","Zone","Sub Zone","State","Sales Office",
    "SAP Dealer code to be mapped Search Term 2",
    "Search Term 1- Old customer code","Search Term 2 - District",
    "Mobile Number","E-Mail ID","Lattitude","Longitude",
    "Name of the Customers (Trade Name or Legal Name)",
    "Mobile Number (Address)","E-mail",
    "Address 1","Address 2","Address 3","Address 4",
    "PIN","City","District","State (Address)",
    "Whatsapp No.","Date of Birth","Date of Anniversary",
    "Counter Potential - Maximum","Counter Potential - Minimum",
    "Is GST Present","GSTIN","Trade Name","Legal Name","Reg Date",
    "City (GST)","Type","Building No.","District Code","State Code",
    "Street","PIN Code",
    "PAN","PAN Holder Name","PAN Status","PAN - Aadhaar Linking Status",
    "IFSC Number","Account Number","Name of Account Holder",
    "Bank Name","Bank Branch",
    "Is Aadhaar Linked with Mobile?","Aadhaar Number","Name","Gender","DOB",
    "Address (Aadhaar)","PIN (Aadhaar)","City (Aadhaar)","State (Aadhaar)",
    "Logistics Transportation Zone","Transportation Zone Description",
    "Transportation Zone Code","Postal Code",
    "Logistics team to vet the T zone selected by Sales Officer",
    "Selection of Available T Zones from T Zone Master list, if found.",
    "If NEW T Zone need to be created, details to be provided by Logistics team",
    "Date of Appointment","Delivering Plant","Plant Name","Plant Code",
    "Incoterns","Incoterns (2)","Incoterns Code",
    "Security Deposit Amount details to filled up, as per checque received by Customer / Dealer",
    "Credit Limit (In Rs.)","Regional Head to be mapped","Zonal Head to be mapped",
    "Sub-Zonal Head (RSM) to be mapped","Area Sales Manager to be mapped",
    "Sales Officer to be mapped","Sales Promoter to be mapped",
    "Sales Promoter Number","Internal control code","SAP CODE",
    "Initiator Name","Initiator Email ID","Initiator Mobile Number",
    "Created By Customer UserID","Sales Head Name","Sales Head Email",
    "Sales Head Mobile Number","Extra2","PAN Result",
    "Mobile Number Result","Email Result","GST Result","Final Result"
]

# =====================================================
# STEP 1: PDF → RAW ROWS (WITH LINE MERGING)
# =====================================================
def pdf_to_raw_rows(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue

            buffer = ""
            for line_no, line in enumerate(text.split("\n"), start=1):
                line = line.strip()

                if buffer:
                    line = buffer + " " + line
                    buffer = ""

                if line.endswith("(") or line.endswith(",") or line.endswith("to"):
                    buffer = line
                    continue

                rows.append({
                    "Page": page_no,
                    "Line No": line_no,
                    "Text": line
                })

    return pd.DataFrame(rows)

# =====================================================
# HELPERS ON RAW DATA
# =====================================================
def same_line(df, label):
    m = df[df["Text"].str.startswith(label + " ")]
    return m.iloc[0]["Text"].replace(label, "").strip() if not m.empty else ""

def next_line(df, label):
    idx = df.index[df["Text"] == label]
    return df.loc[idx[0] + 1, "Text"] if not idx.empty and idx[0] + 1 in df.index else ""

# =====================================================
# STEP 2: RAW ROWS → FINAL ROW
# =====================================================
def build_final_row(df):
    row = {f: "" for f in FIELDS}

    # ---------- HEADER ----------
    for f in [
        "Type of Customer","Name of Customer","Company Code","Customer Group",
        "Sales Group","Region","Zone","Sub Zone","State","Sales Office",
        "Mobile Number","E-Mail ID","Lattitude","Longitude",
        "Whatsapp No.","Date of Birth","Date of Anniversary",
        "Counter Potential - Maximum","Counter Potential - Minimum"
    ]:
        row[f] = same_line(df, f)

    # ---------- SEARCH TERM ----------
    row["SAP Dealer code to be mapped Search Term 2"] = same_line(
        df, "SAP Dealer code to be mapped Search Term"
    )

    # ---------- ADDRESS ----------
    row["Address 1"] = same_line(df, "Address 1")
    row["Address 2"] = same_line(df, "Address 2")
    row["Address 3"] = next_line(df, "Address 3")
    row["Address 4"] = next_line(df, "Address 4")
    row["PIN"] = same_line(df, "PIN")
    row["City"] = same_line(df, "City")
    row["District"] = same_line(df, "District")

    # ---------- GST BLOCK ----------
    gst_idx = df.index[df["Text"] == "GSTIN Details"]
    if not gst_idx.empty:
        i = gst_idx[0]
        seq = [
            "Is GST Present","GSTIN","Trade Name","Legal Name","Reg Date",
            "City (GST)","Type","Building No.","District Code",
            "State Code","Street","PIN Code"
        ]
        for j, field in enumerate(seq, start=1):
            row[field] = df.loc[i + j, "Text"]

    # ---------- PAN ----------
    row["PAN"] = same_line(df, "PAN")
    row["PAN Holder Name"] = same_line(df, "PAN Holder Name")
    row["PAN Status"] = same_line(df, "PAN Status")
    row["PAN - Aadhaar Linking Status"] = same_line(df, "PAN - Aadhaar Linking Status")

    # ---------- BANK ----------
    for f in [
        "IFSC Number","Account Number","Name of Account Holder",
        "Bank Name","Bank Branch"
    ]:
        row[f] = same_line(df, f)

    # ---------- AADHAAR BLOCK ----------
    ad_idx = df.index[df["Text"] == "Aadhaar Details"]
    if not ad_idx.empty:
        i = ad_idx[0]
        seq = [
            "Is Aadhaar Linked with Mobile?","Aadhaar Number","Name",
            "Gender","DOB","Address (Aadhaar)",
            "PIN (Aadhaar)","City (Aadhaar)","State (Aadhaar)"
        ]
        for j, field in enumerate(seq, start=1):
            row[field] = df.loc[i + j, "Text"]

    # ---------- LOGISTICS ----------
    for f in [
        "Logistics Transportation Zone",
        "Transportation Zone Description",
        "Transportation Zone Code"
    ]:
        row[f] = same_line(df, f)

    # ---------- SECURITY DEPOSIT ----------
    sd = df[df["Text"].str.startswith("Security Deposit Amount")]
    if not sd.empty:
        row["Security Deposit Amount details to filled up, as per checque received by Customer / Dealer"] = sd.iloc[0]["Text"]

    # ---------- RESULTS ----------
    for f in [
        "PAN Result","Mobile Number Result",
        "Email Result","GST Result","Final Result"
    ]:
        row[f] = same_line(df, f)

    return row

# =====================================================
# PIPELINE
# =====================================================
if uploaded_files:
    final_rows = []

    for pdf in uploaded_files:
        raw_df = pdf_to_raw_rows(pdf)
        final_row = build_final_row(raw_df)
        final_row["Source File"] = pdf.name
        final_rows.append(final_row)

    final_df = pd.DataFrame(final_rows)

    st.subheader("✅ Final Output (SAP-Ready)")
    st.dataframe(final_df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="SAP_Data")

    st.download_button(
        "⬇️ Download Final Excel",
        data=output.getvalue(),
        file_name="Final_SAP_Customer_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
