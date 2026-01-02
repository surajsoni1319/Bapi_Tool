import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO

# --------------------------------------------------
# Page Config
# --------------------------------------------------
st.set_page_config(page_title="Layout Based PDF Extractor", layout="wide")

# --------------------------------------------------
# Layout-Based Extraction
# --------------------------------------------------
def extract_layout_rows(pdf_bytes, y_tolerance=3):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    rows = []

    for page in doc:
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if "lines" not in block:
                continue

            for line in block["lines"]:
                y = round(line["bbox"][1])

                texts = []
                for span in line["spans"]:
                    texts.append({
                        "text": span["text"].strip(),
                        "x": span["bbox"][0],
                        "y": y
                    })

                if len(texts) >= 2:
                    rows.append(texts)

    doc.close()
    return rows


def detect_field_value_pairs(rows):
    structured_data = {}

    for row in rows:
        row = sorted(row, key=lambda x: x["x"])

        field = row[0]["text"]
        value = " ".join([r["text"] for r in row[1:]])

        if field and value:
            structured_data[field] = value

    return structured_data


def create_excel_rowwise(data):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df = pd.DataFrame(
            list(data.items()),
            columns=["Field Name", "Value"]
        )
        df.to_excel(writer, sheet_name="Extracted Data", index=False)

        ws = writer.sheets["Extracted Data"]
        ws.column_dimensions["A"].width = 45
        ws.column_dimensions["B"].width = 60

    output.seek(0)
    return output

# --------------------------------------------------
# UI
# --------------------------------------------------
st.title("📄 Layout-Based PDF → Row-wise Excel")
st.write("This uses **STRUCTURE DETECTION (coordinates)** like iLovePDF.")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    pdf_bytes = uploaded_file.read()

    if st.button("🚀 Extract Using Layout"):
        layout_rows = extract_layout_rows(pdf_bytes)

        structured_data = detect_field_value_pairs(layout_rows)

        if structured_data:
            st.success("Structure detected successfully")

            st.subheader("📋 Detected Structure (Temp Data)")
            st.dataframe(
                pd.DataFrame(
                    list(structured_data.items()),
                    columns=["Field Name", "Value"]
                ),
                use_container_width=True,
                hide_index=True
            )

            excel = create_excel_rowwise(structured_data)

            st.download_button(
                "📥 Download Row-wise Excel",
                data=excel,
                file_name="layout_based_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            st.error("No structure detected")
