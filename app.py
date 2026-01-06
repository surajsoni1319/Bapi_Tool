import streamlit as st
import pdfplumber
import pandas as pd
import io
from datetime import datetime

st.set_page_config(
    page_title="PDF → Excel Converter",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF to Excel Converter")
st.markdown("""
Convert PDF documents to Excel format with raw text extraction.  
Each line from your PDF will be preserved exactly as-is in the output.
""")

# Sidebar options
with st.sidebar:
    st.header("⚙️ Options")
    
    include_empty_lines = st.checkbox(
        "Include empty lines",
        value=False,
        help="Keep blank lines from the PDF in the output"
    )
    
    separate_sheets = st.checkbox(
        "Separate sheet per file",
        value=False,
        help="Create a different Excel sheet for each PDF file"
    )
    
    add_metadata = st.checkbox(
        "Add extraction metadata",
        value=True,
        help="Include timestamp and file information"
    )
    
    st.divider()
    st.markdown("### 💡 Tips")
    st.markdown("""
    - Upload multiple PDFs at once
    - Large files may take longer
    - Preview shows first 1000 rows
    """)

# File uploader
uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True,
    help="Select one or more PDF files to convert"
)

def pdf_to_raw_excel(pdf_file, include_empty=False):
    """Extract raw text from PDF and return as DataFrame"""
    rows = []
    
    try:
        with pdfplumber.open(pdf_file) as pdf:
            total_pages = len(pdf.pages)
            
            for page_no, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                
                if not text:
                    continue
                
                lines = text.split("\n")
                
                for line_no, line in enumerate(lines, start=1):
                    line_stripped = line.strip()
                    
                    # Skip empty lines if option is disabled
                    if not include_empty and not line_stripped:
                        continue
                    
                    rows.append({
                        "Source_File": pdf_file.name,
                        "Page": page_no,
                        "Total_Pages": total_pages,
                        "Line_No": line_no,
                        "Text": line_stripped if include_empty else line_stripped
                    })
        
        return pd.DataFrame(rows), None
    
    except Exception as e:
        return None, str(e)

if uploaded_files:
    # Progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    all_dfs = []
    errors = []
    
    # Process each file
    for idx, pdf in enumerate(uploaded_files):
        status_text.text(f"Processing: {pdf.name} ({idx + 1}/{len(uploaded_files)})")
        progress_bar.progress((idx + 1) / len(uploaded_files))
        
        df, error = pdf_to_raw_excel(pdf, include_empty_lines)
        
        if error:
            errors.append(f"❌ {pdf.name}: {error}")
        elif df is not None and not df.empty:
            all_dfs.append(df)
        else:
            errors.append(f"⚠️ {pdf.name}: No text content found")
    
    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()
    
    # Display errors if any
    if errors:
        with st.expander("⚠️ Warnings & Errors", expanded=True):
            for error in errors:
                st.warning(error)
    
    # Process results
    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        
        # Display statistics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Files Processed", len(all_dfs))
        with col2:
            st.metric("Total Lines", len(final_df))
        with col3:
            st.metric("Total Pages", final_df['Total_Pages'].sum())
        with col4:
            st.metric("Unique Files", final_df['Source_File'].nunique())
        
        # Preview section
        st.subheader("📊 Data Preview")
        
        # Filter options
        col1, col2 = st.columns([2, 1])
        with col1:
            selected_file = st.selectbox(
                "Filter by file:",
                options=["All Files"] + list(final_df['Source_File'].unique())
            )
        with col2:
            preview_rows = st.number_input(
                "Preview rows:",
                min_value=10,
                max_value=5000,
                value=100,
                step=50
            )
        
        # Apply filter
        if selected_file != "All Files":
            preview_df = final_df[final_df['Source_File'] == selected_file].head(preview_rows)
        else:
            preview_df = final_df.head(preview_rows)
        
        st.dataframe(
            preview_df,
            use_container_width=True,
            height=400
        )
        
        if len(final_df) > preview_rows:
            st.info(f"Showing {len(preview_df)} of {len(final_df)} total rows")
        
        # Export section
        st.subheader("💾 Export Options")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            filename = st.text_input(
                "Output filename:",
                value=f"PDF_Export_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        
        # Generate Excel file
        output = io.BytesIO()
        
        try:
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                if separate_sheets and len(all_dfs) > 1:
                    # Create separate sheets for each file
                    for df in all_dfs:
                        sheet_name = df['Source_File'].iloc[0][:31]  # Excel limit
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                else:
                    # Single sheet with all data
                    final_df.to_excel(writer, index=False, sheet_name="RAW_PDF_TEXT")
                
                # Add metadata sheet if requested
                if add_metadata:
                    metadata_df = pd.DataFrame({
                        "Property": [
                            "Export Date",
                            "Total Files",
                            "Total Lines Extracted",
                            "Total Pages",
                            "Empty Lines Included"
                        ],
                        "Value": [
                            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            len(all_dfs),
                            len(final_df),
                            final_df['Total_Pages'].sum(),
                            "Yes" if include_empty_lines else "No"
                        ]
                    })
                    metadata_df.to_excel(writer, index=False, sheet_name="Metadata")
            
            # Download button
            st.download_button(
                label="⬇️ Download Excel File",
                data=output.getvalue(),
                file_name=f"{filename}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.success("✅ Excel file ready for download!")
            
        except Exception as e:
            st.error(f"Error creating Excel file: {str(e)}")
    
    else:
        st.error("❌ No data could be extracted from the uploaded PDF files.")
        st.info("Please check that your PDFs contain extractable text (not scanned images).")

else:
    # Instructions when no files uploaded
    st.info("👆 Upload one or more PDF files to get started")
    
    with st.expander("ℹ️ How to use"):
        st.markdown("""
        1. **Upload PDF files** using the file uploader above
        2. **Adjust settings** in the sidebar (optional)
        3. **Preview** the extracted data in the table
        4. **Download** your Excel file
        
        **Note:** This tool extracts text-based PDFs. Scanned PDFs (images) may not work without OCR.
        """)
    
    with st.expander("📋 Output Format"):
        st.markdown("""
        The Excel file will contain these columns:
        - **Source_File**: Name of the PDF file
        - **Page**: Page number in the PDF
        - **Total_Pages**: Total pages in that PDF
        - **Line_No**: Line number on that page
        - **Text**: The actual text content
        """)
