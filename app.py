# Streamlit PDF Customer Data Extractor
# File: app.py

import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import os
from datetime import datetime
from pdf_extractor import PDFDataExtractor
import zipfile
import io

# Page configuration
st.set_page_config(
    page_title="PDF Customer Data Extractor",
    page_icon="📄",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        font-weight: bold;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #555;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        border-radius: 4px;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
        border-radius: 4px;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1rem;
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
        border-radius: 4px;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        background-color: #d1ecf1;
        border-left: 4px solid #17a2b8;
        border-radius: 4px;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = None
if 'validations' not in st.session_state:
    st.session_state.validations = None
if 'processed' not in st.session_state:
    st.session_state.processed = False


def main():
    # Header
    st.markdown('<div class="main-header">📄 PDF Customer Data Extractor</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Extract customer creation data from PDF forms with validation</div>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.image("https://via.placeholder.com/200x100/1f77b4/ffffff?text=Star+Cement", use_container_width=True)
        st.markdown("### ⚙️ Settings")
        
        export_format = st.radio(
            "Export Format",
            ["Excel (.xlsx)", "CSV (.csv)", "Both"],
            index=0
        )
        
        show_warnings = st.checkbox("Show Validation Warnings", value=True)
        auto_width = st.checkbox("Auto-adjust Excel columns", value=True)
        
        st.markdown("---")
        st.markdown("### 📊 Statistics")
        if st.session_state.processed and st.session_state.extracted_data is not None:
            df = st.session_state.extracted_data
            validations = st.session_state.validations
            
            st.metric("Total PDFs Processed", len(df))
            st.metric("Valid Records", sum(1 for v in validations if v['is_valid']))
            st.metric("Records with Issues", sum(1 for v in validations if not v['is_valid']))
        
        st.markdown("---")
        st.markdown("### 💡 API Pricing Info")
        st.info("""
        **Claude API (Option 2)**
        - Per PDF: $0.02-0.05
        - 100 PDFs: ~$2-5
        - 1000 PDFs: ~$20-50
        - Accuracy: 98-99%
        """)
    
    # Main content
    tab1, tab2, tab3 = st.tabs(["📤 Upload & Process", "📊 Results & Validation", "📥 Export Data"])
    
    with tab1:
        st.markdown("### Upload PDF Files")
        st.markdown("Upload customer creation PDF forms (ZSUB format)")
        
        uploaded_files = st.file_uploader(
            "Choose PDF files",
            type=['pdf'],
            accept_multiple_files=True,
            help="Select one or more PDF files to process"
        )
        
        if uploaded_files:
            st.success(f"✓ {len(uploaded_files)} file(s) uploaded")
            
            # Show file list
            with st.expander("📋 View Uploaded Files"):
                for i, file in enumerate(uploaded_files, 1):
                    st.text(f"{i}. {file.name} ({file.size / 1024:.2f} KB)")
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button("🚀 Process PDFs", type="primary", use_container_width=True):
                    process_pdfs(uploaded_files, show_warnings)
    
    with tab2:
        if st.session_state.processed:
            show_results()
        else:
            st.info("👈 Upload and process PDF files to see results here")
    
    with tab3:
        if st.session_state.processed:
            show_export_options(export_format, auto_width)
        else:
            st.info("👈 Upload and process PDF files to export data")


def process_pdfs(uploaded_files, show_warnings):
    """Process uploaded PDF files"""
    with st.spinner("Processing PDFs... This may take a moment."):
        try:
            # Create temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                pdf_paths = []
                
                # Save uploaded files
                for uploaded_file in uploaded_files:
                    file_path = temp_path / uploaded_file.name
                    with open(file_path, 'wb') as f:
                        f.write(uploaded_file.getbuffer())
                    pdf_paths.append(str(file_path))
                
                # Initialize extractor and process
                extractor = PDFDataExtractor()
                df, validations = extractor.process_multiple_pdfs(pdf_paths)
                
                # Store in session state
                st.session_state.extracted_data = df
                st.session_state.validations = validations
                st.session_state.processed = True
                
                st.success("✅ Processing completed successfully!")
                st.balloons()
        
        except Exception as e:
            st.error(f"❌ Error processing PDFs: {str(e)}")
            st.exception(e)


def show_results():
    """Display extraction results and validation"""
    df = st.session_state.extracted_data
    validations = st.session_state.validations
    
    st.markdown("### 📊 Extraction Results")
    
    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Records", len(df))
    with col2:
        valid_count = sum(1 for v in validations if v['is_valid'])
        st.metric("Valid Records", valid_count, delta=f"{valid_count/len(df)*100:.1f}%")
    with col3:
        issue_count = sum(v['issue_count'] for v in validations)
        st.metric("Total Issues", issue_count)
    with col4:
        warning_count = sum(v['warning_count'] for v in validations)
        st.metric("Total Warnings", warning_count)
    
    # Validation Summary
    st.markdown("### ✅ Validation Summary")
    
    validation_data = []
    for i, v in enumerate(validations):
        status = "✅ Valid" if v['is_valid'] else "❌ Invalid"
        validation_data.append({
            'File': df.iloc[i]['Source File'],
            'Status': status,
            'Issues': v['issue_count'],
            'Warnings': v['warning_count']
        })
    
    validation_df = pd.DataFrame(validation_data)
    st.dataframe(validation_df, use_container_width=True)
    
    # Detailed records
    st.markdown("### 📋 Detailed Records")
    
    # Filter options
    col1, col2 = st.columns(2)
    with col1:
        filter_status = st.selectbox(
            "Filter by Status",
            ["All", "Valid Only", "Invalid Only"]
        )
    with col2:
        search_term = st.text_input("Search in Customer Name or PAN", "")
    
    # Apply filters
    filtered_indices = range(len(df))
    if filter_status == "Valid Only":
        filtered_indices = [i for i, v in enumerate(validations) if v['is_valid']]
    elif filter_status == "Invalid Only":
        filtered_indices = [i for i, v in enumerate(validations) if not v['is_valid']]
    
    # Show records
    for idx in filtered_indices:
        record = df.iloc[idx]
        validation = validations[idx]
        
        # Apply search filter
        if search_term:
            if search_term.lower() not in str(record.get('Name of Customer', '')).lower() and \
               search_term.lower() not in str(record.get('PAN', '')).lower():
                continue
        
        with st.expander(f"📄 {record['Source File']} - {record.get('Name of Customer', 'N/A')}"):
            # Validation status
            if validation['is_valid']:
                st.markdown('<div class="success-box">✅ <b>Valid Record</b></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="error-box">❌ <b>Invalid Record</b> - {validation["issue_count"]} issue(s) found</div>', unsafe_allow_html=True)
                for issue in validation['issues']:
                    st.markdown(f"- {issue}")
            
            if validation['warnings']:
                st.markdown(f'<div class="warning-box">⚠️ <b>Warnings:</b> {validation["warning_count"]} warning(s)</div>', unsafe_allow_html=True)
                for warning in validation['warnings']:
                    st.markdown(f"- {warning}")
            
            # Key information
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("**Customer Details**")
                st.text(f"Name: {record.get('Name of Customer', 'N/A')}")
                st.text(f"PAN: {record.get('PAN', 'N/A')}")
                st.text(f"Mobile: {record.get('Mobile Number', 'N/A')}")
            with col2:
                st.markdown("**Location**")
                st.text(f"State: {record.get('State', 'N/A')}")
                st.text(f"District: {record.get('District', 'N/A')}")
                st.text(f"City: {record.get('City', 'N/A')}")
            with col3:
                st.markdown("**Business**")
                st.text(f"Sales Group: {record.get('Sales Group', 'N/A')}")
                st.text(f"Region: {record.get('Region', 'N/A')}")
                st.text(f"Zone: {record.get('Zone', 'N/A')}")
            
            # Show all fields
            if st.checkbox(f"Show all fields - {record['Source File']}", key=f"show_all_{idx}"):
                st.dataframe(record.to_frame(), use_container_width=True)


def show_export_options(export_format, auto_width):
    """Show export options and download buttons"""
    df = st.session_state.extracted_data
    validations = st.session_state.validations
    
    st.markdown("### 📥 Export Data")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Export to Excel
    if export_format in ["Excel (.xlsx)", "Both"]:
        st.markdown("#### 📊 Export to Excel")
        
        try:
            extractor = PDFDataExtractor()
            excel_buffer = io.BytesIO()
            
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Customer Data', index=False)
                
                # Validation summary
                validation_df = pd.DataFrame([
                    {
                        'File': df.iloc[i]['Source File'],
                        'Status': 'Valid' if v['is_valid'] else 'Invalid',
                        'Issues': v['issue_count'],
                        'Warnings': v['warning_count'],
                        'Details': '; '.join(v['issues'] + v['warnings'])
                    }
                    for i, v in enumerate(validations)
                ])
                validation_df.to_excel(writer, sheet_name='Validation Summary', index=False)
            
            excel_buffer.seek(0)
            
            st.download_button(
                label="📥 Download Excel File",
                data=excel_buffer,
                file_name=f"customer_data_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Error creating Excel file: {str(e)}")
    
    # Export to CSV
    if export_format in ["CSV (.csv)", "Both"]:
        st.markdown("#### 📄 Export to CSV")
        
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        
        st.download_button(
            label="📥 Download CSV File",
            data=csv_buffer.getvalue(),
            file_name=f"customer_data_{timestamp}.csv",
            mime="text/csv",
            type="secondary",
            use_container_width=True
        )
    
    # Export both as ZIP
    if export_format == "Both":
        st.markdown("#### 📦 Export Both as ZIP")
        
        try:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                # Add Excel
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Customer Data', index=False)
                excel_buffer.seek(0)
                zip_file.writestr(f"customer_data_{timestamp}.xlsx", excel_buffer.getvalue())
                
                # Add CSV
                csv_data = df.to_csv(index=False)
                zip_file.writestr(f"customer_data_{timestamp}.csv", csv_data)
            
            zip_buffer.seek(0)
            
            st.download_button(
                label="📥 Download ZIP (Excel + CSV)",
                data=zip_buffer,
                file_name=f"customer_data_{timestamp}.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Error creating ZIP file: {str(e)}")


if __name__ == "__main__":
    main()
