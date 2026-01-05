import streamlit as st
# PDF Customer Data Extractor - Complete Solution
# File: pdf_extractor.py

import pdfplumber
import pandas as pd
import re
from pathlib import Path
from typing import Dict, List, Tuple
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class PDFDataExtractor:
    """Extract customer data from PDF forms with validation"""
    
    def __init__(self):
        self.field_patterns = self._initialize_patterns()
        self.validation_rules = self._initialize_validation_rules()
        
    def _initialize_patterns(self) -> Dict[str, str]:
        """Initialize regex patterns for all fields"""
        return {
            # Basic Information
            'Type of Customer': r'Type of Customer\s+(.+?)(?:\n|$)',
            'Name of Customer': r'Name of Customer\s+(.+?)(?:\n|$)',
            'Company Code': r'Company Code\s+(.+?)(?:\n|$)',
            'Customer Group': r'Customer Group\s+(.+?)(?:\n|$)',
            'Sales Group': r'Sales Group\s+(.+?)(?:\n|$)',
            'Region': r'Region\s+(.+?)(?:\n|$)',
            'Zone': r'Zone\s+(.+?)(?:\n|$)',
            'Sub Zone': r'Sub Zone\s+(.+?)(?:\n|$)',
            'State': r'(?:^|\n)State\s+(.+?)(?:\n|$)',
            'Sales Office': r'Sales Office\s+(.+?)(?:\n|$)',
            
            # SAP and Search Terms
            'SAP Dealer code': r'SAP Dealer code to be mapped Search\s*(?:Term\s*2?)?\s*(\d+)',
            'Search Term 1': r'Search Term 1-\s*Old customer code\s*(.+?)(?:\n|$)',
            'Search Term 2': r'Search Term 2\s*-\s*District\s*(.+?)(?:\n|$)',
            
            # Contact Information
            'Mobile Number': r'Mobile Number\s+(\d{10})',
            'E-Mail ID': r'E-Mail ID\s+([^\n]*?)(?:\n|$)',
            'Lattitude': r'Lattitude\s+([\d.]+)',
            'Longitude': r'Longitude\s+([\d.]+)',
            
            # Customer Address Section
            'Trade Name or Legal Name': r'Name of the Customers\s*\(Trade Name or\s*Legal Name\)\s+(.+?)(?:\n|$)',
            'Address 1': r'Address 1\s+(.+?)(?:\n|$)',
            'Address 2': r'Address 2\s+(.+?)(?:\n|$)',
            'Address 3': r'Address 3\s+(.+?)(?:\n|$)',
            'Address 4': r'Address 4\s+(.+?)(?:\n|$)',
            'PIN': r'(?:^|\n)PIN\s+(\d{6})',
            'City': r'(?:^|\n)City\s+(.+?)(?:\n|$)',
            'District': r'(?:^|\n)District\s+(.+?)(?:\n|$)',
            'Whatsapp No': r'Whatsapp No\.\s*(\d*)',
            'Date of Birth': r'Date of Birth\s+([0-9]{2}-[0-9]{2}-[0-9]{4})',
            'Date of Anniversary': r'Date of Anniversary\s+([0-9]{2}-[0-9]{2}-[0-9]{4})?',
            
            # Counter Potential
            'Counter Potential Maximum': r'Counter Potential\s*-\s*Maximum\s+(.+?)(?:\n|$)',
            'Counter Potential Minimum': r'Counter Potential\s*-\s*Minimum\s+(.+?)(?:\n|$)',
            
            # GST Details
            'Is GST Present': r'Is GST Present\s+(.+?)(?:\n|$)',
            'GSTIN': r'(?:^|\n)GSTIN\s+([A-Z0-9]{15})?',
            'GST Trade Name': r'Trade Name\s+(.+?)(?:\n|Legal Name)',
            'GST Legal Name': r'Legal Name\s+(.+?)(?:\n|Reg Date)',
            'Reg Date': r'Reg Date\s+(.+?)(?:\n|$)',
            'GST City': r'(?:City\s+.+?\n){1}City\s+(.+?)(?:\n|$)',
            'GST Type': r'Type\s+(.+?)(?:\n|Building)',
            'Building No': r'Building No\.\s+(.+?)(?:\n|$)',
            'District Code': r'District Code\s+(.+?)(?:\n|$)',
            'State Code': r'State Code\s+(.+?)(?:\n|$)',
            'Street': r'Street\s+(.+?)(?:\n|$)',
            'PIN Code': r'PIN Code\s+(\d{6})?',
            
            # PAN Details
            'PAN': r'(?:^|\n)PAN\s+([A-Z]{5}\d{4}[A-Z])',
            'PAN Holder Name': r'PAN Holder Name\s+(.+?)(?:\n|$)',
            'PAN Status': r'PAN Status\s+(.+?)(?:\n|$)',
            'PAN Aadhaar Linking Status': r'PAN\s*-\s*Aadhaar Linking Status\s+(.+?)(?:\n|$)',
            
            # Bank Details
            'IFSC Number': r'IFSC Number\s+([A-Z]{4}[0-9A-Z]{7})',
            'Account Number': r'Account Number\s+(\d+)',
            'Name of Account Holder': r'Name of Account Holder\s+(.+?)(?:\n|$)',
            'Bank Name': r'Bank Name\s+(.+?)(?:\n|$)',
            'Bank Branch': r'Bank Branch\s+(.+?)(?:\n|$)',
            
            # Aadhaar Details
            'Is Aadhaar Linked with Mobile': r'Is Aadhaar Linked with Mobile\?\s+(.+?)(?:\n|$)',
            'Aadhaar Number': r'Aadhaar Number\s+((?:XXXX\s*){2,3}\d{4}|\d{12})',
            'Aadhaar Name': r'(?:Aadhaar Details.*?Name\s+)(.+?)(?:\n|$)',
            'Gender': r'Gender\s+([A-Z]+)',
            'DOB': r'DOB\s+(.+?)(?:\n|$)',
            'Aadhaar Address': r'(?:Aadhaar Details.*?Address\s+)(.+?)(?:\nPIN)',
            'Aadhaar PIN': r'(?:Aadhaar Details.*?PIN\s+)(\d{6})',
            'Aadhaar City': r'(?:Aadhaar Details.*?City\s+)(.+?)(?:\n|$)',
            'Aadhaar State': r'(?:Aadhaar Details.*?State\s+)(.+?)(?:\n|$)',
            
            # Transportation Zone Details
            'Logistics Transportation Zone': r'Logistics Transportation Zone\s+(.+?)(?:\n|$)',
            'Transportation Zone Description': r'Transportation Zone\s*Description\s+(.+?)(?:\n|$)',
            'Transportation Zone Code': r'Transportation Zone Code\s+(\d+)',
            'Postal Code': r'Postal Code\s+(.+?)(?:\n|$)',
            'Logistics team to vet': r'Logistics team to vet the T zone selected by\s*Sales Officer\s+(.+?)(?:\n|$)',
            'Selection of Available T Zones': r'Selection of Available T Zones from T Zone\s*Master list, if found\.\s+(.+?)(?:\n|$)',
            'NEW T Zone': r'If NEW T Zone need to be created, details\s*to be provided by Logistics team\s+(.+?)(?:\n|$)',
            'Date of Appointment': r'Date of Appointment\s+([0-9]{2}-[0-9]{2}-[0-9]{4})?',
            
            # Plant Details
            'Delivering Plant': r'Delivering Plant\s+(.+?)(?:\n|$)',
            'Plant Name': r'Plant Name\s+(.+?)(?:\n|$)',
            'Plant Code': r'Plant Code\s+([A-Z0-9]+)',
            'Incoterns': r'Incoterns\s+([A-Z]+\s*-[^-\n]+?)(?:\n)',
            'Incoterns Description': r'Incoterns\s+[A-Z]+\s*-\s*(.+?)(?:\n|Incoterns Code)',
            'Incoterns Code': r'Incoterns Code\s+([A-Z]+)',
            'Security Deposit Amount': r'Security Deposit Amount.*?(\d+)',
            'Credit Limit': r'Credit Limit\s*\(In Rs\.\)\s*(.+?)(?:\n|$)',
            
            # Mapping Details
            'Regional Head': r'Regional Head to be mapped\s+([E0-9]+)',
            'Zonal Head': r'Zonal Head to be mapped\s+([E0-9]+)',
            'Sub-Zonal Head': r'Sub-Zonal Head \(RSM\) to be mapped\s+([E0-9]+)',
            'Area Sales Manager': r'Area Sales Manager to be mapped\s+([E0-9]+)',
            'Sales Officer': r'Sales Officer to be mapped\s+([E0-9]+)',
            'Sales Promoter': r'Sales Promoter to be mapped\s+(.+?)(?:\n|$)',
            'Sales Promoter Number': r'Sales Promoter Number\s+(.+?)(?:\n|$)',
            'Internal control code': r'Internal control code\s+([A-Z0-9]+)',
            'SAP CODE': r'(?:^|\n)SAP CODE\s+(.+?)(?:\n|$)',
            
            # Initiator Details
            'Initiator Name': r'Initiator Name\s+(.+?)(?:\n|$)',
            'Initiator Email': r'Initiator Email ID\s+(.+?)(?:\n|$)',
            'Initiator Mobile': r'Initiator Mobile Number\s+(\d+)',
            'Created By UserID': r'Created By Customer UserID\s+([E0-9]+)',
            
            # Sales Head Details
            'Sales Head Name': r'Sales Head Name\s+(.+?)(?:\n|$)',
            'Sales Head Email': r'Sales Head Email\s+(.+?)(?:\n|$)',
            'Sales Head Mobile': r'Sales Head Mobile Number\s+(\d+)',
            'Extra2': r'Extra2\s+(.+?)(?:\n|$)',
            
            # Duplicity Check
            'PAN Result': r'PAN Result\s+([NY])',
            'Mobile Number Result': r'Mobile Number Result\s+([NY])',
            'Email Result': r'Email Result\s+([NY])?',
            'GST Result': r'GST Result\s+([NY])?',
            'Final Result': r'Final Result\s+(.+?)(?:\n|$)',
        }
    
    def _initialize_validation_rules(self) -> Dict:
        """Define validation rules for fields"""
        return {
            'PAN': {
                'pattern': r'^[A-Z]{5}\d{4}[A-Z]$',
                'message': 'Invalid PAN format (should be ABCDE1234F)'
            },
            'Mobile Number': {
                'pattern': r'^\d{10}$',
                'message': 'Invalid mobile number (should be 10 digits)'
            },
            'E-Mail ID': {
                'pattern': r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$',
                'message': 'Invalid email format'
            },
            'PIN': {
                'pattern': r'^\d{6}$',
                'message': 'Invalid PIN code (should be 6 digits)'
            },
            'GSTIN': {
                'pattern': r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$',
                'message': 'Invalid GSTIN format'
            },
            'IFSC Number': {
                'pattern': r'^[A-Z]{4}0[A-Z0-9]{6}$',
                'message': 'Invalid IFSC code format'
            },
            'mandatory_fields': [
                'Name of Customer', 'PAN', 'Mobile Number', 'State', 
                'District', 'Bank Name', 'Account Number'
            ]
        }
    
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Extract all text from PDF file"""
        try:
            text = ""
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text += page.extract_text() + "\n"
            return text
        except Exception as e:
            logger.error(f"Error extracting text from {pdf_path}: {str(e)}")
            return ""
    
    def extract_field(self, text: str, field_name: str) -> str:
        """Extract a single field value using regex pattern"""
        pattern = self.field_patterns.get(field_name)
        if not pattern:
            return ""
        
        try:
            match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
            if match:
                value = match.group(1).strip()
                # Clean up common artifacts
                value = re.sub(r'\s+', ' ', value)  # Replace multiple spaces
                return value if value else ""
            return ""
        except Exception as e:
            logger.warning(f"Error extracting {field_name}: {str(e)}")
            return ""
    
    def extract_all_fields(self, pdf_path: str) -> Dict[str, str]:
        """Extract all fields from PDF"""
        logger.info(f"Processing: {pdf_path}")
        text = self.extract_text_from_pdf(pdf_path)
        
        if not text:
            logger.error(f"No text extracted from {pdf_path}")
            return {}
        
        data = {'Source File': Path(pdf_path).name}
        
        # Extract all fields
        for field_name in self.field_patterns.keys():
            data[field_name] = self.extract_field(text, field_name)
        
        return data
    
    def validate_field(self, field_name: str, value: str) -> Tuple[bool, str]:
        """Validate a single field"""
        if not value or value == "":
            if field_name in self.validation_rules.get('mandatory_fields', []):
                return False, f"Missing mandatory field: {field_name}"
            return True, ""
        
        if field_name in self.validation_rules and 'pattern' in self.validation_rules[field_name]:
            rule = self.validation_rules[field_name]
            if not re.match(rule['pattern'], value):
                return False, rule['message']
        
        return True, ""
    
    def validate_data(self, data: Dict[str, str]) -> Dict[str, any]:
        """Validate all extracted data"""
        issues = []
        warnings = []
        
        for field_name, value in data.items():
            if field_name == 'Source File':
                continue
                
            is_valid, message = self.validate_field(field_name, value)
            if not is_valid:
                if field_name in self.validation_rules.get('mandatory_fields', []):
                    issues.append(message)
                else:
                    warnings.append(message)
        
        return {
            'is_valid': len(issues) == 0,
            'issues': issues,
            'warnings': warnings,
            'issue_count': len(issues),
            'warning_count': len(warnings)
        }
    
    def process_multiple_pdfs(self, pdf_paths: List[str]) -> Tuple[pd.DataFrame, List[Dict]]:
        """Process multiple PDF files"""
        all_data = []
        all_validations = []
        
        for pdf_path in pdf_paths:
            try:
                data = self.extract_all_fields(pdf_path)
                if data:
                    validation = self.validate_data(data)
                    all_data.append(data)
                    all_validations.append(validation)
                    
                    if not validation['is_valid']:
                        logger.warning(f"{pdf_path}: {validation['issue_count']} issues found")
            except Exception as e:
                logger.error(f"Error processing {pdf_path}: {str(e)}")
        
        df = pd.DataFrame(all_data)
        return df, all_validations
    
    def export_to_excel(self, df: pd.DataFrame, validations: List[Dict], output_path: str):
        """Export data to Excel with formatting"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Write main data
                df.to_excel(writer, sheet_name='Customer Data', index=False)
                
                # Create validation summary
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
                
                # Format columns
                workbook = writer.book
                worksheet = writer.sheets['Customer Data']
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
            
            logger.info(f"Data exported successfully to {output_path}")
            return True
        except Exception as e:
            logger.error(f"Error exporting to Excel: {str(e)}")
            return False


def main():
    """Main execution function"""
    # Initialize extractor
    extractor = PDFDataExtractor()
    
    # Example: Process all PDFs in a folder
    pdf_folder = Path("./pdfs")  # Change this to your PDF folder
    pdf_files = list(pdf_folder.glob("*.pdf"))
    
    if not pdf_files:
        logger.error(f"No PDF files found in {pdf_folder}")
        return
    
    logger.info(f"Found {len(pdf_files)} PDF files to process")
    
    # Process PDFs
    df, validations = extractor.process_multiple_pdfs([str(f) for f in pdf_files])
    
    # Export to Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"customer_data_extracted_{timestamp}.xlsx"
    
    if extractor.export_to_excel(df, validations, output_file):
        logger.info(f"✓ Processing complete!")
        logger.info(f"✓ Total records: {len(df)}")
        logger.info(f"✓ Valid records: {sum(1 for v in validations if v['is_valid'])}")
        logger.info(f"✓ Output file: {output_file}")
    
    # Also save as CSV
    csv_file = output_file.replace('.xlsx', '.csv')
    df.to_csv(csv_file, index=False)
    logger.info(f"✓ CSV file: {csv_file}")


if __name__ == "__main__":
    main()
