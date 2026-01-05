# Improved PDF Customer Data Extractor - Fixed Version
# File: pdf_extractor_fixed.py

import pdfplumber
import pandas as pd
import re
from pathlib import Path
from typing import Dict, List, Tuple
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class ImprovedPDFDataExtractor:
    """Improved extractor with better text extraction and pattern matching"""
    
    def __init__(self):
        self.field_patterns = self._initialize_patterns()
        self.validation_rules = self._initialize_validation_rules()
        
    def _initialize_patterns(self) -> Dict[str, List[str]]:
        """Initialize multiple regex patterns for each field (fallback patterns)"""
        return {
            # Basic Information - Multiple patterns for robustness
            'Type of Customer': [
                r'Type\s+of\s+Customer\s+(.+?)(?:\n)',
                r'Type of Customer\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Name of Customer': [
                r'Name\s+of\s+Customer\s+(.+?)(?:\n)',
                r'Name of Customer\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Company Code': [
                r'Company\s+Code\s+(.+?)(?:\n)',
                r'Company Code\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Customer Group': [
                r'Customer\s+Group\s+(.+?)(?:\n)',
                r'Customer Group\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Sales Group': [
                r'Sales\s+Group\s+(.+?)(?:\n)',
                r'Sales Group\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Region': [
                r'Region\s+(.+?)(?:\n)',
                r'Region\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Zone': [
                r'(?<!\w)Zone\s+(.+?)(?:\n)',
                r'(?<!\w)Zone\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Sub Zone': [
                r'Sub\s+Zone\s+(.+?)(?:\n)',
                r'Sub Zone\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'State': [
                r'(?:^|\n)State\s+(.+?)(?:\n)',
                r'State\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Sales Office': [
                r'Sales\s+Office\s+(.+?)(?:\n)',
                r'Sales Office\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            
            # SAP and Search Terms
            'SAP Dealer code': [
                r'SAP\s+Dealer\s+code.*?(\d{10})',
                r'SAP Dealer code to be mapped.*?(\d{10})',
            ],
            'Search Term 1': [
                r'Search\s+Term\s+1.*?code\s*(.+?)(?:\n)',
            ],
            'Search Term 2': [
                r'Search\s+Term\s+2.*?District\s*(.+?)(?:\n)',
            ],
            
            # Contact Information - Better mobile number extraction
            'Mobile Number': [
                r'Mobile\s+Number\s+(\d{10})',
                r'Mobile Number\s*[:：]?\s*(\d{10})',
                r'(?:Mobile|Phone|Contact).*?(\d{10})',
            ],
            'E-Mail ID': [
                r'E-Mail\s+ID\s+([^\s\n]+@[^\s\n]+)',
                r'E-?Mail\s*[:：]?\s*([^\s\n]+@[^\s\n]+)',
                r'Email\s*[:：]?\s*([^\s\n]+@[^\s\n]+)',
            ],
            'Lattitude': [
                r'Lattitude\s+([\d.]+)',
                r'Latitude\s*[:：]?\s*([\d.]+)',
            ],
            'Longitude': [
                r'Longitude\s+([\d.]+)',
                r'Longitude\s*[:：]?\s*([\d.]+)',
            ],
            
            # Address fields
            'Address 1': [
                r'Address\s+1\s+(.+?)(?:\n)',
                r'Address 1\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Address 2': [
                r'Address\s+2\s+(.+?)(?:\n)',
                r'Address 2\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Address 3': [
                r'Address\s+3\s+(.+?)(?:\n)',
                r'Address 3\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Address 4': [
                r'Address\s+4\s+(.+?)(?:\n)',
                r'Address 4\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'PIN': [
                r'(?:^|\n)PIN\s+(\d{6})',
                r'PIN\s*[:：]?\s*(\d{6})',
                r'Pin.*?(\d{6})',
            ],
            'City': [
                r'(?:^|\n)City\s+(.+?)(?:\n)',
                r'City\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'District': [
                r'(?:^|\n)District\s+(.+?)(?:\n)',
                r'District\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Whatsapp No': [
                r'Whatsapp\s+No\.\s*(\d{10})',
                r'WhatsApp.*?(\d{10})',
            ],
            'Date of Birth': [
                r'Date\s+of\s+Birth\s+(\d{2}[-/]\d{2}[-/]\d{4})',
                r'DOB.*?(\d{2}[-/]\d{2}[-/]\d{4})',
            ],
            'Date of Anniversary': [
                r'Date\s+of\s+Anniversary\s+(\d{2}[-/]\d{2}[-/]\d{4})',
            ],
            
            # Counter Potential
            'Counter Potential Maximum': [
                r'Counter\s+Potential\s*-\s*Maximum\s+(.+?)(?:\n)',
                r'Maximum.*?(\d+(?:\.\d+)?)\s*(?:MT|mt)?',
            ],
            'Counter Potential Minimum': [
                r'Counter\s+Potential\s*-\s*Minimum\s+(.+?)(?:\n)',
                r'Minimum.*?(\d+(?:\.\d+)?)\s*(?:MT|mt)?',
            ],
            
            # GST Details
            'Is GST Present': [
                r'Is\s+GST\s+Present\s+(Yes|No|Y|N)',
            ],
            'GSTIN': [
                r'(?:^|\n)GSTIN\s+([A-Z0-9]{15})',
                r'GSTIN\s*[:：]?\s*([A-Z0-9]{15})',
            ],
            'GST Trade Name': [
                r'Trade\s+Name\s+(.+?)(?:\n)',
            ],
            'GST Legal Name': [
                r'Legal\s+Name\s+(.+?)(?:\n)',
            ],
            'Reg Date': [
                r'Reg\s+Date\s+(.+?)(?:\n)',
                r'Registration Date.*?(\d{2}[/-]\d{2}[/-]\d{4})',
            ],
            
            # PAN Details - More robust
            'PAN': [
                r'(?:^|\n)PAN\s+([A-Z]{5}\d{4}[A-Z])',
                r'PAN\s*[:：]?\s*([A-Z]{5}\d{4}[A-Z])',
                r'(?:PAN|Pan).*?([A-Z]{5}\d{4}[A-Z])',
            ],
            'PAN Holder Name': [
                r'PAN\s+Holder\s+Name\s+(.+?)(?:\n)',
                r'PAN Holder.*?Name.*?(.+?)(?:\n)',
            ],
            'PAN Status': [
                r'PAN\s+Status\s+(VALID|INVALID|Valid|Invalid)',
            ],
            'PAN Aadhaar Linking Status': [
                r'PAN.*?Aadhaar\s+Linking\s+Status\s+([YNR])',
            ],
            
            # Bank Details
            'IFSC Number': [
                r'IFSC\s+Number\s+([A-Z]{4}[0-9A-Z]{7})',
                r'IFSC.*?([A-Z]{4}[0-9A-Z]{7})',
            ],
            'Account Number': [
                r'Account\s+Number\s+(\d+)',
                r'Account.*?Number.*?(\d{9,18})',
            ],
            'Name of Account Holder': [
                r'Name\s+of\s+Account\s+Holder\s+(.+?)(?:\n)',
                r'Account Holder.*?(.+?)(?:\n)',
            ],
            'Bank Name': [
                r'Bank\s+Name\s+(.+?)(?:\n)',
                r'Bank Name\s*[:：]?\s*(.+?)(?:\n|$)',
            ],
            'Bank Branch': [
                r'Bank\s+Branch\s+(.+?)(?:\n)',
                r'Branch.*?(.+?)(?:\n)',
            ],
            
            # Aadhaar Details
            'Is Aadhaar Linked with Mobile': [
                r'Is\s+Aadhaar\s+Linked.*?Mobile.*?(Yes|No|Y|N)',
            ],
            'Aadhaar Number': [
                r'Aadhaar\s+Number\s+((?:XXXX\s*){2,}\d{4}|\d{12})',
                r'Aadhaar.*?((?:XXXX\s*){2,}\d{4})',
            ],
            'Gender': [
                r'Gender\s+(MALE|FEMALE|M|F|Male|Female)',
            ],
            'DOB': [
                r'DOB\s+(.+?)(?:\n)',
                r'Date of Birth.*?(\d{2}[-/]\d{2}[-/]\d{4})',
            ],
            
            # Transportation Zone
            'Logistics Transportation Zone': [
                r'Logistics\s+Transportation\s+Zone\s+(.+?)(?:\n)',
            ],
            'Transportation Zone Description': [
                r'Transportation\s+Zone\s+Description\s+(.+?)(?:\n)',
            ],
            'Transportation Zone Code': [
                r'Transportation\s+Zone\s+Code\s+(\d+)',
            ],
            'Date of Appointment': [
                r'Date\s+of\s+Appointment\s+(\d{2}[-/]\d{2}[-/]\d{4})',
            ],
            
            # Plant Details
            'Delivering Plant': [
                r'Delivering\s+Plant\s+(.+?)(?:\n)',
            ],
            'Plant Name': [
                r'Plant\s+Name\s+(.+?)(?:\n)',
            ],
            'Plant Code': [
                r'Plant\s+Code\s+([A-Z0-9]+)',
            ],
            'Incoterns Code': [
                r'Incoterns\s+Code\s+([A-Z]+)',
                r'Incoterm.*?Code.*?([A-Z]{3})',
            ],
            'Security Deposit Amount': [
                r'Security\s+Deposit.*?(\d+)',
            ],
            'Credit Limit': [
                r'Credit\s+Limit.*?(\d+)',
            ],
            
            # Mapping Details
            'Regional Head': [
                r'Regional\s+Head.*?mapped\s+([E0-9]+)',
            ],
            'Zonal Head': [
                r'Zonal\s+Head.*?mapped\s+([E0-9]+)',
            ],
            'Sub-Zonal Head': [
                r'Sub-Zonal\s+Head.*?mapped\s+([E0-9]+)',
            ],
            'Area Sales Manager': [
                r'Area\s+Sales\s+Manager.*?mapped\s+([E0-9]+)',
            ],
            'Sales Officer': [
                r'Sales\s+Officer.*?mapped\s+([E0-9]+)',
            ],
            'Internal control code': [
                r'Internal\s+control\s+code\s+([A-Z0-9]+)',
            ],
            'SAP CODE': [
                r'SAP\s+CODE\s+([A-Z0-9]+)',
            ],
            
            # Initiator Details
            'Initiator Name': [
                r'Initiator\s+Name\s+(.+?)(?:\n)',
            ],
            'Initiator Email': [
                r'Initiator\s+Email.*?([^\s\n]+@[^\s\n]+)',
            ],
            'Initiator Mobile': [
                r'Initiator\s+Mobile.*?(\d{10})',
            ],
            'Created By UserID': [
                r'Created\s+By.*?UserID\s+([E0-9]+)',
            ],
            
            # Sales Head
            'Sales Head Name': [
                r'Sales\s+Head\s+Name\s+(.+?)(?:\n)',
            ],
            'Sales Head Email': [
                r'Sales\s+Head\s+Email\s+([^\s\n]+@[^\s\n]+)',
            ],
            'Sales Head Mobile': [
                r'Sales\s+Head\s+Mobile.*?(\d{10})',
            ],
            
            # Duplicity Check
            'PAN Result': [
                r'PAN\s+Result\s+([NY])',
            ],
            'Mobile Number Result': [
                r'Mobile\s+Number\s+Result\s+([NY])',
            ],
            'Final Result': [
                r'Final\s+Result\s+(True|False)',
            ],
        }
    
    def _initialize_validation_rules(self) -> Dict:
        """Define validation rules for fields"""
        return {
            'PAN': {
                'pattern': r'^[A-Z]{5}\d{4}[A-Z]$',
                'message': 'Invalid PAN format'
            },
            'Mobile Number': {
                'pattern': r'^\d{10}$',
                'message': 'Invalid mobile number'
            },
            'E-Mail ID': {
                'pattern': r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$',
                'message': 'Invalid email format'
            },
            'PIN': {
                'pattern': r'^\d{6}$',
                'message': 'Invalid PIN code'
            },
            'mandatory_fields': [
                'Name of Customer', 'PAN', 'Mobile Number', 'State', 'District'
            ]
        }
    
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Extract text with better handling"""
        try:
            text = ""
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    try:
                        # Try standard extraction
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n"
                        else:
                            # Try with layout parameter
                            page_text = page.extract_text(layout=True)
                            if page_text:
                                text += page_text + "\n"
                    except Exception as e:
                        logger.warning(f"Error extracting page {page_num}: {str(e)}")
                        continue
            
            # Clean up the text
            text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces with single space
            text = re.sub(r' (\n) ', '\n', text)  # Clean up newlines
            
            return text
        except Exception as e:
            logger.error(f"Error extracting text from {pdf_path}: {str(e)}")
            return ""
    
    def extract_field(self, text: str, field_name: str) -> str:
        """Extract field using multiple patterns (fallback)"""
        patterns = self.field_patterns.get(field_name, [])
        
        for pattern in patterns:
            try:
                match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE | re.DOTALL)
                if match:
                    value = match.group(1).strip()
                    # Clean up the value
                    value = re.sub(r'\s+', ' ', value)
                    value = value.strip()
                    if value and value != "":
                        logger.debug(f"Extracted {field_name}: {value}")
                        return value
            except Exception as e:
                logger.debug(f"Pattern failed for {field_name}: {str(e)}")
                continue
        
        return ""
    
    def extract_all_fields(self, pdf_path: str) -> Dict[str, str]:
        """Extract all fields from PDF"""
        logger.info(f"Processing: {pdf_path}")
        text = self.extract_text_from_pdf(pdf_path)
        
        if not text:
            logger.error(f"No text extracted from {pdf_path}")
            return {'Source File': Path(pdf_path).name, 'Error': 'No text extracted'}
        
        logger.info(f"Extracted {len(text)} characters of text")
        
        data = {'Source File': Path(pdf_path).name}
        
        # Extract all fields
        extracted_count = 0
        for field_name in self.field_patterns.keys():
            value = self.extract_field(text, field_name)
            data[field_name] = value
            if value:
                extracted_count += 1
        
        logger.info(f"Extracted {extracted_count} out of {len(self.field_patterns)} fields")
        
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
    
    def validate_data(self, data: Dict[str, str]) -> Dict:
        """Validate all extracted data"""
        issues = []
        warnings = []
        
        for field_name, value in data.items():
            if field_name in ['Source File', 'Error']:
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
                if data and 'Error' not in data:
                    validation = self.validate_data(data)
                    all_data.append(data)
                    all_validations.append(validation)
                    
                    if not validation['is_valid']:
                        logger.warning(f"{pdf_path}: {validation['issue_count']} issues found")
                else:
                    logger.error(f"Failed to extract data from {pdf_path}")
            except Exception as e:
                logger.error(f"Error processing {pdf_path}: {str(e)}")
        
        if not all_data:
            logger.error("No data extracted from any PDF!")
            return pd.DataFrame(), []
        
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
    extractor = ImprovedPDFDataExtractor()
    
    # Example: Process all PDFs in a folder
    pdf_folder = Path("./pdfs")  # Change this to your PDF folder
    pdf_files = list(pdf_folder.glob("*.pdf"))
    
    if not pdf_files:
        logger.error(f"No PDF files found in {pdf_folder}")
        print(f"\n❌ No PDF files found in {pdf_folder}")
        print(f"Please place your PDF files in the '{pdf_folder}' folder and try again.")
        return
    
    logger.info(f"Found {len(pdf_files)} PDF files to process")
    print(f"\n✓ Found {len(pdf_files)} PDF files")
    
    # Process PDFs
    print("\n🔄 Processing PDFs...")
    df, validations = extractor.process_multiple_pdfs([str(f) for f in pdf_files])
    
    if df.empty:
        print("\n❌ No data could be extracted from the PDFs")
        print("Possible reasons:")
        print("  1. PDFs might be scanned images (need OCR)")
        print("  2. PDFs might be password-protected")
        print("  3. PDF format might be different from expected")
        return
    
    # Export to Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"customer_data_extracted_{timestamp}.xlsx"
    
    print("\n📊 Extraction Summary:")
    print(f"  Total records: {len(df)}")
    print(f"  Valid records: {sum(1 for v in validations if v['is_valid'])}")
    print(f"  Records with issues: {sum(1 for v in validations if not v['is_valid'])}")
    
    if extractor.export_to_excel(df, validations, output_file):
        print(f"\n✅ Success!")
        print(f"  Excel file: {output_file}")
        
        # Also save as CSV
        csv_file = output_file.replace('.xlsx', '.csv')
        df.to_csv(csv_file, index=False)
        print(f"  CSV file: {csv_file}")
        
        # Show sample of extracted data
        print(f"\n📋 Sample Data (first record):")
        if len(df) > 0:
            first_record = df.iloc[0]
            for field in ['Name of Customer', 'PAN', 'Mobile Number', 'State', 'District']:
                print(f"  {field}: {first_record.get(field, 'N/A')}")
    else:
        print("\n❌ Failed to export data")


if __name__ == "__main__":
    main()
