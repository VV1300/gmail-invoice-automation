#!/usr/bin/env python3
"""
Simple RPA Invoice Processing System
"""

import os
import sys
import time
import re
import pandas as pd
import pypdf
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any

# Import existing email downloader
from email_downloader import GmailDownloader
from config import EMAIL_CONFIG

class SimpleRPA:
    """Simple RPA for invoice processing"""
    
    def __init__(self):
        # Create directories
        self.input_dir = Path("input")
        self.output_dir = Path("output")
        self.input_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)
        
        # Initialize Gmail downloader
        self.gmail_downloader = GmailDownloader(EMAIL_CONFIG)
    
    def download_invoices_from_gmail(self):
        """Download invoice PDFs from Gmail using existing downloader"""
        print("Downloading invoices from Gmail...")
        
        try:
            result = self.gmail_downloader.download_invoices(days_back=30)
            
            if result['success']:
                print(f"Downloaded {result['files_downloaded']} files from {result['emails_found']} emails")
                return result['downloaded_files']
            else:
                print(f"Download failed: {result.get('error', 'Unknown error')}")
                return []
                
        except Exception as e:
            print(f"Error downloading invoices: {e}")
            return []
    
    def extract_invoice_data(self, pdf_path: str) -> Dict[str, Any]:
        """Extract invoice data from PDF"""
        try:
            print(f"Processing PDF: {Path(pdf_path).name}")
            
            # Read PDF
            with open(pdf_path, 'rb') as file:
                pdf_reader = pypdf.PdfReader(file)
                
                # Extract text
                text_content = ""
                for page in pdf_reader.pages:
                    text_content += page.extract_text() + "\n"
            
            # Extract invoice data using simple patterns
            invoice_data = {
                'invoice_number': None,
                'vendor_name': None,
                'invoice_date': None,
                'invoice_amount': None,
                'due_date': None,
                'payment_status': None,
                'file_path': Path(pdf_path).name
            }
            
            lines = text_content.split('\n')
            
            # Extract invoice number
            for line in lines:
                if 'Invoice Number:' in line:
                    parts = line.split(':')
                    if len(parts) > 1:
                        invoice_data['invoice_number'] = parts[1].strip()
                        break
            
            # Extract vendor name
            for line in lines:
                line = line.strip()
                if line and len(line) > 3:
                    skip_keywords = ['invoice', 'bill', 'date', 'phone', 'email', 'address', 'description', 'qty', 'unit', 'total', 'please', 'share', 'form', 'within', 'hours', 'recent', 'infusion']
                    if not any(keyword in line.lower() for keyword in skip_keywords):
                        invoice_data['vendor_name'] = line
                        break
            
            # Extract invoice date
            for line in lines:
                if 'Invoice Date:' in line:
                    parts = line.split(':')
                    if len(parts) > 1:
                        invoice_data['invoice_date'] = parts[1].strip()
                        break
            
            # Extract invoice amount
            for line in lines:
                if '$' in line:
                    amounts = re.findall(r'\$([\d,]+\.?\d*)', line)
                    if amounts:
                        try:
                            amount_str = amounts[-1].replace(',', '')
                            invoice_data['invoice_amount'] = float(amount_str)
                            break
                        except ValueError:
                            continue
            
            # Extract due date
            for line in lines:
                if 'Due Date:' in line:
                    parts = line.split(':')
                    if len(parts) > 1:
                        invoice_data['due_date'] = parts[1].strip()
                        break
            
            # Extract payment status
            payment_status_found = False
            for line in lines:
                if 'Payment Status:' in line:
                    parts = line.split(':')
                    if len(parts) > 1:
                        status = parts[1].strip()
                        if 'unpaid' in status.lower():
                            invoice_data['payment_status'] = 'Unpaid'
                        elif 'paid' in status.lower():
                            invoice_data['payment_status'] = 'Paid'
                        else:
                            invoice_data['payment_status'] = status
                        payment_status_found = True
                        break
            
            # If no payment status found, keep it as None (will be blank in Excel)
            if not payment_status_found:
                invoice_data['payment_status'] = None
            
            print(f"Extracted: Invoice={invoice_data['invoice_number']}, Date={invoice_data['invoice_date']}, Amount=${invoice_data['invoice_amount']}, Vendor={invoice_data['vendor_name']}, Due={invoice_data['due_date']}, Status={invoice_data['payment_status']}")
            return invoice_data
            
        except Exception as e:
            print(f"Error extracting data from {pdf_path}: {e}")
            return {
                'invoice_number': None,
                'vendor_name': None,
                'invoice_date': None,
                'invoice_amount': None,
                'due_date': None,
                'payment_status': None,
                'file_path': Path(pdf_path).name
            }
    
    def process_invoices(self):
        """Main processing function"""
        print("=" * 50)
        print("SIMPLE RPA INVOICE PROCESSING")
        print("=" * 50)
        
        # Step 1: Download invoices from Gmail
        print("\nStep 1: Downloading invoices from Gmail...")
        downloaded_files = self.download_invoices_from_gmail()
        
        if not downloaded_files:
            print("No invoices downloaded from Gmail")
            return
        
        print(f"Downloaded {len(downloaded_files)} PDF files")
        
        # Step 2: Process each PDF
        print("\nStep 2: Processing PDF files...")
        invoice_data_list = []
        
        for pdf_file in downloaded_files:
            invoice_data = self.extract_invoice_data(pdf_file)
            # Only include if we extracted some data
            if invoice_data['invoice_number'] or invoice_data['invoice_amount'] or invoice_data['vendor_name']:
                invoice_data_list.append(invoice_data)
        
        if not invoice_data_list:
            print("No invoice data extracted")
            return
        
        # Step 3: Create Excel report
        print("\nStep 3: Creating Excel report...")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_file = self.output_dir / f"invoice_report_{timestamp}.xlsx"
        
        # Create DataFrame
        df = pd.DataFrame(invoice_data_list)
        
        # Write to Excel
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Invoice_Data', index=False)
            
            # Create summary sheet
            summary_data = []
            total_amount = sum(item.get('invoice_amount', 0) for item in invoice_data_list if item.get('invoice_amount'))
            total_invoices = len(invoice_data_list)
            unique_vendors = len(set(item.get('vendor_name', '') for item in invoice_data_list if item.get('vendor_name')))
            
            summary_data.append({'Metric': 'Total Invoices', 'Value': total_invoices})
            summary_data.append({'Metric': 'Total Amount', 'Value': f"${total_amount:,.2f}"})
            summary_data.append({'Metric': 'Unique Vendors', 'Value': unique_vendors})
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        print(f"Excel report created: {excel_file}")
        print(f"Processed {len(invoice_data_list)} invoices")
        print(f"Total amount: ${total_amount:,.2f}")
        
        print("\n" + "=" * 50)
        print("PROCESSING COMPLETED SUCCESSFULLY!")
        print("=" * 50)

def main():
    """Main function"""
    rpa = SimpleRPA()
    rpa.process_invoices()

if __name__ == "__main__":
    main() 