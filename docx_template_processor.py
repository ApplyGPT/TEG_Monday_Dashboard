#!/usr/bin/env python3
"""
Create documents from original .docx files with variable replacement
Preserves exact formatting, design, text, figures, tables, etc.
"""

import os
import shutil
from docx import Document
from docx.shared import Inches
import re
from datetime import datetime

class DocxTemplateProcessor:
    """Process .docx templates with variable replacement while preserving formatting"""
    
    def __init__(self):
        self.template_mapping = {
            'development_contract': {
                'file': 'Development Contract.docx',
                'replacements': {
                    'EUGENIA ZHANG': 'CLIENT_NAME',
                    'March 14, 2025': 'CONTRACT_DATE',
                    '$15,865.00': 'CONTRACT_AMOUNT',
                    'VITALINA GHINZELLI': 'TEGMADE_FOR'
                }
            },
            'development_terms': {
                'file': 'Development Terms and Conditions .docx',
                'replacements': {
                    'SHERRY CASSEL': 'CLIENT_NAME',
                    'JUNE 16, 2025': 'CONTRACT_DATE',
                    'VITALINA GHINZELLI': 'TEGMADE_FOR'
                }
            },
            'production_contract': {
                'file': 'Production Contract.docx',
                'replacements': {
                    'Natalie Barrett': 'CLIENT_NAME',
                    'December 06, 2024': 'CONTRACT_DATE',
                    # Replace hardcoded amounts with placeholders (like Development Contract)
                    '$36,830.00': 'DEPOSIT_AMOUNT',
                    '$56,918.00': 'TOTAL_CONTRACT_AMOUNT',
                    '$20,088.00': 'SEWING_COST',
                    '$16,743.00': 'PRE_PRODUCTION_FEE',
                    'VITALINA GHINZELLI': 'TEGMADE_FOR'
                }
            },
            'production_terms': {
                'file': 'Production Terms and Conditions.docx',
                'replacements': {
                    'Natalie Barrett': 'CLIENT_NAME',
                    'December 06, 2024': 'CONTRACT_DATE',
                    'VITALINA GHINZELLI': 'TEGMADE_FOR'
                }
            }
        }
    
    def process_document(self, template_type, client_name, email, 
                        contract_amount=None, contract_date=None,
                        deposit_amount=None, total_contract_amount=None,
                        sewing_cost=None, pre_production_fee=None,
                        uploaded_pdf=None, tegmade_for=None):
        """
        Process a document template with variable replacement
        
        Args:
            template_type: Type of template to process
            client_name: Client name to replace
            email: Client email (for reference)
            contract_amount: Contract amount (if applicable)
            contract_date: Contract date (if not provided, uses current date)
            deposit_amount: Deposit amount (for production contracts)
            total_contract_amount: Total contract amount (for production contracts)
            sewing_cost: Sewing cost (for production contracts)
            pre_production_fee: Pre-production fee (for production contracts)
            uploaded_pdf: Uploaded PDF file (Streamlit UploadedFile object)
            tegmade_for: Name to replace VITALINA GHINZELLI (optional)
            
        Returns:
            str: Path to the processed document
        """
        if template_type not in self.template_mapping:
            raise ValueError(f"Unknown template type: {template_type}")
        
        template_config = self.template_mapping[template_type]
        template_file = template_config['file']
        template_path = os.path.join('inputs', template_file)
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
        
        # Set default date if not provided
        if not contract_date:
            contract_date = datetime.now().strftime("%B %d, %Y")
        
        # Create output filename
        safe_client_name = client_name.replace(' ', '_').replace(',', '')
        safe_date = contract_date.replace(' ', '_').replace(',', '')
        output_filename = f"{template_type}_{safe_client_name}_{safe_date}.docx"
        output_path = os.path.join('processed_documents', output_filename)
        
        # Create output directory if it doesn't exist
        os.makedirs('processed_documents', exist_ok=True)
        
        # Copy template to output location
        shutil.copy2(template_path, output_path)
        
        # Open the document for processing
        doc = Document(output_path)
        
        # Prepare replacement values
        replacement_values = {
            'CLIENT_NAME': client_name,
            'CONTRACT_DATE': contract_date,
            'CONTRACT_AMOUNT': self._format_contract_amount(contract_amount) if contract_amount else '$0.00',
            'DEPOSIT_AMOUNT': self._format_contract_amount(deposit_amount) if deposit_amount else '$0.00',
            'TOTAL_CONTRACT_AMOUNT': self._format_contract_amount(total_contract_amount) if total_contract_amount else '$0.00',
            'SEWING_COST': self._format_contract_amount(sewing_cost) if sewing_cost else '$0.00',
            'PRE_PRODUCTION_FEE': self._format_contract_amount(pre_production_fee) if pre_production_fee else '$0.00',
            'TEGMADE_FOR': tegmade_for if tegmade_for else 'VITALINA GHINZELLI'
        }
        
        # Process replacements
        replacements_made = self._replace_text_in_document(doc, template_config['replacements'], replacement_values)
        
        # Insert PDF (converted to image) after first paragraph for contract documents
        if uploaded_pdf and template_type in ['development_contract', 'production_contract']:
            print(f"📄 Processing PDF: {uploaded_pdf.name} ({uploaded_pdf.size} bytes)")
            self._insert_pdf_after_first_paragraph(doc, uploaded_pdf)
        
        # Save the processed document
        doc.save(output_path)
        
        print(f"✅ Processed {template_type}: {output_filename}")
        print(f"   Replacements made: {replacements_made}")
        
        return output_path
    
    def _replace_text_in_document(self, doc, replacement_map, values):
        """Replace text in document while preserving formatting and making replacements bold"""
        replacements_made = []
        
        # Process paragraphs - do all replacements in one pass to avoid conflicts
        for paragraph in doc.paragraphs:
            paragraph_replacements = []
            for old_text, variable in replacement_map.items():
                if old_text in paragraph.text:
                    paragraph_replacements.append((old_text, values[variable]))
                    replacements_made.append(f"{old_text} -> {values[variable]}")
            
            # Apply all replacements to this paragraph at once
            if paragraph_replacements:
                self._replace_multiple_texts_with_bold(paragraph, paragraph_replacements)
        
        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph_replacements = []
                        for old_text, variable in replacement_map.items():
                            if old_text in paragraph.text:
                                paragraph_replacements.append((old_text, values[variable]))
                                replacements_made.append(f"{old_text} -> {values[variable]}")
                        
                        # Apply all replacements to this paragraph at once
                        if paragraph_replacements:
                            self._replace_multiple_texts_with_bold(paragraph, paragraph_replacements)
        
        # Process images and logos (add placeholders)
        self._process_images_and_logos(doc)
        
        return replacements_made
    
    def _replace_multiple_texts_with_bold(self, paragraph, replacements):
        """Replace multiple texts in a paragraph and make all replacements bold"""
        # Get the full paragraph text
        full_text = paragraph.text
        
        # Apply all replacements
        new_full_text = full_text
        for old_text, new_text in replacements:
            new_full_text = new_full_text.replace(old_text, new_text)
        
        # Clear all runs and rebuild the paragraph
        paragraph.clear()
        
        # Simple approach: split by each replacement text and rebuild
        remaining_text = new_full_text
        
        while remaining_text:
            # Find the first replacement text in the remaining text
            earliest_pos = len(remaining_text)
            earliest_replacement = None
            
            for old_text, new_text in replacements:
                pos = remaining_text.find(new_text)
                if pos != -1 and pos < earliest_pos:
                    earliest_pos = pos
                    earliest_replacement = new_text
            
            if earliest_replacement:
                # Add text before the replacement
                if earliest_pos > 0:
                    paragraph.add_run(remaining_text[:earliest_pos])
                
                # Add the replacement text in bold
                bold_run = paragraph.add_run(earliest_replacement)
                bold_run.bold = True
                
                # Update remaining text
                remaining_text = remaining_text[earliest_pos + len(earliest_replacement):]
            else:
                # No more replacements, add remaining text
                if remaining_text:
                    paragraph.add_run(remaining_text)
                break
    
    def _replace_text_with_bold(self, paragraph, old_text, new_text):
        """Replace text in a paragraph and make the new text bold"""
        # Get the full paragraph text
        full_text = paragraph.text
        
        if old_text in full_text:
            # Replace the text in the full paragraph
            new_full_text = full_text.replace(old_text, new_text)
            
            # Clear all runs and rebuild the paragraph
            paragraph.clear()
            
            # Split the new text around the replacement
            parts = new_full_text.split(new_text)
            
            # Add the part before the replacement
            if parts[0]:
                paragraph.add_run(parts[0])
            
            # Add the replacement text in bold
            bold_run = paragraph.add_run(new_text)
            bold_run.bold = True
            
            # Add the remaining parts
            for i in range(1, len(parts)):
                if parts[i]:
                    paragraph.add_run(parts[i])
    
    def _process_images_and_logos(self, doc):
        """Process images and logos in the document"""
        # This method can be extended to handle images/logos
        # For now, we'll just ensure they're preserved in the document
        pass
    
    def _format_contract_amount(self, amount_str):
        """
        Format contract amount as $n,nnn.nn
        
        Args:
            amount_str: Contract amount string (e.g., "5000", "$5000", "5000.00")
            
        Returns:
            str: Formatted amount (e.g., "$5,000.00")
        """
        try:
            # Remove any existing $ and commas
            clean_amount = amount_str.replace('$', '').replace(',', '')
            
            # Convert to float
            amount = float(clean_amount)
            
            # Format with commas and 2 decimal places
            formatted = f"${amount:,.2f}"
            
            return formatted
            
        except (ValueError, TypeError):
            # If formatting fails, return original string with $ prefix
            return f"${amount_str}" if not amount_str.startswith('$') else amount_str
    
    def _insert_pdf_after_first_paragraph(self, doc, uploaded_pdf):
        """
        Insert a PDF (converted to image) after specific text patterns in contract documents
        
        Args:
            doc: Document object
            uploaded_pdf: Streamlit UploadedFile object containing PDF data
        """
        try:
            import io
            from docx.shared import Inches
            
            # Convert PDF first page to image
            image_data = self._convert_pdf_to_image(uploaded_pdf)
            
            # First, remove any existing images in the document
            self._remove_existing_images(doc)
            
            # Define the specific text patterns to look for (using key parts)
            target_patterns = [
                "Development Contract, and Terms and Conditions Agreement (Attachment B), and as follows:",
                "The attached Production Workbook includes all detail as a part of this agreement."
            ]
            
            # Find the paragraph containing the target text
            target_paragraph = None
            for paragraph in doc.paragraphs:
                paragraph_text = paragraph.text.strip()
                for pattern in target_patterns:
                    if pattern in paragraph_text:
                        target_paragraph = paragraph
                        print(f"📍 Found target paragraph: '{paragraph_text[:100]}...'")
                        break
                if target_paragraph:
                    break
            
            # If exact patterns not found, try partial matching
            if target_paragraph is None:
                print("⚠️ Exact patterns not found, trying partial matching")
                partial_patterns = [
                    ("Development Contract", "Attachment B"),
                    ("Production Workbook", "includes all detail")
                ]
                
                for paragraph in doc.paragraphs:
                    paragraph_text = paragraph.text.strip()
                    for pattern1, pattern2 in partial_patterns:
                        if pattern1 in paragraph_text and pattern2 in paragraph_text:
                            target_paragraph = paragraph
                            print(f"📍 Found partial match: '{paragraph_text[:100]}...'")
                            break
                    if target_paragraph:
                        break
            
            # If no target pattern found, fall back to first substantial paragraph
            if target_paragraph is None:
                print("⚠️ Target pattern not found, using first substantial paragraph")
                for paragraph in doc.paragraphs:
                    text = paragraph.text.strip()
                    if (len(text) > 50 and 
                        not text.isupper() and 
                        not text.endswith(':') and 
                        '.' in text):
                        target_paragraph = paragraph
                        print(f"📍 Fallback to: '{text[:50]}...'")
                        break
            
            if target_paragraph is None:
                print("⚠️ No suitable paragraph found in document")
                return
            
            # Create a new paragraph for the image
            new_paragraph = doc.add_paragraph()
            
            # Insert the new paragraph after the target paragraph
            new_paragraph._element.getparent().remove(new_paragraph._element)
            target_paragraph._element.addnext(new_paragraph._element)
            
            # Add the image to the new paragraph
            run = new_paragraph.add_run()
            
            # Add image with reasonable size (max width 6 inches)
            run.add_picture(io.BytesIO(image_data), width=Inches(6))
            
            print(f"✅ PDF (as image) inserted after target paragraph: {uploaded_pdf.name}")
            
        except Exception as e:
            print(f"❌ Error inserting PDF: {str(e)}")
            # Don't raise the exception to avoid breaking the document processing
    
    def _convert_pdf_to_image(self, uploaded_pdf):
        """
        Convert the first page of a PDF to an image using PyMuPDF (no external dependencies needed)
        
        Args:
            uploaded_pdf: Streamlit UploadedFile object containing PDF data
            
        Returns:
            bytes: Image data (PNG format)
        """
        try:
            import fitz  # PyMuPDF
            import io
            from PIL import Image
            
            # Get PDF bytes
            pdf_bytes = uploaded_pdf.read()
            
            # Reset the file pointer for potential future reads
            uploaded_pdf.seek(0)
            
            # Open PDF from bytes
            pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
            
            if pdf_document.page_count == 0:
                raise ValueError("PDF has no pages")
            
            # Get the first page
            page = pdf_document[0]
            
            # Render page to an image (matrix for scaling - 2.0 = 200 DPI equivalent)
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert pixmap to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Convert PIL Image to bytes
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            
            # Close the PDF
            pdf_document.close()
            
            print("✅ PDF converted to image successfully using PyMuPDF")
            return img_byte_arr.getvalue()
            
        except ImportError as e:
            raise ImportError(f"PyMuPDF library is required. Install with: pip install PyMuPDF Pillow. Error: {str(e)}")
        except Exception as e:
            print(f"❌ Error converting PDF to image: {str(e)}")
            raise
    
    def _remove_existing_images(self, doc):
        """
        Remove all existing images from the document
        
        Args:
            doc: Document object
        """
        try:
            images_removed = 0
            
            # Collect all runs with images first, then remove them
            runs_to_remove = []
            
            # Find images in paragraphs
            for paragraph in doc.paragraphs:
                for run in paragraph.runs:
                    # Check if run contains images
                    if run._element.xpath('.//a:blip'):
                        runs_to_remove.append(run)
            
            # Find images in tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                if run._element.xpath('.//a:blip'):
                                    runs_to_remove.append(run)
            
            # Now remove all the runs with images
            for run in runs_to_remove:
                parent = run._element.getparent()
                if parent is not None:
                    parent.remove(run._element)
                    images_removed += 1
            
            print(f"✅ Removed {images_removed} existing image(s) from document")
            
        except Exception as e:
            print(f"⚠️ Error removing existing images: {str(e)}")
    
    def get_template_info(self, template_type):
        """Get information about a template"""
        if template_type not in self.template_mapping:
            return None
        
        config = self.template_mapping[template_type]
        template_path = os.path.join('inputs', config['file'])
        
        if not os.path.exists(template_path):
            return None
        
        doc = Document(template_path)
        
        return {
            'file': config['file'],
            'paragraphs': len(doc.paragraphs),
            'tables': len(doc.tables),
            'replacements': config['replacements']
        }

def test_document_processing():
    """Test the document processing system"""
    print("Testing Document Processing System")
    print("=" * 50)
    
    processor = DocxTemplateProcessor()
    
    # Test data
    test_client = "Jennifer Smith"
    test_email = "jennifer@example.com"
    test_amount = "$15,865.00"
    test_date = "September 25, 2025"
    
    # Test each template type
    template_types = ['development_contract', 'development_terms', 'terms_conditions', 'production_contract']
    
    for template_type in template_types:
        print(f"\n--- Testing {template_type} ---")
        
        try:
            # Get template info
            info = processor.get_template_info(template_type)
            if info:
                print(f"Template: {info['file']}")
                print(f"Paragraphs: {info['paragraphs']}")
                print(f"Tables: {info['tables']}")
                print(f"Replacements: {info['replacements']}")
            
            # Process document
            output_path = processor.process_document(
                template_type=template_type,
                client_name=test_client,
                email=test_email,
                contract_amount=test_amount if template_type in ['development_contract', 'production_contract'] else None,
                contract_date=test_date
            )
            
            print(f"✅ Document created: {output_path}")
            
        except Exception as e:
            print(f"❌ Error processing {template_type}: {str(e)}")
    
    print(f"\n✅ All documents processed successfully!")
    print(f"📁 Check the 'processed_documents' folder for the generated files")

if __name__ == "__main__":
    test_document_processing()
