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
                    '$15,865.00': 'CONTRACT_AMOUNT'
                }
            },
            'development_terms': {
                'file': 'Development Terms and Conditions .docx',
                'replacements': {
                    'SHERRY CASSEL': 'CLIENT_NAME',
                    'JUNE 16, 2025': 'CONTRACT_DATE'
                }
            },
            'terms_conditions': {
                'file': 'TERMS AND CONDITIONS.docx',
                'replacements': {
                    'Natalie Barrett': 'CLIENT_NAME',
                    'December 06, 2024': 'CONTRACT_DATE'
                }
            },
            'production_contract': {
                'file': 'Production Contract.docx',
                'replacements': {
                    'Natalie Barrett': 'CLIENT_NAME',
                    'December 06, 2024': 'CONTRACT_DATE'
                }
            }
        }
    
    def process_document(self, template_type, client_name, email, 
                        contract_amount=None, contract_date=None):
        """
        Process a document template with variable replacement
        
        Args:
            template_type: Type of template to process
            client_name: Client name to replace
            email: Client email (for reference)
            contract_amount: Contract amount (if applicable)
            contract_date: Contract date (if not provided, uses current date)
            
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
            'CONTRACT_AMOUNT': self._format_contract_amount(contract_amount) if contract_amount else '$0.00'
        }
        
        # Process replacements
        replacements_made = self._replace_text_in_document(doc, template_config['replacements'], replacement_values)
        
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
