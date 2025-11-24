"""
SignNow API Integration Module
Handles document creation and sending for contract signing
"""

import requests
import json
import base64
from typing import Dict, Optional, Tuple
import streamlit as st
import os
from docx_template_processor import DocxTemplateProcessor

class SignNowAPI:
    """SignNow API client for document creation and sending"""
    
    def __init__(self, client_id: str, client_secret: str, basic_auth_token: str, 
                 username: str, password: str, api_key: str = None):
        """
        Initialize SignNow API client
        
        Args:
            client_id: SignNow application client ID
            client_secret: SignNow application client secret
            basic_auth_token: SignNow basic authorization token
            username: SignNow account username/email
            password: SignNow account password
            api_key: SignNow API key (optional)
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.basic_auth_token = basic_auth_token
        self.username = username
        self.password = password
        self.api_key = api_key
        self.base_url = "https://api.signnow.com"  # Updated to production URL
        self.access_token = None
        self.user_email = None
        self.docx_processor = DocxTemplateProcessor()
        
    def authenticate(self) -> bool:
        """
        Authenticate with SignNow API using OAuth2 password grant
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        try:
            # OAuth2 token endpoint
            auth_url = "https://api.signnow.com/oauth2/token"
            
            # Headers as per SignNow documentation
            headers = {
                "Authorization": f"Basic {self.basic_auth_token}",
                "Content-Type": "application/x-www-form-urlencoded"
            }
            
            # Request body as per SignNow documentation
            auth_data = {
                "username": self.username,
                "password": self.password,
                "grant_type": "password",
                "scope": "*"
            }
            
            response = requests.post(auth_url, headers=headers, data=auth_data, timeout=30)
            response.raise_for_status()
            
            auth_response = response.json()
            self.access_token = auth_response.get("access_token")
            
            if not self.access_token:
                st.error("Failed to get access token from SignNow")
                return False
            
            # Get user email for document sending
            self._get_user_email()
                
            return True
            
        except requests.exceptions.RequestException as e:
            st.error(f"SignNow authentication failed: {str(e)}")
            return False
        except Exception as e:
            st.error(f"Unexpected error during SignNow authentication: {str(e)}")
            return False
    
    def _get_user_email(self) -> bool:
        """
        Get user's email address for document sending
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            user_url = f"{self.base_url}/user"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json"
            }
            
            response = requests.get(user_url, headers=headers, timeout=30)
            response.raise_for_status()
            
            user_data = response.json()
            emails = user_data.get('emails', [])
            
            if emails:
                self.user_email = emails[0]
                return True
            else:
                st.error("No email found for user")
                return False
                
        except Exception as e:
            st.error(f"Failed to get user email: {str(e)}")
            return False
    
    def create_document_from_template(self, template_type: str, document_name: str, 
                                   client_name: str, email: str, 
                                   contract_amount: str = None, contract_date: str = None) -> Optional[str]:
        """
        Create a new document by processing the original .docx template with exact formatting preserved
        
        Args:
            template_type: Type of template to use (development_contract, development_terms, terms_conditions, production_contract)
            document_name: Name for the new document
            client_name: Client name value
            email: Email value
            contract_amount: Contract amount value (optional)
            contract_date: Contract date value (optional)
            
        Returns:
            str: Document ID if successful, None otherwise
        """
        if not self.access_token:
            if not self.authenticate():
                return None
        
        try:
            # Process the .docx template with exact formatting preserved
            processed_docx_path = self.docx_processor.process_document(
                template_type=template_type,
                client_name=client_name,
                email=email,
                contract_amount=contract_amount,
                contract_date=contract_date
            )
            
            # Convert .docx to PDF for SignNow upload
            pdf_content = self._convert_docx_to_pdf(processed_docx_path)
            
            # Upload the PDF to SignNow
            files = {
                'file': (f'{document_name}.pdf', pdf_content, 'application/pdf')
            }
            
            data = {
                'name': document_name
            }
            
            headers = {
                "Authorization": f"Bearer {self.access_token}"
            }
            
            response = requests.post(
                f"{self.base_url}/document",
                files=files,
                data=data,
                headers=headers,
                timeout=60
            )
            response.raise_for_status()
            
            create_response = response.json()
            document_id = create_response.get("id")
            
            if not document_id:
                st.error("Failed to create document")
                return None
                
            # Clean up temporary file
            if os.path.exists(processed_docx_path):
                os.remove(processed_docx_path)
                
            return document_id
            
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to create document: {str(e)}")
            return None
        except Exception as e:
            st.error(f"Unexpected error creating document: {str(e)}")
            return None
    
    def _convert_docx_to_pdf(self, docx_path: str, highlight_values: list | None = None) -> bytes:
        """
        Convert .docx file to PDF bytes with proper formatting preservation using ReportLab
        
        Args:
            docx_path: Path to the .docx file
            
        Returns:
            bytes: PDF content
        """
        try:
            from docx import Document
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            from reportlab.lib import colors
            from io import BytesIO
            import html as html_escape
            
            doc = Document(docx_path)
            
            # Create PDF buffer
            buffer = BytesIO()
            pdf_doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=1*inch, bottomMargin=1*inch)
            
            # Get styles
            styles = getSampleStyleSheet()
            
            # Create custom styles for better formatting
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=20,
                spaceAfter=20,
                alignment=1,  # Center alignment
                textColor=colors.black,
                fontName='Helvetica-Bold'
            )
            
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontSize=16,
                spaceAfter=12,
                fontName='Helvetica-Bold',
                textColor=colors.black
            )
            
            bold_style = ParagraphStyle(
                'CustomBold',
                parent=styles['Normal'],
                fontSize=12,
                spaceAfter=8,
                fontName='Helvetica-Bold'
            )
            
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontSize=11,
                spaceAfter=6,
                fontName='Helvetica'
            )
            
            amount_style = ParagraphStyle(
                'ContractAmount',
                parent=styles['Normal'],
                fontSize=14,
                spaceAfter=8,
                fontName='Helvetica-Bold',
                textColor=colors.green
            )
            
            placeholder_style = ParagraphStyle(
                'Placeholder',
                parent=styles['Normal'],
                fontSize=16,
                spaceAfter=12,
                alignment=1,  # Center alignment
                fontName='Helvetica-Bold',
                textColor=colors.blue
            )
            
            screenshot_placeholder_style = ParagraphStyle(
                'ScreenshotPlaceholder',
                parent=styles['Normal'],
                fontSize=16,
                spaceAfter=12,
                alignment=1,  # Center alignment
                fontName='Helvetica-Bold',
                textColor=colors.red
            )
            
            # Normalize highlight values
            highlight_values = [v for v in (highlight_values or []) if isinstance(v, str) and v.strip()]

            # Build PDF content
            story = []
            
            # Check for images/logos first
            if hasattr(doc, 'inline_shapes') and doc.inline_shapes:
                story.append(Paragraph("IMAGE/LOGO NOT DISPLAYED IN PREVIEW", placeholder_style))
                story.append(Spacer(1, 20))
            
            # Add document content with exact formatting preservation
            for paragraph in doc.paragraphs:
                # Check if this paragraph contains an image (screenshot placeholder)
                has_image = False
                if paragraph.runs:
                    for run in paragraph.runs:
                        if run._element.xpath('.//a:blip'):
                            has_image = True
                            break
                
                # If paragraph has an image, show red placeholder
                if has_image:
                    story.append(Paragraph("IMAGE ATTACHED GOES HERE", screenshot_placeholder_style))
                    story.append(Spacer(1, 12))
                    continue
                
                if paragraph.text.strip():
                    text = paragraph.text.strip()
                    
                    # Check for actual bold formatting in the document
                    is_bold = paragraph.runs and any(run.bold for run in paragraph.runs)
                    
                    # Check for actual font size in the document
                    font_size = 11  # Default size
                    if paragraph.runs:
                        for run in paragraph.runs:
                            if run.font.size:
                                font_size = run.font.size.pt
                                break
                    
                    # Map font families to ReportLab supported fonts
                    def map_font_family(font_name, is_bold=False):
                        """Map document fonts to ReportLab supported fonts"""
                        font_name = font_name.lower() if font_name else 'helvetica'
                        
                        # Font mapping
                        font_map = {
                            'arial': 'Helvetica',
                            'times': 'Times-Roman',
                            'times new roman': 'Times-Roman',
                            'courier': 'Courier',
                            'helvetica': 'Helvetica',
                            'calibri': 'Helvetica',  # Map Calibri to Helvetica
                            'verdana': 'Helvetica',  # Map Verdana to Helvetica
                        }
                        
                        # Get base font
                        base_font = font_map.get(font_name, 'Helvetica')
                        
                        # Add bold suffix if needed
                        if is_bold:
                            if base_font == 'Helvetica':
                                return 'Helvetica-Bold'
                            elif base_font == 'Times-Roman':
                                return 'Times-Bold'
                            elif base_font == 'Courier':
                                return 'Courier-Bold'
                            else:
                                return 'Helvetica-Bold'  # Fallback
                        else:
                            return base_font
                    
                    # Check for actual font family in the document
                    font_family = 'Helvetica'  # Default font
                    if paragraph.runs:
                        for run in paragraph.runs:
                            if run.font.name:
                                font_family = run.font.name
                                break
                    
                    # Force all paragraphs to normal weight; only bold exact variable values inline
                    style = ParagraphStyle(
                        'DynamicNormal',
                        parent=styles['Normal'],
                        fontSize=font_size,
                        spaceAfter=4,
                        fontName=map_font_family(font_family, False),
                        textColor=colors.black
                    )
                    
                    # Escape text for XML/HTML and bold only highlight values inline
                    escaped_text = html_escape.escape(text)
                    if highlight_values:
                        # Replace longer values first to avoid partial nesting
                        for val in sorted(highlight_values, key=lambda v: len(v or ''), reverse=True):
                            if not val:
                                continue
                            escaped_val = html_escape.escape(val)
                            escaped_text = escaped_text.replace(escaped_val, f"<b>{escaped_val}</b>")

                    story.append(Paragraph(escaped_text, style))
                    
                    # Add spacing based on paragraph formatting
                    if paragraph.paragraph_format.space_after:
                        spacing = paragraph.paragraph_format.space_after.pt
                        story.append(Spacer(1, spacing))
                    else:
                        story.append(Spacer(1, 2))  # Minimal default spacing
            
            # Add table placeholders
            if doc.tables:
                story.append(Spacer(1, 20))
                story.append(Paragraph("TABLE NOT DISPLAYED IN PREVIEW", placeholder_style))
            
            # Build PDF
            pdf_doc.build(story)
            
            # Get PDF bytes
            pdf_bytes = buffer.getvalue()
            buffer.close()
            
            return pdf_bytes
            
        except Exception as e:
            st.error(f"Error converting .docx to PDF: {str(e)}")
            # Fallback to simple PDF
            return b"%PDF-1.4\n1 0 obj\n<<\n/Type /Catalog\n/Pages 2 0 R\n>>\nendobj\n%%EOF"

    def _merge_pdfs(self, pdf_bytes_list: list[bytes]) -> bytes:
        """Merge multiple PDF byte streams into a single PDF bytes using PyPDF2."""
        try:
            from PyPDF2 import PdfReader, PdfWriter
            from io import BytesIO

            writer = PdfWriter()
            for pdf_bytes in pdf_bytes_list:
                reader = PdfReader(BytesIO(pdf_bytes))
                for page in reader.pages:
                    writer.add_page(page)

            out = BytesIO()
            writer.write(out)
            out.seek(0)
            return out.getvalue()
        except Exception as e:
            st.error(f"Failed to merge PDFs: {str(e)}")
            # Return first PDF as fallback
            return pdf_bytes_list[0] if pdf_bytes_list else b""

    def upload_pdf(self, document_name: str, pdf_bytes: bytes) -> str | None:
        """Upload a single PDF to SignNow and return document id."""
        if not self.access_token:
            if not self.authenticate():
                return None
        try:
            files = {
                'file': (f'{document_name}.pdf', pdf_bytes, 'application/pdf')
            }
            data = { 'name': document_name }
            headers = { "Authorization": f"Bearer {self.access_token}" }
            response = requests.post(f"{self.base_url}/document", files=files, data=data, headers=headers, timeout=90)
            response.raise_for_status()
            doc_id = response.json().get("id")
            return doc_id
        except Exception as e:
            st.error(f"Failed to upload merged PDF: {str(e)}")
            return None

    def create_and_send_merged_pair(self, pair_type: str, contract_docx_path: str, terms_docx_path: str,
                                    document_name: str, email: str, highlight_values: list | None = None) -> tuple[bool, str]:
        """Convert both DOCX to PDFs, merge into one, upload, and send invite."""
        try:
            # Convert to PDFs
            contract_pdf = self._convert_docx_to_pdf(contract_docx_path, highlight_values=highlight_values)
            terms_pdf = self._convert_docx_to_pdf(terms_docx_path, highlight_values=highlight_values)
            # Merge PDFs: contract first, then terms
            merged_pdf = self._merge_pdfs([contract_pdf, terms_pdf])
            if not merged_pdf:
                return False, "Failed to generate merged PDF"
            # Upload merged PDF with provided document_name (already formatted by caller)
            doc_id = self.upload_pdf(document_name, merged_pdf)
            if not doc_id:
                return False, "Failed to upload merged document"
            # Send for signing
            if self.send_document_for_signing(doc_id, email, document_name, ""):
                return True, f"Merged document sent successfully to {email}. Document ID: {doc_id}"
            return False, "Merged document uploaded but failed to send for signing"
        except Exception as e:
            return False, f"Error creating/sending merged pair: {str(e)}"

    def _merge_docx(self, first_path: str, second_path: str) -> str:
        """Merge two DOCX files into one DOCX (simple append of body content). Returns merged path."""
        from docx import Document as Docx
        import os
        merged = Docx(first_path)
        second = Docx(second_path)

        # Append paragraphs and tables from second into merged
        for element in second.element.body:
            merged.element.body.append(element)

        out_dir = os.path.join(os.getcwd(), 'processed_documents')
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, 'merged_contract_terms.docx')
        merged.save(out_path)
        return out_path

    def upload_docx(self, document_name: str, docx_path: str) -> str | None:
        """Upload a DOCX to SignNow (let SignNow convert). Returns document id."""
        if not self.access_token:
            if not self.authenticate():
                return None
        try:
            with open(docx_path, 'rb') as f:
                content = f.read()
            files = {
                'file': (f"{document_name}.docx", content, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            }
            data = { 'name': document_name }
            headers = { "Authorization": f"Bearer {self.access_token}" }
            resp = requests.post(f"{self.base_url}/document", files=files, data=data, headers=headers, timeout=180)
            resp.raise_for_status()
            return resp.json().get('id')
        except Exception as e:
            st.error(f"Failed to upload DOCX: {str(e)}")
            return None

    def create_and_send_merged_pair_docx(self, pair_type: str, contract_docx_path: str, terms_docx_path: str,
                                          document_name: str, email: str) -> tuple[bool, str]:
        """Merge two processed DOCX files into one DOCX, upload, and send for signing."""
        try:
            merged_docx = self._merge_docx(contract_docx_path, terms_docx_path)
            if not merged_docx or not os.path.exists(merged_docx):
                return False, "Failed to merge DOCX documents"
            doc_id = self.upload_docx(document_name, merged_docx)
            if not doc_id:
                return False, "Failed to upload merged DOCX"
            if self.send_document_for_signing(doc_id, email, document_name, ""):
                return True, f"Merged document sent successfully to {email}. Document ID: {doc_id}"
            return False, "Merged DOCX uploaded but failed to send for signing"
        except Exception as e:
            return False, f"Error creating/sending merged DOCX: {str(e)}"
    
    def _docx_to_html(self, doc):
        """
        Convert docx Document to HTML with proper formatting preservation
        
        Args:
            doc: python-docx Document object
            
        Returns:
            str: HTML content
        """
        html_parts = []
        html_parts.append('<html><head><meta charset="UTF-8"></head><body>')
        
        # Check for images/logos first
        if hasattr(doc, 'inline_shapes') and doc.inline_shapes:
            html_parts.append('<div class="image-placeholder" style="color: blue; font-weight: bold; text-align: center; font-size: 16px;">IMAGE/LOGO NOT DISPLAYED IN PREVIEW</div>')
        
        for paragraph in doc.paragraphs:
            # Check if this paragraph contains an image (screenshot placeholder)
            has_image = False
            if paragraph.runs:
                for run in paragraph.runs:
                    if run._element.xpath('.//a:blip'):
                        has_image = True
                        break
            
            # If paragraph has an image, show red placeholder
            if has_image:
                html_parts.append('<div class="screenshot-placeholder" style="color: red; font-weight: bold; text-align: center; font-size: 16px;">IMAGE ATTACHED GOES HERE</div>')
                continue
            
            if paragraph.text.strip():
                text = paragraph.text.strip()
                
                # Check for actual bold formatting in the document
                is_bold = paragraph.runs and any(run.bold for run in paragraph.runs)
                
                # Determine formatting based on content and actual formatting
                if any(word in text.upper() for word in ['DEVELOPMENT CONTRACT', 'TERMS AND CONDITIONS', 'PRODUCTION CONTRACT']):
                    html_parts.append(f'<h1>{text}</h1>')
                elif any(word in text.upper() for word in ['CONTRACT', 'AGREEMENT', 'TERMS', 'CONDITIONS']):
                    if len(text) < 100:  # Likely a heading
                        html_parts.append(f'<h2>{text}</h2>')
                    elif is_bold:
                        html_parts.append(f'<p class="bold">{text}</p>')
                    else:
                        html_parts.append(f'<p>{text}</p>')
                elif text.startswith('The following agreement'):
                    html_parts.append(f'<p class="bold">{text}</p>')
                elif 'total contract amount' in text.lower():
                    # Highlight contract amount
                    text_with_highlight = text.replace('$', '<span class="contract-amount">$').replace(' and is due', '</span> and is due')
                    html_parts.append(f'<p class="bold">{text_with_highlight}</p>')
                elif any(word in text.upper() for word in ['SIGNATURE', 'SIGNED', 'DATE', 'VITALINA', 'TEG INTL', 'ON BEHALF']):
                    html_parts.append(f'<p class="bold">{text}</p>')
                elif is_bold:
                    html_parts.append(f'<p class="bold">{text}</p>')
                else:
                    html_parts.append(f'<p>{text}</p>')
        
        # Add table placeholders
        if doc.tables:
            html_parts.append('<div class="table-placeholder">TABLE NOT DISPLAYED IN PREVIEW</div>')
        
        html_parts.append('</body></html>')
        return ''.join(html_parts)
    
    
    def send_document_for_signing(self, document_id: str, email: str, 
                                first_name: str, last_name: str) -> bool:
        """
        Send document for signing to the specified email
        
        Args:
            document_id: ID of the document to send
            email: Email address to send to
            first_name: First name of the signer
            last_name: Last name of the signer
            
        Returns:
            bool: True if successful, False otherwise
        """
        if not self.access_token:
            if not self.authenticate():
                return False
        
        try:
            send_url = f"{self.base_url}/document/{document_id}/invite"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json"
            }
            
            data = {
                "to": email,
                "from": self.user_email
            }
            
            response = requests.post(send_url, json=data, headers=headers)
            response.raise_for_status()
            
            return True
            
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to send document: {str(e)}")
            return False
        except Exception as e:
            st.error(f"Unexpected error sending document: {str(e)}")
            return False
    
    def create_and_send_contract(self, template_type: str, client_name: str, email: str, 
                               contract_amount: str = None, contract_date: str = None) -> Tuple[bool, str]:
        """
        Complete workflow: create document from template and send for signing
        
        Args:
            template_type: Type of template to use (development_contract, development_terms, terms_conditions, production_contract)
            client_name: Client name
            email: Email address of the signer
            contract_amount: Contract amount (optional, required for contract types)
            contract_date: Contract date (optional, defaults to current date)
            
        Returns:
            Tuple[bool, str]: (success, message)
        """
        try:
            # Set default date if not provided
            if not contract_date:
                from datetime import datetime
                contract_date = datetime.now().strftime("%B %d, %Y")
            
            document_name = f"{template_type}_{client_name.replace(' ', '_')}_{contract_date.replace(' ', '_').replace(',', '')}"
            
            # Create document
            document_id = self.create_document_from_template(
                template_type, document_name, client_name, email, contract_amount, contract_date
            )
            
            if not document_id:
                return False, "Failed to create document from template"
            
            # Send for signing
            if self.send_document_for_signing(document_id, email, client_name, ""):
                return True, f"Contract sent successfully to {email}. Document ID: {document_id}"
            else:
                return False, "Document created but failed to send for signing"
                
        except Exception as e:
            return False, f"Error in contract workflow: {str(e)}"
    
    def create_and_send_document_pair(self, pair_type, client_name, email, 
                                    contract_amount=None, contract_date=None,
                                    deposit_amount=None, total_contract_amount=None,
                                    sewing_cost=None, pre_production_fee=None,
                                    tegmade_for=None):
        """
        Create and send a document pair for signing
        
        Args:
            pair_type: Type of document pair ('development_pair' or 'production_pair')
            client_name: Client name
            email: Email address of the signer
            contract_amount: Contract amount (for development contracts)
            contract_date: Contract date (optional, defaults to current date)
            deposit_amount: Deposit amount (for production contracts)
            total_contract_amount: Total contract amount (for production contracts)
            sewing_cost: Sewing cost (for production contracts)
            pre_production_fee: Pre-production fee (for production contracts)
            tegmade_for: Name to replace VITALINA GHINZELLI (optional)
            
        Returns:
            Tuple[bool, str]: (success, message)
        """
        try:
            # Set default date if not provided
            if not contract_date:
                from datetime import datetime
                contract_date = datetime.now().strftime("%B %d, %Y")
            
            # Determine which documents to process
            if pair_type == 'development_pair':
                # Process development contract and terms
                contract_template = 'development_contract'
                terms_template = 'development_terms'
            elif pair_type == 'production_pair':
                # Process production contract and terms
                contract_template = 'production_contract'
                terms_template = 'production_terms'
            else:
                return False, f"Unknown pair type: {pair_type}"
            
            # Process both documents
            from docx_template_processor import DocxTemplateProcessor
            processor = DocxTemplateProcessor()
            
            # Process contract document
            contract_path = processor.process_document(
                template_type=contract_template,
                client_name=client_name,
                email=email,
                contract_amount=contract_amount,
                contract_date=contract_date,
                deposit_amount=deposit_amount,
                total_contract_amount=total_contract_amount,
                sewing_cost=sewing_cost,
                pre_production_fee=pre_production_fee,
                tegmade_for=tegmade_for
            )
            
            # Process terms document
            terms_path = processor.process_document(
                template_type=terms_template,
                client_name=client_name,
                email=email,
                contract_amount=None,
                contract_date=contract_date,
                tegmade_for=tegmade_for
            )
            
            # Upload both documents to SignNow
            document_ids = []
            
            for doc_path, doc_type in [(contract_path, "Contract"), (terms_path, "Terms")]:
                with open(doc_path, 'rb') as docx_file:
                    docx_content = docx_file.read()
                
                actual_filename = os.path.basename(doc_path)
                files = {
                    'file': (actual_filename, docx_content, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                }
                
                data = {
                    'name': f'{pair_type}_{doc_type} - {client_name}'
                }
                
                headers = {
                    "Authorization": f"Bearer {self.access_token}"
                }
                
                # Create document in SignNow
                response = requests.post(
                    f"{self.base_url}/document",
                    files=files,
                    data=data,
                    headers=headers,
                    timeout=90
                )
                response.raise_for_status()
                
                create_response = response.json()
                document_id = create_response.get("id")
                document_ids.append(document_id)
                
                # Clean up processed document
                os.remove(doc_path)
            
            # Send both documents for signing
            for document_id in document_ids:
                send_url = f"{self.base_url}/document/{document_id}/invite"
                send_data = {
                    "to": email,
                    "from": self.user_email
                }
                
                send_response = requests.post(send_url, json=send_data, headers=headers)
                send_response.raise_for_status()
            
            return True, f"Document pair sent successfully to {email}. Document IDs: {', '.join(document_ids)}"
                
        except Exception as e:
            return False, f"Error creating and sending document pair: {str(e)}"


def load_signnow_credentials(account_name: str = None) -> Dict[str, str]:
    """
    Load SignNow credentials from Streamlit secrets
    
    Args:
        account_name: Name of the account to load ('heather', 'jennifer', 'anthony', or None for default)
    
    Returns:
        Dict containing SignNow credentials
    """
    try:
        # Determine which secrets section to use
        if account_name:
            # Use specific account section
            secrets_key = f'signnow - {account_name.lower()}'
        else:
            # Fallback to default signnow section if no account specified
            secrets_key = 'signnow'
        
        if secrets_key not in st.secrets:
            if account_name:
                st.error(f"SignNow configuration for '{account_name}' not found in secrets.toml")
            else:
                st.error("SignNow configuration not found in secrets.toml")
            return {}
        
        signnow_config = st.secrets[secrets_key]
        
        required_fields = ['client_id', 'client_secret', 'basic_auth_token', 'username', 'password']
        for field in required_fields:
            if field not in signnow_config:
                st.error(f"SignNow {field} not found in secrets.toml for account '{account_name}'")
                return {}
        
        return signnow_config
        
    except Exception as e:
        st.error(f"Error reading SignNow secrets: {str(e)}")
        return {}