"""
SignNow API Integration Module
Handles document creation and sending for contract signing
"""

import requests
import json
import base64
from typing import Dict, Optional, Tuple
import streamlit as st

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
    
    def create_document_from_template(self, template_id: str, document_name: str, 
                                   first_name: str, last_name: str, email: str, 
                                   contract_amount: str) -> Optional[str]:
        """
        Create a new document by uploading a PDF template
        
        Args:
            template_id: Not used (kept for compatibility)
            document_name: Name for the new document
            first_name: First name value
            last_name: Last name value
            email: Email value
            contract_amount: Contract amount value
            
        Returns:
            str: Document ID if successful, None otherwise
        """
        if not self.access_token:
            if not self.authenticate():
                return None
        
        try:
            # Create a simple PDF contract
            pdf_content = self._create_contract_pdf(first_name, last_name, email, contract_amount)
            
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
                
            return document_id
            
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to create document: {str(e)}")
            return None
        except Exception as e:
            st.error(f"Unexpected error creating document: {str(e)}")
            return None
    
    def _create_contract_pdf(self, first_name: str, last_name: str, email: str, contract_amount: str) -> bytes:
        """
        Create a simple PDF contract with the provided information
        
        Args:
            first_name: First name
            last_name: Last name
            email: Email address
            contract_amount: Contract amount
            
        Returns:
            bytes: PDF content
        """
        # Simple PDF content - in production you'd use a proper PDF library like reportlab
        pdf_content = f"""%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
/Resources <<
/Font <<
/F1 5 0 R
>>
>>
>>
endobj

4 0 obj
<<
/Length 200
>>
stream
BT
/F1 16 Tf
100 700 Td
(CONTRACT AGREEMENT) Tj
0 -30 Td
/F1 12 Tf
(Client: {first_name} {last_name}) Tj
0 -20 Td
(Email: {email}) Tj
0 -20 Td
(Contract Amount: ${contract_amount}) Tj
0 -40 Td
(This contract is generated automatically.) Tj
0 -20 Td
(Please review and sign this document.) Tj
ET
endstream
endobj

5 0 obj
<<
/Type /Font
/Subtype /Type1
/BaseFont /Helvetica
>>
endobj

xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000274 00000 n 
0000000368 00000 n 
trailer
<<
/Size 6
/Root 1 0 R
>>
startxref
590
%%EOF"""
        
        return pdf_content.encode('utf-8')
    
    def _fill_document_fields(self, document_id: str, first_name: str, last_name: str, 
                            email: str, contract_amount: str) -> bool:
        """
        Fill in document fields with provided values
        
        Args:
            document_id: ID of the document to fill
            first_name: First name value
            last_name: Last name value
            email: Email value
            contract_amount: Contract amount value
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Get document structure to find field IDs
            doc_url = f"{self.base_url}/document/{document_id}"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json"
            }
            
            response = requests.get(doc_url, headers=headers)
            response.raise_for_status()
            
            doc_data = response.json()
            
            # Prepare field values (this would need to be customized based on actual template)
            field_values = {
                "first_name": first_name,
                "last_name": last_name,
                "email": email,
                "contract_amount": contract_amount
            }
            
            # Update document fields
            update_url = f"{self.base_url}/document/{document_id}/field"
            
            for field_name, field_value in field_values.items():
                field_data = {
                    "field_name": field_name,
                    "field_value": field_value
                }
                
                field_response = requests.put(update_url, json=field_data, headers=headers)
                field_response.raise_for_status()
            
            return True
            
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to fill document fields: {str(e)}")
            return False
        except Exception as e:
            st.error(f"Unexpected error filling fields: {str(e)}")
            return False
    
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
    
    def create_and_send_contract(self, first_name: str, last_name: str, email: str, 
                               contract_amount: str, template_id: str = None) -> Tuple[bool, str]:
        """
        Complete workflow: create document from template and send for signing
        
        Args:
            first_name: First name of the signer
            last_name: Last name of the signer
            email: Email address of the signer
            contract_amount: Contract amount
            template_id: Template ID to use (optional, will use default if not provided)
            
        Returns:
            Tuple[bool, str]: (success, message)
        """
        try:
            # Use default template if none provided
            if not template_id:
                template_id = "default_contract_template"  # This would be replaced with actual template ID
            
            document_name = f"Contract_{first_name}_{last_name}_{contract_amount}"
            
            # Create document
            document_id = self.create_document_from_template(
                template_id, document_name, first_name, last_name, email, contract_amount
            )
            
            if not document_id:
                return False, "Failed to create document from template"
            
            # Send for signing
            if self.send_document_for_signing(document_id, email, first_name, last_name):
                return True, f"Contract sent successfully to {email}. Document ID: {document_id}"
            else:
                return False, "Document created but failed to send for signing"
                
        except Exception as e:
            return False, f"Error in contract workflow: {str(e)}"


def load_signnow_credentials() -> Dict[str, str]:
    """
    Load SignNow credentials from Streamlit secrets
    
    Returns:
        Dict containing SignNow credentials
    """
    try:
        if 'signnow' not in st.secrets:
            st.error("SignNow configuration not found in secrets.toml")
            return {}
        
        signnow_config = st.secrets['signnow']
        
        required_fields = ['client_id', 'client_secret', 'basic_auth_token', 'username', 'password']
        for field in required_fields:
            if field not in signnow_config:
                st.error(f"SignNow {field} not found in secrets.toml")
                return {}
        
        return signnow_config
        
    except Exception as e:
        st.error(f"Error reading SignNow secrets: {str(e)}")
        return {}


def create_sample_contract_template() -> str:
    """
    Create a sample contract template (this would typically be done through SignNow UI)
    For now, return a placeholder template ID
    
    Returns:
        str: Template ID
    """
    # In a real implementation, you would:
    # 1. Create a PDF template with form fields
    # 2. Upload it to SignNow
    # 3. Get the template ID
    # 4. Return that ID
    
    return "sample_contract_template_id"
