"""
QuickBooks API Integration Module
Handles invoice creation and sending
"""

import requests
import json
from typing import Dict, Optional, Tuple
import streamlit as st
from datetime import datetime

class QuickBooksAPI:
    """QuickBooks API client for invoice creation and sending"""
    
    def __init__(self, client_id: str, client_secret: str, 
                 refresh_token: str, company_id: str, sandbox: bool = True):
        """
        Initialize QuickBooks API client
        
        Args:
            client_id: QuickBooks application client ID
            client_secret: QuickBooks application client secret
            refresh_token: OAuth refresh token
            company_id: QuickBooks company ID
            sandbox: Whether to use sandbox environment (default: True)
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.refresh_token = refresh_token
        self.company_id = company_id
        self.sandbox = sandbox
        
        # Set base URL based on environment
        if sandbox:
            self.base_url = "https://sandbox-quickbooks.api.intuit.com"
        else:
            self.base_url = "https://quickbooks.api.intuit.com"
        
        self.access_token = None
    
    def authenticate(self) -> bool:
        """
        Authenticate with QuickBooks API using refresh token
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        try:
            auth_url = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
            
            headers = {
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json"
            }
            
            data = {
                "grant_type": "refresh_token",
                "refresh_token": self.refresh_token
            }
            
            # Use basic auth with client credentials
            auth = requests.auth.HTTPBasicAuth(self.client_id, self.client_secret)
            
            response = requests.post(auth_url, data=data, headers=headers, auth=auth)
            response.raise_for_status()
            
            auth_response = response.json()
            self.access_token = auth_response.get("access_token")
            
            if not self.access_token:
                st.error("Failed to get access token from QuickBooks")
                return False
                
            return True
            
        except requests.exceptions.RequestException as e:
            st.error(f"QuickBooks authentication failed: {str(e)}")
            return False
        except Exception as e:
            st.error(f"Unexpected error during QuickBooks authentication: {str(e)}")
            return False
    
    def create_customer(self, first_name: str, last_name: str, email: str) -> Optional[str]:
        """
        Create a customer in QuickBooks or use existing customer in sandbox
        
        Args:
            first_name: Customer's first name
            last_name: Customer's last name
            email: Customer's email address
            
        Returns:
            str: Customer ID if successful, None otherwise
        """
        if not self.access_token:
            if not self.authenticate():
                return None
        
        # In sandbox mode, we can't create customers, so use existing ones
        if self.sandbox:
            st.info("🔧 Sandbox Mode: Using existing customer for testing")
            return self._get_or_create_sandbox_customer(first_name, last_name, email)
        
        # First check if customer already exists
        existing_customer = self._find_customer_by_email(email)
        if existing_customer:
            st.info(f"✅ Customer already exists with email: {email}")
            return existing_customer
        
        try:
            customer_url = f"{self.base_url}/v3/company/{self.company_id}/customer"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Accept-Encoding": "identity"  # Disable gzip to avoid decoding issues
            }
            
            # Use the correct format based on QuickBooks API documentation
            customer_data = {
                "DisplayName": f"{first_name} {last_name}",
                "GivenName": first_name,
                "FamilyName": last_name,
                "PrimaryEmailAddr": {
                    "Address": email
                }
            }
            
            # Customer object should be at root level, not wrapped
            payload = customer_data
            
            response = requests.post(customer_url, json=payload, headers=headers)
            response.raise_for_status()
            
            customer_response = response.json()
            customer = customer_response.get("Customer")
            
            if customer:
                return customer.get("Id")
            else:
                # Customer might already exist, try to find them
                return self._find_customer_by_email(email)
                
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to create customer: {str(e)}")
            return None
        except Exception as e:
            st.error(f"Unexpected error creating customer: {str(e)}")
            return None
    
    def _get_or_create_sandbox_customer(self, first_name: str, last_name: str, email: str) -> Optional[str]:
        """
        Get or create a customer for sandbox testing
        Since sandbox doesn't allow customer creation, we'll use existing customers
        """
        try:
            # First try to find existing customer by email
            existing_customer = self._find_customer_by_email(email)
            if existing_customer:
                return existing_customer
            
            # If not found, get the first available customer for testing
            query_url = f"{self.base_url}/v3/company/{self.company_id}/query?query=SELECT * FROM Customer"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json",
                "Accept-Encoding": "identity"  # Disable gzip to avoid decoding issues
            }
            
            response = requests.get(query_url, headers=headers)
            response.raise_for_status()
            
            customers_data = response.json()
            customers = customers_data.get('QueryResponse', {}).get('Customer', [])
            
            if customers:
                # Use the first customer for testing
                customer = customers[0]
                customer_id = customer.get('Id')
                customer_name = customer.get('Name', 'Unknown')
                
                st.info(f"🔧 Using existing customer '{customer_name}' (ID: {customer_id}) for testing")
                return customer_id
            else:
                st.error("No customers found in sandbox")
                return None
                
        except Exception as e:
            st.error(f"Failed to get sandbox customer: {str(e)}")
            return None
    
    def _find_customer_by_email(self, email: str) -> Optional[str]:
        """
        Find existing customer by email address
        
        Args:
            email: Customer's email address
            
        Returns:
            str: Customer ID if found, None otherwise
        """
        try:
            query_url = f"{self.base_url}/v3/company/{self.company_id}/query"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Accept-Encoding": "identity"  # Disable gzip to avoid decoding issues
            }
            
            # Query to find customer by email
            query = f"SELECT * FROM Customer WHERE PrimaryEmailAddr = '{email}'"
            
            params = {"query": query}
            
            response = requests.get(query_url, params=params, headers=headers)
            response.raise_for_status()
            
            query_response = response.json()
            customers = query_response.get("QueryResponse", {}).get("Customer", [])
            
            if customers:
                return customers[0].get("Id")
            
            return None
            
        except Exception as e:
            st.error(f"Error finding customer: {str(e)}")
            return None
    
    def create_invoice(self, customer_id: str, first_name: str, last_name: str, 
                     email: str, contract_amount: str, description: str = "Contract Services") -> Optional[str]:
        """
        Create an invoice in QuickBooks or simulate in sandbox
        
        Args:
            customer_id: ID of the customer
            first_name: Customer's first name
            last_name: Customer's last name
            email: Customer's email address
            contract_amount: Contract amount (will be converted to float)
            description: Description of the service
            
        Returns:
            str: Invoice ID if successful, None otherwise
        """
        if not self.access_token:
            if not self.authenticate():
                return None
        
        # In sandbox mode, we can't create invoices, so simulate the process
        if self.sandbox:
            st.info("🔧 Sandbox Mode: Simulating invoice creation")
            return self._simulate_invoice_creation(customer_id, first_name, last_name, email, contract_amount, description)
        
        try:
            # Convert contract amount to float
            amount_str = contract_amount.replace('$', '').replace(',', '')
            amount = float(amount_str)
            
            invoice_url = f"{self.base_url}/v3/company/{self.company_id}/invoice"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Accept-Encoding": "identity"  # Disable gzip to avoid decoding issues
            }
            
            # Create invoice data using correct format
            invoice_data = {
                "CustomerRef": {
                    "value": customer_id
                },
                "TxnDate": datetime.now().strftime("%Y-%m-%d"),
                "DueDate": datetime.now().strftime("%Y-%m-%d"),
                "Line": [
                    {
                        "DetailType": "SalesItemLineDetail",
                        "Amount": amount,
                        "SalesItemLineDetail": {
                            "ItemRef": {
                                "value": "1"  # Default service item - would need to be configured
                            },
                            "Qty": 1,
                            "UnitPrice": amount
                        },
                        "Description": description
                    }
                ],
                "EmailStatus": "NotSet"
            }
            
            # Invoice object should be at root level, not wrapped
            payload = invoice_data
            
            response = requests.post(invoice_url, json=payload, headers=headers)
            response.raise_for_status()
            
            invoice_response = response.json()
            invoice = invoice_response.get("Invoice")
            
            if invoice:
                return invoice.get("Id")
            else:
                st.error("Failed to create invoice")
                return None
                
        except ValueError:
            st.error("Invalid contract amount format")
            return None
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to create invoice: {str(e)}")
            return None
        except Exception as e:
            st.error(f"Unexpected error creating invoice: {str(e)}")
            return None
    
    def _simulate_invoice_creation(self, customer_id: str, first_name: str, last_name: str, 
                                 email: str, contract_amount: str, description: str) -> str:
        """
        Simulate invoice creation for sandbox testing
        """
        try:
            # Convert contract amount to float
            amount_str = contract_amount.replace('$', '').replace(',', '')
            amount = float(amount_str)
            
            # Generate a simulated invoice ID
            simulated_invoice_id = f"SIM_{customer_id}_{int(datetime.now().timestamp())}"
            
            st.success(f"✅ Invoice simulation successful!")
            st.info(f"📋 Simulated Invoice Details:")
            st.info(f"   • Customer: {first_name} {last_name}")
            st.info(f"   • Email: {email}")
            st.info(f"   • Amount: ${amount:,.2f}")
            st.info(f"   • Description: {description}")
            st.info(f"   • Simulated Invoice ID: {simulated_invoice_id}")
            st.info(f"   • Date: {datetime.now().strftime('%Y-%m-%d')}")
            
            st.warning("🔧 Note: This is a sandbox simulation. In production, a real invoice would be created.")
            
            return simulated_invoice_id
            
        except ValueError:
            st.error("Invalid contract amount format")
            return None
        except Exception as e:
            st.error(f"Error in invoice simulation: {str(e)}")
            return None
    
    def send_invoice(self, invoice_id: str, email: str) -> bool:
        """
        Send invoice to customer via email or simulate in sandbox
        
        Args:
            invoice_id: ID of the invoice to send
            email: Email address to send to
            
        Returns:
            bool: True if successful, False otherwise
        """
        if not self.access_token:
            if not self.authenticate():
                return False
        
        # In sandbox mode, simulate sending the invoice
        if self.sandbox:
            st.info("🔧 Sandbox Mode: Simulating invoice email")
            st.success(f"✅ Invoice email simulation successful!")
            st.info(f"📧 Simulated sending invoice {invoice_id} to {email}")
            st.warning("🔧 Note: This is a sandbox simulation. In production, a real email would be sent.")
            return True
        
        try:
            # Use the correct endpoint format from QuickBooks API documentation
            send_url = f"{self.base_url}/v3/company/{self.company_id}/invoice/{invoice_id}/send?sendTo={email}"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/octet-stream",
                "Accept": "application/json",
                "Accept-Encoding": "identity"  # Disable gzip to avoid decoding issues
            }
            
            # Send empty body as per API documentation
            response = requests.post(send_url, headers=headers)
            response.raise_for_status()
            
            return True
            
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to send invoice: {str(e)}")
            return False
        except Exception as e:
            st.error(f"Unexpected error sending invoice: {str(e)}")
            return False
    
    def create_and_send_invoice(self, first_name: str, last_name: str, email: str, 
                              contract_amount: str, description: str = "Contract Services") -> Tuple[bool, str]:
        """
        Complete workflow: create customer, create invoice, and send it
        
        Args:
            first_name: Customer's first name
            last_name: Customer's last name
            email: Customer's email address
            contract_amount: Contract amount
            description: Description of the service
            
        Returns:
            Tuple[bool, str]: (success, message)
        """
        try:
            # Create or find customer
            customer_id = self.create_customer(first_name, last_name, email)
            if not customer_id:
                return False, "Failed to create or find customer"
            
            # Create invoice
            invoice_id = self.create_invoice(customer_id, first_name, last_name, email, contract_amount, description)
            if not invoice_id:
                return False, "Failed to create invoice"
            
            # Send invoice
            if self.send_invoice(invoice_id, email):
                return True, f"Invoice created and sent successfully to {email}. Invoice ID: {invoice_id}"
            else:
                return False, "Invoice created but failed to send"
                
        except Exception as e:
            return False, f"Error in invoice workflow: {str(e)}"


def load_quickbooks_credentials() -> Dict[str, str]:
    """
    Load QuickBooks credentials from Streamlit secrets
    
    Returns:
        Dict containing QuickBooks credentials
    """
    try:
        if 'quickbooks' not in st.secrets:
            st.error("QuickBooks configuration not found in secrets.toml")
            return {}
        
        quickbooks_config = st.secrets['quickbooks']
        
        required_fields = ['client_id', 'client_secret', 'refresh_token', 'company_id']
        for field in required_fields:
            if field not in quickbooks_config:
                st.error(f"QuickBooks {field} not found in secrets.toml")
                return {}
        
        return quickbooks_config
        
    except Exception as e:
        st.error(f"Error reading QuickBooks secrets: {str(e)}")
        return {}


def setup_quickbooks_oauth() -> str:
    """
    Instructions for setting up QuickBooks OAuth
    
    Returns:
        str: Instructions for OAuth setup
    """
    return """
    To set up QuickBooks OAuth:
    
    1. Go to https://developer.intuit.com/
    2. Create a new app or use existing app
    3. Get your Client ID and Client Secret
    4. Set up OAuth redirect URI
    5. Use OAuth flow to get refresh token
    6. Get your Company ID from QuickBooks
    7. Add all credentials to secrets.toml
    
    For testing, you can use the sandbox environment.
    """
