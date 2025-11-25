"""
QuickBooks API Integration Module
Clean version for Production Cloud Deployment.
Removes all SSL/DNS hacks to resolve '400 No Handler' errors on cloud platforms.

FIXES APPLIED:
1. Bill To Address: Improved BillAddr handling to ensure it shows in PDF
2. Due Date: Fixed to respect the selected invoice date instead of calculating from payment terms
3. Email Message: Changed DisplayName to use client name instead of company name
"""

import requests
import requests.exceptions
import json
import os
import toml
import streamlit as st
from datetime import datetime
from typing import Dict, Optional, Tuple

# Standard session configuration with proper User-Agent
def get_qb_session():
    session = requests.Session()
    # Proper identification avoids WAF/Firewall blocks
    session.headers.update({
        'User-Agent': 'StreamlitQuickBooksApp/1.0',
        'Accept': 'application/json'
    })
    return session

class QuickBooksAPI:
    """QuickBooks API client for invoice creation and sending"""
    
    def __init__(self, client_id: str, client_secret: str, 
                 refresh_token: str, company_id: str, sandbox: bool = False):
        self.client_id = client_id
        self.client_secret = client_secret
        self.refresh_token = refresh_token
        self.company_id = company_id
        self.sandbox = sandbox
        
        # Official Production URL (No cluster manipulation)
        if sandbox:
            self.base_url = "https://sandbox-quickbooks.api.intuit.com"
        else:
            self.base_url = "https://quickbooks.api.intuit.com"
            
        self.access_token = None
        self.session = get_qb_session()

    def _get_headers(self):
        """Returns standard headers with current token"""
        if not self.access_token:
            return {}
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
            "Content-Type": "application/json",
            "User-Agent": "StreamlitQuickBooksApp/1.0"
        }

    def authenticate(self, force_refresh: bool = False) -> bool:
        """Authenticates with QuickBooks API using the refresh token"""
        if force_refresh:
            self.access_token = None
            
        try:
            auth_url = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
            
            data = {
                "grant_type": "refresh_token",
                "refresh_token": self.refresh_token
            }
            
            auth = requests.auth.HTTPBasicAuth(self.client_id, self.client_secret)
            
            # In cloud environments, verify=True is correct. SSL is trusted.
            response = self.session.post(auth_url, data=data, auth=auth, timeout=30)
            
            if response.status_code != 200:
                st.error(f"Authentication error: {response.text}")
                return False
            
            auth_response = response.json()
            self.access_token = auth_response.get("access_token")
            new_refresh_token = auth_response.get("refresh_token")
            
            if not self.access_token:
                return False
            
            # Update refresh token if it changed
            if new_refresh_token and new_refresh_token != self.refresh_token:
                self.refresh_token = new_refresh_token
                self._update_secrets_file(new_refresh_token)
                
            return True
            
        except Exception as e:
            st.error(f"Connection error during authentication: {str(e)}")
            return False

    def _update_secrets_file(self, new_token):
        """Attempts to update secrets.toml (best effort)"""
        try:
            secrets_path = os.path.join('.streamlit', 'secrets.toml')
            if os.path.exists(secrets_path):
                with open(secrets_path, 'r') as f:
                    config = toml.load(f)
                
                if 'quickbooks' in config:
                    config['quickbooks']['refresh_token'] = new_token
                    with open(secrets_path, 'w') as f:
                        toml.dump(config, f)
        except:
            # On cloud deploy, usually we cannot write to files.
            # Just log it internally or ignore.
            pass

    def _make_request(self, method, endpoint, data=None, params=None):
        """Centralized request wrapper with retry logic and minorversion"""
        if not self.access_token:
            if not self.authenticate():
                return None

        # Set default minorversion if not provided in params
        if params is None:
            params = {}
        # Only set default minorversion if not already specified
        if 'minorversion' not in params:
            params['minorversion'] = '40'

        url = f"{self.base_url}/v3/company/{self.company_id}/{endpoint}"
        
        try:
            if method.lower() == 'get':
                response = self.session.get(url, headers=self._get_headers(), params=params, timeout=45)
            else:
                response = self.session.post(url, headers=self._get_headers(), params=params, json=data, timeout=45)

            # If 401, try refreshing token once
            if response.status_code == 401:
                st.info("Token expired, refreshing...")
                if self.authenticate(force_refresh=True):
                    # Retry
                    if method.lower() == 'get':
                        response = self.session.get(url, headers=self._get_headers(), params=params, timeout=45)
                    else:
                        response = self.session.post(url, headers=self._get_headers(), params=params, json=data, timeout=45)
            
            return response
        except requests.exceptions.Timeout as e:
            st.error(f"⏰ Request timeout ({endpoint}): {str(e)}")
            return None
        except requests.exceptions.ConnectionError as e:
            st.error(f"🔌 Connection error ({endpoint}): {str(e)}")
            return None
        except Exception as e:
            st.error(f"❌ Request error ({endpoint}): {str(e)}")
            return None

    # --- Business Methods (Customer, Invoice, etc.) ---
    
    def _get_payment_term_id(self, payment_terms: str) -> str:
        """
        Get payment term ID for QuickBooks API
        This would need to be configured based on your QuickBooks setup
        """
        # This is a placeholder - in production, you'd query QuickBooks for actual term IDs
        # These IDs match your actual QuickBooks payment terms
        term_mapping = {
            "Due on receipt": "1",
            "Net 15": "2",
            "Net 30": "3", 
            "Net 60": "4"
        }
        return term_mapping.get(payment_terms, "1")  # Default to Due on receipt instead of Net 30
    
    def _parse_bill_address(self, company_name: Optional[str], client_address: Optional[str]) -> Optional[Dict[str, str]]:
        """
        Convert company/address strings into QuickBooks BillAddr structure (Line1-Line5).
        Fix: Now handles newlines properly from Streamlit text_area.
        """
        lines: list[str] = []
        
        if company_name:
            lines.append(company_name.strip())
        
        if client_address:
            # Normalize newlines
            clean_address = client_address.replace('\r\n', '\n').replace('\r', '\n')
            raw_lines = [part.strip() for part in clean_address.split('\n') if part.strip()]
            
            address_segments: list[str] = []
            for raw_line in raw_lines:
                # Split each line by comma to create separate address segments
                comma_parts = [seg.strip() for seg in raw_line.split(',') if seg.strip()]
                if comma_parts:
                    address_segments.extend(comma_parts)
                else:
                    address_segments.append(raw_line)
            
            # Merge state abbreviation + ZIP into single line (e.g., "CT" + "02703" -> "CT 02703")
            merged_segments: list[str] = []
            idx = 0
            while idx < len(address_segments):
                current = address_segments[idx]
                if (
                    idx + 1 < len(address_segments)
                    and len(current) <= 3
                    and current.replace('.', '').isalpha()
                    and address_segments[idx + 1].replace(' ', '').replace('-', '').isdigit()
                ):
                    merged_segments.append(f"{current} {address_segments[idx + 1]}")
                    idx += 2
                    continue
                merged_segments.append(current)
                idx += 1
            
            lines.extend(merged_segments)
        
        if not lines:
            return None
        
        bill_addr: Dict[str, str] = {}
        # QuickBooks allows up to 5 lines for address
        for idx, value in enumerate(lines[:5]):
            bill_addr[f"Line{idx + 1}"] = value
        
        return bill_addr
    
    def _get_or_create_service_item(self, service_name: str) -> str:
        """
        Get or create a service item for the given service name
        Returns the item ID to use in invoice line items
        """
        # First, try to find existing item
        existing_id = self._find_item_by_name(service_name)
        if existing_id:
            return existing_id
        
        # Create new service item
        return self._create_service_item(service_name)
    
    def _find_item_by_name(self, item_name: str) -> Optional[str]:
        """Find an existing item by name"""
        try:
            # Escape single quotes for QB SQL
            safe_name = item_name.replace("'", "\\'")
            query = f"SELECT * FROM Item WHERE Name = '{safe_name}' AND Type = 'Service'"
            response = self._make_request('GET', 'query', params={'query': query})
            
            if response and response.status_code == 200:
                data = response.json()
                items = data.get('QueryResponse', {}).get('Item', [])
                if items:
                    return items[0]['Id']
        except Exception as e:
            st.warning(f"⚠️ Error searching for item '{item_name}': {str(e)}")
        
        return None
    
    def _create_service_item(self, service_name: str) -> str:
        """Create a new service item"""
        try:
            payload = {
                "Name": service_name,
                "Type": "Service",
                "IncomeAccountRef": {"value": "1"}  # Default income account
            }
            
            response = self._make_request('POST', 'item', data=payload)
            
            if response and response.status_code == 200:
                item = response.json().get("Item", {})
                return item.get("Id")
            else:
                st.warning(f"⚠️ Failed to create service item '{service_name}', using default")
                return "1"  # Fallback to default service item
                
        except Exception as e:
            st.warning(f"⚠️ Error creating service item '{service_name}': {str(e)}, using default")
            return "1"  # Fallback to default service item
    
    def _update_invoice_doc_number(self, invoice_id: str, invoice_data: dict):
        """Update invoice DocNumber to use the invoice ID"""
        try:
            # Prepare update payload with mandatory QB fields
            update_payload = {
                "Id": invoice_id,
                "SyncToken": invoice_data.get("SyncToken", "0"),
                "DocNumber": invoice_id  # Use QB ID as doc number
            }
            
            # QuickBooks requires several fields to be present on update
            for field in ["CustomerRef", "TxnDate", "Line", "DueDate", "SalesTermRef", "BillAddr"]:
                if field in invoice_data:
                    update_payload[field] = invoice_data[field]
            
            response = self._make_request('POST', 'invoice', data=update_payload)
            if not (response and response.status_code == 200):
                st.warning("⚠️ Could not update invoice number, but invoice was created successfully")
                
        except Exception as e:
            st.warning(f"⚠️ Error updating invoice DocNumber: {str(e)}")
            # Don't fail the whole process for this
    
    def _update_customer_for_company_billing(self, customer_id: str, company_name: str):
        """Update existing customer to show only company name in BILL TO (remove person name)"""
        try:
            # First get the current customer data
            response = self._make_request('GET', f'customer/{customer_id}')
            if response and response.status_code == 200:
                customer_data = response.json().get("Customer", {})
                
                # Update to show only company name in BILL TO
                update_payload = {
                    "Id": customer_id,
                    "SyncToken": customer_data.get("SyncToken", "0"),
                    "DisplayName": company_name,  # Use company name as DisplayName
                    "CompanyName": company_name,
                    # Explicitly remove GivenName and FamilyName to prevent showing person name
                    "GivenName": "",
                    "FamilyName": ""
                }
                
                # Keep essential fields but do NOT copy BillAddr from customer
                # We want the invoice BillAddr to override the customer's default
                if "PrimaryEmailAddr" in customer_data:
                    update_payload["PrimaryEmailAddr"] = customer_data["PrimaryEmailAddr"]
                
                # Update customer
                update_response = self._make_request('POST', 'customer', data=update_payload)
                if update_response and update_response.status_code == 200:
                    st.info(f"✅ Updated customer to show only company name in BILL TO")
                
        except Exception as e:
            st.warning(f"⚠️ Error updating customer for company billing: {str(e)}")
            # Don't fail the whole process for this
    
    def _clear_customer_billing_address(self, customer_id: str):
        """Clear customer's default billing address to prevent invoice conflicts"""
        try:
            # Get current customer data
            response = self._make_request('GET', f'customer/{customer_id}')
            if response and response.status_code == 200:
                customer_data = response.json().get("Customer", {})
                
                # Remove BillAddr from customer if it exists
                if "BillAddr" in customer_data:
                    update_payload = {
                        "Id": customer_id,
                        "SyncToken": customer_data.get("SyncToken", "0"),
                        "DisplayName": customer_data.get("DisplayName"),
                        "CompanyName": customer_data.get("CompanyName"),
                        "GivenName": "",
                        "FamilyName": ""
                    }
                    
                    # Keep essential fields but remove BillAddr
                    if "PrimaryEmailAddr" in customer_data:
                        update_payload["PrimaryEmailAddr"] = customer_data["PrimaryEmailAddr"]
                    
                    # Update customer without BillAddr
                    update_response = self._make_request('POST', 'customer', data=update_payload)
                    if update_response and update_response.status_code == 200:
                        st.info(f"✅ Cleared customer default address to prevent conflicts")
                
        except Exception as e:
            st.warning(f"⚠️ Error clearing customer billing address: {str(e)}")
            # Don't fail the whole process for this

    def create_customer(self, first_name: str, last_name: str, email: str, company_name: str = None, cc_email: str = None) -> Optional[str]:
        """
        Create or find a customer
        FIX #3: Changed DisplayName logic to use client name instead of company name
        This ensures the email says "Dear [Client Name]" not "Dear [Company Name]"
        
        SOLUTION B: Add SecondaryEmailAddr if cc_email is provided
        This ensures CC is sent even if QuickBooks ignores the send endpoint CC parameter
        """
        # FIX #3: Always use person's name as DisplayName for email greeting
        # Company name will be stored in CompanyName field and shown in Bill To address
        display_name = f"{first_name} {last_name}"
        
        # Try to find existing customer by email
        existing_id = self._find_customer_by_email(email)
        if existing_id:
            # If customer exists and we have CC email, try to update SecondaryEmailAddr
            if cc_email and cc_email.strip():
                self._update_customer_secondary_email(existing_id, cc_email)
            return existing_id

        # Create new customer
        payload = {
            "DisplayName": display_name,  # FIX #3: Always use person's name for email greeting
            "GivenName": first_name,
            "FamilyName": last_name,
            "PrimaryEmailAddr": {"Address": email}
        }
        
        # Add company name to CompanyName field (will show in Bill To address)
        if company_name:
            payload["CompanyName"] = company_name
        
        # SOLUTION B: Add SecondaryEmailAddr if CC email is provided
        # QuickBooks will automatically CC secondary email when sending invoices
        if cc_email and cc_email.strip():
            payload["SecondaryEmailAddr"] = {"Address": cc_email.strip()}

        response = self._make_request('POST', 'customer', data=payload)
        
        if response and response.status_code in [200, 201]:
            customer = response.json().get("Customer", {})
            person_name = f"{first_name} {last_name}"
            st.success(f"✅ Customer created: {person_name}")
            return customer.get("Id")
        
        # Specific error handling for duplicates (if search failed but they exist)
        if response and "Duplicate Name Exists Error" in response.text:
            st.warning("⚠️ Customer name already exists. Trying to find by name...")
            return self._find_customer_by_display_name(display_name)

        st.error(f"Error creating customer: {response.text if response else 'No response'}")
        return None

    def _find_customer_by_email(self, email: str) -> Optional[str]:
        query = f"SELECT * FROM Customer WHERE PrimaryEmailAddr = '{email}'"
        response = self._make_request('GET', 'query', params={'query': query})
        
        if response and response.status_code == 200:
            data = response.json()
            customers = data.get('QueryResponse', {}).get('Customer', [])
            if customers:
                return customers[0]['Id']
        return None

    def _find_customer_by_display_name(self, display_name: str) -> Optional[str]:
        # Escape single quotes for QB SQL
        safe_name = display_name.replace("'", "\\'")
        query = f"SELECT * FROM Customer WHERE DisplayName = '{safe_name}'"
        response = self._make_request('GET', 'query', params={'query': query})
        
        if response and response.status_code == 200:
            data = response.json()
            customers = data.get('QueryResponse', {}).get('Customer', [])
            if customers:
                return customers[0]['Id']
        return None
    
    def _update_customer_secondary_email(self, customer_id: str, cc_email: str) -> bool:
        """
        Update existing customer's SecondaryEmailAddr for Solution B
        This ensures CC is sent even if QuickBooks ignores send endpoint CC parameter
        """
        try:
            # First, get the customer to retrieve SyncToken (required for updates)
            response = self._make_request('GET', f'customer/{customer_id}')
            if not response or response.status_code != 200:
                return False
            
            customer = response.json().get("Customer", {})
            sync_token = customer.get("SyncToken", "0")
            
            # Update with SecondaryEmailAddr
            payload = {
                "Id": customer_id,
                "SyncToken": sync_token,
                "SecondaryEmailAddr": {"Address": cc_email.strip()},
                "sparse": True  # Only update specified fields
            }
            
            update_response = self._make_request('POST', 'customer', data=payload)
            return update_response and update_response.status_code in [200, 201]
        except Exception as e:
            st.warning(f"Could not update customer secondary email: {e}")
            return False

    def create_invoice(self, customer_id: str, first_name: str, last_name: str, 
                      email: str, company_name: str = None, client_address: str = None,
                      contract_amount: str = "0", description: str = "Contract Services",
                      line_items: list = None, payment_terms: str = "Due in Full",
                      enable_payment_link: bool = True, invoice_date: str = None, cc_email: str = None) -> Optional[str]:
        """
        Create an invoice for a customer
        FIX #1: Improved BillAddr handling to ensure it shows in PDF
        FIX #2: Fixed DueDate to respect selected invoice date
        """
        
        # Build line items
        lines = []
        
        if line_items:
            for item in line_items:
                item_type = item.get('type', 'Service')
                item_description = item.get('description', item_type)
                line_description = item.get('line_description', '')
                quantity = float(item.get('quantity', 1) or 1)
                unit_price = float(item.get('unit_price', item.get('amount', 0)) or 0)
                amount = quantity * unit_price
                
                if amount < 0:
                    # Use DiscountLineDetail so QuickBooks shows it beneath TAX
                    discount_line = {
                        "Amount": abs(amount),
                        "DetailType": "DiscountLineDetail",
                        "Description": line_description or item_description or item_type,
                        "DiscountLineDetail": {
                            "PercentBased": False
                        }
                    }
                    lines.append(discount_line)
                else:
                    # Get or create service item
                    service_item_id = self._get_or_create_service_item(item_type)
                    
                    line_item = {
                        "Amount": amount,
                        "DetailType": "SalesItemLineDetail",
                        "SalesItemLineDetail": {
                            "ItemRef": {"value": service_item_id},
                            "Qty": quantity,
                            "UnitPrice": unit_price
                        }
                    }
                    
                    # Add description if provided
                    if line_description:
                        line_item["Description"] = line_description
                    
                    lines.append(line_item)
        else:
            # Fallback to contract amount if no line items
            if contract_amount and float(contract_amount.replace('$', '').replace(',', '')) > 0:
                amount = float(contract_amount.replace('$', '').replace(',', ''))
                lines.append({
                    "Amount": amount,
                    "DetailType": "SalesItemLineDetail",
                    "SalesItemLineDetail": {
                        "ItemRef": {"value": "1"},
                        "Qty": 1,
                        "UnitPrice": amount
                    },
                    "Description": description
                })
        
        # Handle invoice date
        if invoice_date:
            try:
                txn_date = invoice_date.strftime("%Y-%m-%d") if hasattr(invoice_date, 'strftime') else str(invoice_date)
            except:
                txn_date = datetime.now().strftime("%Y-%m-%d")
        else:
            txn_date = datetime.now().strftime("%Y-%m-%d")

        # Build invoice data structure
        # Note: Don't set TotalAmt - let QuickBooks calculate it automatically
        # Note: Don't set EmailStatus - it conflicts with ToBeEmailed
        invoice_data = {
            "CustomerRef": {"value": customer_id},
            "TxnDate": txn_date,
            "Line": lines
        }
        
        # FIX #1: Improved BillAddr handling to ensure it shows in PDF
        bill_addr = self._parse_bill_address(company_name, client_address)
        if bill_addr:
            invoice_data["BillAddr"] = bill_addr
        
        # FIX #2: Set payment terms - don't set both DueDate and SalesTermRef together
        # QuickBooks will calculate DueDate from SalesTermRef, or we can set DueDate directly
        if payment_terms in ["Due on receipt", "Due in Full"]:
            # For immediate payment, set DueDate directly (don't set SalesTermRef)
            invoice_data["DueDate"] = txn_date
        elif payment_terms:
            # For Net 15, Net 30, etc., set the payment term and let QB calculate DueDate
            invoice_data["SalesTermRef"] = {
                "value": self._get_payment_term_id(payment_terms)
            }
        else:
            # Default to due on receipt with invoice date
            invoice_data["DueDate"] = txn_date
        
        payload = invoice_data
        
        # SOLUTION A: Don't set BillEmail in invoice creation
        # QuickBooks ignores CC/BCC query parameters when BillEmail is set in the invoice
        # By removing BillEmail, we force QuickBooks to use ONLY the sendTo parameter
        # This ensures CC/BCC in the send endpoint are respected
        # The email will be sent via: /invoice/{id}/send?sendTo=email&cc=cc_email
        
        # Don't set ToBeEmailed - we'll send manually via send endpoint with CC
        # This gives us control over CC/BCC addresses
        
        # Use standard minorversion (no special requirements since CC is in send endpoint)
        params = {"minorversion": "8"}
        response = self._make_request('POST', 'invoice', data=payload, params=params)
        
        if response is None:
            st.error("❌ No response received from QuickBooks API - possible network or timeout issue")
            return None
            
        if response.status_code == 200:
            try:
                invoice_resp = response.json()
                invoice_id = invoice_resp.get("Invoice", {}).get("Id")
                self._update_invoice_doc_number(invoice_id, invoice_resp.get("Invoice", {}))
                success_msg = f"✅ Invoice created successfully (ID: {invoice_id})"
                if cc_email and cc_email.strip():
                    success_msg += f" (CC: {cc_email})"
                st.success(success_msg)
                return invoice_id
            except Exception as e:
                st.error(f"❌ Error parsing invoice response: {str(e)}")
                return None
        else:
            st.error(f"❌ Invoice creation failed - Status: {response.status_code}")
            try:
                error_text = response.text
                st.error(f"Full Response: {error_text}")
                
                # Try to parse and show detailed error
                try:
                    error_json = response.json()
                    if "Fault" in error_json:
                        errors = error_json.get("Fault", {}).get("Error", [])
                        for err in errors:
                            st.error(f"Error Code: {err.get('code', 'Unknown')}")
                            st.error(f"Error Message: {err.get('Message', 'Unknown')}")
                            detail = err.get('Detail', '')
                            st.error(f"Error Detail: {detail}")
                            
                            # Highlight BillEmailCc if mentioned
                            if "BillEmailCc" in detail or "BillEmail" in detail:
                                st.warning("💡 Issue might be with BillEmail or BillEmailCc structure. Check the payload above.")
                except:
                    pass
            except Exception as e:
                st.error(f"Could not parse error response: {e}")
            return None

    def send_invoice(self, invoice_id: str, email: str, cc_email: str = None, bcc_email: str = None) -> bool:
        """
        Send the created invoice via email using the SEND endpoint.
        If CC email is provided, send two separate emails (one to client, one to salesman).
        This is more reliable than using CC functionality which QuickBooks doesn't always respect.
        
        Endpoint format:
        POST /v3/company/{companyId}/invoice/{invoiceId}/send?sendTo=<email>
        """
        url = f"{self.base_url}/v3/company/{self.company_id}/invoice/{invoice_id}/send"
        
        if not self.access_token:
            self.authenticate()
            
        headers = self._get_headers()
        # The send endpoint expects a specific content-type or no body
        headers["Content-Type"] = "application/octet-stream"
        
        # Send email to primary recipient (client)
        primary_sent = False
        try:
            params = {
                "sendTo": email,
                "minorversion": "8"
            }
            response = self.session.post(url, headers=headers, params=params)
            if response.status_code == 200:
                primary_sent = True
            else:
                st.warning(f"⚠️ Failed to send invoice to client ({email}): {response.status_code}")
                if response.text:
                    st.warning(f"Error details: {response.text[:200]}")
        except Exception as e:
            st.warning(f"⚠️ Error sending invoice to client: {str(e)}")
        
        # If CC email is provided, send a second email to the salesman
        cc_sent = False
        if cc_email and cc_email.strip():
            try:
                params = {
                    "sendTo": cc_email.strip(),
                    "minorversion": "8"
                }
                response = self.session.post(url, headers=headers, params=params)
                if response.status_code == 200:
                    cc_sent = True
                else:
                    st.warning(f"⚠️ Failed to send invoice to salesman ({cc_email}): {response.status_code}")
                    if response.text:
                        st.warning(f"Error details: {response.text[:200]}")
            except Exception as e:
                st.warning(f"⚠️ Error sending invoice to salesman: {str(e)}")
        
        # Report results
        if primary_sent:
            if cc_email and cc_email.strip():
                if cc_sent:
                    st.success(f"📧 Invoice sent to {email} and {cc_email}")
                else:
                    st.success(f"📧 Invoice sent to {email} (salesman email failed)")
            else:
                st.success(f"📧 Invoice sent to {email}")
            return True
        else:
            return False

    def create_and_send_invoice(self, first_name: str, last_name: str, email: str, company_name: str = None,
                              client_address: str = None, contract_amount: str = "0", description: str = "Contract Services",
                              line_items: list = None, payment_terms: str = "Due in Full",
                              enable_payment_link: bool = True, invoice_date: str = None, cc_email: str = None) -> Tuple[bool, str]:
        """
        Orchestrator: Create Customer -> Create Invoice -> Send Invoice
        
        CC Email Implementation:
        - Sends two separate emails: one to client email, one to salesman email
        - This is more reliable than CC functionality which QuickBooks doesn't always respect
        """
        
        customer_id = self.create_customer(first_name, last_name, email, company_name, cc_email=cc_email)
        if not customer_id:
            return False, "Failed to create or find customer."
            
        invoice_id = self.create_invoice(customer_id, first_name, last_name, email, company_name, client_address,
                                       contract_amount, description, line_items, payment_terms, enable_payment_link, invoice_date, cc_email=cc_email)
        
        if not invoice_id:
            return False, "Failed to create invoice."
        
        # STEP 2: Send invoice via SEND endpoint
        # If CC email is provided, send two separate emails (one to client, one to salesman)
        sent = self.send_invoice(invoice_id, email, cc_email=cc_email)
        msg = f"Invoice {invoice_id} created successfully!"
        if sent:
            if cc_email and cc_email.strip():
                msg += f" Sent via email to {email} and {cc_email}."
            else:
                msg += f" Sent via email to {email}."
        else:
            msg += " (Email sending failed, but invoice exists)."
            
        return True, msg

# --- Helper functions compatible with quickbooks_form.py ---

def load_quickbooks_credentials() -> Dict[str, str]:
    """Load credentials from secrets.toml"""
    if 'quickbooks' not in st.secrets:
        st.error("QuickBooks config not found in secrets.toml")
        return {}
    return st.secrets['quickbooks']

def setup_quickbooks_oauth() -> str:
    return "Please check your Streamlit Secrets configuration."

def verify_production_credentials(api) -> Tuple[bool, str]:
    """Verify connection by checking a simple endpoint"""
    st.info("🔍 Testing connection to QuickBooks Production...")
    res = api._make_request('GET', 'preferences')
    
    if res and res.status_code == 200:
        return True, "✅ Connection successful! Credentials are valid."
    
    code = res.status_code if res else "Error"
    
    if code == 401:
        return False, "❌ Error 401: Unauthorized. Check your Client ID/Secret and if keys are for Production."
    if code == 403:
        return False, "❌ Error 403: Access Denied. App may not have permission."
        
    return False, f"❌ Connection failed. Status: {code}"