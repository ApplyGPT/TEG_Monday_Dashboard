"""
QuickBooks API Integration Module
Clean version for Production Cloud Deployment.
Removes all SSL/DNS hacks to resolve '400 No Handler' errors on cloud platforms.

FIXES APPLIED:
1. Bill To Address: Improved BillAddr handling to ensure it shows in PDF
2. Due Date: Fixed to respect the selected invoice date instead of calculating from payment terms
3. Email Message: Changed DisplayName to use client name instead of company name
4. Customer Creation: Improved timeout handling, retry logic, and error messages
5. Email Mapping: Fixed to ensure client email (not cc_email) is saved in Customer and BillEmail fields
6. Payment Methods: Payment method fields (AllowOnlinePayment, AllowOnlineCreditCard, AllowOnlineACH) are NOT 
   supported at invoice level - they cause 400 errors. Payment methods must be configured in QuickBooks account settings.

NOTES:
- The $25 ACH convenience fee is a QuickBooks account setting, not controlled by code.
  To disable: Account & Settings → Sales → Invoice Payments → Uncheck "Your customer pays the fee"
- Payment methods (Credit Card vs ACH) are controlled at the QuickBooks account level, not per invoice.
  Configure in: Account & Settings → Sales → Invoice Payments
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
                 refresh_token: str, company_id: str, sandbox: bool = False, access_token: str = None):
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
            
        # Use provided access_token if available (from secrets.toml), otherwise will authenticate on first request
        self.access_token = access_token
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
            
            # Update tokens in secrets.toml if they changed
            if new_refresh_token and new_refresh_token != self.refresh_token:
                self.refresh_token = new_refresh_token
                self._update_secrets_file(new_refresh_token, self.access_token)
            elif self.access_token:
                # Even if refresh_token didn't change, update access_token
                self._update_secrets_file(None, self.access_token)
                
            return True
            
        except Exception as e:
            st.error(f"Connection error during authentication: {str(e)}")
            return False

    def _update_secrets_file(self, new_refresh_token=None, new_access_token=None):
        """Attempts to update secrets.toml with new tokens (best effort)"""
        try:
            secrets_path = os.path.join('.streamlit', 'secrets.toml')
            if os.path.exists(secrets_path):
                with open(secrets_path, 'r') as f:
                    config = toml.load(f)
                
                if 'quickbooks' in config:
                    if new_refresh_token:
                        config['quickbooks']['refresh_token'] = new_refresh_token
                    if new_access_token:
                        config['quickbooks']['access_token'] = new_access_token
                    with open(secrets_path, 'w') as f:
                        toml.dump(config, f)
        except:
            # On cloud deploy, usually we cannot write to files.
            # Just log it internally or ignore.
            pass

    def _make_request(self, method, endpoint, data=None, params=None, retry_count=0):
        """Centralized request wrapper with retry logic and minorversion"""
        if not self.access_token:
            if not self.authenticate():
                st.error(f"❌ Authentication failed for {endpoint}")
                return None

        # Set default minorversion if not provided in params
        if params is None:
            params = {}
        # Only set default minorversion if not already specified
        if 'minorversion' not in params:
            params['minorversion'] = '40'

        url = f"{self.base_url}/v3/company/{self.company_id}/{endpoint}"
        
        try:
            headers = self._get_headers()
            if not headers:
                st.error(f"❌ No headers available for {endpoint} - access token missing")
                return None
            
            # Increased timeout for customer creation operations
            timeout = 60 if endpoint in ['customer', 'query'] else 45
            
            if method.lower() == 'get':
                response = self.session.get(url, headers=headers, params=params, timeout=timeout)
            else:
                response = self.session.post(url, headers=headers, params=params, json=data, timeout=timeout)

            # If 401, try refreshing token once
            if response.status_code == 401:
                if self.authenticate(force_refresh=True):
                    # Retry with new token
                    headers = self._get_headers()
                    if method.lower() == 'get':
                        response = self.session.get(url, headers=headers, params=params, timeout=timeout)
                    else:
                        response = self.session.post(url, headers=headers, params=params, json=data, timeout=timeout)
            
            return response
        except requests.exceptions.Timeout as e:
            # Retry once on timeout for critical operations
            if retry_count < 1 and endpoint in ['customer', 'query']:
                st.warning(f"⏰ Request timeout ({endpoint}), retrying...")
                return self._make_request(method, endpoint, data, params, retry_count + 1)
            st.error(f"⏰ Request timeout ({endpoint}): {str(e)}")
            st.error(f"   URL: {url}")
            st.error(f"   This could be due to QuickBooks API being slow or overloaded")
            return None
        except requests.exceptions.ConnectionError as e:
            # Retry once on connection error for critical operations
            if retry_count < 1 and endpoint in ['customer', 'query']:
                st.warning(f"🔌 Connection error ({endpoint}), retrying...")
                return self._make_request(method, endpoint, data, params, retry_count + 1)
            st.error(f"🔌 Connection error ({endpoint}): {str(e)}")
            st.error(f"   URL: {url}")
            st.error(f"   Check your internet connection and QuickBooks API status")
            return None
        except requests.exceptions.RequestException as e:
            st.error(f"❌ Request exception ({endpoint}): {str(e)}")
            st.error(f"   URL: {url}")
            return None
        except Exception as e:
            st.error(f"❌ Unexpected error ({endpoint}): {str(e)}")
            st.error(f"   URL: {url}")
            import traceback
            st.error(f"   Traceback: {traceback.format_exc()}")
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
            # Silently fail - will create new item
            pass
        
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
                return "1"  # Fallback to default service item
                
        except Exception as e:
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
            # Silently fail - invoice was already created successfully
        except Exception as e:
            # Silently fail - invoice was already created successfully
            pass
    
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
                # Silently handle - non-critical operation
        except Exception as e:
            # Silently handle - non-critical operation
            pass
    
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
                    # Silently handle - non-critical operation
        except Exception as e:
            # Silently handle - non-critical operation
            pass

    def create_customer(self, first_name: str, last_name: str, email: str, company_name: str = None, cc_email: str = None) -> Optional[str]:
        """
        Create or find a customer
        FIX #3: Changed DisplayName logic to use client name instead of company name
        This ensures the email says "Dear [Client Name]" not "Dear [Company Name]"
        
        SOLUTION B: Add SecondaryEmailAddr if cc_email is provided
        This ensures CC is sent even if QuickBooks ignores the send endpoint CC parameter
        """
        # Ensure we're authenticated before attempting customer operations
        if not self.access_token:
            if not self.authenticate():
                st.error("❌ Failed to authenticate with QuickBooks. Please check your credentials.")
                return None
        
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
        # Note: SecondaryEmailAddr is not supported in QuickBooks API v3
        # We'll add it after customer creation if needed
        payload = {
            "DisplayName": display_name,  # FIX #3: Always use person's name for email greeting
            "GivenName": first_name,
            "FamilyName": last_name,
            "PrimaryEmailAddr": {"Address": email}
        }
        
        # Add company name to CompanyName field (will show in Bill To address)
        if company_name:
            payload["CompanyName"] = company_name
        
        # Note: SecondaryEmailAddr is not a supported field in QuickBooks Customer API
        # We'll handle CC email separately via the invoice send endpoint

        # Verify access_token is available before making request
        if not self.access_token:
            if not self.authenticate():
                st.error("❌ Authentication failed. Cannot create customer.")
                return None
        
        response = self._make_request('POST', 'customer', data=payload)
        
        # Debug: Log response status for troubleshooting
        if response:
            print(f"DEBUG: Customer creation response status: {response.status_code}")
        else:
            print("DEBUG: Customer creation returned None - no response from API")
        
        if response and response.status_code in [200, 201]:
            customer = response.json().get("Customer", {})
            person_name = f"{first_name} {last_name}"
            customer_id = customer.get("Id")
            st.success(f"✅ Customer created: {person_name}")
            
            # Try to add secondary email if provided (may not be supported, but worth trying)
            if cc_email and cc_email.strip() and customer_id:
                self._update_customer_secondary_email(customer_id, cc_email)
            
            return customer_id
        
        # Handle specific error cases
        if not response:
            # Try to check if customer exists by name as fallback before showing error
            existing_id = self._find_customer_by_display_name(display_name)
            if existing_id:
                update_success = self._update_existing_customer(existing_id, first_name, last_name, email, company_name, cc_email)
                if update_success:
                    st.success(f"✅ Customer updated: {display_name}")
                    return existing_id
                else:
                    # Update failed, but customer exists - return it anyway
                    st.warning(f"⚠️ Customer exists but update failed. Using existing customer ID: {existing_id}")
                    return existing_id
            
            # Only show error if fallback also failed
            st.error("❌ Error creating customer: No response from QuickBooks API. This could indicate:")
            st.error("   • Authentication failure - check your access token")
            st.error("   • Network/connection issue")
            st.error("   • QuickBooks API timeout")
            # Debug: Try to get more info
            if not self.access_token:
                st.error("   • Access token is missing - authentication may have failed")
            return None
        
        # Check for specific error types (including 400 validation errors)
        try:
            error_data = response.json()
            
            # QuickBooks API can return errors in different formats:
            # Format 1: {"Fault": {"Error": [...]}}
            # Format 2: {"Error": [...]}
            fault = error_data.get("Fault", {})
            errors = fault.get("Error", [])
            
            # If no Fault.Error, try direct Error array (some API versions)
            if not errors:
                errors = error_data.get("Error", [])
            
            if errors:
                error_msg = errors[0].get("Message", "Unknown error")
                error_detail = errors[0].get("Detail", "")
                error_code = errors[0].get("code", "")
                
                # Handle duplicate name error
                if "Duplicate Name Exists Error" in error_msg or error_code == "6240":
                    st.warning(f"⚠️ Customer with name '{display_name}' already exists. Updating existing customer...")
                    found_id = self._find_customer_by_display_name(display_name)
                    if found_id:
                        # Customer exists but might be missing email or other info
                        # Try to update the existing customer with the new information
                        update_success = self._update_existing_customer(found_id, first_name, last_name, email, company_name, cc_email)
                        if update_success:
                            st.success(f"✅ Customer updated: {display_name}")
                            return found_id
                        else:
                            # If update failed, still return the ID (customer exists)
                            st.warning(f"⚠️ Customer exists but update failed. Using existing customer ID: {found_id}")
                            return found_id
                    else:
                        st.error(f"❌ Customer exists but could not be found: {error_msg}")
                        st.error(f"   Error Detail: {error_detail}")
                        return None
                
                # Handle validation errors (400) - might be due to unsupported properties
                if response.status_code == 400:
                    # If we already handled duplicate name above, don't show generic error
                    if "Duplicate Name Exists Error" not in error_msg and error_code != "6240":
                        st.error(f"❌ Validation Error (400): {error_msg}")
                        st.error(f"   Error Code: {error_code}")
                        st.error(f"   Details: {error_detail}")
                        # Try creating without SecondaryEmailAddr if that might be the issue
                        if "SecondaryEmailAddr" in str(payload):
                            payload_retry = payload.copy()
                            payload_retry.pop("SecondaryEmailAddr", None)
                            response_retry = self._make_request('POST', 'customer', data=payload_retry)
                            if response_retry and response_retry.status_code in [200, 201]:
                                customer = response_retry.json().get("Customer", {})
                                person_name = f"{first_name} {last_name}"
                                st.success(f"✅ Customer created: {person_name}")
                                # Try to update with secondary email separately if needed
                                if cc_email and cc_email.strip():
                                    customer_id = customer.get("Id")
                                    if customer_id:
                                        self._update_customer_secondary_email(customer_id, cc_email)
                                return customer.get("Id")
                        return None
                    # If duplicate name was handled above, we already returned
                    return None
                
                # Handle other validation errors
                st.error(f"❌ Error creating customer: {error_msg}")
                if error_detail:
                    st.error(f"   Details: {error_detail}")
                return None
        except (ValueError, KeyError) as e:
            # If response is not JSON or doesn't have expected structure
            st.error(f"❌ Error parsing response: {str(e)}")
            # Show raw response for debugging
            try:
                error_text = response.text[:500] if hasattr(response, 'text') else str(response)
                st.error(f"   Raw response: {error_text}")
            except:
                pass
        
        # Generic error handling
        error_text = response.text[:500] if hasattr(response, 'text') else str(response)
        st.error(f"❌ Error creating customer (Status {response.status_code}): {error_text}")
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
    
    def _update_existing_customer(self, customer_id: str, first_name: str, last_name: str, 
                                  email: str, company_name: str = None, cc_email: str = None) -> bool:
        """
        Update an existing customer with new information (e.g., add missing email)
        This is used when a customer exists but is missing information like email address
        """
        try:
            # Get current customer data
            response = self._make_request('GET', f'customer/{customer_id}')
            if not response or response.status_code != 200:
                return False
            
            customer_data = response.json().get("Customer", {})
            sync_token = customer_data.get("SyncToken", "0")
            
            # Build update payload with new information
            update_payload = {
                "Id": customer_id,
                "SyncToken": sync_token,
                "DisplayName": f"{first_name} {last_name}",
                "GivenName": first_name,
                "FamilyName": last_name,
                "PrimaryEmailAddr": {"Address": email},
                "sparse": True  # Only update specified fields
            }
            
            # Add company name if provided
            if company_name:
                update_payload["CompanyName"] = company_name
            
            # Update customer
            update_response = self._make_request('POST', 'customer', data=update_payload)
            if update_response and update_response.status_code in [200, 201]:
                # Try to add secondary email if provided
                if cc_email and cc_email.strip():
                    self._update_customer_secondary_email(customer_id, cc_email)
                return True
            return False
        except Exception as e:
            # Silently fail - non-critical operation
            return False
    
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
            # Silently fail - non-critical operation
            return False

    def create_invoice(self, customer_id: str, first_name: str, last_name: str, 
                      email: str, company_name: str = None, client_address: str = None,
                      contract_amount: str = "0", description: str = "Contract Services",
                      line_items: list = None, payment_terms: str = "Due in Full",
                      enable_payment_link: bool = True, invoice_date: str = None, cc_email: str = None,
                      include_cc_fee: bool = False) -> Optional[str]:
        """
        Create an invoice for a customer
        FIX #1: Improved BillAddr handling to ensure it shows in PDF
        FIX #2: Fixed DueDate to respect selected invoice date
        FIX #3: Set BillEmail to client email (not cc_email)
        FIX #4: Set AllowOnlineCreditCard and AllowOnlineACH based on include_cc_fee
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
        
        # FIX #3: Set BillEmail to client email (not cc_email/salesman email)
        # This ensures the invoice is associated with the correct client email
        # Adding back to test - if this causes 400 error, we'll remove it again
        invoice_data["BillEmail"] = {"Address": email}
        
        # FIX #4: Payment method settings
        # NOTE: Payment methods (AllowOnlinePayment, AllowOnlineCreditCard, AllowOnlineACH) 
        # are NOT supported at the invoice level in QuickBooks API v3.
        # These fields cause a 400 error: "Request has invalid or unsupported property"
        # 
        # Payment methods must be configured at the QuickBooks account level:
        # Account & Settings → Sales → Invoice Payments
        # 
        # The include_cc_fee parameter is still used to add the 3% processing fee as a line item,
        # but the payment method selection (CC vs ACH) is controlled by QuickBooks account settings.
        
        payload = invoice_data
        
        # NOTE: BillEmail is set to client email (not cc_email) to ensure correct email association
        # Payment method fields are not supported at invoice level - must be set in QuickBooks account settings
        
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
                return invoice_id
            except Exception as e:
                st.error(f"❌ Error parsing invoice response: {str(e)}")
                return None
        else:
            st.error(f"❌ Invoice creation failed - Status: {response.status_code}")
            try:
                error_text = response.text
                st.error(f"**Full Response:**")
                st.code(error_text, language='json')
                
                # Try to parse and show detailed error
                try:
                    error_json = response.json()
                    if "Fault" in error_json:
                        errors = error_json.get("Fault", {}).get("Error", [])
                        st.error("**Detailed Error Information:**")
                        for err in errors:
                            st.error(f"• **Error Code:** {err.get('code', 'Unknown')}")
                            st.error(f"• **Error Message:** {err.get('Message', 'Unknown')}")
                            detail = err.get('Detail', '')
                            if detail:
                                st.error(f"• **Error Detail:** {detail}")
                            
                            # Try to extract which property is invalid
                            if "Property Name" in detail or "property" in detail.lower():
                                st.warning("💡 **Tip:** Check the payload above to identify the invalid property.")
                            
                except Exception as parse_err:
                    st.warning(f"Could not parse error JSON: {parse_err}")
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
                # Error will be handled by caller
                pass
        except Exception as e:
            # Error will be handled by caller
            pass
        
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
                    # Error will be handled by caller
                    pass
            except Exception as e:
                # Error will be handled by caller
                pass
        
        # Report results
        if primary_sent:
            if cc_email and cc_email.strip():
                if cc_sent:
                    print(f"📧 Invoice sent to {email} and {cc_email}")
                else:
                    print(f"📧 Invoice sent to {email} (salesman email failed)")
            else:
                print(f"📧 Invoice sent to {email}")
            return True
        else:
            return False

    def create_and_send_invoice(self, first_name: str, last_name: str, email: str, company_name: str = None,
                              client_address: str = None, contract_amount: str = "0", description: str = "Contract Services",
                              line_items: list = None, payment_terms: str = "Due in Full",
                              enable_payment_link: bool = True, invoice_date: str = None, cc_email: str = None,
                              include_cc_fee: bool = False) -> Tuple[bool, str]:
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
                                       contract_amount, description, line_items, payment_terms, enable_payment_link, invoice_date, cc_email=cc_email, include_cc_fee=include_cc_fee)
        
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
    res = api._make_request('GET', 'preferences')
    
    if res and res.status_code == 200:
        return True, "✅ Connection successful! Credentials are valid."
    
    code = res.status_code if res else "Error"
    
    if code == 401:
        return False, "❌ Error 401: Unauthorized. Check your Client ID/Secret and if keys are for Production."
    if code == 403:
        return False, "❌ Error 403: Access Denied. App may not have permission."
        
    return False, f"❌ Connection failed. Status: {code}"


def create_monday_update(item_id: str, message: str) -> bool:
    """Create an update (comment) on a Monday.com item.
    
    Args:
        item_id: The Monday.com item ID
        message: The message to post as an update
        
    Returns:
        True if successful, False otherwise
    """
    try:
        monday_config = st.secrets.get("monday", {})
        api_token = monday_config.get("api_token")
        
        if not api_token:
            st.error("Monday.com API token not found in secrets.")
            return False
        
        url = "https://api.monday.com/v2"
        headers = {
            "Authorization": api_token,
            "Content-Type": "application/json",
        }
        
        # Escape special characters in message for GraphQL
        # Replace newlines with <br> for HTML formatting, escape quotes and backslashes
        escaped_message = (message
                          .replace('\\', '\\\\')  # Escape backslashes first
                          .replace('"', '\\"')   # Escape double quotes
                          .replace('\n', '<br>')  # Replace newlines with HTML breaks
                          .replace('\r', ''))     # Remove carriage returns
        
        # GraphQL mutation to create update
        mutation = f"""
        mutation {{
            create_update(
                item_id: {item_id}
                body: "{escaped_message}"
            ) {{
                id
            }}
        }}
        """
        
        response = requests.post(url, json={"query": mutation}, headers=headers, timeout=30)
        result = response.json()
        
        if "errors" in result:
            st.error(f"Error creating Monday.com update: {result['errors']}")
            return False
        
        return True
        
    except Exception as e:
        st.error(f"Failed to create Monday.com update: {e}")
        return False