"""
QuickBooks API Integration Module
Handles invoice creation and sending
"""

import requests
import json
import os
import toml
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
        self.items_cache = None  # Cache for QuickBooks items
    
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
            
            # Enhanced error reporting
            if response.status_code != 200:
                try:
                    error_details = response.json()
                    error_msg = error_details.get('error_description', error_details.get('error', 'Unknown error'))
                    st.error(f"QuickBooks authentication failed: {error_msg}")
                    
                    if response.status_code == 400:
                        st.warning("⚠️ Refresh token expired or invalid. Run 'quickbooks_refresh_token.py' to get a new one.")
                    
                    return False
                except ValueError:
                    st.error(f"QuickBooks authentication failed with status code: {response.status_code}")
                    return False
            
            response.raise_for_status()
            
            auth_response = response.json()
            self.access_token = auth_response.get("access_token")
            new_refresh_token = auth_response.get("refresh_token")
            
            if not self.access_token:
                st.error("Failed to get access token from QuickBooks")
                return False
            
            # IMPORTANT: Automatically update refresh token if a new one is provided
            if new_refresh_token and new_refresh_token != self.refresh_token:
                try:
                    # Update the refresh token in memory
                    old_refresh_token = self.refresh_token
                    self.refresh_token = new_refresh_token
                    
                    # Update secrets.toml file with new refresh token
                    secrets_path = os.path.join('.streamlit', 'secrets.toml')
                    
                    # Read current secrets
                    with open(secrets_path, 'r') as f:
                        secrets_config = toml.load(f)
                    
                    # Update the refresh token
                    if 'quickbooks' in secrets_config:
                        secrets_config['quickbooks']['refresh_token'] = new_refresh_token
                        
                        # Write back to file
                        with open(secrets_path, 'w') as f:
                            toml.dump(secrets_config, f)
                        
                        st.success("🔄 QuickBooks refresh token automatically updated in secrets.toml")
                    else:
                        st.warning("⚠️ New refresh token received but could not update secrets.toml")
                        st.code(f"refresh_token = \"{new_refresh_token}\"", language="toml")
                        
                except Exception as e:
                    st.warning(f"⚠️ Could not auto-update refresh token: {str(e)}")
                    st.info("Please manually update secrets.toml with this new refresh token:")
                    st.code(f"refresh_token = \"{new_refresh_token}\"", language="toml")
                
            return True
            
        except requests.exceptions.RequestException as e:
            st.error(f"QuickBooks authentication failed: {str(e)}")
            return False
        except Exception as e:
            st.error(f"Unexpected error during QuickBooks authentication: {str(e)}")
            return False
    
    def create_customer(self, first_name: str, last_name: str, email: str, company_name: str = None) -> Optional[str]:
        """
        Create a customer in QuickBooks or use existing customer in sandbox
        
        Args:
            first_name: Customer's first name
            last_name: Customer's last name
            email: Customer's email address
            company_name: Customer's company name (optional)
            
        Returns:
            str: Customer ID if successful, None otherwise
        """
        if not self.access_token:
            if not self.authenticate():
                return None
        
        # In sandbox mode, we can't create customers, so use existing ones
        if self.sandbox:
            st.info("🔧 Sandbox Mode: Using existing customer for testing")
            return self._get_or_create_sandbox_customer(first_name, last_name, email, company_name)
        
        # First check if customer already exists
        existing_customer_id = self._find_customer_by_email(email)
        if existing_customer_id:
            # Get customer details to show the name
            existing_customer_info = self._get_customer_info(existing_customer_id)
            if existing_customer_info:
                existing_name = existing_customer_info.get('DisplayName', 'Unknown')
                input_name = f"{first_name} {last_name}"
                if existing_name != input_name:
                    st.warning(f"⚠️ Customer with email {email} already exists as '{existing_name}'")
                    st.info("💡 Invoice will be sent to the existing customer. To use a different name, use a different email or update the customer in QuickBooks.")
                else:
                    st.info(f"✅ Customer already exists with email: {email}")
            else:
                st.info(f"✅ Customer already exists with email: {email}")
            return existing_customer_id
        
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
            
            # Add company name if provided
            if company_name:
                customer_data["CompanyName"] = company_name
            
            # Customer object should be at root level, not wrapped
            payload = customer_data
            
            response = requests.post(customer_url, json=payload, headers=headers)
            
            # If we get a 401, try to refresh the token and retry
            if response.status_code == 401:
                st.info("🔄 Access token expired, refreshing...")
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.post(customer_url, json=payload, headers=headers)
                else:
                    st.error("Failed to refresh access token")
                    return None
            
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
    
    def _get_or_create_sandbox_customer(self, first_name: str, last_name: str, email: str, company_name: str = None) -> Optional[str]:
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
            
            # If we get a 401, try to refresh the token and retry
            if response.status_code == 401:
                st.info("🔄 Access token expired, refreshing...")
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.get(query_url, headers=headers)
                else:
                    st.error("Failed to refresh access token")
                    return None
            
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
    
    def _get_customer_info(self, customer_id: str) -> Optional[Dict]:
        """
        Get customer information by ID
        
        Args:
            customer_id: Customer ID
            
        Returns:
            Dict: Customer information if found, None otherwise
        """
        if not self.access_token:
            if not self.authenticate():
                return None
        
        try:
            customer_url = f"{self.base_url}/v3/company/{self.company_id}/customer/{customer_id}"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json",
                "Accept-Encoding": "identity"
            }
            
            response = requests.get(customer_url, headers=headers)
            
            if response.status_code == 401:
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.get(customer_url, headers=headers)
                else:
                    return None
            
            response.raise_for_status()
            customer_data = response.json()
            return customer_data.get("Customer")
            
        except Exception as e:
            return None
    
    def _find_customer_by_email(self, email: str) -> Optional[str]:
        """
        Find existing customer by email address
        
        Args:
            email: Customer's email address
            
        Returns:
            str: Customer ID if found, None otherwise
        """
        # Ensure we have a valid access token
        if not self.access_token:
            if not self.authenticate():
                return None
        
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
            
            # If we get a 401, try to refresh the token and retry
            if response.status_code == 401:
                st.info("🔄 Access token expired, refreshing...")
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.get(query_url, params=params, headers=headers)
                else:
                    st.error("Failed to refresh access token")
                    return None
            
            response.raise_for_status()
            
            query_response = response.json()
            customers = query_response.get("QueryResponse", {}).get("Customer", [])
            
            if customers:
                return customers[0].get("Id")
            
            return None
            
        except Exception as e:
            st.error(f"Error finding customer: {str(e)}")
            return None
    
    def _get_all_items(self) -> list:
        """
        Fetch all items (Service and Non-Inventory) from QuickBooks and cache them
        
        Returns:
            list: List of item dictionaries
        """
        # Return cached items if available
        if self.items_cache is not None:
            return self.items_cache
        
        if not self.access_token:
            if not self.authenticate():
                return []
        
        try:
            query_url = f"{self.base_url}/v3/company/{self.company_id}/query"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json",
                "Accept-Encoding": "identity"
            }
            
            # Query for all active items (includes Service and NonInventory)
            query = "SELECT * FROM Item WHERE Active = true"
            params = {"query": query}
            
            response = requests.get(query_url, params=params, headers=headers)
            
            if response.status_code == 401:
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.get(query_url, params=params, headers=headers)
                else:
                    return []
            
            response.raise_for_status()
            data = response.json()
            items = data.get("QueryResponse", {}).get("Item", [])
            
            # Cache the items
            self.items_cache = items
            
            return items
            
        except Exception as e:
            print(f"Error fetching items: {str(e)}")
            return []
    
    def _get_default_income_account(self) -> str:
        """Get the default income account ID for services"""
        try:
            # Query for income accounts - try multiple types
            queries = [
                "SELECT * FROM Account WHERE AccountType = 'Income' MAXRESULTS 5",
                "SELECT * FROM Account WHERE Classification = 'Revenue' MAXRESULTS 5"
            ]
            
            for query in queries:
                url = f"{self.base_url}/v3/company/{self.company_id}/query?query={query}"
                headers = {"Authorization": f"Bearer {self.access_token}", "Accept": "application/json"}
                
                response = requests.get(url, headers=headers)
                if response.status_code == 200:
                    data = response.json()
                    accounts = data.get("QueryResponse", {}).get("Account", [])
                    if accounts:
                        # Try to find "Sales" or "Service" income account
                        for acc in accounts:
                            acc_name = acc.get("Name", "").lower()
                            if "sales" in acc_name or "service" in acc_name or "income" in acc_name:
                                return acc.get("Id")
                        # If no match, use first available
                        return accounts[0].get("Id", "1")
            return "1"  # Fallback ID
        except Exception as e:
            print(f"Error getting income account: {e}")
            return "1"
    
    def _get_default_expense_account(self) -> str:
        """Get the default expense account ID for services"""
        try:
            # Query for expense accounts
            queries = [
                "SELECT * FROM Account WHERE AccountType = 'Cost of Goods Sold' MAXRESULTS 5",
                "SELECT * FROM Account WHERE AccountType = 'Expense' MAXRESULTS 5"
            ]
            
            for query in queries:
                url = f"{self.base_url}/v3/company/{self.company_id}/query?query={query}"
                headers = {"Authorization": f"Bearer {self.access_token}", "Accept": "application/json"}
                
                response = requests.get(url, headers=headers)
                if response.status_code == 200:
                    data = response.json()
                    accounts = data.get("QueryResponse", {}).get("Account", [])
                    if accounts:
                        # Try to find COGS or expense account
                        for acc in accounts:
                            acc_name = acc.get("Name", "").lower()
                            if "cost" in acc_name or "expense" in acc_name or "cogs" in acc_name:
                                return acc.get("Id")
                        # If no match, use first available
                        return accounts[0].get("Id", "1")
            return "1"  # Fallback ID
        except Exception as e:
            print(f"Error getting expense account: {e}")
            return "1"
    
    def _create_service_item(self, item_name: str, description: str = "") -> str:
        """
        Create a new service item in QuickBooks
        
        Args:
            item_name: Name of the item to create
            description: Optional description for the item
            
        Returns:
            str: Created item ID or "2" if creation fails
        """
        if not self.access_token:
            if not self.authenticate():
                return "2"  # Fallback to generic item
        
        try:
            url = f"{self.base_url}/v3/company/{self.company_id}/item"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json"
            }
            
            # Create service item following QuickBooks API format
            # Use account IDs from your QuickBooks company:
            # ID 1 = Services (Income)
            # ID 12 = Cost of Goods Sold (Expense)
            item_data = {
                "Name": item_name,
                "Type": "Service",
                "Active": True,
                "IncomeAccountRef": {
                    "value": "1",  # Services income account
                    "name": "Services"
                },
                "ExpenseAccountRef": {
                    "value": "12",  # Cost of Goods Sold
                    "name": "Cost of Goods Sold"
                }
            }
            
            if description:
                item_data["Description"] = description
            
            response = requests.post(url, headers=headers, json=item_data)
            
            if response.status_code == 401:
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.post(url, headers=headers, json=item_data)
                else:
                    return "2"
            
            if response.status_code == 200 or response.status_code == 201:
                result = response.json()
                created_item = result.get("Item", {})
                item_id = created_item.get("Id")
                
                if item_id:
                    # Clear cache so new item is included in future queries
                    self.items_cache = None
                    return item_id
            
            # If creation failed, fallback to item "2"
            return "2"
            
        except Exception as e:
            return "2"

    def _find_best_item_match(self, line_item_name: str, show_match_info: bool = False) -> str:
        """
        Find the best matching QuickBooks item for a line item name
        Uses exact match, then partial match, then creates new item if needed
        
        Args:
            line_item_name: Name of the line item to match
            show_match_info: Whether to show matching info in UI (default: False)
            
        Returns:
            str: QuickBooks item ID
        """
        items = self._get_all_items()
        
        if not items:
            # Create the item if we can't fetch existing items
            item_id = self._create_service_item(line_item_name)
            if show_match_info:
                st.info(f"✓ '{line_item_name}' → Created new item")
            return item_id
        
        line_item_lower = line_item_name.lower().strip()
        
        # 1. Try exact match (case insensitive)
        for item in items:
            if item.get('Active', False):
                item_name = item.get('Name', '').lower().strip()
                if item_name == line_item_lower:
                    if show_match_info:
                        st.info(f"✓ '{line_item_name}' → Exact match: '{item.get('Name')}'")
                    return item.get('Id')
        
        # 2. Try partial match - check if line item name contains item name or vice versa
        for item in items:
            if item.get('Active', False):
                item_name = item.get('Name', '').lower().strip()
                # Skip the generic "-" item for partial matching
                if item_name in ['-', '']:
                    continue
                # Check both directions
                if item_name in line_item_lower or line_item_lower in item_name:
                    if show_match_info:
                        st.info(f"✓ '{line_item_name}' → Partial match: '{item.get('Name')}'")
                    return item.get('Id')
        
        # 3. No match found - try to create the item
        if show_match_info:
            st.info(f"📝 No match found for '{line_item_name}', attempting to create...")
        
        item_id = self._create_service_item(line_item_name)
        
        # If creation was successful, return the new item ID
        if item_id != "2":
            return item_id
        
        # 4. Creation failed - find and use the "-" generic item as fallback
        for item in items:
            if item.get('Active', False):
                item_name = item.get('Name', '').strip()
                item_type = item.get('Type', '')
                # Only use "-" if it's a Service or NonInventory item, not a Category
                if item_name == '-' and item_type in ['Service', 'NonInventory']:
                    if show_match_info:
                        st.warning(f"⚠️ '{line_item_name}' → Item creation failed, using generic '-' item")
                    return item.get('Id')
        
        # 5. Last resort: Fallback to item ID "2" if "-" not found
        if show_match_info:
            st.warning(f"⚠️ '{line_item_name}' → Using default item ID 2")
        return "2"
    
    def create_invoice(self, customer_id: str, first_name: str, last_name: str, 
                     email: str, company_name: str = None, client_address: str = None,
                     contract_amount: str = "0", description: str = "Contract Services", 
                     line_items: list = None, payment_terms: str = "Due in Full", 
                     enable_payment_link: bool = True, invoice_date: str = None) -> Optional[str]:
        """
        Create an invoice in QuickBooks or simulate in sandbox
        
        Args:
            customer_id: ID of the customer
            first_name: Customer's first name
            last_name: Customer's last name
            email: Customer's email address
            company_name: Customer's company name (optional)
            client_address: Customer's billing address (optional)
            contract_amount: Contract amount (will be converted to float)
            description: Description of the service
            line_items: List of line items with type, amount, and description
            payment_terms: Payment terms for the invoice
            enable_payment_link: Whether to enable online payment link
            invoice_date: Date of the invoice
            
        Returns:
            str: Invoice ID if successful, None otherwise
        """
        # Always re-authenticate to ensure we have a fresh access token
        if not self.authenticate():
            st.error("❌ Failed to authenticate with QuickBooks")
            return None
        
        # In sandbox mode, we can't create invoices, so simulate the process
        if self.sandbox:
            st.info("🔧 Sandbox Mode: Simulating invoice creation")
            return self._simulate_invoice_creation(customer_id, first_name, last_name, email, company_name, client_address,
                                                 contract_amount, description, line_items, 
                                                 payment_terms, enable_payment_link, invoice_date)
        
        try:
            # Use line items if provided, otherwise use contract amount
            if line_items:
                total_amount = sum(item['amount'] * item.get('quantity', 1) for item in line_items)
                invoice_lines = []
                
                for item in line_items:
                    quantity = item.get('quantity', 1)
                    unit_price = item['amount']
                    line_total = unit_price * quantity
                    
                    # Get line item details
                    line_item_type = item.get('type', item.get('description', 'Service'))  # Fee type (main item name)
                    line_item_description = item.get('line_description', item.get('name', ''))  # Optional description
                    
                    # Determine what to show based on the line item type
                    if line_item_type == "Contract Services":
                        # Main contract: Use "Contract Services" as item, show custom description
                        item_id = self._find_best_item_match(line_item_type, show_match_info=False)
                        full_description = line_item_description if line_item_description else ""  # "Summer Collection"
                    elif line_item_type == "Credit Card Processing Fee":
                        # CC Fee: Use "Credit Card Processing Fee" as item, add description
                        item_id = self._find_best_item_match(line_item_type, show_match_info=False)
                        full_description = "3% processing fee for credit card payments"
                    elif line_item_type == "Credits & Discounts":
                        # Credits: Use "Credits & Discounts" as item, credit description in description field
                        item_id = self._find_best_item_match(line_item_type, show_match_info=False)
                        full_description = line_item_description if line_item_description else item.get('description', '')
                    else:
                        # Additional line items: Use fee type as item, optional description
                        item_id = self._find_best_item_match(line_item_type, show_match_info=False)
                        full_description = line_item_description if line_item_description else ""
                    
                    invoice_lines.append({
                        "Amount": line_total,
                        "DetailType": "SalesItemLineDetail",
                        "Description": full_description,  # Fee type only
                        "SalesItemLineDetail": {
                            "Qty": quantity,
                            "UnitPrice": unit_price,
                            "ItemRef": {
                                "value": item_id  # Use "-" item
                            }
                        }
                    })
            else:
                # Convert contract amount to float
                amount_str = contract_amount.replace('$', '').replace(',', '')
                amount = float(amount_str)
                total_amount = amount
                
                # Find or create item for the description (e.g., "Contract Services")
                item_id = self._find_best_item_match(description, show_match_info=False)
                
                invoice_lines = [
                    {
                        "DetailType": "SalesItemLineDetail",
                        "Amount": amount,
                        "SalesItemLineDetail": {
                            "ItemRef": {
                                "value": item_id  # Use matched or created item
                            },
                            "Qty": 1,
                            "UnitPrice": amount
                        },
                        "Description": ""  # No additional description needed
                    }
                ]
            
            invoice_url = f"{self.base_url}/v3/company/{self.company_id}/invoice"
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Accept-Encoding": "identity"  # Disable gzip to avoid decoding issues
            }
            
            # Set invoice date - this shows in the DATE column for each line
            if invoice_date:
                txn_date = invoice_date.strftime("%Y-%m-%d") if hasattr(invoice_date, 'strftime') else str(invoice_date)
            else:
                txn_date = datetime.now().strftime("%Y-%m-%d")
            
            # Create invoice data using correct format
            invoice_data = {
                "CustomerRef": {
                    "value": customer_id
                },
                "TxnDate": txn_date,
                "DueDate": txn_date,  # Due in Full means due immediately
                "Line": invoice_lines,
                "EmailStatus": "NotSet",
                "AllowOnlinePayment": enable_payment_link,
                "AllowOnlineCreditCardPayment": enable_payment_link,
                "AllowOnlineACHPayment": enable_payment_link
            }
            
            # Add company name and address to BillToAddr so it appears between person's name and address on invoice
            if company_name or client_address:
                bill_to = {}
                if company_name:
                    bill_to["CompanyName"] = company_name
                if client_address:
                    # Parse the address - assume single line format
                    bill_to["Line1"] = client_address
                invoice_data["BillToAddr"] = bill_to
            
            # Add payment terms - always set them explicitly
            if payment_terms == "Due in Full":
                # For "Due in Full", we don't set SalesTermRef (defaults to Due on Receipt)
                pass
            else:
                # For Net terms, set the appropriate term
                invoice_data["SalesTermRef"] = {
                    "value": self._get_payment_term_id(payment_terms)
                }
            
            # Invoice object should be at root level, not wrapped
            payload = invoice_data
            
            response = requests.post(invoice_url, json=payload, headers=headers)
            
            # If we get a 401, try to refresh the token and retry
            if response.status_code == 401:
                st.info("🔄 Access token expired, refreshing...")
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.post(invoice_url, json=payload, headers=headers)
                else:
                    st.error("Failed to refresh access token")
                    return None
            
            # Check response status
            if response.status_code != 200:
                st.error(f"❌ QuickBooks API Error (Status {response.status_code})")
                st.error(f"Response: {response.text[:500]}")
                return None
            
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
                                 email: str, company_name: str = None, client_address: str = None,
                                 contract_amount: str = "0", description: str = "Contract Services", 
                                 line_items: list = None, payment_terms: str = "Due in Full", 
                                 enable_payment_link: bool = True, invoice_date: str = None) -> str:
        """
        Simulate invoice creation for sandbox testing
        """
        try:
            # Calculate total amount
            if line_items:
                total_amount = sum(item['amount'] * item.get('quantity', 1) for item in line_items)
            else:
                # Convert contract amount to float
                amount_str = contract_amount.replace('$', '').replace(',', '')
                total_amount = float(amount_str)
            
            # Generate a simulated invoice ID
            simulated_invoice_id = f"SIM_{customer_id}_{int(datetime.now().timestamp())}"
            
            st.success(f"✅ Invoice simulation successful!")
            st.info(f"📋 Simulated Invoice Preview (Your Custom Template):")
            st.info(f"")
            st.info(f"   BILL TO: {first_name} {last_name}")
            if company_name:
                st.info(f"   COMPANY: {company_name}")
            if client_address:
                st.info(f"   ADDRESS: {client_address}")
            st.info(f"   EMAIL: {email}")
            st.info(f"   TERMS: {payment_terms}")
            
            if invoice_date:
                date_str = invoice_date.strftime('%m/%d/%Y') if hasattr(invoice_date, 'strftime') else str(invoice_date)
            else:
                date_str = datetime.now().strftime('%m/%d/%Y')
            st.info(f"   DATE: {date_str}")
            st.info(f"")
            
            # Display line items in Standard template format
            if line_items:
                st.info(f"   DATE        ACTIVITY                          QTY    RATE        AMOUNT")
                st.info(f"   " + "-" * 70)
                
                subtotal = 0
                for item in line_items:
                    quantity = item.get('quantity', 1)
                    unit_price = item['amount']
                    line_total = unit_price * quantity
                    subtotal += line_total
                    
                    item_name = item.get('type', item.get('description', 'Service'))
                    line_desc = item.get('line_description', '')
                    
                    # Format the display like the Standard template
                    if unit_price < 0:
                        # Credit/Discount
                        st.info(f"   {date_str}  {item_name}")
                        if line_desc:
                            st.info(f"              {line_desc}")
                        st.info(f"   {'':>10} {quantity:>3}  ${abs(unit_price):>8,.2f}  -${abs(line_total):>8,.2f}")
                    else:
                        st.info(f"   {date_str}  {item_name}")
                        if line_desc:
                            st.info(f"              {line_desc}")
                        st.info(f"   {'':>10} {quantity:>3}  ${unit_price:>8,.2f}   ${line_total:>8,.2f}")
                    st.info(f"")
                
                st.info(f"   " + "-" * 60)
                st.info(f"   SUBTOTAL: ${subtotal:>10,.2f}")
            else:
                st.info(f"   • Amount: ${total_amount:,.2f}")
                st.info(f"   • Description: {description}")
            
            st.info(f"")
            if enable_payment_link:
                st.info(f"   🔗 Online Payment: **ENABLED** - Credit Card and ACH")
            else:
                st.info(f"   Online Payment: Disabled")
            
            st.info(f"   • **Amount Due: ${total_amount:,.2f}**")
            st.info(f"   • Simulated Invoice ID: {simulated_invoice_id}")
            
            if invoice_date:
                date_str = invoice_date.strftime('%Y-%m-%d') if hasattr(invoice_date, 'strftime') else str(invoice_date)
                st.info(f"   • Invoice Date: {date_str}")
            else:
                st.info(f"   • Date: {datetime.now().strftime('%Y-%m-%d')}")
            
            st.warning("🔧 Note: This is a sandbox simulation. In production, a real invoice would be created.")
            
            return simulated_invoice_id
            
        except ValueError:
            st.error("Invalid contract amount format")
            return None
        except Exception as e:
            st.error(f"Error in invoice simulation: {str(e)}")
            return None
    
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
        return term_mapping.get(payment_terms, "2")  # Default to Net 30
    
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
            
            # If we get a 401, try to refresh the token and retry
            if response.status_code == 401:
                st.info("🔄 Access token expired, refreshing...")
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.post(send_url, headers=headers)
                else:
                    st.error("Failed to refresh access token")
                    return False
            
            response.raise_for_status()
            
            return True
            
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to send invoice: {str(e)}")
            return False
        except Exception as e:
            st.error(f"Unexpected error sending invoice: {str(e)}")
            return False
    
    def create_and_send_invoice(self, first_name: str, last_name: str, email: str, company_name: str = None,
                              client_address: str = None, contract_amount: str = "0", description: str = "Contract Services",
                              line_items: list = None, payment_terms: str = "Due in Full",
                              enable_payment_link: bool = True, invoice_date: str = None) -> Tuple[bool, str]:
        """
        Complete workflow: create customer, create invoice, and send it
        
        Args:
            first_name: Customer's first name
            last_name: Customer's last name
            email: Customer's email address
            company_name: Customer's company name (optional)
            client_address: Customer's billing address (optional)
            contract_amount: Contract amount
            description: Description of the service
            line_items: List of line items with type, amount, and description
            payment_terms: Payment terms for the invoice
            enable_payment_link: Whether to enable online payment link
            invoice_date: Date of the invoice
            
        Returns:
            Tuple[bool, str]: (success, message)
        """
        try:
            # Create or find customer
            customer_id = self.create_customer(first_name, last_name, email, company_name)
            if not customer_id:
                return False, "Failed to create or find customer"
            
            # Create invoice
            invoice_id = self.create_invoice(customer_id, first_name, last_name, email, company_name, client_address,
                                           contract_amount, description, line_items,
                                           payment_terms, enable_payment_link, invoice_date)
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
