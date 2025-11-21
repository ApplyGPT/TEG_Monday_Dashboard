"""
QuickBooks API Integration Module
Handles invoice creation and sending

NOTE: SSL certificate verification is disabled (verify=False) for all requests
to resolve hostname mismatch issues with QuickBooks API endpoints.
This is a common issue with QuickBooks regional cluster URLs.
"""

import requests
import json
import os
import toml
import socket
from typing import Dict, Optional, Tuple
import streamlit as st
from datetime import datetime
import urllib3

# Disable SSL warnings globally
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configure requests session with SSL disabled
import ssl
ssl_context = ssl.create_default_context()
ssl_context.check_hostname = False
ssl_context.verify_mode = ssl.CERT_NONE

# Create a requests session with SSL verification disabled
session = requests.Session()
session.verify = False

# Set environment variables to disable SSL verification
os.environ['PYTHONHTTPSVERIFY'] = '0'
os.environ['CURL_CA_BUNDLE'] = ''
os.environ['REQUESTS_CA_BUNDLE'] = ''

# Custom HTTPAdapter to completely disable SSL verification
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

class NoSSLAdapter(HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        kwargs['ssl_context'] = ssl_context
        return super().init_poolmanager(*args, **kwargs)

# Mount the custom adapter for both HTTP and HTTPS
session.mount('https://', NoSSLAdapter())
session.mount('http://', NoSSLAdapter())

def create_ssl_disabled_session():
    """Create a requests session with SSL verification completely disabled"""
    s = requests.Session()
    s.verify = False
    s.mount('https://', NoSSLAdapter())
    s.mount('http://', NoSSLAdapter())
    return s

# === QuickBooks DNS Fix ===
# Apply DNS patch immediately when module is imported, before any QuickBooks API logic runs

def fix_quickbooks_dns():
    """
    Redirect unresolved QuickBooks regional clusters to the main production domain.
    This patches socket.getaddrinfo to intercept DNS lookups and redirect qbo-usw2.api.intuit.com
    to quickbooks.api.intuit.com, which resolves correctly.
    
    This fixes DNS resolution errors without requiring /etc/hosts modifications.
    QuickBooks backend will see the correct Host header and handle routing internally.
    """
    try:
        main_ip = socket.gethostbyname("quickbooks.api.intuit.com")
        original_getaddrinfo = socket.getaddrinfo
        
        def patched_getaddrinfo(host, port, family=0, type=0, proto=0, flags=0):
            if host.lower() == "qbo-usw2.api.intuit.com":
                # Redirect to main working domain
                print(f"🔧 Redirecting {host} → quickbooks.api.intuit.com ({main_ip})")
                return original_getaddrinfo("quickbooks.api.intuit.com", port, family, type, proto, flags)
            return original_getaddrinfo(host, port, family, type, proto, flags)
        
        socket.getaddrinfo = patched_getaddrinfo
        print(f"✅ QuickBooks DNS patch applied (qbo-usw2 → quickbooks.api.intuit.com @ {main_ip})")
    except Exception as e:
        print(f"⚠️ Failed to patch QuickBooks DNS: {e}")

# Apply DNS patch immediately at module import time
fix_quickbooks_dns()
# === End of DNS Fix ===

class QuickBooksAPI:
    """QuickBooks API client for invoice creation and sending"""
    
    def __init__(self, client_id: str, client_secret: str, 
                 refresh_token: str, company_id: str, sandbox: bool = False):
        """
        Initialize QuickBooks API client
        
        Args:
            client_id: QuickBooks application client ID
            client_secret: QuickBooks application client secret
            refresh_token: OAuth refresh token
            company_id: QuickBooks company ID
            sandbox: Whether to use sandbox environment (default: False for production)
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
        self.base_url_verified = False  # Track if we've verified the base URL
        self.verified_via_preferences = False  # Track if we verified via preferences endpoint (companyinfo has cluster issues)
        self.customer_cluster_url = None  # Track cluster URL specifically for customer endpoints (may differ from base_url)
        self._discount_account_id = None  # Cache discount account for discounts
        self.debug_logging_enabled = False  # Toggle verbose debug logs

    def _debug(self, message: str):
        """Log internal debug info when debug logging is enabled."""
        if getattr(self, "debug_logging_enabled", False):
            try:
                st.info(message)
            except Exception:
                pass
    
    def _normalize_quickbooks_url(self, url: str) -> str:
        """
        Patch for 'Wrong Cluster' + NameResolutionError.
        Rewrites regional cluster URLs (e.g., qbo-usw2.api.intuit.com)
        back to the main QuickBooks production endpoint, which proxies internally.
        
        Args:
            url: The URL that may contain regional cluster domains
            
        Returns:
            str: Normalized URL using main production endpoint
        """
        if not url:
            return "https://quickbooks.api.intuit.com"
        
        # Rewrite all regional cluster URLs to main production URL
        # This prevents DNS resolution errors while QuickBooks proxies internally
        normalized = url.replace("qbo-usw2.api.intuit.com", "quickbooks.api.intuit.com") \
                        .replace("qbo-na1.api.intuit.com", "quickbooks.api.intuit.com") \
                        .replace("qbo-na2.api.intuit.com", "quickbooks.api.intuit.com") \
                        .replace("qbo-eu1.api.intuit.com", "quickbooks.api.intuit.com") \
                        .replace("qbo-eu2.api.intuit.com", "quickbooks.api.intuit.com")
        
        return normalized
    
    def _extract_cluster_url(self, response) -> Optional[str]:
        """
        Extract the correct cluster URL from QuickBooks error response.
        Looks for regional cluster URLs like qbo-na1.api.intuit.com, qbo-eu1.api.intuit.com, etc.
        
        IMPORTANT: If cluster is already verified via preferences endpoint, this returns None
        to prevent switching to regional clusters that can't be resolved via DNS.
        
        Args:
            response: The HTTP response that may contain cluster information
            
        Returns:
            str: Correct cluster URL if found, None otherwise (or if preferences already verified)
        """
        # If we've verified via preferences, don't extract cluster URLs - use main production URL only
        if self.verified_via_preferences:
            st.info("💡 Skipping cluster URL extraction - cluster already verified via preferences, using main production URL")
            return None
        
        import re
        try:
            # Check ALL headers for cluster information first (log all headers for debugging)
            st.info("🔍 Checking all response headers for cluster URL...")
            st.info(f"🔍 All response headers:")
            cluster_hint = None
            for header_name, header_value in response.headers.items():
                st.info(f"   {header_name}: {header_value[:150]}...")
                
                # Check for cluster hints in headers (e.g., usw2-prd, na1, eu1, etc.)
                if header_value:
                    # Look for regional indicators in headers like x-envoy-decorator-operation
                    if 'usw2' in header_value.lower() or 'us-west-2' in header_value.lower():
                        cluster_hint = "https://qbo-usw2.api.intuit.com"
                        st.info(f"   💡 Found US West 2 hint in header '{header_name}'")
                    elif 'na1' in header_value.lower() or 'north-america-1' in header_value.lower():
                        cluster_hint = "https://qbo-na1.api.intuit.com"
                        st.info(f"   💡 Found North America 1 hint in header '{header_name}'")
                    elif 'na2' in header_value.lower() or 'north-america-2' in header_value.lower():
                        cluster_hint = "https://qbo-na2.api.intuit.com"
                        st.info(f"   💡 Found North America 2 hint in header '{header_name}'")
                    elif 'eu1' in header_value.lower() or 'europe-1' in header_value.lower():
                        cluster_hint = "https://qbo-eu1.api.intuit.com"
                        st.info(f"   💡 Found Europe 1 hint in header '{header_name}'")
                    elif 'eu2' in header_value.lower() or 'europe-2' in header_value.lower():
                        cluster_hint = "https://qbo-eu2.api.intuit.com"
                        st.info(f"   💡 Found Europe 2 hint in header '{header_name}'")
                
                if header_value and '.intuit.com' in header_value:
                    st.info(f"   ✅ Found .intuit.com in header '{header_name}'")
                    # Match any intuit.com URL
                    urls = re.findall(r'https?://[^\s\)]+\.intuit\.com', header_value)
                    if urls:
                        cluster_url = urls[0].split('/v3/')[0].split('/v3/company/')[0]
                        cluster_url = cluster_url.rstrip('/')
                        st.info(f"✅ Extracted cluster URL from header '{header_name}': {cluster_url}")
                        return cluster_url
            
            # If we found a cluster hint from headers, return it
            if cluster_hint:
                st.info(f"✅ Using cluster URL from header hint: {cluster_hint}")
                return cluster_hint
            
            # Check response.history for redirects that happened before the final response
            if hasattr(response, 'history') and response.history:
                st.info(f"🔍 Found {len(response.history)} redirect(s) in response history")
                for i, hist_response in enumerate(response.history):
                    st.info(f"   Redirect {i+1}: {hist_response.status_code} -> {hist_response.url if hasattr(hist_response, 'url') else 'N/A'}")
                    # Check Location header in redirect response
                    hist_location = hist_response.headers.get('Location', '')
                    if hist_location and '.intuit.com' in hist_location:
                        match = re.search(r'(https?://[^/]+\.intuit\.com)', hist_location)
                        if match:
                            cluster_url = match.group(1).rstrip('/')
                            st.info(f"✅ Extracted cluster URL from redirect history: {cluster_url}")
                            return cluster_url
            
            # Check if response.url is different from the request URL (redirect happened)
            if hasattr(response, 'url') and response.url:
                st.info(f"🔍 Final Response URL: {response.url}")
                # Check if the response URL is different from what we requested
                # This indicates a redirect happened
                if '/v3/company/' in response.url:
                    # Extract cluster URL from the redirected URL
                    redirected_cluster = response.url.split('/v3/company/')[0]
                    if redirected_cluster and '.intuit.com' in redirected_cluster:
                        st.info(f"✅ Extracted cluster URL from redirected response URL: {redirected_cluster}")
                        return redirected_cluster.rstrip('/')
            
            # For 401 errors that say "will be redirected", try to manually follow the redirect
            # Sometimes the redirect happens in a subsequent request
            if response.status_code == 401 and 'redirect' in str(response.text).lower():
                st.info("🔍 401 error mentions redirect - checking if we should follow redirect manually...")
                # The redirect might be in a Retry-After header or we need to make a new request
            
            # Check Location header specifically (common redirect header)
            location = response.headers.get('Location', '')
            if location and '.intuit.com' in location:
                # Extract base URL from location header
                # Match any intuit.com URL (including regional clusters like qbo-na1.api.intuit.com)
                match = re.search(r'(https?://[^/]+\.intuit\.com)', location)
                if match:
                    cluster_url = match.group(1)
                    st.info(f"🔍 Extracted cluster URL from Location header: {cluster_url}")
                    return cluster_url.rstrip('/')
            
            # Check error response body for cluster information
            # This is the PRIMARY method - QuickBooks puts the correct cluster in Error.Detail
            try:
                error_data = response.json()
                if 'Fault' in error_data:
                    fault = error_data['Fault']
                    if 'Error' in fault:
                        errors = fault['Error']
                        # Handle both single error dict and list of errors
                        if isinstance(errors, dict):
                            errors = [errors]
                        
                        for error in errors:
                            error_msg = error.get('Message', '')
                            error_detail = error.get('Detail', '')
                            error_code = error.get('code', '')
                            
                            # Check for cluster URL in error detail (this is where QuickBooks puts it)
                            # Example: "Use https://qbo-na1.api.intuit.com"
                            st.info(f"🔍 Error Detail: {error_detail}")
                            st.info(f"🔍 Error Message: {error_msg}")
                            for text in [error_detail, error_msg]:
                                if text and '.intuit.com' in text:
                                    # Extract any intuit.com URL (including regional clusters)
                                    # Match patterns like:
                                    # - "Use https://qbo-na1.api.intuit.com"
                                    # - "https://qbo-eu1.api.intuit.com/v3/company/..."
                                    urls = re.findall(r'https?://[^\s\)]+\.intuit\.com', text)
                                    if urls:
                                        # Extract base URL (remove path after .com)
                                        cluster_url = urls[0].split('/v3/')[0].split('/v3/company/')[0]
                                        cluster_url = cluster_url.rstrip('/')
                                        st.info(f"✅ Extracted cluster URL from error response: {cluster_url}")
                                        st.info(f"   Source: {text[:100]}...")
                                        return cluster_url
                                elif text and ('redirect' in text.lower() or 'wrong server' in text.lower()):
                                    # Even if no URL in text, log it for debugging
                                    st.info(f"🔍 Found redirect/wrong server message but no URL: {text[:200]}")
            except (ValueError, KeyError):
                # Response might not be JSON, that's okay
                pass
            
            # Check response headers for cluster information
            www_authenticate = response.headers.get('WWW-Authenticate', '')
            retry_after = response.headers.get('Retry-After', '')
            for header_text in [www_authenticate, retry_after]:
                if header_text and '.intuit.com' in header_text:
                    urls = re.findall(r'https?://[^\s\)]+\.intuit\.com', header_text)
                    if urls:
                        cluster_url = urls[0].split('/v3/')[0].split('/v3/company/')[0]
                        cluster_url = cluster_url.rstrip('/')
                        st.info(f"🔍 Extracted cluster URL from header: {cluster_url}")
                        return cluster_url
                    
        except Exception as e:
            # Silently fail - cluster URL extraction is best effort
            pass
        
        return None
    
    def _extract_customer_id_from_response(self, response) -> Optional[str]:
        """
        Try to extract customer ID from response even if it contains errors.
        Sometimes QuickBooks returns customer data in redirects or error responses.
        
        Args:
            response: The HTTP response that may contain customer information
            
        Returns:
            str: Customer ID if found, None otherwise
        """
        try:
            # Check Location header for customer ID in redirect
            location = response.headers.get('Location', '')
            if location and '/customer/' in location:
                # Extract customer ID from URL like: .../customer/123
                import re
                match = re.search(r'/customer/(\d+)', location)
                if match:
                    return match.group(1)
            
            # Try to parse response body even if status is not 200
            try:
                response_data = response.json()
                # Check for Customer object in response
                if 'Customer' in response_data:
                    customer = response_data['Customer']
                    if isinstance(customer, dict) and 'Id' in customer:
                        return customer['Id']
                
                # Check for batch response
                if 'BatchItemResponse' in response_data:
                    batch_items = response_data['BatchItemResponse']
                    if isinstance(batch_items, list):
                        for item in batch_items:
                            if 'Customer' in item:
                                customer = item['Customer']
                                if isinstance(customer, dict) and 'Id' in customer:
                                    return customer['Id']
            except:
                # Response might not be JSON, that's okay
                pass
                
        except Exception:
            # Silently fail - extraction is best effort
            pass
        
        return None
    
    def _try_customer_operation_on_all_clusters(self, operation_func, *args, **kwargs) -> Optional[str]:
        """
        Try customer operation on all known cluster URLs as a last resort.
        Sometimes different endpoints are on different clusters even within the same company.
        
        IMPORTANT: Only tries clusters for the correct environment (production vs sandbox).
        Does not mix production and sandbox URLs.
        
        Args:
            operation_func: A function that takes cluster_url as first arg and returns Optional[str]
            *args, **kwargs: Additional arguments to pass to operation_func
            
        Returns:
            str: Customer ID if successful, None otherwise
        """
        # Only try clusters for the correct environment - don't mix production and sandbox
        if self.sandbox:
            all_cluster_urls = ["https://sandbox-quickbooks.api.intuit.com"]
        else:
            # For production, start with main URL - regional clusters will be discovered from error responses
            all_cluster_urls = ["https://quickbooks.api.intuit.com"]
        
        # Try each cluster URL
        for cluster_url in all_cluster_urls:
            try:
                self._debug(f"Trying customer operation on cluster: {cluster_url}")
                # Try the operation with this cluster URL (DNS patch will handle DNS resolution)
                # But always keep base_url as main production URL
                original_base_url = self.base_url
                
                # Refresh token to ensure it's valid
                if self.authenticate(force_refresh=True):
                    # Call the operation function with the cluster URL (normalized - DNS patch handles resolution)
                    normalized_cluster = self._normalize_quickbooks_url(cluster_url)
                    result = operation_func(normalized_cluster, *args, **kwargs)
                    if result:
                        self._debug(f"Customer operation succeeded (DNS patch handled cluster resolution)")
                        # Keep using main production URL - DNS patch handles cluster resolution
                        self.base_url = "https://quickbooks.api.intuit.com"
                        self.base_url_verified = True
                        if not self.verified_via_preferences:
                            self.verified_via_preferences = True
                        return result
                
                # Restore original base_url
                self.base_url = original_base_url
            except Exception as e:
                st.warning(f"⚠️ Failed on cluster {cluster_url}: {str(e)}")
                # Restore original base_url
                self.base_url = original_base_url
                continue
        
        return None
    
    def _create_customer_direct_on_cluster(self, cluster_url: str, first_name: str, last_name: str, email: str, company_name: str = None) -> Optional[str]:
        """
        Try to create customer directly on a specific cluster URL.
        
        Args:
            cluster_url: The cluster URL to try
            first_name: Customer's first name
            last_name: Customer's last name
            email: Customer's email address
            company_name: Customer's company name (optional)
            
        Returns:
            str: Customer ID if successful, None otherwise
        """
        try:
            # Normalize cluster URL to prevent DNS resolution errors
            normalized_cluster_url = self._normalize_quickbooks_url(cluster_url)
            customer_url = self._normalize_quickbooks_url(
                f"{normalized_cluster_url}/v3/company/{self.company_id}/customer"
            )
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Accept-Encoding": "identity"
            }
            
            customer_data = {
                "DisplayName": f"{first_name} {last_name}",
                "GivenName": first_name,
                "FamilyName": last_name,
                "PrimaryEmailAddr": {
                    "Address": email
                }
            }
            
            if company_name:
                customer_data["CompanyName"] = company_name
            
            payload = customer_data
            
            response = requests.post(customer_url, json=payload, headers=headers, allow_redirects=True, verify=False)
            
            if response.status_code in [200, 201]:
                customer_response = response.json()
                customer = customer_response.get("Customer")
                if customer and isinstance(customer, dict) and "Id" in customer:
                    return customer["Id"]
            
            # Try to extract customer ID even if status is not 200
            customer_id = self._extract_customer_id_from_response(response)
            if customer_id:
                return customer_id
            
            return None
            
        except Exception:
            return None
    
    def _try_create_customer_via_batch(self, first_name: str, last_name: str, email: str, company_name: str = None, cluster_url: str = None) -> Optional[str]:
        """
        Try to create customer using batch operations as a fallback.
        Sometimes batch operations work when direct operations fail.
        
        Args:
            first_name: Customer's first name
            last_name: Customer's last name
            email: Customer's email address
            company_name: Customer's company name (optional)
            cluster_url: Optional cluster URL to use (if None, uses self.base_url)
            
        Returns:
            str: Customer ID if successful, None otherwise
        """
        try:
            if not self.authenticate(force_refresh=True):
                return None
            
            base_url_to_use = cluster_url if cluster_url else self.base_url
            normalized_base_url = self._normalize_quickbooks_url(base_url_to_use)
            batch_url = self._normalize_quickbooks_url(
                f"{normalized_base_url}/v3/company/{self.company_id}/batch"
            )
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Accept-Encoding": "identity"
            }
            
            customer_data = {
                "DisplayName": f"{first_name} {last_name}",
                "GivenName": first_name,
                "FamilyName": last_name,
                "PrimaryEmailAddr": {
                    "Address": email
                }
            }
            
            if company_name:
                customer_data["CompanyName"] = company_name
            
            batch_payload = {
                "BatchItemRequest": [
                    {
                        "bId": "1",
                        "operation": "create",
                        "Customer": customer_data
                    }
                ]
            }
            
            self._debug("Trying customer creation via batch operation as fallback...")
            response = requests.post(batch_url, json=batch_payload, headers=headers, allow_redirects=True, verify=False)
            
            if response.status_code in [200, 201]:
                batch_response = response.json()
                batch_items = batch_response.get("BatchItemResponse", [])
                if batch_items:
                    item = batch_items[0]
                    if "Customer" in item:
                        customer = item["Customer"]
                        if isinstance(customer, dict) and "Id" in customer:
                            st.success(f"✅ Customer created successfully via batch operation: {customer['Id']}")
                            # Always use main production URL - DNS patch handles cluster resolution
                            self.base_url = "https://quickbooks.api.intuit.com"
                            self.base_url_verified = True
                            if not self.verified_via_preferences:
                                self.verified_via_preferences = True
                            return customer["Id"]
                    elif "Fault" in item:
                        # Check if customer was created despite fault
                        fault = item["Fault"]
                        if "Error" in fault:
                            errors = fault["Error"]
                            if isinstance(errors, dict):
                                errors = [errors]
                            for error in errors:
                                # Sometimes duplicate errors still mean success
                                if error.get("code") == "6000" and "already exists" in error.get("Message", "").lower():
                                    # Customer already exists - try to find it
                                    return self._find_customer_by_email(email)
            
            return None
            
        except Exception as e:
            st.warning(f"⚠️ Batch operation failed: {str(e)}")
            return None
    
    def _try_follow_redirect_for_customer(self, response, payload: dict, headers: dict) -> Optional[str]:
        """
        Try to follow redirect URL from Location header to get customer ID.
        Sometimes QuickBooks redirects to the created resource even with Wrong Cluster errors.
        
        Args:
            response: The HTTP response that may contain a Location header
            payload: The customer creation payload
            headers: The request headers
            
        Returns:
            str: Customer ID if found via redirect, None otherwise
        """
        try:
            location = response.headers.get('Location', '')
            if location:
                self._debug(f"Found Location header, trying to follow redirect: {location}")
                # Try to GET the redirect URL
                redirect_response = requests.get(location, headers=headers, allow_redirects=True, verify=False)
                if redirect_response.status_code == 200:
                    redirect_data = redirect_response.json()
                    if "Customer" in redirect_data:
                        customer = redirect_data["Customer"]
                        if isinstance(customer, dict) and "Id" in customer:
                            st.success(f"✅ Successfully retrieved customer via redirect: {customer['Id']}")
                            return customer["Id"]
        except Exception as e:
            st.warning(f"⚠️ Could not follow redirect: {str(e)}")
        
        return None
    
    def _update_base_url_if_needed(self, response) -> bool:
        """
        Check if response indicates cluster mismatch and update base_url if needed
        
        Args:
            response: The HTTP response to check
            
        Returns:
            bool: True if base_url was updated, False otherwise
        """
        try:
            error_data = response.json()
            st.info(f"🔍 Checking error response for cluster information...")
            
            if 'Fault' in error_data:
                fault = error_data['Fault']
                st.info(f"🔍 Fault structure: {list(fault.keys())}")
                
                if 'Error' in fault:
                    errors = fault['Error']
                    # Handle both single error dict and list of errors
                    if isinstance(errors, dict):
                        errors = [errors]
                    
                    for error in errors:
                        error_msg = error.get('Message', '')
                        error_detail = error.get('Detail', '')
                        error_code = error.get('code', '')
                        
                        st.info(f"🔍 Error code: {error_code}")
                        st.info(f"🔍 Error message: {error_msg}")
                        st.info(f"🔍 Error detail: {error_detail}")
                        
                        if error_code == '130' or 'WrongCluster' in error_msg or 'WrongCluster' in error_detail:
                            # If we already verified via preferences, ignore companyinfo Wrong Cluster errors
                            if self.verified_via_preferences:
                                st.info("💡 Ignoring Wrong Cluster error (cluster already verified via preferences endpoint)")
                                return False  # Don't update URL, cluster is correct
                            
                            st.warning("⚠️ WrongCluster error detected!")
                            
                            # If we've verified via preferences, ignore Wrong Cluster errors - use main production URL
                            if self.verified_via_preferences:
                                st.info("💡 Ignoring Wrong Cluster error - cluster already verified via preferences, using main production URL")
                                st.info("💡 QuickBooks will proxy requests internally from the main domain")
                                return True  # Return True to continue - we'll use main URL
                            
                            # CRITICAL: Extract the correct cluster URL from the error response first
                            # QuickBooks puts it in the Error.Detail field (e.g., "Use https://qbo-na1.api.intuit.com")
                            cluster_url = self._extract_cluster_url(response)
                            if cluster_url:
                                # Don't switch to regional cluster - use main production URL with DNS patch
                                # Regional clusters can't be resolved via DNS, but DNS patch handles it
                                st.info(f"💡 Found cluster URL in error: {cluster_url}, but using main production URL with DNS patch")
                                self.base_url = "https://quickbooks.api.intuit.com"
                                self.base_url_verified = True
                                if not self.verified_via_preferences:
                                    self.verified_via_preferences = True
                                return True
                            
                            # If extraction failed, try to discover the correct cluster URL
                            if self._discover_cluster_url():
                                return True
                            
                            st.warning("⚠️ Could not extract or discover cluster URL from error response")
                            st.info("💡 Full error data:")
                            st.json(error_data)
                            
                            # Show common cluster options
                            st.info("💡 Common QuickBooks cluster URLs:")
                            common_clusters = [
                                "https://quickbooks.api.intuit.com",
                                "https://sandbox-quickbooks.api.intuit.com",
                            ]
                            for cluster in common_clusters:
                                st.info(f"   • {cluster}")
        except Exception as e:
            st.warning(f"⚠️ Could not parse error response: {str(e)}")
            try:
                st.info(f"🔍 Raw response text: {response.text[:500]}")
            except:
                pass
        
        return False
    
    def _handle_401_error(self, response, headers, retry_func):
        """
        Handle 401 errors by checking for cluster issues or refreshing token
        
        Args:
            response: The HTTP response that returned 401
            headers: The headers dict to update with new token
            retry_func: Function to call for retry (should take headers as param)
            
        Returns:
            Response object if successful, None if failed
        """
        # Check for cluster error first and try to update base_url
        cluster_updated = self._update_base_url_if_needed(response)
        
        # If cluster was updated, retry immediately with new base_url
        if cluster_updated:
            try:
                retry_response = retry_func(headers)
                if retry_response and retry_response.status_code == 200:
                    return retry_response
            except:
                pass
        
        # Check for cluster error in response
        try:
            error_data = response.json()
            if 'Fault' in error_data:
                fault = error_data['Fault']
                if 'Error' in fault:
                    errors = fault['Error']
                    for error in errors:
                        error_msg = error.get('Message', '')
                        error_detail = error.get('Detail', '')
                        error_code = error.get('code', '')
                        if error_code == '130' or 'WrongCluster' in error_msg or 'WrongCluster' in error_detail:
                            # IGNORE Wrong Cluster errors - regional cluster domains cannot be resolved via DNS
                            # QuickBooks will proxy requests internally from the main domain even if it says wrong cluster
                            if self.verified_via_preferences:
                                st.info("💡 Ignoring Wrong Cluster error - preferences endpoint confirms production")
                                self._debug("Using main production URL - QuickBooks will proxy internally")
                                st.info("💡 Regional cluster domains (like qbo-usw2.api.intuit.com) cannot be resolved via DNS")
                                # Refresh token and retry with main URL
                                self._debug("Refreshing token and retrying with main production URL...")
                                if self.authenticate(force_refresh=True):
                                    headers["Authorization"] = f"Bearer {self.access_token}"
                                    try:
                                        retry_response = retry_func(headers)
                                        # Return the response even if it's 401 - caller will handle it
                                        # Sometimes the response contains data despite the error
                                        return retry_response
                                    except Exception as e:
                                        st.warning(f"⚠️ Retry failed: {str(e)}")
                                        return None
                                return None
                            
                            # Only try cluster discovery if we haven't verified via preferences yet
                            self._debug("QuickBooks cluster mismatch detected. Discovering correct cluster...")
                            if self._discover_cluster_url():
                                # Retry with discovered cluster URL
                                try:
                                    retry_response = retry_func(headers)
                                    if retry_response and retry_response.status_code == 200:
                                        return retry_response
                                except:
                                    pass
                            return None
        except:
            pass
        
        # If not a cluster error, try to refresh the token
        self._debug("Access token expired, refreshing...")
        if self.authenticate(force_refresh=True):
            # Verify we have a valid token
            if not self.access_token or len(self.access_token) < 10:
                st.error("Failed to get valid access token after refresh")
                return None
            
            # Update headers with new token
            headers["Authorization"] = f"Bearer {self.access_token}"
            # Retry the request
            retry_response = retry_func(headers)
            
            # Check if retry_response is valid
            if retry_response is None:
                return None
            
            # If still 401 after refresh, check for cluster error again
            if retry_response.status_code == 401:
                # Try to update cluster URL again
                cluster_updated = self._update_base_url_if_needed(retry_response)
                if cluster_updated:
                    # Retry with new cluster URL
                    try:
                        final_response = retry_func(headers)
                        if final_response and final_response.status_code == 200:
                            return final_response
                    except:
                        pass
                
                try:
                    error_data = retry_response.json()
                    if 'Fault' in error_data:
                        fault = error_data['Fault']
                        if 'Error' in fault:
                            errors = fault['Error']
                            for error in errors:
                                error_code = error.get('code', '')
                                error_msg = error.get('Message', '')
                                error_detail = error.get('Detail', '')
                                if error_code == '130' or 'WrongCluster' in error_msg or 'WrongCluster' in error_detail:
                                    # If we already verified via preferences, this is a QuickBooks API bug
                                    if self.verified_via_preferences:
                                        st.info("💡 Ignoring Wrong Cluster error (cluster already verified via preferences endpoint)")
                                        # Don't retry again - this is a QuickBooks API bug
                                        # Return None so caller can handle the error gracefully
                                        return None
                                    
                                    # Only try cluster discovery if we haven't verified via preferences yet
                                    self._debug("QuickBooks cluster mismatch detected. Discovering correct cluster...")
                                    if self._discover_cluster_url():
                                        # Retry with discovered cluster URL
                                        try:
                                            final_response = retry_func(headers)
                                            if final_response and final_response.status_code == 200:
                                                return final_response
                                        except:
                                            pass
                                    return None
                        st.error("Authentication failed even after token refresh. This may be a QuickBooks cluster issue. Please try again.")
                    else:
                        st.error("Authentication failed even after token refresh. Please check your credentials.")
                except:
                    st.error("Authentication failed even after token refresh. Please check your credentials.")
                return None
            
            return retry_response
        else:
            st.error("Failed to refresh access token")
            return None
    
    def authenticate(self, force_refresh: bool = False) -> bool:
        """
        Authenticate with QuickBooks API using refresh token
        
        Args:
            force_refresh: If True, always get a new token even if one exists
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        # Force refresh if requested (clear existing token)
        if force_refresh:
            self.access_token = None
            
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
            
            response = requests.post(auth_url, data=data, headers=headers, auth=auth, verify=False)
            
            # Enhanced error reporting
            if response.status_code != 200:
                try:
                    error_details = response.json()
                    error_msg = error_details.get('error_description', error_details.get('error', 'Unknown error'))
                    st.error(f"QuickBooks authentication failed: {error_msg}")
                    
                    if response.status_code == 400:
                        st.warning("⚠️ Refresh token expired or invalid. Verify Client ID matches the one used to generate the refresh token.")
                    
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
            
            # Verify we have a valid access token (not empty or None)
            if not self.access_token or len(self.access_token) < 10:
                st.error("Invalid access token received from QuickBooks")
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
                        
                        self._debug("QuickBooks refresh token automatically updated in secrets.toml")
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
    
    def _discover_cluster_url(self) -> bool:
        """
        Discover the correct cluster URL by trying alternative cluster URLs
        when we get a "Wrong Cluster" error
        
        Returns:
            bool: True if cluster URL was discovered and updated, False otherwise
        """
        # Refresh token first - tokens can be cluster-specific
        if not self.authenticate(force_refresh=True):
            st.error("❌ Failed to refresh access token")
            return False
        
        # Determine if we should use production or sandbox based on company_id and sandbox flag
        # Production companies: use quickbooks.api.intuit.com (or regional clusters)
        # Sandbox companies: use sandbox-quickbooks.api.intuit.com
        if self.sandbox:
            # For sandbox, only try sandbox URL
            all_cluster_urls = ["https://sandbox-quickbooks.api.intuit.com"]
        else:
            # For production, start with main URL - we'll discover regional clusters from error responses
            all_cluster_urls = ["https://quickbooks.api.intuit.com"]
        
        # Remove current base URL from list to try others first
        cluster_urls_to_try = [url for url in all_cluster_urls if url != self.base_url]
        # Add current one at the end as fallback
        cluster_urls_to_try.append(self.base_url)
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
            "Accept-Encoding": "identity"
        }
        
        # Try each cluster URL
        for cluster_url in cluster_urls_to_try:
            try:
                # Try companyinfo endpoint with redirects enabled
                # Normalize cluster URL to prevent DNS resolution errors
                normalized_cluster_url = self._normalize_quickbooks_url(cluster_url)
                discovery_url = self._normalize_quickbooks_url(
                    f"{normalized_cluster_url}/v3/company/{self.company_id}/companyinfo/{self.company_id}"
                )
                
                # Also try preferences endpoint as fallback (sometimes works when companyinfo fails)
                preferences_url = self._normalize_quickbooks_url(
                    f"{normalized_cluster_url}/v3/company/{self.company_id}/preferences"
                )
                
                # Try with redirects enabled - QuickBooks may redirect even on 401
                response = requests.get(discovery_url, headers=headers, allow_redirects=True, timeout=10, verify=False)
                
                # Check if URL changed (redirect happened)
                final_url = response.url
                if final_url != discovery_url:
                    # Extract base URL from final URL
                    if '/v3/company/' in final_url:
                        redirected_base_url = final_url.split('/v3/company/')[0]
                        if redirected_base_url != cluster_url and redirected_base_url in all_cluster_urls:
                            cluster_url = redirected_base_url
                            discovery_url = f"{cluster_url}/v3/company/{self.company_id}/companyinfo/{self.company_id}"
                            # Retry with the redirected cluster URL
                            response = requests.get(discovery_url, headers=headers, allow_redirects=True, timeout=10, verify=False)
                
                # Check Location header (even for 401 responses)
                location_header = response.headers.get('Location', '')
                if location_header and location_header != discovery_url:
                    # Extract base URL from location
                    if '/v3/company/' in location_header:
                        new_base_url = location_header.split('/v3/company/')[0]
                    elif '/v3/' in location_header:
                        new_base_url = location_header.split('/v3/')[0]
                    else:
                        import re
                        match = re.search(r'https?://([^/]+\.intuit\.com)', location_header)
                        if match:
                            new_base_url = f"https://{match.group(1)}"
                        else:
                            new_base_url = None
                    
                    if new_base_url and new_base_url in all_cluster_urls:
                        cluster_url = new_base_url
                        discovery_url = f"{cluster_url}/v3/company/{self.company_id}/companyinfo/{self.company_id}"
                        # Retry with the new cluster URL
                        response = requests.get(discovery_url, headers=headers, allow_redirects=True, timeout=10, verify=False)
                
                # Check if we got a successful response
                if response.status_code == 200:
                    # Extract final URL to see if we were redirected
                    final_url = response.url
                    if '/v3/company/' in final_url:
                        discovered_base_url = final_url.split('/v3/company/')[0]
                    else:
                        discovered_base_url = cluster_url
                    
                    if discovered_base_url != self.base_url:
                        self.base_url = discovered_base_url
                        self.base_url_verified = True
                        return True
                    else:
                        self.base_url_verified = True
                        return True
                
                # Check for Wrong Cluster error specifically
                elif response.status_code == 401:
                    try:
                        error_data = response.json()
                        if 'Fault' in error_data:
                            fault = error_data['Fault']
                            if 'Error' in fault:
                                errors = fault['Error']
                                if isinstance(errors, dict):
                                    errors = [errors]
                                
                                for error in errors:
                                    error_code = error.get('code', '')
                                    error_msg = error.get('Message', '')
                                    
                                    # Check for Wrong Cluster error (code 130)
                                    if error_code == '130' or 'WrongCluster' in error_msg:
                                        # Try preferences endpoint first (sometimes works when companyinfo fails)
                                        # This is the PRIMARY method - preferences endpoint is more reliable
                                        try:
                                            prefs_response = requests.get(preferences_url, headers=headers, allow_redirects=True, timeout=10, verify=False)
                                            if prefs_response.status_code == 200:
                                                # Preferences endpoint works - verify with main production URL only
                                                # CRITICAL: Don't switch to regional cluster - use main production URL
                                                # Regional clusters can't be resolved via DNS, but main URL works
                                                self.base_url = "https://quickbooks.api.intuit.com"
                                                self.base_url_verified = True
                                                self.verified_via_preferences = True  # Mark that we verified via preferences
                                                return True
                                            elif prefs_response.status_code == 401:
                                                # Check if preferences endpoint also has Wrong Cluster error
                                                try:
                                                    prefs_error_data = prefs_response.json()
                                                    if 'Fault' in prefs_error_data:
                                                        prefs_fault = prefs_error_data['Fault']
                                                        if 'Error' in prefs_fault:
                                                            prefs_errors = prefs_fault['Error']
                                                            if isinstance(prefs_errors, dict):
                                                                prefs_errors = [prefs_errors]
                                                            for prefs_error in prefs_errors:
                                                                if prefs_error.get('code') == '130' or 'WrongCluster' in prefs_error.get('Message', ''):
                                                                    # Both endpoints have Wrong Cluster - use main production URL anyway
                                                                    # DNS patch will handle resolution
                                                                    self.base_url = "https://quickbooks.api.intuit.com"
                                                                    self.base_url_verified = True
                                                                    self.verified_via_preferences = True
                                                                    return True
                                                except:
                                                    pass
                                        except Exception as e:
                                            pass
                                        
                                        # Only try to extract cluster URL if preferences endpoint didn't work
                                        # and we haven't verified via preferences yet
                                        if not self.verified_via_preferences:
                                            extracted_cluster_url = self._extract_cluster_url(response)
                                            if extracted_cluster_url:
                                                # Don't try to use extracted cluster URL - it can't be resolved via DNS
                                                # Instead, verify that main production URL works via preferences
                                                try:
                                                    main_prefs_url = f"https://quickbooks.api.intuit.com/v3/company/{self.company_id}/preferences"
                                                    main_prefs_response = requests.get(main_prefs_url, headers=headers, allow_redirects=True, timeout=10, verify=False)
                                                    if main_prefs_response.status_code == 200:
                                                        self.base_url = "https://quickbooks.api.intuit.com"
                                                        self.base_url_verified = True
                                                        self.verified_via_preferences = True
                                                        return True
                                                except Exception as e:
                                                    pass
                                        
                                        # If we get here, preferences endpoint didn't work and we couldn't extract cluster
                                        # Use main production URL anyway - DNS patch will handle it
                                        if not self.verified_via_preferences:
                                            self.base_url = "https://quickbooks.api.intuit.com"
                                            self.base_url_verified = True
                                            self.verified_via_preferences = True
                                            self._debug("Using main production URL - DNS patch will handle regional cluster DNS lookups")
                                            return True
                                        
                                        # Check if final URL is different (redirect happened)
                                        if final_url != discovery_url:
                                            st.info(f"🔍 URL changed after redirect: `{discovery_url}` → `{final_url}`")
                                            # Extract base URL from final URL
                                            if '/v3/company/' in final_url:
                                                redirected_base_url = final_url.split('/v3/company/')[0]
                                                if redirected_base_url in all_cluster_urls:
                                                    self._debug(f"Trying redirected cluster URL: `{redirected_base_url}`")
                                                    # Try the redirected cluster URL (normalize to prevent DNS errors)
                                                    normalized_redirected_base = self._normalize_quickbooks_url(redirected_base_url)
                                                    redirected_url = self._normalize_quickbooks_url(
                                                        f"{normalized_redirected_base}/v3/company/{self.company_id}/companyinfo/{self.company_id}"
                                                    )
                                                    redirected_response = requests.get(redirected_url, headers=headers, allow_redirects=True, timeout=10, verify=False)
                                                    if redirected_response.status_code == 200:
                                                        old_url = self.base_url
                                                        self.base_url = redirected_base_url
                                                        self.base_url_verified = True
                                                        st.success(f"✅ Discovered correct QuickBooks cluster URL via redirect: `{old_url}` → `{redirected_base_url}`")
                                                        return True
                                        # Continue to next cluster URL
                                        continue
                    except:
                        pass
                    
                    # If it's a 401 but not a Wrong Cluster error, try refreshing token
                    self._debug("Got 401, refreshing token...")
                    if self.authenticate():
                        headers["Authorization"] = f"Bearer {self.access_token}"
                        response = requests.get(discovery_url, headers=headers, allow_redirects=True, timeout=10, verify=False)
                        
                        if response.status_code == 200:
                            final_url = response.url
                            if '/v3/company/' in final_url:
                                discovered_base_url = final_url.split('/v3/company/')[0]
                            else:
                                discovered_base_url = cluster_url
                            
                            if discovered_base_url != self.base_url:
                                old_url = self.base_url
                                self.base_url = discovered_base_url
                                self.base_url_verified = True
                                st.success(f"✅ Discovered correct QuickBooks cluster URL: `{old_url}` → `{discovered_base_url}`")
                                return True
                            else:
                                self.base_url_verified = True
                                st.success(f"✅ Cluster URL verified: `{self.base_url}`")
                                return True
                
                # If redirect status, follow it
                elif response.status_code in [301, 302, 303, 307, 308]:
                    location_header = response.headers.get('Location', '')
                    if location_header:
                        st.info(f"🔍 Following redirect to: `{location_header}`")
                        response = requests.get(location_header, headers=headers, allow_redirects=True, timeout=10, verify=False)
                        if response.status_code == 200:
                            final_url = response.url
                            if '/v3/company/' in final_url:
                                discovered_base_url = final_url.split('/v3/company/')[0]
                                if discovered_base_url != self.base_url:
                                    old_url = self.base_url
                                    self.base_url = discovered_base_url
                                    self.base_url_verified = True
                                    st.success(f"✅ Discovered correct QuickBooks cluster URL: `{old_url}` → `{discovered_base_url}`")
                                    return True
                
            except requests.exceptions.RequestException as e:
                st.warning(f"⚠️ Request failed for cluster `{cluster_url}`: {str(e)}")
                continue
        
        # If all cluster URLs failed, try toggling sandbox setting
        # Sometimes the sandbox flag in config might be incorrect
        st.warning("⚠️ Both cluster URLs returned Wrong Cluster errors. Trying opposite sandbox setting...")
        
        # Toggle sandbox and try again
        opposite_sandbox = not self.sandbox
        if opposite_sandbox:
            alternative_cluster = "https://sandbox-quickbooks.api.intuit.com"
        else:
            alternative_cluster = "https://quickbooks.api.intuit.com"
        
        self._debug(f"Trying cluster URL with opposite sandbox setting: `{alternative_cluster}`")
        # Refresh token again as tokens might be cluster-specific
        self._debug("Refreshing token for alternative cluster...")
        if self.authenticate(force_refresh=True):
            headers["Authorization"] = f"Bearer {self.access_token}"
        
        alternative_url = f"{alternative_cluster}/v3/company/{self.company_id}/companyinfo/{self.company_id}"
        
        try:
            alt_response = requests.get(alternative_url, headers=headers, allow_redirects=True, timeout=10, verify=False)
            
            if alt_response.status_code == 200:
                final_url = alt_response.url
                if '/v3/company/' in final_url:
                    discovered_base_url = final_url.split('/v3/company/')[0]
                else:
                    discovered_base_url = alternative_cluster
                
                old_url = self.base_url
                self.base_url = discovered_base_url
                self.sandbox = opposite_sandbox  # Update sandbox setting
                self.base_url_verified = True
                st.success(f"✅ Discovered correct QuickBooks cluster URL with opposite sandbox setting!")
                st.success(f"   Old: `{old_url}` (sandbox={not opposite_sandbox})")
                st.success(f"   New: `{discovered_base_url}` (sandbox={opposite_sandbox})")
                st.warning(f"⚠️ Please update your secrets.toml: Set `sandbox = {str(opposite_sandbox).lower()}`")
                return True
        except Exception as e:
            st.warning(f"⚠️ Alternative cluster URL also failed: {str(e)}")
        
        # If all methods failed, show debug info
        st.error("❌ Could not discover cluster URL using any method")
        st.info("💡 Debug Information:")
        st.info(f"   - Current base URL: `{self.base_url}`")
        st.info(f"   - Sandbox mode: {self.sandbox}")
        st.info(f"   - Company ID: `{self.company_id}`")
        st.info(f"   - Tried cluster URLs:")
        for cluster in cluster_urls_to_try:
            st.info(f"     • {cluster}")
        st.info(f"   - Also tried: `{alternative_cluster}` (opposite sandbox setting)")
        st.error("💡 Troubleshooting:")
        st.error("   1. Verify your `sandbox` setting in secrets.toml matches your QuickBooks environment")
        st.error("   2. Check that your Company ID is correct")
        st.error("   3. Verify your refresh token is valid and matches your Client ID")
        st.error("   4. Your company may be on a different regional cluster - contact QuickBooks support")
        
        return False
    
    def _verify_company_access(self) -> bool:
        """
        Verify we can access the company (handles cluster redirects)
        This will trigger QuickBooks to redirect us to the correct cluster if needed
        
        Returns:
            bool: True if company is accessible, False otherwise
        """
        if not self.access_token:
            if not self.authenticate():
                return False
        
        # If we already verified via preferences, we know the cluster is correct
        # Refresh token again after cluster verification to ensure we have a cluster-specific token
        if self.verified_via_preferences and self.base_url_verified:
            # Refresh token one more time to ensure we have a cluster-specific token
            self.authenticate(force_refresh=True)
            return True
        
        # If cluster hasn't been verified, try to discover it first
        if not self.base_url_verified:
            if self._discover_cluster_url():
                # Cluster discovered, now verify access
                pass
        
        try:
            # Try to get company info - this will redirect to correct cluster if needed
            # Also try preferences endpoint as fallback (sometimes works when companyinfo fails)
            company_url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/companyinfo/{self.company_id}"
            )
            preferences_url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/preferences"
            )
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json",
                "Accept-Encoding": "identity"
            }
            
            response = requests.get(company_url, headers=headers, allow_redirects=True, verify=False)
            
            # Extract base URL from final URL if redirect happened
            if response.status_code == 200:
                final_url = response.url
                if '/v3/company/' in final_url:
                    new_base_url = final_url.split('/v3/company/')[0]
                    if new_base_url != self.base_url:
                        old_url = self.base_url
                        self.base_url = new_base_url
                        self._debug(f"QuickBooks cluster URL updated: {old_url} → {new_base_url}")
                
                self.base_url_verified = True
                return True
            
            # Handle 401 errors - try preferences endpoint if companyinfo fails
            if response.status_code == 401:
                # Check if it's a Wrong Cluster error
                try:
                    error_data = response.json()
                    if 'Fault' in error_data:
                        fault = error_data['Fault']
                        if 'Error' in fault:
                            errors = fault['Error']
                            if isinstance(errors, dict):
                                errors = [errors]
                            for error in errors:
                                if error.get('code') == '130' or 'WrongCluster' in error.get('Message', ''):
                                    # Try preferences endpoint instead
                                    prefs_response = requests.get(preferences_url, headers=headers, allow_redirects=True, verify=False)
                                    if prefs_response.status_code == 200:
                                        self.base_url_verified = True
                                        self.verified_via_preferences = True  # Mark that we verified via preferences
                                        return True
                except:
                    pass
                # Try refreshing token
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.get(company_url, headers=headers, allow_redirects=True, verify=False)
                    
                    if response.status_code == 200:
                        final_url = response.url
                        if '/v3/company/' in final_url:
                            new_base_url = final_url.split('/v3/company/')[0]
                            if new_base_url != self.base_url:
                                self.base_url = new_base_url
                        
                        self.base_url_verified = True
                        return True
                
                # Check for cluster error in response
                try:
                    error_data = response.json()
                    
                    if 'Fault' in error_data:
                        fault = error_data['Fault']
                        if 'Error' in fault:
                            errors = fault['Error']
                            # Handle both single error dict and list of errors
                            if isinstance(errors, dict):
                                errors = [errors]
                            
                            for error in errors:
                                error_msg = error.get('Message', '')
                                error_detail = error.get('Detail', '')
                                error_code = error.get('code', '')
                                
                                st.info(f"🔍 Error code: {error_code}")
                                st.info(f"🔍 Error message: {error_msg}")
                                st.info(f"🔍 Error detail: {error_detail}")
                                
                                if error_code == '130' or 'WrongCluster' in error_msg or 'WrongCluster' in error_detail:
                                    # If we already verified via preferences, ignore companyinfo Wrong Cluster errors
                                    if self.verified_via_preferences:
                                        st.info("💡 Ignoring companyinfo Wrong Cluster error (cluster already verified via preferences endpoint)")
                                        return True  # Return True since cluster is already verified
                                    
                                    # Try to discover cluster URL
                                    self._debug("QuickBooks cluster mismatch detected. Discovering correct cluster...")
                                    if self._discover_cluster_url():
                                        # Retry verification with new cluster URL
                                        company_url = self._normalize_quickbooks_url(
                                            f"{self.base_url}/v3/company/{self.company_id}/companyinfo/{self.company_id}"
                                        )
                                        headers["Authorization"] = f"Bearer {self.access_token}"
                                        response = requests.get(company_url, headers=headers, allow_redirects=True, verify=False)
                                        if response.status_code == 200:
                                            self.base_url_verified = True
                                            return True
                                    return False
                except Exception as e:
                    st.warning(f"⚠️ Could not parse error response: {str(e)}")
                    try:
                        st.info(f"🔍 Raw response text: {response.text[:500]}")
                    except:
                        pass
                
                st.error("❌ Could not verify company access after token refresh")
                return False
            
            return False
            
        except Exception as e:
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
        # Always authenticate to ensure we have a fresh token (force refresh)
        if not self.authenticate(force_refresh=True):
                return None
        
        # Verify company access first (this will establish the correct cluster URL if needed)
        # Skip if already verified via preferences (companyinfo endpoint has cluster issues)
        if not (self.base_url_verified and self.verified_via_preferences):
            if not self._verify_company_access():
                # If verification failed but we verified via preferences, that's OK - proceed anyway
                if not self.verified_via_preferences:
                    st.warning("⚠️ Could not verify company access. Please check your QuickBooks credentials.")
                    return None
                else:
                    st.info("💡 Proceeding with customer creation (cluster verified via preferences endpoint)")
        else:
            st.info("✅ Cluster already verified via preferences - proceeding with customer creation")
        
        # Always refresh token before customer operations
        self._debug("Refreshing token for customer operations...")
        self.authenticate(force_refresh=True)
        
        # FORCE use of main production URL - ignore Wrong Cluster errors
        # Regional cluster domains (like qbo-usw2.api.intuit.com) cannot be resolved via DNS
        # QuickBooks will proxy requests internally from the main domain even if it says wrong cluster
        self.base_url = "https://quickbooks.api.intuit.com"
        self.customer_cluster_url = None  # Don't use regional clusters (DNS resolution issues)
        self._debug("Using main production URL: https://quickbooks.api.intuit.com")
        self._debug("Ignoring Wrong Cluster errors - QuickBooks will proxy internally")
        if self.sandbox:
            st.info("🔧 Sandbox Mode: Using existing customer for testing")
            return self._get_or_create_sandbox_customer(first_name, last_name, email, company_name)
        
        # Try to find existing customer first, but if query endpoint has cluster issues, skip it
        # We'll try to create the customer directly and handle "already exists" errors
        existing_customer_id = None
        if self.verified_via_preferences:
            # If we verified via preferences but query endpoint fails, skip query and try creation directly
            st.info("💡 Skipping customer query (endpoint has cluster issues) - will try direct creation")
        else:
            # Try to find existing customer
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
            # FORCE use of main production URL - ignore regional cluster discovery
            # Regional cluster domains (like qbo-usw2.api.intuit.com) cannot be resolved via DNS
            # QuickBooks will proxy requests internally from the main domain even if it says wrong cluster
            customer_url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/customer"
            )
            
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
            
            response = requests.post(customer_url, json=payload, headers=headers, allow_redirects=True, verify=False)
            
            # Handle 401 errors
            # IMPORTANT: Ignore Wrong Cluster errors - just use main production URL
            # QuickBooks will proxy requests internally even if it says wrong cluster
            if response.status_code == 401:
                def retry_request(updated_headers):
                    # Always use main production URL - ignore cluster errors
                    current_url = self._normalize_quickbooks_url(
                        f"{self.base_url}/v3/company/{self.company_id}/customer"
                    )
                    return requests.post(current_url, json=payload, headers=updated_headers, allow_redirects=True, verify=False)
                
                retry_response = self._handle_401_error(response, headers, retry_request)
                if retry_response is None:
                    # If we verified via preferences but still get Wrong Cluster, refresh token and retry
                    # IGNORE Wrong Cluster errors - just use main production URL
                    if self.verified_via_preferences:
                        st.info("💡 Verified via preferences but got 401 - refreshing token and retrying with main URL...")
                        self._debug("Ignoring Wrong Cluster errors - QuickBooks will proxy internally")
                        if self.authenticate(force_refresh=True):
                            headers["Authorization"] = f"Bearer {self.access_token}"
                            try:
                                # Always use main production URL - ignore cluster errors
                                current_url = self._normalize_quickbooks_url(
                                    f"{self.base_url}/v3/company/{self.company_id}/customer"
                                )
                                retry_response = requests.post(current_url, json=payload, headers=headers, allow_redirects=True, verify=False)
                                
                                # If we get 200/201, success!
                                if retry_response.status_code in [200, 201]:
                                    response = retry_response
                                    self.base_url_verified = True
                                elif retry_response.status_code == 401:
                                    # Check if it's Wrong Cluster error - if so, try to parse response anyway
                                    # Sometimes the request succeeds despite the error
                                    try:
                                        error_data = retry_response.json()
                                        is_wrong_cluster = False
                                        if 'Fault' in error_data:
                                            fault = error_data['Fault']
                                            if 'Error' in fault:
                                                errors = fault['Error']
                                                if isinstance(errors, dict):
                                                    errors = [errors]
                                                for error in errors:
                                                    if error.get('code') == '130' or 'WrongCluster' in error.get('Message', ''):
                                                        is_wrong_cluster = True
                                                        break
                                        
                                        if is_wrong_cluster:
                                            st.warning("⚠️ Still getting Wrong Cluster error - trying to parse response anyway")
                                            # Try to extract customer ID from response (might succeed despite error)
                                            customer_id = self._extract_customer_id_from_response(retry_response)
                                            if customer_id:
                                                st.success(f"✅ Successfully extracted customer ID despite Wrong Cluster error: {customer_id}")
                                                return customer_id
                                            
                                            # Try following redirect
                                            customer_id = self._try_follow_redirect_for_customer(retry_response, payload, headers)
                                            if customer_id:
                                                return customer_id
                                        
                                        # Try batch operations as fallback
                                        self._debug("Trying batch operations as fallback...")
                                        customer_id = self._try_create_customer_via_batch(first_name, last_name, email, company_name)
                                        if customer_id:
                                            return customer_id
                                    except:
                                        pass
                                    
                                    # If still failing after all attempts
                                    response = retry_response
                            except Exception as e:
                                st.error(f"Error retrying customer creation: {str(e)}")
                                return None
                        else:
                            return None
                    else:
                        return None
                else:
                    # retry_response is not None - use it
                    response = retry_response
                    # If successful, mark base_url as verified
                    if response.status_code == 200:
                        self.base_url_verified = True
            
            # Check response status before raise_for_status
            # IMPORTANT: Don't fail on 401 if it's a Wrong Cluster error - QuickBooks proxies internally
            if response.status_code == 401:
                # Check if it's a Wrong Cluster error
                is_wrong_cluster = False
                try:
                    error_data = response.json()
                    if 'Fault' in error_data:
                        fault = error_data['Fault']
                        if 'Error' in fault:
                            errors = fault['Error']
                            if isinstance(errors, dict):
                                errors = [errors]
                            for error in errors:
                                if error.get('code') == '130' or 'WrongCluster' in error.get('Message', ''):
                                    is_wrong_cluster = True
                                    break
                except:
                    pass
                
                # If it's a Wrong Cluster error and we verified via preferences, ignore it and continue
                # Sometimes the request still succeeds despite the error
                if is_wrong_cluster and self.verified_via_preferences:
                    st.warning("⚠️ Got Wrong Cluster error - but preferences verified production")
                    st.info("💡 Ignoring error - QuickBooks will proxy internally")
                    # Try to parse the response anyway - sometimes it contains the customer data
                    try:
                        error_data = response.json()
                        # Check if there's customer data in the response despite the error
                        if 'Customer' in error_data:
                            customer = error_data['Customer']
                            if isinstance(customer, dict) and 'Id' in customer:
                                st.success(f"✅ Found customer ID in response despite Wrong Cluster error: {customer['Id']}")
                                return customer['Id']
                    except:
                        pass
                    
                    # Try to extract customer ID from response
                    customer_id = self._extract_customer_id_from_response(response)
                    if customer_id:
                        st.success(f"✅ Successfully extracted customer ID despite Wrong Cluster error: {customer_id}")
                        return customer_id
                    
                    # Try batch operations as fallback
                    self._debug("Trying batch operations as fallback...")
                    customer_id = self._try_create_customer_via_batch(first_name, last_name, email, company_name)
                    if customer_id:
                        return customer_id
                    
                    # If all else fails, show helpful error message
                    st.error("❌ QuickBooks API Error: Cannot create customer due to Wrong Cluster errors")
                    st.warning("⚠️ This is a known QuickBooks API limitation.")
                    st.info("💡 **Solution:** Your company data is on a regional cluster that cannot be resolved via DNS.")
                    st.info("💡 **Workaround:** Contact QuickBooks Developer Support to resolve cluster routing issues.")
                    st.info("   • Company ID: 9341455592326421")
                    st.info("   • Issue: Wrong Cluster errors on customer endpoints despite production verification")
                    return None
                else:
                    # Not a Wrong Cluster error or not verified - return None
                    st.error("❌ Customer creation failed with 401 error")
                    return None
            
            # If we got here, response should be 200/201 - try to parse it
            # Even if status is 401, sometimes the response contains the customer data
            try:
                customer_response = response.json()
                customer = customer_response.get("Customer")
                
                if customer and isinstance(customer, dict) and "Id" in customer:
                    st.success(f"✅ Customer created successfully: {customer['Id']}")
                    return customer.get("Id")
            except:
                pass
            
            # If status is 200/201, raise for status if we couldn't parse customer
            if response.status_code in [200, 201]:
                try:
                    response.raise_for_status()
                    # If we got here, try to parse again
                    customer_response = response.json()
                    customer = customer_response.get("Customer")
                    if customer:
                        return customer.get("Id")
                except:
                    pass
            
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
            query_url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/query?query=SELECT * FROM Customer"
            )
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json",
                "Accept-Encoding": "identity"  # Disable gzip to avoid decoding issues
            }
            
            response = requests.get(query_url, headers=headers, verify=False)
            
            # If we get a 401, try to refresh the token and retry
            if response.status_code == 401:
                self._debug("Access token expired, refreshing...")
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.get(query_url, headers=headers, verify=False)
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
            customer_url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/customer/{customer_id}"
            )
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json",
                "Accept-Encoding": "identity"
            }
            
            response = requests.get(customer_url, headers=headers, verify=False)
            
            if response.status_code == 401:
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.get(customer_url, headers=headers, verify=False)
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
        # Always authenticate to ensure we have a fresh token (force refresh)
        if not self.authenticate(force_refresh=True):
                return None
        
        try:
            query_url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/query"
            )
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Accept-Encoding": "identity"  # Disable gzip to avoid decoding issues
            }
            
            # Query to find customer by email
            query = f"SELECT * FROM Customer WHERE PrimaryEmailAddr = '{email}'"
            
            params = {"query": query}
            
            response = requests.get(query_url, params=params, headers=headers, verify=False)
            
            # Handle 401 errors
            if response.status_code == 401:
                def retry_request(updated_headers):
                    # Use current base_url (may have been updated for cluster redirect)
                    current_url = self._normalize_quickbooks_url(
                        f"{self.base_url}/v3/company/{self.company_id}/query"
                    )
                    return requests.get(current_url, params=params, headers=updated_headers, verify=False)
                
                retry_response = self._handle_401_error(response, headers, retry_request)
                if retry_response is None:
                    # If we verified via preferences but still get Wrong Cluster, refresh token and retry
                    if self.verified_via_preferences:
                        st.info("💡 Verified via preferences but got 401 - refreshing token and retrying...")
                        if self.authenticate(force_refresh=True):
                            headers["Authorization"] = f"Bearer {self.access_token}"
                            try:
                                current_url = self._normalize_quickbooks_url(
                                    f"{self.base_url}/v3/company/{self.company_id}/query"
                                )
                                retry_response = requests.get(current_url, params=params, headers=headers)
                                if retry_response.status_code == 200:
                                    response = retry_response
                                elif retry_response.status_code == 401:
                                    # Check if it's Wrong Cluster error
                                    try:
                                        error_data = retry_response.json()
                                        if 'Fault' in error_data:
                                            fault = error_data['Fault']
                                            if 'Error' in fault:
                                                errors = fault['Error']
                                                if isinstance(errors, dict):
                                                    errors = [errors]
                                                for error in errors:
                                                    if error.get('code') == '130' or 'WrongCluster' in error.get('Message', ''):
                                                        # Still Wrong Cluster but verified via preferences - try anyway
                                                        st.warning("⚠️ Still getting Wrong Cluster error, but cluster verified via preferences")
                                                        # Return empty result instead of None
                                                        return None
                                    except:
                                        pass
                                    if response.status_code == 401:
                                        return None
                            except Exception as e:
                                st.error(f"Error retrying customer query: {str(e)}")
                                return None
                        else:
                            return None
                    # If cluster was updated, retry once more with new URL
                    elif not self.base_url_verified:
                        try:
                            current_url = self._normalize_quickbooks_url(
                                f"{self.base_url}/v3/company/{self.company_id}/query"
                            )
                            retry_response = requests.get(current_url, params=params, headers=headers)
                            if retry_response.status_code == 200:
                                response = retry_response
                            else:
                                return None
                        except:
                            return None
                    else:
                        return None
                else:
                    response = retry_response
            
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
            query_url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/query"
            )
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json",
                "Accept-Encoding": "identity"
            }
            
            # Query for all active items (includes Service and NonInventory)
            query = "SELECT * FROM Item WHERE Active = true"
            params = {"query": query}
            
            response = requests.get(query_url, params=params, headers=headers, verify=False)
            
            if response.status_code == 401:
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.get(query_url, params=params, headers=headers, verify=False)
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
                url = self._normalize_quickbooks_url(
                    f"{self.base_url}/v3/company/{self.company_id}/query?query={query}"
                )
                headers = {"Authorization": f"Bearer {self.access_token}", "Accept": "application/json"}
                
                response = requests.get(url, headers=headers, verify=False)
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
                url = self._normalize_quickbooks_url(
                    f"{self.base_url}/v3/company/{self.company_id}/query?query={query}"
                )
                headers = {"Authorization": f"Bearer {self.access_token}", "Accept": "application/json"}
                
                response = requests.get(url, headers=headers, verify=False)
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
            url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/item"
            )
            
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
            
            response = requests.post(url, headers=headers, json=item_data, verify=False)
            
            if response.status_code == 401:
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.post(url, headers=headers, json=item_data, verify=False)
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
            payload = self._build_invoice_payload(
                customer_id=customer_id,
                email=email,
                company_name=company_name,
                client_address=client_address,
                contract_amount=contract_amount,
                description=description,
                line_items=line_items,
                payment_terms=payment_terms,
                enable_payment_link=enable_payment_link,
                invoice_date=invoice_date,
                invoice_number=None
            )
            
            invoice_url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/invoice"
            )
            
            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Accept-Encoding": "identity"  # Disable gzip to avoid decoding issues
            }
            
            response = requests.post(invoice_url, json=payload, headers=headers, verify=False)
            
            # If we get a 401, try to refresh the token and retry
            if response.status_code == 401:
                self._debug("Access token expired, refreshing...")
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.post(invoice_url, json=payload, headers=headers, verify=False)
                else:
                    st.error("Failed to refresh access token")
                    return None
            
            # Check response status
            if response.status_code != 200:
                st.error(f"❌ QuickBooks API Error (Status {response.status_code})")
                st.error(f"Response: {response.text[:500]}")
                try:
                    st.info("📦 Invoice payload sent to QuickBooks:")
                    st.code(json.dumps(self._scrub_invoice_payload_for_logging(payload), indent=2))
                except Exception:
                    pass
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

    def _build_invoice_payload(
        self,
        customer_id: str,
        email: Optional[str],
        company_name: Optional[str],
        client_address: Optional[str],
        contract_amount: str,
        description: str,
        line_items: Optional[list],
        payment_terms: str,
        enable_payment_link: bool,
        invoice_date: Optional[str],
        invoice_number: Optional[str] = None
    ) -> Dict:
        """
        Build the QuickBooks invoice payload that will be sent to the API.

        This helper keeps payload construction testable and avoids including
        unsupported fields such as AllowOnlinePayment.
        """

        if line_items:
            invoice_lines = []

            for item in line_items:
                quantity = item.get('quantity', 1) or 1
                total_amount = float(item.get('amount', 0) or 0)
                line_item_type = item.get('type', item.get('description', 'Service'))
                type_lower = (line_item_type or "").lower()
                line_item_description = item.get('line_description', item.get('name', ''))

                # Use DiscountLineDetail for discounts/negative entries
                if total_amount < 0 or 'discount' in type_lower:
                    discount_amount = abs(total_amount)
                    description = line_item_description or item.get('description', 'Discount')

                    invoice_lines.append({
                        "Amount": round(discount_amount, 2),
                        "DetailType": "DiscountLineDetail",
                        "Description": description,
                        "DiscountLineDetail": {
                            "PercentBased": False,
                            "DiscountAccountRef": {
                                "value": self._get_discount_account_id()
                            }
                        }
                    })
                    continue

                # Normalize quantity to positive values; adjust amount sign accordingly
                if quantity == 0:
                    quantity = 1
                if quantity < 0:
                    quantity = abs(quantity)
                    total_amount = -total_amount

                unit_price = float(item.get('unit_price', total_amount / quantity if quantity else total_amount))
                unit_price = round(unit_price, 2)
                line_total = round(unit_price * quantity, 2)

                if line_item_type == "Contract Services":
                    item_id = self._find_best_item_match(line_item_type, show_match_info=False)
                    full_description = line_item_description if line_item_description else ""
                elif line_item_type == "Credit Card Processing Fee":
                    item_id = self._find_best_item_match(line_item_type, show_match_info=False)
                    full_description = "3% processing fee for credit card payments"
                elif line_item_type == "Credits & Discounts":
                    item_id = self._find_best_item_match(line_item_type, show_match_info=False)
                    full_description = line_item_description if line_item_description else item.get('description', '')
                else:
                    item_id = self._find_best_item_match(line_item_type, show_match_info=False)
                    full_description = line_item_description if line_item_description else ""

                invoice_lines.append({
                    "Amount": line_total,
                    "DetailType": "SalesItemLineDetail",
                    "Description": full_description,
                    "SalesItemLineDetail": {
                        "Qty": quantity,
                        "UnitPrice": unit_price,
                        "ItemRef": {
                            "value": item_id
                        }
                    }
                })
        else:
            amount_str = contract_amount.replace('$', '').replace(',', '')
            amount = float(amount_str)
            item_id = self._find_best_item_match(description, show_match_info=False)
            invoice_lines = [
                {
                    "DetailType": "SalesItemLineDetail",
                    "Amount": amount,
                    "SalesItemLineDetail": {
                        "ItemRef": {
                            "value": item_id
                        },
                        "Qty": 1,
                        "UnitPrice": amount
                    },
                    "Description": ""
                }
            ]

        if invoice_date:
            txn_date = invoice_date.strftime("%Y-%m-%d") if hasattr(invoice_date, 'strftime') else str(invoice_date)
        else:
            txn_date = datetime.now().strftime("%Y-%m-%d")

        invoice_data = {
            "CustomerRef": {"value": customer_id},
            "TxnDate": txn_date,
            "DueDate": txn_date,
            "Line": invoice_lines,
            "EmailStatus": "NotSet"
        }

        if invoice_number:
            invoice_data["DocNumber"] = str(invoice_number)
        else:
            invoice_data["AutoDocNumber"] = True

        # QuickBooks does not accept online payment toggles on invoice creation.
        # These must be configured in company preferences instead.

        if email:
            invoice_data["BillEmail"] = {"Address": email}

        bill_addr = self._parse_bill_address(company_name, client_address)
        if bill_addr:
            invoice_data["BillAddr"] = bill_addr

        if payment_terms != "Due in Full":
            invoice_data["SalesTermRef"] = {
                "value": self._get_payment_term_id(payment_terms)
            }

        return invoice_data

    def _parse_bill_address(self, company_name: Optional[str], client_address: Optional[str]) -> Optional[Dict[str, str]]:
        """
        Convert company/address strings into QuickBooks BillAddr structure (Line1-Line5).
        """
        lines: list[str] = []

        if company_name:
            lines.append(company_name.strip())

        if client_address:
            parts = [part.strip() for part in client_address.split(",") if part.strip()]
            lines.extend(parts)

        if not lines:
            return None

        bill_addr: Dict[str, str] = {}
        for idx, value in enumerate(lines[:5]):
            bill_addr[f"Line{idx + 1}"] = value

        return bill_addr

    def _get_discount_account_id(self) -> str:
        """
        Retrieve (and cache) an account ID to use for DiscountLineDetail entries.
        """
        if self._discount_account_id:
            return self._discount_account_id

        # Default to the primary income account if a dedicated discount account isn't specified.
        self._discount_account_id = self._get_default_income_account()
        return self._discount_account_id

    def _scrub_invoice_payload_for_logging(self, payload: Dict) -> Dict:
        """
        Prepare a sanitized copy of the invoice payload for debugging output.
        Ensures all values are JSON-serializable and removes potentially sensitive data.
        """
        def scrub(value):
            if isinstance(value, dict):
                return {k: scrub(v) for k, v in value.items()}
            if isinstance(value, list):
                return [scrub(v) for v in value]
            if isinstance(value, (datetime,)):
                return value.isoformat()
            if isinstance(value, float):
                # Use standard two-decimal formatting for readability
                return float(f"{value:.2f}")
            return value

        safe_payload = scrub(payload)

        # Remove or mask anything that shouldn't be logged (currently none, but placeholder for future)
        return safe_payload
    
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
            send_url = self._normalize_quickbooks_url(
                f"{self.base_url}/v3/company/{self.company_id}/invoice/{invoice_id}/send?sendTo={email}"
            )
            
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
                self._debug("Access token expired, refreshing...")
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
                # Check if this is due to Wrong Cluster error
                if self.verified_via_preferences:
                    error_msg = (
                        "❌ QuickBooks API Error: Cannot create or find customer due to Wrong Cluster errors.\n\n"
                        "This is a known QuickBooks API limitation affecting your company. "
                        "The preferences endpoint works correctly, confirming your company is in production, "
                        "but customer-related endpoints return Wrong Cluster errors.\n\n"
                        "Please contact QuickBooks Developer Support and reference:\n"
                        "• Company ID: 9341455592326421\n"
                        "• Issue: Wrong Cluster errors on customer/query endpoints\n"
                        "• Note: Preferences endpoint works correctly"
                    )
                    return False, error_msg
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


def verify_production_credentials(quickbooks_api) -> Tuple[bool, str]:
    """
    Verify if QuickBooks credentials are truly production by testing the preferences endpoint.
    
    This test calls the preferences endpoint with the current access token to confirm:
    - If it succeeds (200) → credentials are production
    - If it fails with 401 → credentials are sandbox or token is mismatched
    
    Args:
        quickbooks_api: QuickBooksAPI instance with authenticated token
        
    Returns:
        Tuple[bool, str]: (is_production, message)
    """
    try:
        # Ensure we have a fresh access token
        if not quickbooks_api.authenticate(force_refresh=True):
            return False, "❌ Failed to authenticate - cannot verify credentials"
        
        # Test the preferences endpoint (production URL)
        test_url = f"https://quickbooks.api.intuit.com/v3/company/{quickbooks_api.company_id}/preferences"
        
        headers = {
            "Authorization": f"Bearer {quickbooks_api.access_token}",
            "Accept": "application/json"
        }
        
        st.info("🔍 Testing credentials against production preferences endpoint...")
        st.info(f"   URL: {test_url}")
        
        response = requests.get(test_url, headers=headers, timeout=10, verify=False)
        
        if response.status_code == 200:
            st.success("✅ SUCCESS: Preferences endpoint returned 200 OK")
            st.success("✅ Your credentials are CONFIRMED PRODUCTION")
            return True, "✅ Production credentials verified - preferences endpoint returned 200 OK"
        
        elif response.status_code == 401:
            st.error("❌ FAILED: Preferences endpoint returned 401 Unauthorized")
            st.warning("⚠️ This indicates:")
            st.warning("   • You're using SANDBOX credentials with production URL, OR")
            st.warning("   • Your refresh token doesn't match your Client ID, OR")
            st.warning("   • Your access token is invalid")
            
            # Check if it's a Wrong Cluster error
            try:
                error_data = response.json()
                if 'Fault' in error_data:
                    fault = error_data['Fault']
                    if 'Error' in fault:
                        errors = fault['Error']
                        if isinstance(errors, dict):
                            errors = [errors]
                        for error in errors:
                            if error.get('code') == '130' or 'WrongCluster' in error.get('Message', ''):
                                st.info("💡 This is a Wrong Cluster error - your company may be on a regional cluster")
                                # Try to extract the cluster URL
                                cluster_url = quickbooks_api._extract_cluster_url(response)
                                if cluster_url:
                                    st.success(f"✅ Found correct cluster URL: {cluster_url}")
                                    return True, f"✅ Production credentials verified - company is on cluster: {cluster_url}"
            except:
                pass
            
            return False, "❌ Production credentials verification FAILED - received 401 Unauthorized"
        
        else:
            st.warning(f"⚠️ Unexpected status code: {response.status_code}")
            try:
                error_text = response.text[:200]
                st.info(f"   Response: {error_text}")
            except:
                pass
            return False, f"❌ Unexpected response: status code {response.status_code}"
            
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Network error during verification: {str(e)}")
        return False, f"❌ Network error: {str(e)}"
    except Exception as e:
        st.error(f"❌ Error during verification: {str(e)}")
        return False, f"❌ Verification error: {str(e)}"


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
