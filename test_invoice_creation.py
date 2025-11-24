"""
Test script to debug QuickBooks invoice creation with CC email
"""
import sys
import os
import json
import toml

# Add parent directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from quickbooks_integration import QuickBooksAPI

def load_quickbooks_credentials_local():
    """Load credentials from secrets.toml without streamlit"""
    secrets_path = os.path.join(os.path.dirname(__file__), '.streamlit', 'secrets.toml')
    if not os.path.exists(secrets_path):
        print(f"❌ Secrets file not found at: {secrets_path}")
        return {}
    
    with open(secrets_path, 'r') as f:
        secrets = toml.load(f)
    
    if 'quickbooks' not in secrets:
        print("❌ QuickBooks config not found in secrets.toml")
        return {}
    
    return secrets['quickbooks']

def test_minimal_invoice():
    """Test creating a minimal invoice to identify the problem"""
    
    # Load credentials
    credentials = load_quickbooks_credentials_local()
    if not credentials:
        print("❌ Failed to load credentials")
        return
    
    # Initialize API
    api = QuickBooksAPI(
        client_id=credentials['client_id'],
        client_secret=credentials['client_secret'],
        refresh_token=credentials['refresh_token'],
        company_id=credentials['company_id'],
        sandbox=credentials.get('sandbox', False)
    )
    
    # Authenticate
    print("🔐 Authenticating...")
    if not api.authenticate():
        print("❌ Authentication failed")
        return
    print("✅ Authenticated successfully")
    
    # Test 1: Create a customer first
    print("\n📝 Creating test customer...")
    customer_id = api.create_customer(
        first_name="Test",
        last_name="Customer",
        email="test@example.com",
        company_name="Test Company"
    )
    
    if not customer_id:
        print("❌ Failed to create customer")
        return
    
    print(f"✅ Customer created: {customer_id}")
    
    # Test 2: Minimal invoice payload (no CC)
    print("\n🧪 Test 1: Minimal invoice WITHOUT CC")
    minimal_payload = {
        "CustomerRef": {"value": customer_id},
        "TxnDate": "2025-11-24",
        "Line": [
            {
                "Amount": 100.0,
                "DetailType": "SalesItemLineDetail",
                "SalesItemLineDetail": {
                    "ItemRef": {"value": "1"},
                    "Qty": 1.0,
                    "UnitPrice": 100.0
                },
                "Description": "Test Service"
            }
        ],
        "BillEmail": {"Address": "test@example.com"}
    }
    
    print("\n📤 Payload (Test 1):")
    print(json.dumps(minimal_payload, indent=2))
    
    params = {"minorversion": "75"}
    response = api._make_request('POST', 'invoice', data=minimal_payload, params=params)
    
    if response and response.status_code == 200:
        invoice_data = response.json().get("Invoice", {})
        invoice_id = invoice_data.get("Id")
        print(f"✅ Test 1 SUCCESS - Invoice created: {invoice_id}")
        
        # Test 3: Add CC to the same invoice
        print("\n🧪 Test 2: Invoice WITH CC email")
        minimal_payload_with_cc = minimal_payload.copy()
        minimal_payload_with_cc["BillEmailCc"] = {"Address": "cc@example.com"}
        minimal_payload_with_cc["ToBeEmailed"] = True
        
        print("\n📤 Payload (Test 2):")
        print(json.dumps(minimal_payload_with_cc, indent=2))
        
        response2 = api._make_request('POST', 'invoice', data=minimal_payload_with_cc, params=params)
        
        if response2 and response2.status_code == 200:
            invoice_data2 = response2.json().get("Invoice", {})
            invoice_id2 = invoice_data2.get("Id")
            print(f"✅ Test 2 SUCCESS - Invoice with CC created: {invoice_id2}")
        else:
            print(f"❌ Test 2 FAILED - Status: {response2.status_code if response2 else 'None'}")
            if response2:
                print(f"Response: {response2.text}")
        
    else:
        print(f"❌ Test 1 FAILED - Status: {response.status_code if response else 'None'}")
        if response:
            print(f"Response: {response.text}")
            try:
                error_json = response.json()
                print(f"\n📋 Parsed Error:")
                print(json.dumps(error_json, indent=2))
            except:
                pass

def test_full_payload():
    """Test with the actual payload structure from the code"""
    
    credentials = load_quickbooks_credentials_local()
    if not credentials:
        print("❌ Failed to load credentials")
        return
    
    api = QuickBooksAPI(
        client_id=credentials['client_id'],
        client_secret=credentials['client_secret'],
        refresh_token=credentials['refresh_token'],
        company_id=credentials['company_id'],
        sandbox=credentials.get('sandbox', False)
    )
    
    if not api.authenticate():
        print("❌ Authentication failed")
        return
    
    print("\n🔍 Testing full payload structure...")
    
    # Get a customer ID
    customer_id = api.create_customer(
        first_name="Test",
        last_name="Customer2",
        email="test2@example.com"
    )
    
    if not customer_id:
        print("❌ Failed to create customer")
        return
    
    # Build payload similar to what create_invoice does
    invoice_data = {
        "CustomerRef": {"value": customer_id},
        "TxnDate": "2025-11-24",
        "Line": [
            {
                "Amount": 100.0,
                "DetailType": "SalesItemLineDetail",
                "SalesItemLineDetail": {
                    "ItemRef": {"value": "1"},
                    "Qty": 1.0,
                    "UnitPrice": 100.0
                },
                "Description": "Test Service"
            }
        ]
    }
    
    # Calculate TotalAmt
    total_amt = sum(line.get("Amount", 0) for line in invoice_data["Line"] if isinstance(line.get("Amount"), (int, float)))
    if total_amt > 0:
        invoice_data["TotalAmt"] = float(total_amt)
    
    # Add email fields
    invoice_data["BillEmail"] = {"Address": "test2@example.com"}
    invoice_data["BillEmailCc"] = {"Address": "cc@example.com"}
    invoice_data["ToBeEmailed"] = True
    
    # Add payment terms
    invoice_data["DueDate"] = "2025-11-24"
    invoice_data["SalesTermRef"] = {"value": "1"}
    
    print("\n📤 Full Payload:")
    print(json.dumps(invoice_data, indent=2))
    
    params = {"minorversion": "75"}
    response = api._make_request('POST', 'invoice', data=invoice_data, params=params)
    
    if response and response.status_code == 200:
        invoice_id = response.json().get("Invoice", {}).get("Id")
        print(f"\n✅ SUCCESS - Invoice created: {invoice_id}")
    else:
        print(f"\n❌ FAILED - Status: {response.status_code if response else 'None'}")
        if response:
            print(f"\n📋 Full Error Response:")
            print(response.text)
            try:
                error_json = response.json()
                print(f"\n📋 Parsed Error:")
                print(json.dumps(error_json, indent=2))
                
                # Try to identify problematic field
                errors = error_json.get("Fault", {}).get("Error", [])
                for error in errors:
                    detail = error.get("Detail", "")
                    print(f"\n🔍 Error Detail: {detail}")
                    
            except Exception as e:
                print(f"Could not parse error JSON: {e}")

if __name__ == "__main__":
    print("=" * 60)
    print("QuickBooks Invoice Creation Test")
    print("=" * 60)
    
    # Test with minimal payload first
    test_minimal_invoice()
    
    print("\n" + "=" * 60)
    
    # Test with full payload structure
    test_full_payload()
    
    print("\n" + "=" * 60)
    print("Test completed")

