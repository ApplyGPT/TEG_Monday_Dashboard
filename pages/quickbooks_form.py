"""
QuickBooks Invoice Form Page
Allows users to review and create invoices
"""

import streamlit as st
import sys
import os

# Add parent directory to path to import quickbooks_integration
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from quickbooks_integration import QuickBooksAPI, load_quickbooks_credentials, setup_quickbooks_oauth
import html

def get_decoded_query_params():
    """Get query parameters and decode HTML entities"""
    query_params = st.query_params
    decoded_params = {}
    
    for key, value in query_params.items():
        # Decode HTML entities (like &amp; -> &)
        decoded_value = html.unescape(str(value))
        decoded_params[key] = decoded_value
    
    return decoded_params

# Page configuration
st.set_page_config(
    page_title="QuickBooks Invoice Form",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to hide all pages from sidebar navigation
st.markdown("""
<style>
    /* Hide the entire sidebar navigation */
    [data-testid="stSidebarNav"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main QuickBooks form function"""
    st.title("💰 QuickBooks Invoice Form")
    st.markdown("Review and create invoices for clients")
    
    # Load QuickBooks credentials
    credentials = load_quickbooks_credentials()
    if not credentials:
        st.error("QuickBooks credentials not configured. Please check your secrets.toml file.")
        
        # Show setup instructions
        with st.expander("🔧 QuickBooks Setup Instructions"):
            st.markdown(setup_quickbooks_oauth())
        
        st.stop()
    
    # Initialize QuickBooks API
    quickbooks_api = QuickBooksAPI(
        client_id=credentials['client_id'],
        client_secret=credentials['client_secret'],
        refresh_token=credentials['refresh_token'],
        company_id=credentials['company_id'],
        sandbox=credentials.get('sandbox', True)  # Default to sandbox
    )
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ QuickBooks Settings")
        environment = "Sandbox" if credentials.get('sandbox', True) else "Production"
        st.info(f"Environment: {environment}")
        
        # Test connection button
        if st.button("🔗 Test QuickBooks Connection"):
            with st.spinner("Testing connection..."):
                if quickbooks_api.authenticate():
                    st.success("✅ Connection successful!")
                else:
                    st.error("❌ Connection failed!")
    
    # Main form
    st.subheader("Invoice Information")
    
    # Get data from session state (from redirect) or URL parameters
    if 'quickbooks_data' in st.session_state:
        # Data from redirect page
        data = st.session_state['quickbooks_data']
        first_name_default = data.get('first_name', '')
        last_name_default = data.get('last_name', '')
        email_default = data.get('email', '')
        contract_amount_default = data.get('contract_amount', '')
        st.success("✅ Data loaded from Monday.com")
    else:
        # Fallback to URL parameters
        query_params = get_decoded_query_params()
        first_name_default = query_params.get('first_name', '')
        last_name_default = query_params.get('last_name', '')
        email_default = query_params.get('email', '')
        contract_amount_default = query_params.get('contract_amount', '')
    
    # Parse contract amount for Adjustments line item
    adjustments_amount = 0.00
    if contract_amount_default:
        try:
            # Remove $ and commas for parsing
            amount_str = contract_amount_default.replace('$', '').replace(',', '')
            adjustments_amount = float(amount_str)
        except ValueError:
            adjustments_amount = 0.00
    
    # Form fields with auto-filled values
    col1, col2 = st.columns(2)
    
    with col1:
        first_name = st.text_input(
            "First Name",
            value=first_name_default,
            help="Enter the client's first name"
        )
        
        email = st.text_input(
            "Email Address",
            value=email_default,
            help="Enter the client's email address"
        )
        
        company_name = st.text_input(
            "Company Name",
            value="",
            help="Enter the client's company name (optional)"
        )
        
        # Credit Card Processing Fee checkbox
        include_cc_fee = st.checkbox(
            "Include 3% Credit Card Processing Fee",
            help="Add a 3% processing fee for credit card payments"
        )

    
    with col2:
        last_name = st.text_input(
            "Last Name", 
            value=last_name_default,
            help="Enter the client's last name"
        )
        
        invoice_date = st.date_input(
            "Invoice Date",
            value=st.session_state.get('current_date', None),
            help="Date of the invoice"
        )
        # Enable payment link checkbox
        enable_payment_link = st.checkbox(
            "Enable Online Payment Link (ACH Only)",
            value=True,
            help="Allow client to pay online through QuickBooks via ACH"
        )
    
    # Client address field
    client_address = st.text_area(
        "Client Address",
        value="",
        height=15,
        help="Enter the client's billing address",
        placeholder="Street Address, City, State ZIP, Country"
    )
    
    # Invoice Details (Line Items)
    st.subheader("📝 Invoice Details")
    
    # Initialize session state for line items if not exists
    if 'line_items' not in st.session_state:
        st.session_state['line_items'] = []
        
        # Auto-add Adjustments line item if contract amount is provided
        if adjustments_amount > 0:
            st.session_state['line_items'].append({
                'type': 'Adjustments',
                'name': 'Contract Amount',
                'quantity': 1,
                'amount': adjustments_amount
            })
            st.info(f"✅ Auto-added Adjustments line item with contract amount: ${adjustments_amount:,.2f}")
    
    # Display existing line items
    if st.session_state['line_items']:
        st.markdown("**<span style='font-size: 18px'>Current Line Items:</span>**", unsafe_allow_html=True)
        for i, item in enumerate(st.session_state['line_items']):
            col1, col2, col3, col4 = st.columns([3, 1, 2, 1])
            with col1:
                # Show fee type as main name, optional description if provided
                item_name = item.get('type', 'Line Item')
                item_desc = item.get('name', '')
                display_text = f"{item_name}" + (f" ({item_desc})" if item_desc else "")
                st.markdown(f"<span style='font-size: 16px'>• {display_text}</span>", unsafe_allow_html=True)
            with col2:
                st.markdown(f"<span style='font-size: 16px'>Qty: {item['quantity']}</span>", unsafe_allow_html=True)
            with col3:
                item_total = item['amount'] * item['quantity']
                if item['quantity'] > 1:
                    st.write(f"$ {item['amount']:,.2f} × {item['quantity']} = $ {item_total:,.2f}")
                else:
                    st.write(f"$ {item_total:,.2f}")
            with col4:
                if st.button("🗑️", key=f"remove_{i}", help="Remove this line item"):
                    st.session_state['line_items'].pop(i)
                    st.rerun()
    
    # Add new line item interface
    col1, col2, col3, col4, col5 = st.columns([2, 2, 1, 2, 1])
    
    # Default values for each fee type (all set to 0.00 except Adjustments)
    fee_defaults = {
        "Adjustments": adjustments_amount,  # Use contract amount from URL
        "Bagging & Tagging": 0.00,
        "Consulting": 0.00,
        "Costing": 0.00,
        "Cutting": 0.00,
        "Design": 0.00,
        "Development - LA": 0.00,
        "Digitizing": 0.00,
        "Fabric": 0.00,
        "Fitting": 0.00,
        "Grading": 0.00,
        "Marketing": 0.00,
        "Patternmaking": 0.00,
        "Pre-production": 0.00,
        "Print-outs": 0.00,
        "Production COD": 0.00,
        "Production Sewing": 0.00,
        "Sample Production": 0.00,
        "Send Out": 0.00,
        "Shipping/Freight": 0.00,
        "Sourcing": 0.00,
        "Trim": 0.00
    }
    
    with col1:
        fee_types = [
            "Adjustments", "Bagging & Tagging", "Consulting", "Costing", "Cutting",
            "Design", "Development - LA", "Digitizing", "Fabric", "Fitting",
            "Grading", "Marketing", "Patternmaking", "Pre-production", "Print-outs",
            "Production COD", "Production Sewing", "Sample Production", "Send Out",
            "Shipping/Freight", "Sourcing", "Trim"
        ]
        
        new_fee_type = st.selectbox(
            "Line Item Name",
            options=fee_types,
            index=0,  # Default to "Adjustments" (first item)
            key="new_fee_type",
            help="Select the type of service (will appear as the line item name on invoice)"
        )
    
    with col2:
        new_line_description = st.text_input(
            "Line Item Description",
            value="",
            placeholder="e.g., Custom details about this item",
            key="new_line_description",
            help="Optional description that will appear below the line item name on invoice"
        )
    
    with col3:
        new_fee_quantity = st.number_input(
            "Quantity",
            min_value=1,
            value=1,
            step=1,
            key="new_fee_quantity",
            help="Enter the quantity for this line item"
        )
    
    with col4:
        # Get default amount for selected fee type
        default_amount = fee_defaults.get(new_fee_type, 0.00)
        
        new_fee_amount = st.number_input(
            "Amount ($)",
            min_value=0.00,
            value=default_amount,
            step=0.01,
            format="%.2f",
            key="new_fee_amount",
            help=f"Enter the dollar amount"
        )
    
    with col5:
        st.markdown("<div style='height: 30px'></div>", unsafe_allow_html=True)
        if st.button("➕ Add", help="Add this line item to the invoice"):
            if new_fee_amount > 0:
                st.session_state['line_items'].append({
                    'type': new_fee_type,  # Fee type is now the main name
                    'name': new_line_description,  # Description (optional)
                    'quantity': new_fee_quantity,
                    'amount': new_fee_amount
                })
                st.rerun()
            else:
                st.warning("Please enter an amount greater than $0.00")
    
    
    
    # Credits and Discounts Section
    st.subheader("💳 Credits & Discounts")
    
    # Initialize session state for credits if not exists
    if 'credits' not in st.session_state:
        st.session_state['credits'] = []
    
    # Display existing credits
    if st.session_state['credits']:
        st.markdown("**<span style='font-size: 18px'>Applied Credits:</span>**", unsafe_allow_html=True)
        for i, credit in enumerate(st.session_state['credits']):
            col1, col2, col3 = st.columns([3, 2, 1])
            with col1:
                st.markdown(f"<span style='font-size: 16px; color: green'>• {credit['description']}</span>", unsafe_allow_html=True)
            with col2:
                st.markdown(f"<span style='font-size: 16px; color: green'>-$ {credit['amount']:,.2f}</span>", unsafe_allow_html=True)
            with col3:
                if st.button("🗑️", key=f"remove_credit_{i}", help="Remove this credit"):
                    st.session_state['credits'].pop(i)
                    st.rerun()
    
    # Add new credit interface
    col1, col2, col3 = st.columns([2, 2, 1])
    
    with col1:
        new_credit_description = st.text_input(
            "Credit Description",
            value="",
            placeholder="e.g., Promotional Discount, Account Credit",
            key="new_credit_description",
            help="Description for the credit/discount"
        )
    
    with col2:
        new_credit_amount = st.number_input(
            "Credit Amount ($)",
            min_value=0.00,
            value=0.00,
            step=0.01,
            format="%.2f",
            key="new_credit_amount",
            help="Enter the credit/discount amount (will be subtracted from total)"
        )
    
    with col3:
        st.markdown("<div style='height: 30px'></div>", unsafe_allow_html=True)
        if st.button("➕ Add Credit", help="Add this credit to the invoice"):
            if new_credit_amount > 0 and new_credit_description:
                st.session_state['credits'].append({
                    'description': new_credit_description,
                    'amount': new_credit_amount
                })
                st.rerun()
            elif not new_credit_description:
                st.warning("Please enter a credit description")
            else:
                st.warning("Please enter a credit amount greater than $0.00")
    
    # Calculate and display total
    st.subheader("Invoice Summary")
    
    # Calculate totals from line items only
    additional_items_total = sum(item['amount'] * item['quantity'] for item in st.session_state['line_items'])
    credits_total = sum(credit['amount'] for credit in st.session_state.get('credits', []))
    
    # Calculate credit card fee if enabled
    cc_fee = 0
    if include_cc_fee and additional_items_total > 0:
        cc_fee = additional_items_total * 0.03
    
    total_amount = additional_items_total + cc_fee - credits_total
    
    # Display breakdown
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**<span style='font-size: 18px'>Amount Breakdown:</span>**", unsafe_allow_html=True)
        if st.session_state['line_items']:
            for item in st.session_state['line_items']:
                item_total = item['amount'] * item['quantity']
                item_name = item.get('type', 'Line Item')
                item_desc = item.get('name', '')
                display_text = f"{item_name}" + (f" ({item_desc})" if item_desc else "")
                if item['quantity'] > 1:
                    st.write(f"• {display_text}: $ {item['amount']:,.2f} × {item['quantity']} = $ {item_total:,.2f}")
                else:
                    st.write(f"• {display_text}: $ {item_total:,.2f}")
        else:
            st.write("• No line items added yet")
        
        # Display credit card fee if enabled
        if include_cc_fee and cc_fee > 0:
            st.write(f"• Credit Card Processing Fee (3%): $ {cc_fee:,.2f}")
        
        # Display credits
        for credit in st.session_state.get('credits', []):
            st.markdown(f"<span style='color: green'>• {credit['description']}: -$ {credit['amount']:,.2f}</span>", unsafe_allow_html=True)
    
    with col2:
        st.subheader(f"**Amount Due: $ {total_amount:,.2f}**")
    
    # Validation
    if not all([first_name, last_name, email]):
        st.warning("⚠️ Please fill in all required fields before proceeding.")
        return
    
    # Email validation
    if "@" not in email or "." not in email:
        st.error("❌ Please enter a valid email address.")
        return
    
    # Check if at least one line item is added
    if not st.session_state.get('line_items', []):
        st.error("❌ Please add at least one line item to create an invoice.")
        return
    
    # Action button
    st.subheader("🚀 Actions")
    
    if st.button("💰 Create & Send Invoice", type="primary", use_container_width=False):
        with st.spinner("Creating and sending invoice..."):
            # Prepare line items data
            line_items_data = []
            
            # Add line items from the form
            for item in st.session_state['line_items']:
                item_type = item.get('type', 'Line Item')  # Fee type (main name)
                item_description = item.get('name', '')  # Optional description
                line_items_data.append({
                    'type': item_type,  # Fee type as the main item name
                    'amount': item['amount'],
                    'quantity': item['quantity'],
                    'description': item_type,
                    'line_description': item_description  # Optional description
                })
            
            # Add credit card processing fee if enabled
            if include_cc_fee:
                # Calculate fee based on line items total
                line_items_total = sum(item['amount'] * item['quantity'] for item in st.session_state['line_items'])
                if line_items_total > 0:
                    cc_fee_amount = line_items_total * 0.03
                    line_items_data.append({
                        'type': 'Credit Card Processing Fee',
                        'amount': cc_fee_amount,
                        'quantity': 1,
                        'description': '3% Credit Card Processing Fee',
                        'line_description': 'Processing fee for credit card payments'
                    })
            
            # Add credits as negative line items
            for credit in st.session_state.get('credits', []):
                line_items_data.append({
                    'type': 'Credits & Discounts',  # Item name
                    'amount': -credit['amount'],  # Negative amount for credit
                    'quantity': 1,
                    'description': credit['description'],  # Credit description
                    'line_description': credit['description']  # Shows as description on invoice
                })
            
            # Create and send invoice
            success, message = quickbooks_api.create_and_send_invoice(
                first_name=first_name,
                last_name=last_name,
                email=email,
                company_name=company_name if company_name else None,
                client_address=client_address if client_address else None,
                contract_amount="0",  # No base contract amount
                description="Invoice",
                line_items=line_items_data,
                payment_terms="Due on receipt",
                enable_payment_link=enable_payment_link,
                invoice_date=invoice_date
            )
            
            if success:
                st.success(f"✅ {message}")
                if enable_payment_link:
                    st.success("🔗 Payment link included in invoice - customer can pay online via ACH")
                st.balloons()
                # Clear line items and credits after successful creation
                st.session_state['line_items'] = []
                st.session_state['credits'] = []
            else:
                st.error(f"❌ {message}")

if __name__ == "__main__":
    main()
