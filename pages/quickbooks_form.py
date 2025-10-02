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
    
    with col2:
        last_name = st.text_input(
            "Last Name", 
            value=last_name_default,
            help="Enter the client's last name"
        )
        
        contract_amount = st.text_input(
            "Contract Amount",
            value=contract_amount_default,
            help="Enter the contract amount (e.g., $10,000)"
        )
    
    # Additional invoice details
    st.subheader("Invoice Details")
    
    col1, col2 = st.columns(2)
    
    with col1:
        line_item_name = st.text_input(
            "Line Item Name",
            value="Contract Services",
            help="Name of the service/product (appears on invoice)"
        )
        
        # Add description field for main line item
        main_line_description = st.text_area(
            "Description (Optional)",
            value="",
            height=80,
            help="Additional description for the main contract line item",
            placeholder="E.g., Development services for Spring 2025 collection"
        )
        
        # Payment terms - fixed to "Due on receipt" (hidden from UI)
        custom_terms = "Due on receipt"
    
    with col2:
        invoice_date = st.date_input(
            "Invoice Date",
            value=st.session_state.get('current_date', None),
            help="Date of the invoice"
        )
        
        # Enable payment link checkbox
        enable_payment_link = st.checkbox(
            "Enable Online Payment Link",
            value=True,
            help="Allow client to pay online through QuickBooks"
        )
        
        if enable_payment_link:
            st.info("✅ Invoice will include payment link for credit card and ACH payments")
    
    # Credit Card Processing Fee
    st.subheader("Payment Processing")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        include_cc_fee = st.checkbox(
            "Include 3% Credit Card Processing Fee",
            help="Add a 3% processing fee for credit card payments"
        )
    
    with col2:
        if include_cc_fee and contract_amount:
            try:
                amount_str = contract_amount.replace('$', '').replace(',', '')
                base_amount = float(amount_str)
                cc_fee_amount = base_amount * 0.03
                st.info(f"**CC Fee:** ${cc_fee_amount:,.2f}")
            except ValueError:
                st.warning("Enter valid contract amount to calculate CC fee")
    
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
    
    # Additional Line Items
    st.subheader("📝 Additional Line Items")
    
    # Initialize session state for line items if not exists
    if 'line_items' not in st.session_state:
        st.session_state['line_items'] = []
    
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
    
    # Default values for each fee type
    fee_defaults = {
        "Adjustments": 50.00,
        "Bagging & Tagging": 25.00,
        "Consulting": 150.00,
        "Costing": 75.00,
        "Cutting": 100.00,
        "Design": 200.00,
        "Development - LA": 300.00,
        "Digitizing": 125.00,
        "Fabric": 0.00,
        "Fitting": 80.00,
        "Grading": 90.00,
        "Marketing": 175.00,
        "Patternmaking": 120.00,
        "Pre-production": 150.00,
        "Print-outs": 30.00,
        "Production COD": 200.00,
        "Production Sewing": 180.00,
        "Sample Production": 250.00,
        "Send Out": 40.00,
        "Shipping/Freight": 0.00,
        "Sourcing": 100.00,
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
            "Fee Type",
            options=fee_types,
            key="new_fee_type",
            help="Select the type of fee (will appear as the line item name on invoice)"
        )
    
    with col2:
        new_line_description = st.text_input(
            "Line Item Description",
            value="",
            placeholder="e.g., Custom details about this item",
            key="new_line_description",
            help="Optional description that will appear below the fee type on invoice"
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
    
    # Calculate and display total
    st.subheader("Invoice Summary")
    
    if contract_amount:
        try:
            amount_str = contract_amount.replace('$', '').replace(',', '')
            base_amount = float(amount_str)
            
            # Calculate totals
            subtotal = base_amount
            cc_fee = base_amount * 0.03 if include_cc_fee else 0
            additional_items_total = sum(item['amount'] * item['quantity'] for item in st.session_state['line_items'])
            credits_total = sum(credit['amount'] for credit in st.session_state.get('credits', []))
            total_amount = subtotal + cc_fee + additional_items_total - credits_total
            
            # Display breakdown
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**<span style='font-size: 18px'>Amount Breakdown:</span>**", unsafe_allow_html=True)
                st.write(f"• Contract Amount: $ {subtotal:,.2f}")
                if include_cc_fee:
                    st.write(f"• Credit Card Fee (3%): $ {cc_fee:,.2f}")
                for item in st.session_state['line_items']:
                    item_total = item['amount'] * item['quantity']
                    item_name = item.get('type', 'Line Item')
                    item_desc = item.get('name', '')
                    display_text = f"{item_name}" + (f" ({item_desc})" if item_desc else "")
                    if item['quantity'] > 1:
                        st.write(f"• {display_text}: $ {item['amount']:,.2f} × {item['quantity']} = $ {item_total:,.2f}")
                    else:
                        st.write(f"• {display_text}: $ {item_total:,.2f}")
                
                # Display credits
                for credit in st.session_state.get('credits', []):
                    st.markdown(f"<span style='color: green'>• {credit['description']}: -$ {credit['amount']:,.2f}</span>", unsafe_allow_html=True)
            
            with col2:
                st.subheader(f"**Amount Due: $ {total_amount:,.2f}**")
                
        except ValueError:
            st.warning("Enter valid contract amount to calculate totals")
    
    # Validation
    if not all([first_name, last_name, email, contract_amount]):
        st.warning("⚠️ Please fill in all required fields before proceeding.")
        return
    
    # Email validation
    if "@" not in email or "." not in email:
        st.error("❌ Please enter a valid email address.")
        return
    
    # Contract amount validation
    try:
        # Remove $ and commas for validation
        amount_str = contract_amount.replace('$', '').replace(',', '')
        amount = float(amount_str)
        if amount <= 0:
            st.error("❌ Contract amount must be greater than zero.")
            return
    except ValueError:
        st.error("❌ Please enter a valid contract amount (e.g., $10,000).")
        return
    
    # Action button
    st.subheader("🚀 Actions")
    
    if st.button("💰 Create & Send Invoice", type="primary", use_container_width=False):
        with st.spinner("Creating and sending invoice..."):
            # Prepare line items data
            line_items_data = []
            
            # Add main contract line item
            try:
                amount_str = contract_amount.replace('$', '').replace(',', '')
                base_amount = float(amount_str)
                line_items_data.append({
                    'type': line_item_name,
                    'amount': base_amount,
                    'description': line_item_name,
                    'line_description': main_line_description if main_line_description else ''
                })
                
                # Add credit card fee if selected
                if include_cc_fee:
                    cc_fee_amount = base_amount * 0.03
                    line_items_data.append({
                        'type': 'Credit Card Processing Fee',
                        'amount': cc_fee_amount,
                        'description': '3% Credit Card Processing Fee'
                    })
                
                # Add additional line items
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
                
                # Add credits as negative line items
                for credit in st.session_state.get('credits', []):
                    line_items_data.append({
                        'type': 'Credits & Discounts',  # Item name
                        'amount': -credit['amount'],  # Negative amount for credit
                        'quantity': 1,
                        'description': credit['description'],  # Credit description
                        'line_description': credit['description']  # Shows as description on invoice
                    })
                
            except ValueError:
                st.error("❌ Invalid contract amount format")
                return
            
            # Create and send invoice
            success, message = quickbooks_api.create_and_send_invoice(
                first_name=first_name,
                last_name=last_name,
                email=email,
                contract_amount=contract_amount,
                description=line_item_name,
                line_items=line_items_data,
                payment_terms=custom_terms,
                enable_payment_link=enable_payment_link,
                invoice_date=invoice_date
            )
            
            if success:
                st.success(f"✅ {message}")
                if enable_payment_link:
                    st.success("🔗 Payment link included in invoice - customer can pay online via credit card or ACH")
                st.balloons()
                # Clear line items and credits after successful creation
                st.session_state['line_items'] = []
                st.session_state['credits'] = []
            else:
                st.error(f"❌ {message}")

if __name__ == "__main__":
    main()
