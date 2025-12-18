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

# Custom CSS to hide all non-tool pages from sidebar navigation
st.markdown("""
<style>
/* Hide ALL sidebar list items by default */
[data-testid="stSidebarNav"] li {
    display: none !important;
}

/* Show list items that contain allowed tool pages using :has() selector */
[data-testid="stSidebarNav"] li:has(a[href*="quickbooks"]),
[data-testid="stSidebarNav"] li:has(a[href*="signnow"]),
[data-testid="stSidebarNav"] li:has(a[href*="/tools"]),
[data-testid="stSidebarNav"] li:has(a[href*="workbook"]),
[data-testid="stSidebarNav"] li:has(a[href*="deck_creator"]),
[data-testid="stSidebarNav"] li:has(a[href*="dev_inspection"]) {
    display: block !important;
}
</style>
<script>
// JavaScript to show only tool pages and hide everything else
(function() {
    function showToolPagesOnly() {
        const navItems = document.querySelectorAll('[data-testid="stSidebarNav"] li');
        const allowedPages = ['quickbooks', 'signnow', 'tools', 'workbook', 'deck_creator', 'dev_inspection'];
        
        // Check if we're currently on an ads dashboard page
        const currentUrl = window.location.href.toLowerCase();
        const currentPath = window.location.pathname.toLowerCase();
        const isOnAdsDashboard = currentUrl.includes('ads') && currentUrl.includes('dashboard') ||
                                 currentPath.includes('ads') && currentPath.includes('dashboard');
        
        navItems.forEach(item => {
            const link = item.querySelector('a');
            if (!link) {
                item.style.setProperty('display', 'none', 'important');
                return;
            }
            
            const href = (link.getAttribute('href') || '').toLowerCase();
            const text = link.textContent.trim().toLowerCase();
            
            // Check if this is an allowed tool page
            const isToolPage = allowedPages.some(page => {
                return href.includes(page) || text.includes(page.toLowerCase());
            });
            
            // Make sure it's not ads dashboard or other dashboards
            const isDashboard = (text.includes('ads') && text.includes('dashboard')) || 
                              (href.includes('ads') && href.includes('dashboard'));
            
            // Hide dev_inspection if we're on an ads dashboard page
            const isDevInspection = href.includes('dev_inspection') || text.includes('dev_inspection');
            if (isOnAdsDashboard && isDevInspection) {
                item.style.setProperty('display', 'none', 'important');
                return;
            }
            
            if (isToolPage && !isDashboard) {
                item.style.setProperty('display', 'block', 'important');
                link.style.setProperty('display', 'block', 'important');
            } else {
                item.style.setProperty('display', 'none', 'important');
            }
        });
    }
    
    // Run immediately and on load
    showToolPagesOnly();
    window.addEventListener('load', function() {
        setTimeout(showToolPagesOnly, 50);
        setTimeout(showToolPagesOnly, 200);
        setTimeout(showToolPagesOnly, 500);
    });
    
    // Watch for DOM changes
    const observer = new MutationObserver(function() {
        showToolPagesOnly();
    });
    
    setTimeout(function() {
        const sidebar = document.querySelector('[data-testid="stSidebarNav"]');
        if (sidebar) {
            observer.observe(sidebar, { 
                childList: true, 
                subtree: true,
                attributes: true
            });
        }
    }, 100);
})();
</script>
""", unsafe_allow_html=True)

def main():
    """Main QuickBooks form function"""
    st.title("💰 QuickBooks Invoice Form")
    st.markdown("Review and create invoices for clients")
    
    # Get data from session state (from redirect) or URL parameters
    # This must be done BEFORE any form fields that use these defaults
    item_id = ''
    if 'quickbooks_data' in st.session_state:
        # Data from redirect page
        data = st.session_state['quickbooks_data']
        first_name_default = data.get('first_name', '')
        last_name_default = data.get('last_name', '')
        email_default = data.get('email', '')
        contract_amount_default = data.get('contract_amount', '')
        cc_email_default = data.get('cc_email', '')
        company_name_default = data.get('company_name', '')
        item_id = data.get('item_id', '').strip()
        st.success("✅ Data loaded from Monday.com")
    else:
        # Fallback to URL parameters
        query_params = get_decoded_query_params()
        first_name_default = query_params.get('first_name', '')
        last_name_default = query_params.get('last_name', '')
        email_default = query_params.get('email', '')
        contract_amount_default = query_params.get('contract_amount', '')
        cc_email_default = query_params.get('cc_email', '')
        company_name_default = query_params.get('company_name', '')
        item_id = query_params.get('item_id', '').strip()
    
    # Salesman Email Address (CC) field
    cc_col1, cc_col2 = st.columns(2)
    with cc_col1:
        cc_email = st.text_input(
            "Salesman Email Address (CC)",
            value=cc_email_default,
            help="Email address to CC on the invoice (optional)"
        )
    with cc_col2:
        st.empty()  # Empty column to maintain layout
    
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
        sandbox=credentials.get('sandbox', False),  # Default to False for production
        access_token=credentials.get('access_token')  # Use access_token from secrets.toml if available
    )
    
    # Note: SSL verification is disabled to resolve QuickBooks API hostname mismatch issues
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ QuickBooks Settings")
        environment = "Sandbox" if credentials.get('sandbox', False) else "Production"
        st.info(f"Environment: {environment}")
        
        # Test connection button
        if st.button("🔗 Test QuickBooks Connection"):
            with st.spinner("Testing connection..."):
                if quickbooks_api.authenticate():
                    st.success("✅ Connection successful!")
                else:
                    st.error("❌ Connection failed!")
        
        # Verify production credentials button
        st.divider()
        if st.button("🔍 Verify Production Credentials"):
            with st.spinner("Verifying production credentials..."):
                from quickbooks_integration import verify_production_credentials
                is_production, message = verify_production_credentials(quickbooks_api)
                if is_production:
                    st.success(message)
                else:
                    st.error(message)
    
    # Main form
    st.subheader("Invoice Information")
    
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
            value=company_name_default,
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
            "Enable Online Payment Link",
            value=True,
            help="Allow client to pay online through QuickBooks."
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
    
    # Helper to build line items payload (used for preview and sending)
    def build_line_items_payload(include_credits=True):
        payload = []
        base_subtotal = 0.0
        
        for item in st.session_state['line_items']:
            quantity = item.get('quantity', 1) or 1
            unit_price = float(item.get('amount', 0) or 0)
            line_total = unit_price * quantity
            base_subtotal += line_total
            
            payload.append({
                'type': item.get('type', 'Line Item'),
                'amount': line_total,
                'quantity': quantity,
                'unit_price': unit_price,
                'description': item.get('type', 'Line Item'),
                'line_description': item.get('name', '')
            })
        
        cc_fee_amount = 0.0
        if include_cc_fee and base_subtotal > 0:
            cc_fee_amount = round(base_subtotal * 0.03, 2)
            payload.append({
                'type': 'Credit Card Processing Fee (3%)',
                'amount': cc_fee_amount,
                'quantity': 1,
                'unit_price': cc_fee_amount,
                'description': '3% Credit Card Processing Fee',
                'line_description': 'Processing fee for credit card payments'
            })
        
        credits_total = 0.0
        if include_credits:
            for credit in st.session_state.get('credits', []):
                amount = float(credit.get('amount', 0) or 0)
                credits_total += amount
                payload.append({
                    'type': 'Credits & Discounts',
                    'amount': -amount,
                    'quantity': 1,
                    'unit_price': -amount,
                    'description': credit.get('description', 'Discount'),
                    'line_description': credit.get('description', 'Discount')
                })
        
        return payload, base_subtotal, cc_fee_amount, credits_total
    
    # Build preview payload and totals
    preview_line_items, base_subtotal, cc_fee_amount, credits_total = build_line_items_payload()
    total_amount = round(base_subtotal + cc_fee_amount - credits_total, 2)
    
    # Calculate and display total
    st.subheader("Invoice Summary")
    
    # Display breakdown
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**<span style='font-size: 18px'>Amount Breakdown:</span>**", unsafe_allow_html=True)
        if st.session_state['line_items']:
            for item in st.session_state['line_items']:
                quantity = item.get('quantity', 1) or 1
                unit_price = float(item.get('amount', 0) or 0)
                line_total = unit_price * quantity
                item_name = item.get('type', 'Line Item')
                item_desc = item.get('name', '')
                display_text = f"{item_name}" + (f" ({item_desc})" if item_desc else "")
                if item['quantity'] > 1:
                    st.write(f"• {display_text}: $ {unit_price:,.2f} × {item['quantity']} = $ {line_total:,.2f}")
                else:
                    st.write(f"• {display_text}: $ {unit_price:,.2f}")
        else:
            st.write("• No line items added yet")
        
        # Display credit card fee if enabled
        if include_cc_fee and cc_fee_amount > 0:
            st.write(f"• Credit Card Processing Fee (3%): $ {cc_fee_amount:,.2f}")
        
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
        st.warning("⚠️ Please add at least one line item to create an invoice.")
        return
    
    # Action button
    st.subheader("🚀 Actions")
    
    if st.button("💰 Create & Send Invoice", type="primary", use_container_width=False):
        with st.spinner("Creating and sending invoice..."):
            # Prepare line items data using the shared builder
            line_items_data, _, _, _ = build_line_items_payload(include_credits=True)
            
            # Create and send invoice
            success, message = quickbooks_api.create_and_send_invoice(
                first_name=first_name,
                last_name=last_name,
                email=email,  # Client email (not cc_email)
                company_name=company_name if company_name else None,
                client_address=client_address if client_address else None,
                contract_amount="0",  # No base contract amount
                description="Invoice",
                line_items=line_items_data,
                payment_terms="Due on receipt",
                enable_payment_link=enable_payment_link,
                invoice_date=invoice_date,
                cc_email=cc_email if cc_email else None,  # Salesman email for CC
                include_cc_fee=include_cc_fee  # Pass the CC fee checkbox state
            )
            
            if success:
                st.success(f"✅ {message}")
                if enable_payment_link:
                    print("🔗 Payment link included in invoice - customer can pay online via ACH")
                
                # Post update to Monday.com if item_id is provided
                if item_id:
                    from quickbooks_integration import create_monday_update
                    if create_monday_update(item_id, message):
                        print("✅ Update posted to Monday.com item!")
                    else:
                        st.warning("⚠️ Invoice created successfully, but failed to post update to Monday.com.")
                
                st.balloons()
                # Clear line items and credits after successful creation
                st.session_state['line_items'] = []
                st.session_state['credits'] = []
            else:
                st.error(f"❌ {message}")

if __name__ == "__main__":
    main()