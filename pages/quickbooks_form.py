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
    
    # Get data from URL parameters (if coming from Monday.com)
    query_params = st.query_params
    
    # Form fields with auto-filled values from URL parameters
    col1, col2 = st.columns(2)
    
    with col1:
        first_name = st.text_input(
            "First Name",
            value=query_params.get('first_name', ''),
            help="Enter the client's first name"
        )
        
        email = st.text_input(
            "Email Address",
            value=query_params.get('email', ''),
            help="Enter the client's email address"
        )
    
    with col2:
        last_name = st.text_input(
            "Last Name", 
            value=query_params.get('last_name', ''),
            help="Enter the client's last name"
        )
        
        contract_amount = st.text_input(
            "Contract Amount",
            value=query_params.get('contract_amount', ''),
            help="Enter the contract amount (e.g., $10,000)"
        )
    
    # Additional invoice details
    st.subheader("Invoice Details")
    
    col1, col2 = st.columns(2)
    
    with col1:
        description = st.text_area(
            "Service Description",
            value="Contract Services",
            help="Description of the services provided"
        )
        
        due_days = st.number_input(
            "Payment Terms (Days)",
            min_value=1,
            max_value=365,
            value=30,
            help="Number of days until payment is due"
        )
    
    with col2:
        invoice_date = st.date_input(
            "Invoice Date",
            value=st.session_state.get('current_date', None),
            help="Date of the invoice"
        )
        
        # Calculate due date
        if invoice_date:
            from datetime import timedelta
            due_date = invoice_date + timedelta(days=due_days)
            st.info(f"**Due Date:** {due_date.strftime('%Y-%m-%d')}")
    
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
    
    # Display summary
    st.subheader("📋 Invoice Summary")
    
    summary_col1, summary_col2 = st.columns(2)
    
    with summary_col1:
        st.info(f"**Client:** {first_name} {last_name}")
        st.info(f"**Email:** {email}")
        st.info(f"**Invoice Date:** {invoice_date}")
    
    with summary_col2:
        st.info(f"**Amount:** ${contract_amount}")
        st.info(f"**Description:** {description}")
        st.info(f"**Due Date:** {due_date}")
    
    # Action buttons
    st.subheader("🚀 Actions")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("💰 Create & Send Invoice", type="primary", use_container_width=True):
            with st.spinner("Creating and sending invoice..."):
                # Create and send invoice
                success, message = quickbooks_api.create_and_send_invoice(
                    first_name=first_name,
                    last_name=last_name,
                    email=email,
                    contract_amount=contract_amount,
                    description=description
                )
                
                if success:
                    st.success(f"✅ {message}")
                    st.balloons()
                else:
                    st.error(f"❌ {message}")
    
    with col2:
        if st.button("📄 Preview Invoice", use_container_width=True):
            st.info("📄 Invoice preview would be displayed here")
            st.markdown(f"""
            **Invoice Preview:**
            
            **Bill To:** {first_name} {last_name}
            **Email:** {email}
            **Invoice Date:** {invoice_date}
            **Due Date:** {due_date}
            **Amount:** ${contract_amount}
            **Description:** {description}
            
            *This is a preview. The actual invoice will be created in QuickBooks.*
            """)

if __name__ == "__main__":
    main()
