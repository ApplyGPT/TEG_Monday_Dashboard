"""
SignNow Contract Form Page
Allows users to review and send contracts for signing
"""

import streamlit as st
import sys
import os

# Add parent directory to path to import signnow_integration
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from signnow_integration import SignNowAPI, load_signnow_credentials, create_sample_contract_template
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
    page_title="SignNow Contract Form",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

def main():
    """Main SignNow form function"""
    st.title("📝 SignNow Contract Form")
    st.markdown("Review and send contracts for signing")
    
    # Load SignNow credentials
    credentials = load_signnow_credentials()
    if not credentials:
        st.error("SignNow credentials not configured. Please check your secrets.toml file.")
        st.stop()
    
    # Initialize SignNow API
    signnow_api = SignNowAPI(
        client_id=credentials['client_id'],
        client_secret=credentials['client_secret'],
        basic_auth_token=credentials['basic_auth_token'],
        username=credentials['username'],
        password=credentials['password'],
        api_key=credentials.get('api_key')
    )
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ SignNow Settings")
        st.info("Configured with SignNow API")
        
        # Test connection button
        if st.button("🔗 Test SignNow Connection"):
            with st.spinner("Testing connection..."):
                if signnow_api.authenticate():
                    st.success("✅ Connection successful!")
                else:
                    st.error("❌ Connection failed!")
    
    # Main form
    st.subheader("Contract Information")
    
    # Get data from session state (from redirect) or URL parameters
    if 'signnow_data' in st.session_state:
        # Data from redirect page
        data = st.session_state['signnow_data']
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
    
    # Validation
    if not all([first_name, last_name, email, contract_amount]):
        st.warning("⚠️ Please fill in all fields before proceeding.")
        return
    
    # Email validation
    if "@" not in email or "." not in email:
        st.error("❌ Please enter a valid email address.")
        return
    
    # Contract amount validation
    try:
        # Remove $ and commas for validation
        amount_str = contract_amount.replace('$', '').replace(',', '')
        float(amount_str)
    except ValueError:
        st.error("❌ Please enter a valid contract amount (e.g., $10,000).")
        return
    
    # Action button
    st.subheader("🚀 Actions")
    
    if st.button("📝 Create & Send Contract", type="primary", use_container_width=False):
        with st.spinner("Creating and sending contract..."):
            # Get template ID
            template_id = create_sample_contract_template()
            
            # Create and send contract
            success, message = signnow_api.create_and_send_contract(
                first_name=first_name,
                last_name=last_name,
                email=email,
                contract_amount=contract_amount,
                template_id=template_id
            )
            
            if success:
                st.success(f"✅ {message}")
                st.balloons()
            else:
                st.error(f"❌ {message}")

if __name__ == "__main__":
    main()
