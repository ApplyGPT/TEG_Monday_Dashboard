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
    
    # Display summary
    st.subheader("📋 Contract Summary")
    
    summary_col1, summary_col2 = st.columns(2)
    
    with summary_col1:
        st.info(f"**Client:** {first_name} {last_name}")
        st.info(f"**Email:** {email}")
    
    with summary_col2:
        st.info(f"**Contract Amount:** ${contract_amount}")
        st.info(f"**Status:** Ready to send")
    
    # Action buttons
    st.subheader("🚀 Actions")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📝 Create & Send Contract", type="primary", use_container_width=True):
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
    
    with col2:
        if st.button("📄 Preview Contract", use_container_width=True):
            st.info("📄 Contract preview would be displayed here")
            st.markdown(f"""
            **Contract Preview:**
            
            **Client Name:** {first_name} {last_name}
            **Email:** {email}
            **Contract Amount:** ${contract_amount}
            **Date:** {st.session_state.get('current_date', 'Today')}
            
            *This is a sample contract template. The actual contract would be generated from your SignNow template.*
            """)

if __name__ == "__main__":
    main()
