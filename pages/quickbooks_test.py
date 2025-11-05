"""
QuickBooks Production Credentials Test Page
Simple form to test QuickBooks production credentials without modifying secrets.toml
"""

import streamlit as st
import sys
import os

# Add parent directory to path to import quickbooks_integration
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from quickbooks_integration import QuickBooksAPI

# Page configuration
st.set_page_config(
    page_title="QuickBooks Credentials Test",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to hide sidebar navigation
st.markdown("""
<style>
    [data-testid="stSidebarNav"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main QuickBooks credentials test function"""
    st.title("🔍 QuickBooks Production Credentials Test")
    st.markdown("Test QuickBooks production credentials without modifying secrets.toml")
    st.info("💡 This page is for testing only. Credentials are not saved.")
    
    st.divider()
    
    # Form fields
    st.subheader("📝 Enter Production Credentials")
    
    col1, col2 = st.columns(2)
    
    with col1:
        client_id = st.text_input(
            "Client ID",
            value="",
            help="QuickBooks OAuth Client ID",
            type="default"
        )
        
        refresh_token = st.text_input(
            "Refresh Token",
            value="",
            help="QuickBooks OAuth Refresh Token",
            type="password"
        )
    
    with col2:
        client_secret = st.text_input(
            "Client Secret",
            value="",
            help="QuickBooks OAuth Client Secret",
            type="password"
        )
        
        company_id = st.text_input(
            "Company ID",
            value="",
            help="QuickBooks Company ID (also known as Realm ID)",
            type="default"
        )
    
    st.divider()
    
    # Test button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        test_button = st.button(
            "🔗 Test Connection",
            type="primary",
            use_container_width=True
        )
    
    # Test connection
    if test_button:
        # Validate inputs
        if not all([client_id, client_secret, refresh_token, company_id]):
            st.error("❌ Please fill in all fields before testing.")
            return
        
        # Initialize QuickBooks API with test credentials (sandbox=False for production)
        with st.spinner("Testing QuickBooks production connection..."):
            try:
                quickbooks_api = QuickBooksAPI(
                    client_id=client_id,
                    client_secret=client_secret,
                    refresh_token=refresh_token,
                    company_id=company_id,
                    sandbox=False  # Production mode
                )
                
                # Test authentication
                if quickbooks_api.authenticate():
                    st.success("✅ Connection successful!")
                    
                    # Try to verify company access
                    try:
                        if quickbooks_api._verify_company_access():
                            st.info("✅ Successfully verified company access")
                            st.success("🎉 All credentials are valid and working!")
                        else:
                            st.warning("⚠️ Authentication succeeded but company access verification failed")
                    except Exception as e:
                        st.warning(f"⚠️ Connection works but couldn't verify company access: {str(e)}")
                
                else:
                    st.error("❌ Connection failed!")
                    st.error("Please check your credentials and try again.")
                    
            except Exception as e:
                st.error(f"❌ Error testing connection: {str(e)}")
                st.error("Please verify your credentials are correct.")

if __name__ == "__main__":
    main()

