"""
SEO Metrics Dashboard
Embeds Google Analytics reporting dashboard
"""
import streamlit as st

# Page configuration
st.set_page_config(
    page_title="SEO Metrics",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
<style>
    /* Hide QuickBooks and SignNow pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"] {
        display: none !important;
    }
    
    .embed-container {
        position: relative;
        padding-bottom: 100%;
        height: 0;
        overflow: hidden;
        max-width: 100%;
        background: #f0f0f0;
    }
    
    .embed-container iframe,
    .embed-container object,
    .embed-container embed {
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        border: none;
    }
    
    .stMain .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
    }
    
    h1 {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main SEO Metrics Dashboard"""
    
    # Header
    st.markdown('<div style="padding-bottom: 1rem;"><h1>📈 SEO Metrics Dashboard</h1></div>', unsafe_allow_html=True)
    
    # Google Analytics embed URL
    ga_url = "https://analytics.google.com/analytics/web/?authuser=2#/a252660567p347392975/reports/intelligenthome"
    
    st.markdown("---")
    
    # Display the Google Analytics dashboard in an iframe
    st.markdown("### 📊 Google Analytics Intelligence Home Report")
    st.warning("⚠️ **Note:** Due to Google Analytics security policies (X-Frame-Options), the dashboard cannot be embedded directly.")
    
    # Create iframe with appropriate attributes
    st.components.v1.html("""
    <iframe 
        src="https://analytics.google.com/analytics/web/?authuser=2#/a252660567p347392975/reports/intelligenthome" 
        width="100%" 
        height="900" 
        frameborder="0"
        allowfullscreen
        style="border: 2px solid #1f77b4; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
    </iframe>
    
    <script>
        // Configure GA4 cookies for cross-origin iframe support
        if (typeof gtag !== 'undefined') {
            gtag('config', 'GA_MEASUREMENT_ID', {
                cookie_flags: 'SameSite=None;Secure',
                cookie_update: false
            });
        }
    </script>
    """, height=920)
    
    st.markdown("---")

if __name__ == "__main__":
    main()

