"""
SEO Metrics Dashboard
Embeds Looker Studio reporting dashboard
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
    /* Hide tool pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="tools"],
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"],
    [data-testid="stSidebarNav"] a[href*="workbook_creator"],
    [data-testid="stSidebarNav"] a[href*="deck_creator"],
    [data-testid="stSidebarNav"] a[href*="dev_inspection"] {
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
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main SEO Metrics Dashboard"""
    
    # Header
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown('<div style="padding-bottom: 1rem;"><h1>📈 SEO Metrics Dashboard</h1></div>', unsafe_allow_html=True)
    
    # Display goal metrics image below the title and before the dashboard
    col_img1, col_img2 = st.columns(2, gap="medium")
    with col_img1:
        st.image("inputs/goal_metrics.png", use_container_width=True)
    with col_img2:
        st.image("inputs/image (11).png", use_container_width=True)
    
    st.markdown("---")
    
    # Display the Looker Studio dashboard in an iframe
    
    # Create iframe with Looker Studio embed code
    st.components.v1.html("""
    <iframe 
        width="100%" 
        height="900" 
        src="https://lookerstudio.google.com/embed/reporting/049afda9-f443-4e7f-94de-ea58ffa143a2/page/ZQadF" 
        frameborder="0" 
        style="border:0" 
        allowfullscreen 
        sandbox="allow-storage-access-by-user-activation allow-scripts allow-same-origin allow-popups allow-popups-to-escape-sandbox">
    </iframe>
    """, height=920)
    st.components.v1.html("""
    <iframe 
        src="https://docs.google.com/spreadsheets/d/11ZymjS8bUognebcbbXLTgr9vViXTWk6yrct0HOLBTCQ/edit?usp=sharing&embedded=true" 
        width="100%" 
        height="800" 
        frameborder="0" 
        marginheight="0" 
        marginwidth="0">
        Loading...
    </iframe>
    """, height=820)
    
    st.markdown("---")

if __name__ == "__main__":
    main()

