"""
Employee Tools Hub
Central hub for accessing employee tools: Sign Now, QuickBooks, and Workbook Creator
"""

import streamlit as st

# Page configuration
st.set_page_config(
    page_title="Employee Tools",
    page_icon="🔧",
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
    """Main tools hub function"""
    st.title("🔧 Employee Tools Hub")
    st.markdown("Access employee tools for contracts, invoices, and workbooks")
    
    # Create columns for tool cards
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown("#### 📝 Sign Now")
        st.markdown("Send contracts for signature")
        if st.button("🚀 Open Sign Now", key="signnow_button", use_container_width=True):
            st.switch_page("pages/signnow_form.py")
    
    with col2:
        st.markdown("#### 💰 QuickBooks")
        st.markdown("Send invoices to clients")
        if st.button("🚀 Open QuickBooks", key="quickbooks_button", use_container_width=True):
            st.switch_page("pages/quickbooks_form.py")
    
    with col3:
        st.markdown("#### 📊 Workbook Creator")
        st.markdown("Generate Excel workbooks")
        if st.button("🚀 Open Workbook Creator", key="workbook_button", use_container_width=True):
            st.switch_page("pages/workbook_creator.py")

    with col4:
        st.markdown("#### 📽️ Deck Creator")
        st.markdown("Generate Google Slides decks")
        if st.button("🚀 Open Deck Creator", key="deck_button", use_container_width=True):
            st.switch_page("pages/deck_creator.py")
    
    with col5:
        st.markdown("#### 🔍 Dev & Inspection")
        st.markdown("Dev & Inspection workbooks")
        if st.button("🚀 Open Dev & Inspection", key="dev_inspection_button", use_container_width=True):
            st.switch_page("pages/dev_inspection.py")

if __name__ == "__main__":
    main()

