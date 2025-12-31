"""
Employee Tools Hub
Central hub for accessing employee tools: Sign Now and Workbook Creator
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
[data-testid="stSidebarNav"] li:has(a[href*="signnow"]),
[data-testid="stSidebarNav"] li:has(a[href*="/tools"]),
[data-testid="stSidebarNav"] li:has(a[href*="workbook"]),
[data-testid="stSidebarNav"] li:has(a[href*="deck_creator"]),
[data-testid="stSidebarNav"] li:has(a[href*="a_la_carte"]) {
    display: block !important;
}
</style>
<script>
// JavaScript to show only tool pages and hide everything else
(function() {
    function showToolPagesOnly() {
        const navItems = document.querySelectorAll('[data-testid="stSidebarNav"] li');
        const allowedPages = ['signnow', 'tools', 'workbook', 'deck_creator', 'a_la_carte'];
        
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
            
            // Hide a_la_carte if we're on an ads dashboard page
            const isDevInspection = href.includes('a_la_carte') || text.includes('a_la_carte');
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
    """Main tools hub function"""
    st.title("🔧 Employee Tools Hub")
    st.markdown("Access employee tools for contracts, invoices, and workbooks")
    
    # Create columns for tool cards
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("#### 📝 Sign Now")
        st.markdown("Send contracts for signature")
        if st.button("🚀 Open Sign Now", key="signnow_button", use_container_width=True):
            st.switch_page("pages/signnow_form.py")
    
    with col2:
        st.markdown("#### 📊 Workbook Creator")
        st.markdown("Generate Excel workbooks")
        if st.button("🚀 Open Workbook Creator", key="workbook_button", use_container_width=True):
            st.switch_page("pages/workbook_creator.py")

    with col3:
        st.markdown("#### 📽️ Deck Creator")
        st.markdown("Generate Google Slides decks")
        if st.button("🚀 Open Deck Creator", key="deck_button", use_container_width=True):
            st.switch_page("pages/deck_creator.py")
    
    with col4:
        st.markdown("#### 🔍 A La Carte")
        st.markdown("A La Carte workbooks")
        if st.button("🚀 Open A La Carte", key="a_la_carte_button", use_container_width=True):
            st.switch_page("pages/a_la_carte.py")

if __name__ == "__main__":
    main()

