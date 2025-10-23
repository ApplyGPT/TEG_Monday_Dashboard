import streamlit as st

# Page configuration
st.set_page_config(
    page_title="Katya SEO Tracking",
    page_icon="📊",
    layout="wide"
)

# Custom CSS to hide QuickBooks and SignNow pages from sidebar
st.markdown("""
<style>
    /* Hide QuickBooks and SignNow pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.title("📊 Katya SEO Tracking")
    
    # Google Sheets embed URL
    embed_url = "https://docs.google.com/spreadsheets/d/11ZymjS8bUognebcbbXLTgr9vViXTWk6yrct0HOLBTCQ/edit?gid=0#gid=0"
    
    # Convert to embed format
    embed_id = "11ZymjS8bUognebcbbXLTgr9vViXTWk6yrct0HOLBTCQ"
    embed_src = f"https://docs.google.com/spreadsheets/d/{embed_id}/edit?usp=sharing&embedded=true"
    
    # Embed the Google Sheet
    st.markdown(f"""
    <iframe 
        src="{embed_src}" 
        width="100%" 
        height="800" 
        frameborder="0" 
        marginheight="0" 
        marginwidth="0">
        Loading...
    </iframe>
    """, unsafe_allow_html=True)
    
    st.markdown("---")

if __name__ == "__main__":
    main()
