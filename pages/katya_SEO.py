import streamlit as st
import pandas as pd

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

@st.cache_data
def load_seo_data():
    """Load SEO tracking data from Excel file"""
    try:
        # Load both sheets
        df_keywords = pd.read_excel('inputs/Copy of TEG Keyword Tracking - Katya.xlsx', sheet_name=0)
        df_pages = pd.read_excel('inputs/Copy of TEG Keyword Tracking - Katya.xlsx', sheet_name=1)
        
        return df_keywords, df_pages
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None, None

def format_keywords_dataframe(df_keywords):
    """Format the keywords dataframe with proper column names"""
    df_formatted = df_keywords.copy()
    
    # Remove 'Sacha - 2020.1' column if it exists
    if 'Sacha - 2020.1' in df_formatted.columns:
        df_formatted = df_formatted.drop(columns=['Sacha - 2020.1'])
    
    # Rename columns
    column_mapping = {}
    for i, col in enumerate(df_formatted.columns):
        if col == 'Unnamed: 3':
            column_mapping[col] = 'Graph'
        elif hasattr(col, 'strftime'):  # Check if it's a datetime object
            try:
                column_mapping[col] = col.strftime('%m/%d')
            except:
                column_mapping[col] = str(col)
        elif isinstance(col, str) and col.startswith('2025-'):
            try:
                date_obj = pd.to_datetime(col)
                column_mapping[col] = date_obj.strftime('%m/%d')
            except:
                column_mapping[col] = col
    
    # Apply column renaming
    df_formatted = df_formatted.rename(columns=column_mapping)
    return df_formatted


def format_pages_dataframe(df_pages):
    """Format the pages dataframe with proper column names"""
    df_formatted = df_pages.copy()
    
    # Rename columns with readable dates
    column_mapping = {}
    for i, col in enumerate(df_formatted.columns):
        if hasattr(col, 'strftime'):
            try:
                column_mapping[col] = col.strftime('%b. %d, %Y')
            except:
                column_mapping[col] = str(col)
        elif isinstance(col, str) and col.startswith('2025-'):
            try:
                date_obj = pd.to_datetime(col)
                column_mapping[col] = date_obj.strftime('%b. %d, %Y')
            except:
                column_mapping[col] = col
    
    df_formatted = df_formatted.rename(columns=column_mapping)
    return df_formatted


def main():
    # Header
    st.title("📊 Katya SEO Tracking")
    
    # Load data
    df_keywords, df_pages = load_seo_data()
    
    if df_keywords is None or df_pages is None:
        st.error("Failed to load SEO data. Please check the file path.")
        return
    
    # Format data
    df_keywords_formatted = format_keywords_dataframe(df_keywords)
    df_pages_formatted = format_pages_dataframe(df_pages)
    
    # Tabs
    tab1, tab2 = st.tabs(["By Keyword", "Numbers of Kwds By Page"])
    
    with tab1:
        # Detect date columns (for ranking trend)
        date_columns = [
            c for c in df_keywords_formatted.columns
            if isinstance(c, str) and "/" in c and all(ch.isdigit() or ch == "/" for ch in c)
        ]

        df_keywords_display = df_keywords_formatted.copy()

        # Create sparkline data list for each row
        if date_columns:
            df_keywords_display["Graph"] = df_keywords_display[date_columns].values.tolist()

        # Display dataframe with sparklines
        st.dataframe(
            df_keywords_display,
            use_container_width=True,
            column_config={
                "URL": st.column_config.LinkColumn(
                    "URL",
                    help="Click to open URL",
                    display_text="Open Link"
                ),
                "Graph": st.column_config.LineChartColumn(
                    "Ranking Trend",
                    help=f"Ranking trend from {date_columns[0]} to {date_columns[-1]}" if date_columns else "Ranking trend",
                    y_min=0,
                    y_max=df_keywords_display[date_columns].max().max() if date_columns else None,
                ),
            }
        )
    
    with tab2:
        # Pages tab showing actual URL text
        st.dataframe(
            df_pages_formatted, 
            use_container_width=True,
            column_config={
                "Page": st.column_config.LinkColumn(
                    "Page",
                    help="Click to open page",
                    display_text=None
                )
            }
        )


if __name__ == "__main__":
    main()
