import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import sys
import os

# Add parent directory to path to import database_utils
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from database_utils import get_discovery_call_dates, check_database_exists

# Page configuration
st.set_page_config(
    page_title="Burki Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for strong colors
st.markdown("""
<style>
    .main {
        padding: 1rem;
    }
    .stMetric {
        background-color: #f8f9fa;
        padding: 0.5rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .metric-container {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
    }
    .embed-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
        text-align: center;
    }
    /* Hide QuickBooks and SignNow pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"] {
        display: none !important;
    }
    @media (max-width: 768px) {
        .embed-header {
            font-size: 1.2rem;
        }
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main application function"""
    # Header
    st.title("📊 Burki Dashboard")
    
    # Check if database exists and has data
    db_exists, db_message = check_database_exists()
    
    if not db_exists:
        st.error(f"❌ Database not ready: {db_message}")
        st.info("💡 Please go to the 'Database Refresh' page to initialize the database with Monday.com data.")
        return
    
    # Load discovery call data from database
    with st.spinner("Loading discovery call data from database..."):
        all_discovery_dates = get_discovery_call_dates()
    
    if not all_discovery_dates:
        st.warning("No discovery call dates found for 2025. Please check your data and refresh the database.")
        return
    
    # Convert to DataFrame
    df_2025 = pd.DataFrame(all_discovery_dates)
    
    # Add month information
    df_2025['month'] = df_2025['date'].dt.month
    df_2025['month_name'] = df_2025['date'].dt.strftime('%B')
    
    # Count calls by month
    monthly_counts = df_2025.groupby(['month', 'month_name']).size().reset_index(name='count')
    monthly_counts = monthly_counts.sort_values('month')
    
    # Create the bar graph
    
    # Create bar chart with strong colors
    fig = px.bar(
        monthly_counts, 
        x='month_name', 
        y='count',
        title="Number of Discovery Calls Qualified by Month",
        labels={'count': 'Number of Calls', 'month_name': 'Month'}
    )
    
    # Customize the chart with strong colors
    fig.update_layout(
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(size=12),
        title_font_size=16,
        xaxis_title_font_size=14,
        yaxis_title_font_size=14,
        height=500
    )
    
    # Update bar colors to be more distinct
    fig.update_traces(
        marker_color='#1f77b4',
        marker_line_color='#0d47a1',
        marker_line_width=1,
        text=monthly_counts['count'],  # Show numbers above bars
        textposition='outside'  # Position text above the bars
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Show data summary
    st.info(f"📊 Showing {len(df_2025)} discovery calls from 2025 across all boards")

if __name__ == "__main__":
    main()
