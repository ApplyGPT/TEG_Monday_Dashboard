import streamlit as st
import requests
import pandas as pd
import os
from datetime import datetime
import plotly.express as px

# Page configuration
st.set_page_config(
    page_title="Monday.com Data Viewer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Monday.com API settings from Streamlit secrets
API_TOKEN = st.secrets["monday"]["api_token"]
BOARD_ID = st.secrets["monday"]["board_id"]

# Custom CSS for better embedding and responsive design
st.markdown("""
<style>
    .main {
        padding: 1rem;
    }
    .stDataFrame {
        font-size: 12px;
    }
    .stDataFrame > div {
        max-height: 600px;
        overflow-y: auto;
    }
    .stButton > button {
        width: 100%;
        margin-top: 1rem;
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
    .stMetric {
        background-color: #f8f9fa;
        padding: 0.5rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    @media (max-width: 768px) {
        .embed-header {
            font-size: 1.2rem;
        }
        .stDataFrame {
            font-size: 10px;
        }
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_monday_data():
    """Get all items from Monday.com board with caching"""
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json",
    }
    
    query = f"""
    query {{
        boards(ids: {BOARD_ID}) {{
            items_page(limit: 100) {{
                items {{
                    id
                    name
                    state
                    created_at
                    updated_at
                    column_values {{
                        id
                        text
                        value
                    }}
                }}
            }}
        }}
    }}
    """
    
    try:
        response = requests.post(url, json={"query": query}, headers=headers, timeout=30)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.Timeout:
        st.error("Request timed out. Please try again.")
        return None
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching data: {str(e)}")
        return None
    except Exception as e:
        st.error(f"Unexpected error: {str(e)}")
        return None

def format_data(data):
    """Convert Monday.com data to pandas DataFrame"""
    if not data or "data" not in data or "boards" not in data["data"]:
        return pd.DataFrame()
    
    boards = data["data"]["boards"]
    if not boards or not boards[0].get("items_page"):
        return pd.DataFrame()
    
    items_page = boards[0]["items_page"]
    if not items_page or not items_page.get("items"):
        return pd.DataFrame()
    
    items = items_page["items"]
    if not items:
        return pd.DataFrame()
    
    # Convert to DataFrame
    records = []
    for item in items:
        record = {
            "Item": item.get("name", ""),
            "Attribution Date": "",
            "Google Adspend": ""
        }
        
        # Add specific columns we want to display
        for col_val in item.get("column_values", []):
            col_id = col_val.get("id", "")
            text = col_val.get("text", "")
            value = col_val.get("value", "")
            
            # Map specific column IDs to our desired column names
            if col_id == "name":  # Item name
                record["Item"] = item.get("name", "")
            elif col_id == "date_mkv81p3z":  # Attribution Date
                record["Attribution Date"] = text if text else ""
            elif col_id == "numeric_mkv863mb":  # Google Adspend (the actual column with data)
                record["Google Adspend"] = text if text else ""
        
        records.append(record)
    
    df = pd.DataFrame(records)
    
    # Convert date columns and create month/year column
    df['Attribution Date'] = pd.to_datetime(df['Attribution Date'], errors='coerce')
    
    # Create Month/Year column for x-axis
    df['Month Year'] = df['Attribution Date'].dt.strftime('%B %Y')
    
    # Convert Google Adspend to numeric
    df['Google Adspend'] = pd.to_numeric(df['Google Adspend'], errors='coerce')
    
    # Sort by attribution date
    df = df.sort_values('Attribution Date')
    
    return df

def main():
    """Main application function"""
    # Header
    st.markdown('<div class="embed-header">📊 Google Ads Attribution Dashboard</div>', unsafe_allow_html=True)
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Settings")
        st.info(f"Board ID: {BOARD_ID}")
        st.info(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Refresh button
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
    
    # Load data
    with st.spinner("Loading data from Monday.com..."):
        data = get_monday_data()
        df = format_data(data)

    if df.empty:
        st.warning("No records found in the board. Add some items to Monday.com to see them here.")
        st.info("💡 **Tip**: Make sure your Monday.com board has items and your API token has the correct permissions.")
    else:
        # Year filter
        st.subheader("📅 Filter by Year")
        
        # Get unique years from the data
        df_with_dates = df.dropna(subset=['Attribution Date'])
        if not df_with_dates.empty:
            df_with_dates['Year'] = df_with_dates['Attribution Date'].dt.year
            available_years = sorted(df_with_dates['Year'].unique())
            
            # Add "All Years" option
            year_options = ["All Years"] + [str(year) for year in available_years]
            selected_year = st.selectbox("Select Year:", year_options)
            
            # Filter data based on selected year
            if selected_year == "All Years":
                df_filtered = df_with_dates
                year_label = "All Years"
            else:
                df_filtered = df_with_dates[df_with_dates['Year'] == int(selected_year)]
                year_label = selected_year
        else:
            df_filtered = df_with_dates
            year_label = "All Years"
            selected_year = "All Years"
        
        # Total Adspend metric
        st.subheader("💰 Total Ad Spend")
        total_adspend = df_filtered['Google Adspend'].sum()
        st.metric("Total Ad Spend", f"${total_adspend:,.2f}", delta=None)
        
        # Create the bar chart
        st.subheader(f"📊 Adspend by Month - {year_label}")
        
        # Filter out rows with missing data for charting
        df_chart = df_filtered.dropna(subset=['Attribution Date', 'Google Adspend'])
        
        if not df_chart.empty:
            # Create bar chart
            fig = px.bar(
                df_chart,
                x='Month Year',
                y='Google Adspend',
                title=f'Google Adspend by Month - {year_label}',
                labels={'Google Adspend': 'Adspend ($)', 'Month Year': 'Month'},
                color_discrete_sequence=['#1f77b4']  # Same blue color for all bars
            )
            
            # Update layout
            fig.update_layout(
                xaxis_title="Month",
                yaxis_title="Adspend ($)",
                height=500,
                showlegend=False
            )
            
            # Rotate x-axis labels and make them bigger
            fig.update_xaxes(
                tickangle=45,
                tickfont=dict(size=14),  # Bigger x-axis labels
                title_font=dict(size=16)  # Bigger x-axis title
            )
            
            # Make y-axis labels bigger too
            fig.update_yaxes(
                tickfont=dict(size=12),  # Bigger y-axis labels
                title_font=dict(size=16)  # Bigger y-axis title
            )
            
            # Display the chart
            st.plotly_chart(fig, use_container_width=True)
            
            # Show data table below chart
            st.subheader("📋 Data Table")
            display_columns = ['Item', 'Attribution Date', 'Google Adspend']
            if 'Year' in df_chart.columns:
                display_columns = ['Item', 'Attribution Date', 'Year', 'Google Adspend']
            
            st.dataframe(
                df_chart[display_columns],
                width='stretch',
                hide_index=True
            )
            
            # Download button
            csv = df_chart.to_csv(index=False)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            st.download_button(
                label="📥 Download CSV",
                data=csv,
                file_name=f"google_adspend_data_{timestamp}.csv",
                mime="text/csv"
            )
        else:
            st.warning("No data available for charting. Please check your attribution dates and adspend values.")

if __name__ == "__main__":
    main()
