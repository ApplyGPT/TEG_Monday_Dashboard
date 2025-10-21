import streamlit as st
import requests
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

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

# Monday.com API settings from Streamlit secrets
def load_credentials():
    """Load credentials from Streamlit secrets"""
    try:
        # Access secrets from Streamlit
        if 'monday' not in st.secrets:
            st.error("Monday.com configuration not found in secrets.toml. Please check your configuration.")
            st.stop()
        
        monday_config = st.secrets['monday']
        
        if 'api_token' not in monday_config:
            st.error("API token not found in secrets.toml. Please add your Monday.com API token.")
            st.stop()
            
        required_board_ids = [
            'sales_board_id', 
            'new_leads_board_id', 
            'discovery_call_board_id', 
            'design_review_board_id'
        ]
        
        board_ids = {}
        for board_id_key in required_board_ids:
            if board_id_key not in monday_config:
                st.error(f"{board_id_key} not found in secrets.toml. Please add the board ID.")
                st.stop()
            board_ids[board_id_key] = int(monday_config[board_id_key])
        
        return {
            'api_token': monday_config['api_token'],
            **board_ids
        }
    except Exception as e:
        st.error(f"Error reading secrets: {str(e)}")
        st.stop()

credentials = load_credentials()
API_TOKEN = credentials['api_token']
SALES_BOARD_ID = credentials['sales_board_id']
NEW_LEADS_BOARD_ID = credentials['new_leads_board_id']
DISCOVERY_CALL_BOARD_ID = credentials['discovery_call_board_id']
DESIGN_REVIEW_BOARD_ID = credentials['design_review_board_id']

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_board_data(board_id, board_name):
    """Get all items from a specific Monday.com board with caching"""
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json",
    }
    
    all_items = []
    cursor = None
    limit = 500
    
    with st.spinner(f"🔄 Fetching data from {board_name} board..."):
        while True:
            if cursor:
                query = f"""
                query {{
                    boards(ids: [{board_id}]) {{
                        items_page(limit: {limit}, cursor: "{cursor}") {{
                            cursor
                            items {{
                                id
                                name
                                column_values {{
                                    id
                                    text
                                    type
                                }}
                            }}
                        }}
                    }}
                }}
                """
            else:
                query = f"""
                query {{
                    boards(ids: [{board_id}]) {{
                        items_page(limit: {limit}) {{
                            cursor
                            items {{
                                id
                                name
                                column_values {{
                                    id
                                    text
                                    type
                                }}
                            }}
                        }}
                    }}
                }}
                """
            
            try:
                response = requests.post(url, json={"query": query}, headers=headers, timeout=30)
                
                if response.status_code == 401:
                    st.error("401 Unauthorized: Check your API token and permissions.")
                    return None
                
                response.raise_for_status()
                data = response.json()
                
                if "errors" in data:
                    st.error(f"GraphQL errors: {data['errors']}")
                    return None
                
                boards = data.get("data", {}).get("boards", [])
                if not boards:
                    break
                
                items_page = boards[0].get("items_page", {})
                items = items_page.get("items", [])
                
                if not items:
                    break
                
                all_items.extend(items)
                cursor = items_page.get("cursor")
                
                if not cursor:
                    break
                    
            except requests.exceptions.Timeout:
                st.error("Request timed out. Please try again.")
                return None
            except requests.exceptions.RequestException as e:
                st.error(f"Error fetching data from {board_name}: {str(e)}")
                return None
            except Exception as e:
                st.error(f"Unexpected error: {str(e)}")
                return None
    
    return all_items

def extract_discovery_call_dates(items, board_name):
    """Extract discovery call dates from board items"""
    discovery_dates = []
    
    for item in items:
        item_name = item.get("name", "")
        
        # Filter out items that start with "No", "Not", or "Spam"
        if item_name.lower().startswith(('no ', 'not ', 'spam')):
            continue
        
        for col_val in item.get("column_values", []):
            col_id = col_val.get("id", "")
            text = col_val.get("text", "")
            col_type = col_val.get("type", "")
            
            # Look for date columns that might contain discovery call dates
            # Common patterns for discovery call date columns
            if col_type == "date" and text and text.strip():
                try:
                    # Try to parse the date
                    parsed_date = pd.to_datetime(text, errors='coerce')
                    if not pd.isna(parsed_date) and parsed_date.year == 2025:
                        discovery_dates.append({
                            'date': parsed_date,
                            'item_name': item_name,
                            'board': board_name,
                            'column_id': col_id
                        })
                except:
                    continue
    
    return discovery_dates

def main():
    """Main application function"""
    # Header
    st.title("📊 Burki Dashboard")
    
    # Load and process data from all four boards
    with st.spinner("Loading discovery call data from all boards..."):
        # Get data from all boards (with full pagination to get all items)
        sales_items = get_board_data(SALES_BOARD_ID, "Sales")
        new_leads_items = get_board_data(NEW_LEADS_BOARD_ID, "New Leads")
        discovery_call_items = get_board_data(DISCOVERY_CALL_BOARD_ID, "Discovery Call")
        design_review_items = get_board_data(DESIGN_REVIEW_BOARD_ID, "Design Review")
        
        if not all([sales_items, new_leads_items, discovery_call_items, design_review_items]):
            st.error("Failed to load data from one or more boards. Please check your configuration.")
            return
        
        # Extract discovery call dates from all boards
        all_discovery_dates = []
        
        all_discovery_dates.extend(extract_discovery_call_dates(sales_items, "Sales"))
        all_discovery_dates.extend(extract_discovery_call_dates(new_leads_items, "New Leads"))
        all_discovery_dates.extend(extract_discovery_call_dates(discovery_call_items, "Discovery Call"))
        all_discovery_dates.extend(extract_discovery_call_dates(design_review_items, "Design Review"))
    
    if not all_discovery_dates:
        st.warning("No discovery call dates found for 2025. Please check your data and column configurations.")
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
        marker_line_width=1
    )
    
    st.plotly_chart(fig, use_container_width=True)

if __name__ == "__main__":
    main()
