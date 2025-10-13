import streamlit as st
import requests
import pandas as pd
from datetime import datetime, date
import plotly.express as px

# Page configuration
st.set_page_config(
    page_title="New Leads Check",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to hide QuickBooks and SignNow pages from sidebar
st.markdown("""
<style>
    /* Hide QuickBooks and SignNow pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"] {
        display: none !important;
    }
    .embed-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
        text-align: center;
    }
    .stDataFrame {
        font-size: 12px;
    }
    .stDataFrame > div {
        max-height: 600px;
        overflow-y: auto;
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
            'new_leads_board_id', 'discovery_call_board_id', 
            'design_review_board_id', 'sales_board_id'
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

# Board configurations - will fetch all groups except those starting with No, Not, Spam
BOARDS = {
    'New Leads v2': credentials['new_leads_board_id'],
    'Discovery Call v2': credentials['discovery_call_board_id'],
    'Design Review v2': credentials['design_review_board_id'],
    'Sales v2': credentials['sales_board_id']
}

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_board_groups(board_id, board_name):
    """Get all groups from a Monday.com board and filter out unwanted ones"""
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json",
    }
    
    # GraphQL query to get all groups from the board
    query = f"""
    query {{
        boards(ids: [{board_id}]) {{
            groups {{
                id
                title
            }}
        }}
    }}
    """
    
    try:
        response = requests.post(url, json={"query": query}, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()
        
        # Check for API errors
        if "errors" in data and data["errors"]:
            return []
        
        if ("data" in data and "boards" in data["data"] and 
            data["data"]["boards"] and data["data"]["boards"][0].get("groups")):
            
            all_groups = data["data"]["boards"][0]["groups"]
            
            # Filter out groups that start with "No", "Not", or "Spam" (case insensitive)
            filtered_groups = []
            excluded_groups = []
            
            for group in all_groups:
                group_title = group.get("title", "").strip()  # Remove leading/trailing whitespace
                group_id = group.get("id", "")
                if group_title and group_id:
                    # Check if group starts with excluded prefixes (case insensitive)
                    if (not group_title.lower().startswith(('no', 'not', 'spam'))):
                        filtered_groups.append({"id": group_id, "title": group_title})
                    else:
                        excluded_groups.append(group_title)
            
            return filtered_groups
        else:
            return []
            
    except requests.exceptions.Timeout:
        return []
    except requests.exceptions.RequestException as e:
        return []
    except Exception as e:
        return []

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_board_data(board_id, board_name, groups):
    """Get items from specific groups in a Monday.com board with caching"""
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json",
    }
    
    all_items = []
    
    for group in groups:
        group_id = group["id"]
        group_title = group["title"]
        
        # GraphQL query to fetch data from specific board and group
        query = f"""
        query {{
            boards(ids: [{board_id}]) {{
                groups(ids: ["{group_id}"]) {{
                    id
                    title
                    items_page(limit: 500) {{
                        items {{
                            id
                            name
                            created_at
                            column_values {{
                                id
                                text
                                value
                                type
                            }}
                        }}
                    }}
                }}
            }}
        }}
        """
        
        try:
            response = requests.post(url, json={"query": query}, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            
            # Check for API errors
            if "errors" in data and data["errors"]:
                continue
            
            if ("data" in data and "boards" in data["data"] and 
                data["data"]["boards"] and data["data"]["boards"][0].get("groups")):
                
                board_info = data["data"]["boards"][0]
                groups_data = board_info.get("groups", [])
                
                for group_data in groups_data:
                    items_page = group_data.get("items_page", {})
                    items = items_page.get("items", [])
                    
                    # Add board name and group to each item
                    for item in items:
                        item['board_name'] = board_name
                        item['group_name'] = group_title
                    
                    all_items.extend(items)
                
        except requests.exceptions.Timeout:
            continue
        except requests.exceptions.RequestException as e:
            continue
        except Exception as e:
            continue
    
    return all_items

def get_all_leads_data():
    """Get data from all four boards"""
    all_leads = []
    
    for board_name, board_id in BOARDS.items():
        with st.spinner(f"Getting groups from {board_name}..."):
            # First, get all groups and filter out unwanted ones
            valid_groups = get_board_groups(board_id, board_name)
            
            if valid_groups:
                board_data = get_board_data(board_id, board_name, valid_groups)
                all_leads.extend(board_data)
    
    return all_leads

def format_leads_data(leads_data):
    """Convert Monday.com leads data to pandas DataFrame"""
    if not leads_data:
        return pd.DataFrame()

    records = []
    debug_columns = []  # Store column info for debugging

    for item in leads_data:
        date_created = None

        # Try to get the custom 'Date Created' field from column_values
        for col in item.get("column_values") or []:
            col_id = (col.get("id") or "").lower()
            text = (col.get("text") or "").strip()
            col_type = col.get("type") or ""

            # Debug: collect column information
            if col_id and col_type == "date":
                debug_columns.append({
                    "col_id": col_id,
                    "text": text,
                    "col_type": col_type
                })

            # Try multiple patterns to match "Date Created" column
            if (col_type == "date" and text and 
                ("created" in col_id or "date" in col_id) and
                "new lead form fill date" not in col_id.lower()):
                date_created = text  # This is typically in "YYYY-MM-DD" format
                break

        record = {
            "Item Name": item.get("name", ""),
            "Current Board": item.get("board_name", ""),
            "Group": item.get("group_name", ""),
            "Created At": item.get("created_at", ""),  # Fallback
            "Date Created (Custom)": date_created,
            "Raw Created At": item.get("created_at", "")  # Keep raw for debugging
        }

        records.append(record)

    df = pd.DataFrame(records)
    
    # Store debug info for later use
    if debug_columns:
        df.attrs['debug_columns'] = debug_columns

    if not df.empty:
        # Try to parse the custom 'Date Created' first
        df['Effective Date'] = pd.to_datetime(df['Date Created (Custom)'], errors='coerce')

        # Fallback to 'Created At' if 'Date Created' is missing
        missing_dates = df['Effective Date'].isna()
        df.loc[missing_dates, 'Effective Date'] = pd.to_datetime(df.loc[missing_dates, 'Created At'], errors='coerce', utc=True).dt.tz_convert('America/New_York').dt.tz_localize(None)

        # Add column just for comparing date part
        df['Effective Date Date'] = df['Effective Date'].dt.date

        # For display - show which date source was used
        df['Date Created Display'] = df['Effective Date'].dt.strftime('%Y-%m-%d')
        
        # Add indicator of which date source was used
        df['Date Source'] = df.apply(lambda row: 
            'Custom Date Created' if pd.notna(row['Date Created (Custom)']) and pd.notna(row['Effective Date'])
            else 'Fallback Created At' if pd.notna(row['Raw Created At']) and pd.notna(row['Effective Date'])
            else 'No Date', axis=1)

    return df

def filter_leads_by_date(df, selected_date):
    """Filter leads by the selected date"""
    if df.empty:
        return df
    
    # Convert selected_date to datetime for comparison
    if isinstance(selected_date, str):
        selected_date = pd.to_datetime(selected_date).date()
    elif isinstance(selected_date, date):
        selected_date = selected_date
    
    # Use the already calculated 'Effective Date Date' column if it exists
    if 'Effective Date Date' in df.columns:
        # Filter using the existing date column
        filtered_df = df[df['Effective Date Date'] == selected_date].copy()
    else:
        # Fallback: recalculate if needed
        df_copy = df.copy()
        df_copy['Effective Date Date'] = df_copy['Effective Date'].apply(
            lambda x: x.date() if pd.notna(x) else None
        )
        filtered_df = df_copy[df_copy['Effective Date Date'] == selected_date].copy()
    
    return filtered_df

def main():
    """Main application function"""
    # Header
    st.markdown('<div class="embed-header">🔍 NEW LEADS CHECK</div>', unsafe_allow_html=True)
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Settings")
        st.info(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Refresh button
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
        
        # Display board information
        st.subheader("📋 Boards")
        for board_name, board_id in BOARDS.items():
            st.text(f"{board_name}: {board_id}")
        
        st.info("📌 Fetching ALL groups from each board, excluding groups that start with 'No', 'Not', or 'Spam'")
    
    # Date selector
    st.subheader("📅 Select Date")
    selected_date = st.date_input(
        "Choose a date to view leads created on that day:",
        value=date.today(),
        help="Select the date to filter leads by their creation date"
    )
    
    # Load data
    with st.spinner("Loading leads data from all Monday.com boards..."):
        leads_data = get_all_leads_data()
        df = format_leads_data(leads_data)
    
    if df.empty:
        st.warning("No leads data found. Please check your board configurations and API permissions.")
        return
    
    # Filter data by selected date
    filtered_df = filter_leads_by_date(df, selected_date)
    
    if filtered_df.empty:
        st.info(f"No leads were created on {selected_date.strftime('%B %d, %Y')}.")
        
        # Debug information to help troubleshoot
        if not df.empty:
            st.subheader("🔍 Debug Information")
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Date Range in Data:**")
                if 'Effective Date Date' in df.columns:
                    date_counts = df['Effective Date Date'].value_counts().head(10)
                    if not date_counts.empty:
                        st.write(date_counts)
                    else:
                        st.write("No valid dates found")
                else:
                    st.write("No date column found")
            
            with col2:
                st.write(f"**Selected Date:** {selected_date}")
                st.write(f"**Selected Date Type:** {type(selected_date)}")
                if not df.empty and 'Effective Date Date' in df.columns:
                    sample_dates = df['Effective Date Date'].dropna().head(5).tolist()
                    st.write(f"**Sample dates in data:** {sample_dates}")
            
            # Show column debugging info
            if hasattr(df, 'attrs') and 'debug_columns' in df.attrs:
                st.subheader("📋 Column Debug Information")
                debug_df = pd.DataFrame(df.attrs['debug_columns'])
                if not debug_df.empty:
                    # Show unique column IDs and their sample values
                    unique_cols = debug_df.groupby('col_id').agg({
                        'text': lambda x: list(x.dropna().unique())[:3],  # First 3 unique values
                        'col_type': 'first'
                    }).reset_index()
                    st.write("**Date columns found in Monday.com data:**")
                    st.dataframe(unique_cols, hide_index=True)
                else:
                    st.write("No date columns found in the data")
        
    else:
        # Show filtered results
        total_leads = len(filtered_df)
        st.metric("Leads Found", total_leads)
        
        # Display the filtered data with date information
        display_columns = ['Item Name', 'Current Board', 'Group']     
        
        st.dataframe(
            filtered_df[display_columns],
            width='stretch',
            hide_index=True,
            column_config={
                "Item Name": "Item Name",
                "Current Board": "Current Board", 
                "Group": "Group"
            }
        )

if __name__ == "__main__":
    main()
