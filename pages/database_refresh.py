import streamlit as st
import requests
import pandas as pd
import sqlite3
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# Page configuration
st.set_page_config(
    page_title="Database Refresh",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
<style>
    .main {
        padding: 1rem;
    }
    .stButton > button {
        width: 100%;
        margin-top: 1rem;
        background-color: #1f77b4;
        color: white;
    }
    .status-success {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .status-error {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    /* Hide QuickBooks and SignNow pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

# Database setup
DB_PATH = "monday_data.db"

def init_database():
    """Initialize SQLite database with tables for all boards"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Create tables for each board
    boards = [
        ('sales_board', 'Sales Board'),
        ('new_leads_board', 'New Leads Board'),
        ('discovery_call_board', 'Discovery Call Board'),
        ('design_review_board', 'Design Review Board'),
        ('ads_board', 'Ads Board')
    ]
    
    for table_name, description in boards:
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {table_name} (
                id TEXT PRIMARY KEY,
                name TEXT,
                board_type TEXT,
                column_values TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
    
    conn.commit()
    conn.close()

def load_credentials():
    """Load credentials from Streamlit secrets"""
    try:
        if 'monday' not in st.secrets:
            st.error("Monday.com configuration not found in secrets.toml.")
            return None
        
        monday_config = st.secrets['monday']
        
        if 'api_token' not in monday_config:
            st.error("API token not found in secrets.toml.")
            return None
            
        required_board_ids = [
            'sales_board_id', 
            'new_leads_board_id', 
            'discovery_call_board_id', 
            'design_review_board_id',
            'ads_board_id'
        ]
        
        board_ids = {}
        for board_id_key in required_board_ids:
            if board_id_key not in monday_config:
                st.error(f"{board_id_key} not found in secrets.toml.")
                return None
            board_ids[board_id_key] = int(monday_config[board_id_key])
        
        return {
            'api_token': monday_config['api_token'],
            **board_ids
        }
    except Exception as e:
        st.error(f"Error reading secrets: {str(e)}")
        return None

def get_board_data_from_monday(board_id, board_name, api_token, timeout=60):
    """Fetch all data from a Monday.com board with configurable timeout"""
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": api_token,
        "Content-Type": "application/json",
    }
    
    all_items = []
    cursor = None
    limit = 500
    
    # Special handling for Sales board - use smaller batches and longer timeout
    if board_name.lower() == "sales":
        limit = 200  # Smaller batches for Sales board
        timeout = 120  # Longer timeout for Sales board
    
    page_count = 0
    max_pages = 50  # Safety limit to prevent infinite loops
    
    while page_count < max_pages:
        page_count += 1
        
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
                                value
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
                                value
                                type
                            }}
                        }}
                    }}
                }}
            }}
            """
        
        try:
            response = requests.post(url, json={"query": query}, headers=headers, timeout=timeout)
            
            if response.status_code == 401:
                return None, f"401 Unauthorized: Check API token for {board_name}"
            
            response.raise_for_status()
            data = response.json()
            
            if "errors" in data:
                return None, f"GraphQL errors for {board_name}: {data['errors']}"
            
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
            return None, f"Request timed out for {board_name} (page {page_count}, {len(all_items)} items so far)"
        except requests.exceptions.RequestException as e:
            return None, f"Error fetching data from {board_name}: {str(e)}"
        except Exception as e:
            return None, f"Unexpected error for {board_name}: {str(e)}"
    
    return all_items, None

def save_board_data_to_db(board_data, table_name, board_type):
    """Save board data to SQLite database"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Clear existing data for this board
    cursor.execute(f"DELETE FROM {table_name}")
    
    # Insert new data
    for item in board_data:
        item_id = item.get("id", "")
        name = item.get("name", "")
        column_values = str(item.get("column_values", []))
        
        cursor.execute(f'''
            INSERT INTO {table_name} (id, name, board_type, column_values, updated_at)
            VALUES (?, ?, ?, ?, ?)
        ''', (item_id, name, board_type, column_values, datetime.now()))
    
    conn.commit()
    conn.close()

def refresh_database():
    """Refresh all board data from Monday.com"""
    credentials = load_credentials()
    if not credentials:
        return False, "Failed to load credentials"
    
    api_token = credentials['api_token']
    
    # Board configurations
    boards_config = [
        (credentials['sales_board_id'], 'sales_board', 'Sales'),
        (credentials['new_leads_board_id'], 'new_leads_board', 'New Leads'),
        (credentials['discovery_call_board_id'], 'discovery_call_board', 'Discovery Call'),
        (credentials['design_review_board_id'], 'design_review_board', 'Design Review'),
        (credentials['ads_board_id'], 'ads_board', 'Ads')
    ]
    
    success_count = 0
    errors = []
    detailed_results = []
    
    for board_id, table_name, board_name in boards_config:
        try:
            # Special progress indicator for Sales board
            if board_name.lower() == "sales":
                progress_bar = st.progress(0)
                status_text = st.empty()
                status_text.text(f"🔄 Fetching {board_name} board (this may take longer due to large dataset)...")
            
            # Fetch data from Monday.com
            data, error = get_board_data_from_monday(board_id, board_name, api_token)
            
            # Clear progress indicators
            if board_name.lower() == "sales":
                progress_bar.empty()
                status_text.empty()
            
            if error:
                errors.append(f"{board_name}: {error}")
                detailed_results.append(f"{board_name}: ERROR - {error}")
                continue
            
            # Check if we got data
            if not data:
                errors.append(f"{board_name}: No data returned from API")
                detailed_results.append(f"{board_name}: WARNING - No data returned from API")
                continue
            
            # Save to database
            save_board_data_to_db(data, table_name, board_name)
            success_count += 1
            detailed_results.append(f"{board_name}: SUCCESS - {len(data)} items saved")
            
        except Exception as e:
            errors.append(f"{board_name}: {str(e)}")
            detailed_results.append(f"{board_name}: EXCEPTION - {str(e)}")
    
    return success_count, errors, detailed_results

def get_database_status():
    """Get status of database tables"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    tables = ['sales_board', 'new_leads_board', 'discovery_call_board', 'design_review_board', 'ads_board']
    status = {}
    
    for table in tables:
        try:
            cursor.execute(f"SELECT COUNT(*) FROM {table}")
            count = cursor.fetchone()[0]
            
            cursor.execute(f"SELECT MAX(updated_at) FROM {table}")
            last_updated = cursor.fetchone()[0]
            
            status[table] = {
                'count': count,
                'last_updated': last_updated
            }
        except:
            status[table] = {
                'count': 0,
                'last_updated': 'Never'
            }
    
    conn.close()
    return status

def main():
    """Main application function"""
    st.title("🔄 Database Refresh")
    st.markdown("### Monday.com Data Management")
    
    # Initialize database
    init_database()
    
    # Get current database status
    status = get_database_status()
    
    # Display current status
    st.subheader("📊 Current Database Status")
    
    for table, info in status.items():
        board_name = table.replace('_board', '').replace('_', ' ').title()
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(f"{board_name} Records", info['count'])
        
        with col2:
            st.metric("Last Updated", info['last_updated'])
        
        with col3:
            if info['count'] > 0:
                st.success("✅ Data Available")
            else:
                st.error("❌ No Data")
    
    st.divider()
    
    # Refresh button
    st.subheader("🔄 Refresh Database")
    st.markdown("Click the button below to fetch fresh data from Monday.com and update the local database.")
    
    if st.button("🔄 Refresh All Data", type="primary"):
        with st.spinner("Refreshing database from Monday.com..."):
            success_count, errors, detailed_results = refresh_database()
        
        if success_count > 0:
            st.markdown(f"""
            <div class="status-success">
                <h4>✅ Refresh Complete!</h4>
                <p>Successfully updated {success_count} out of 5 boards.</p>
                <p>Database has been refreshed with latest data from Monday.com.</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Show detailed results
        st.subheader("📋 Detailed Results")
        for result in detailed_results:
            if "SUCCESS" in result:
                st.success(f"✅ {result}")
            elif "ERROR" in result or "EXCEPTION" in result:
                st.error(f"❌ {result}")
            elif "WARNING" in result:
                st.warning(f"⚠️ {result}")
        
        if errors:
            st.markdown(f"""
            <div class="status-error">
                <h4>⚠️ Some Errors Occurred:</h4>
                <ul>
                    {''.join([f'<li>{error}</li>' for error in errors])}
                </ul>
            </div>
            """, unsafe_allow_html=True)
        
        # Refresh the page to show updated status
        st.rerun() 
    
    # Database file info
    if os.path.exists(DB_PATH):
        file_size = os.path.getsize(DB_PATH)
        st.info(f"📁 Database file: `{DB_PATH}` ({file_size:,} bytes)")

if __name__ == "__main__":
    main()
