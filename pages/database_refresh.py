import streamlit as st
import requests
import pandas as pd
import sqlite3
import os
import sys
import subprocess
import time
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
    /* Hide tool pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="tools"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"],
    [data-testid="stSidebarNav"] a[href*="workbook_creator"],
    [data-testid="stSidebarNav"] a[href*="deck_creator"],
    [data-testid="stSidebarNav"] a[href*="a_la_carte"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

# Database setup
MONDAY_DB_PATH = "monday_data.db"
CALENDLY_DB_PATH = "calendly_data.db"

def init_monday_database():
    """Initialize SQLite database with tables for all Monday.com boards"""
    conn = sqlite3.connect(MONDAY_DB_PATH)
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

def init_calendly_database():
    """Initialize SQLite database for Calendly data"""
    conn = sqlite3.connect(CALENDLY_DB_PATH)
    cursor = conn.cursor()
    
    # Create table for Calendly events
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS calendly_events (
            uri TEXT PRIMARY KEY,
            name TEXT,
            start_time TEXT,
            end_time TEXT,
            status TEXT,
            event_type TEXT,
            invitee_name TEXT,
            invitee_email TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    conn.commit()
    conn.close()

def load_monday_credentials():
    """Load Monday.com credentials from Streamlit secrets"""
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
        st.error(f"Error reading Monday secrets: {str(e)}")
        return None

def load_calendly_credentials():
    """Load Calendly credentials from Streamlit secrets"""
    try:
        if 'calendly' not in st.secrets:
            st.error("Calendly configuration not found in secrets.toml.")
            return None
        
        calendly_config = st.secrets['calendly']
        
        if 'calendly_api_key' not in calendly_config:
            st.error("Calendly API key not found in secrets.toml.")
            return None
            
        return {
            'api_key': calendly_config['calendly_api_key']
        }
    except Exception as e:
        st.error(f"Error reading Calendly secrets: {str(e)}")
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
                                ... on BoardRelationValue {{
                                    linked_item_ids
                                    display_value
                                    linked_items {{
                                        id
                                        name
                                    }}
                                }}
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
                                ... on BoardRelationValue {{
                                    linked_item_ids
                                    display_value
                                    linked_items {{
                                        id
                                        name
                                    }}
                                }}
                            }}
                        }}
                    }}
                }}
            }}
            """
        
        # Retry logic for rate limit and concurrency errors
        # Use same max_retries as refresh_database.py which is working
        max_retries = 5
        retry_count = 0
        data = None
        
        while retry_count < max_retries:
            try:
                response = requests.post(url, json={"query": query}, headers=headers, timeout=timeout)
                
                if response.status_code == 401:
                    return None, f"401 Unauthorized: Check API token for {board_name}"
                
                # Parse JSON response (same as refresh_database.py)
                data = response.json()
                
                if "errors" in data:
                    errors = data['errors']
                    # Check if it's a rate limit, concurrency limit, or server error
                    # Use same logic as refresh_database.py which is working
                    is_retryable_error = False
                    retry_seconds = 0
                    
                    for error in errors:
                        extensions = error.get('extensions', {})
                        status_code = extensions.get('status_code')
                        code = extensions.get('code', '')
                        error_code = extensions.get('error_code', '')
                        
                        # Check for retryable errors: 429 (rate limit), LIMIT_EXCEEDED, or 500 (server error)
                        # Option 4: Also retry on 500 Internal Server Errors
                        if (status_code == 429 or 
                            status_code == 500 or
                            'RATE_LIMIT' in str(code) or 
                            'LIMIT_EXCEEDED' in str(code) or
                            'INTERNAL_SERVER_ERROR' in str(error_code)):
                            is_retryable_error = True
                            # Get retry_in_seconds from error, default based on error type
                            if status_code == 500:
                                retry_seconds = max(retry_seconds, extensions.get('retry_in_seconds', 10))
                            else:
                                retry_seconds = max(retry_seconds, extensions.get('retry_in_seconds', 15))
                    
                    if is_retryable_error and retry_count < max_retries - 1:
                        retry_count += 1
                        # Option 2: Increased exponential backoff multiplier from 2 to 5
                        # This gives more backpressure: if Monday says "retry in 15s", we wait 15 + (retry_count * 5)
                        # Example: retry_in_seconds=15, retry_count=1 → wait 20s, retry_count=2 → wait 25s, etc.
                        wait_time = retry_seconds + (retry_count * 5)
                        time.sleep(wait_time)
                        continue
                    else:
                        # Not a retryable error, or max retries reached
                        return None, f"GraphQL errors for {board_name}: {errors}"
                else:
                    # Success - break out of retry loop
                    break
                    
            except Exception as e:
                if retry_count < max_retries - 1:
                    retry_count += 1
                    # Use same exponential backoff as refresh_database.py
                    wait_time = (retry_count * 2) + 5
                    time.sleep(wait_time)
                    continue
                else:
                    return None, f"Error fetching data from {board_name}: {str(e)}"
        
        if data is None or "errors" in data:
            return None, f"Failed to fetch data for {board_name} after {max_retries} retries"
        
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
        
        # Small delay between pages to avoid rate limits
        # Longer delay for Sales board due to its size
        delay = 1.0 if board_name.lower() == "sales" else 0.5
        time.sleep(delay)
    
    return all_items, None

def save_calendly_data_to_db(events_data):
    """Save Calendly events data to SQLite database"""
    conn = sqlite3.connect(CALENDLY_DB_PATH)
    cursor = conn.cursor()
    
    # Clear existing data
    cursor.execute("DELETE FROM calendly_events")
    
    # Insert new data
    for event in events_data:
        uri = event.get('uri', '')
        name = event.get('name', '')
        start_time = event.get('start_time', '')
        end_time = event.get('end_time', '')
        status = event.get('status', '')
        event_type = event.get('event_type', '')
        
        # Get invitee info
        invitees = event.get('invitees', [])
        invitee_name = ""
        invitee_email = ""
        if invitees:
            invitee = invitees[0]
            invitee_name = invitee.get('name', '')
            invitee_email = invitee.get('email', '')
        
        cursor.execute('''
            INSERT INTO calendly_events 
            (uri, name, start_time, end_time, status, event_type, invitee_name, invitee_email, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (uri, name, start_time, end_time, status, event_type, invitee_name, invitee_email, datetime.now()))
    
    conn.commit()
    conn.close()

def save_board_data_to_db(board_data, table_name, board_type):
    """Save board data to SQLite database"""
    conn = sqlite3.connect(MONDAY_DB_PATH)
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

def refresh_monday_database():
    """Refresh all board data from Monday.com"""
    credentials = load_monday_credentials()
    if not credentials:
        return 0, ["Failed to load credentials"], ["ERROR - Failed to load credentials"]
    
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
    
    # Add initial delay before starting to avoid immediate concurrency limits
    time.sleep(2)
    
    for idx, (board_id, table_name, board_name) in enumerate(boards_config):
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
                # Add delay after errors to avoid compounding issues
                if idx < len(boards_config) - 1:
                    time.sleep(10)
                continue
            
            # Check if we got data
            if not data:
                errors.append(f"{board_name}: No data returned from API")
                detailed_results.append(f"{board_name}: WARNING - No data returned from API")
                if idx < len(boards_config) - 1:
                    time.sleep(10)
                continue
            
            # Save to database
            save_board_data_to_db(data, table_name, board_name)
            success_count += 1
            detailed_results.append(f"{board_name}: SUCCESS - {len(data)} items saved")
            
            # Option 1: Stagger board fetches - wait 10 seconds between boards
            # This alone often eliminates FIELD_LIMIT_EXCEEDED errors
            if idx < len(boards_config) - 1:
                time.sleep(10)
            
        except Exception as e:
            errors.append(f"{board_name}: {str(e)}")
            detailed_results.append(f"{board_name}: EXCEPTION - {str(e)}")
    
    return success_count, errors, detailed_results

def generate_new_leads_cache():
    """Generate the new leads month cache by running the cache generation script."""
    try:
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scripts", "generate_new_leads_month_cache.py")
        if os.path.exists(script_path):
            result = subprocess.run(
                [sys.executable, script_path],
                capture_output=True,
                text=True,
                cwd=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            )
            if result.returncode == 0:
                return True, result.stdout
            else:
                error_msg = result.stderr if result.stderr else "Unknown error"
                return False, error_msg
        else:
            return False, f"Cache script not found at {script_path}"
    except Exception as e:
        return False, f"Error generating cache: {str(e)}"

def refresh_calendly_database():
    """Refresh Calendly data from API"""
    credentials = load_calendly_credentials()
    if not credentials:
        return False, "Failed to load Calendly credentials"
    
    api_key = credentials['api_key']
    
    try:
        # Get user info
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        
        # Get user info
        user_response = requests.get('https://api.calendly.com/users/me', headers=headers)
        if user_response.status_code != 200:
            return False, f"Failed to get user info: {user_response.status_code}"
        
        user_data = user_response.json()
        user_uri = user_data['resource']['uri']
        
        # Get event types
        event_types_response = requests.get(f'https://api.calendly.com/event_types?user={user_uri}', headers=headers)
        if event_types_response.status_code != 200:
            return False, f"Failed to get event types: {event_types_response.status_code}"
        
        event_types_data = event_types_response.json()
        event_types = event_types_data.get('collection', [])
        
        # Find ONLY "TEG - Let's Chat" event type (filter out "Intro call with Teg")
        teg_event_types = []
        for event_type in event_types:
            event_name = event_type.get('name', '')
            # Only include "TEG - Let's Chat", explicitly exclude "Intro call with Teg"
            if event_name == "TEG - Let's Chat":
                teg_event_types.append(event_type)
        
        if not teg_event_types:
            return False, "No 'TEG - Let's Chat' event type found"
        
        # Fetch events for "TEG - Let's Chat" only
        all_events = []
        event_names = []
        
        # Add progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_event_types = len(teg_event_types)
        
        for i, event_type in enumerate(teg_event_types):
            event_type_uri = event_type['uri']
            event_type_uuid = event_type_uri.split('/')[-1]
            event_name = event_type['name']
            event_names.append(event_name)
            
            # Update progress
            progress = i / total_event_types
            progress_bar.progress(progress)
            status_text.text(f"Fetching events for: {event_name}")
            
            # Fetch scheduled events with pagination for this event type
            page_count = 0
            next_page_token = None
            
            # Date range for 2025
            min_start_time = "2025-01-01T00:00:00.000000Z"
            # Use current date plus 30 days to include future events
            from datetime import datetime, timedelta
            max_date = (datetime.now() + timedelta(days=30)).strftime('%Y-%m-%dT23:59:59.999999Z')
            max_start_time = max_date
            
            while page_count < 100:  # Increased to fetch more events (up to 10,000 events)
                page_count += 1
                
                params = {
                    'user': user_uri,
                    'event_type': event_type_uuid,
                    'min_start_time': min_start_time,
                    'max_start_time': max_start_time,
                    'count': 100
                }
                
                if next_page_token:
                    params['page_token'] = next_page_token
                
                events_response = requests.get('https://api.calendly.com/scheduled_events', 
                                            headers=headers, params=params)
                
                if events_response.status_code != 200:
                    error_msg = f"Failed to get events for {event_name}: {events_response.status_code}"
                    if events_response.status_code == 404:
                        error_msg += f" - Response: {events_response.text}"
                    return False, error_msg
                
                events_data = events_response.json()
                events = events_data.get('collection', [])
                
                if not events:
                    break
                
                all_events.extend(events)
                
                pagination = events_data.get('pagination', {})
                next_page_token = pagination.get('next_page_token')
                
                if not next_page_token:
                    break
        
        # Clear progress indicators
        progress_bar.empty()
        status_text.empty()
        
        # Save to database
        save_calendly_data_to_db(all_events)
        
        return True, f"Successfully saved {len(all_events)} Calendly events for: {', '.join(event_names)}"
        
    except Exception as e:
        return False, f"Error refreshing Calendly data: {str(e)}"

def get_database_status():
    """Get status of both Monday and Calendly database tables"""
    status = {}
    
    # Monday database status
    try:
        conn = sqlite3.connect(MONDAY_DB_PATH)
        cursor = conn.cursor()
        
        tables = ['sales_board', 'new_leads_board', 'discovery_call_board', 'design_review_board', 'ads_board']
        
        for table in tables:
            try:
                cursor.execute(f"SELECT COUNT(*) FROM {table}")
                count = cursor.fetchone()[0]
                
                cursor.execute(f"SELECT MAX(updated_at) FROM {table}")
                last_updated = cursor.fetchone()[0]
                
                status[f"monday_{table}"] = {
                    'count': count,
                    'last_updated': last_updated
                }
            except:
                status[f"monday_{table}"] = {
                    'count': 0,
                    'last_updated': 'Never'
                }
        
        conn.close()
    except:
        pass
    
    # Calendly database status
    try:
        conn = sqlite3.connect(CALENDLY_DB_PATH)
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT COUNT(*) FROM calendly_events")
            count = cursor.fetchone()[0]
            
            cursor.execute("SELECT MAX(updated_at) FROM calendly_events")
            last_updated = cursor.fetchone()[0]
            
            status['calendly_events'] = {
                'count': count,
                'last_updated': last_updated
            }
        except:
            status['calendly_events'] = {
                'count': 0,
                'last_updated': 'Never'
            }
        
        conn.close()
    except:
        pass
    
    return status

def main():
    """Main application function"""
    st.title("🔄 Database Refresh")
    
    # Initialize both databases
    init_monday_database()
    init_calendly_database()
    
    # Two separate refresh buttons
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📊 Monday.com Data")
        st.markdown("Refresh all Monday.com board data (Sales, Leads, Discovery Calls, etc.)")
        
        # Initialize session state for results if not exists
        if 'monday_refresh_results' not in st.session_state:
            st.session_state.monday_refresh_results = None
        if 'monday_refresh_errors' not in st.session_state:
            st.session_state.monday_refresh_errors = None
        if 'monday_refresh_detailed_results' not in st.session_state:
            st.session_state.monday_refresh_detailed_results = None
        if 'monday_refresh_success_count' not in st.session_state:
            st.session_state.monday_refresh_success_count = None
        if 'monday_cache_success' not in st.session_state:
            st.session_state.monday_cache_success = None
        if 'monday_cache_message' not in st.session_state:
            st.session_state.monday_cache_message = None
        
        if st.button("🔄 Refresh All Monday Data", type="primary", use_container_width=True):
            try:
                with st.spinner("Refreshing Monday.com database..."):
                    success_count, errors, detailed_results = refresh_monday_database()
                
                # Store results in session state
                st.session_state.monday_refresh_success_count = success_count
                st.session_state.monday_refresh_errors = errors
                st.session_state.monday_refresh_detailed_results = detailed_results
                
                # Generate new leads cache after successful Monday refresh
                cache_success = False
                cache_message = ""
                if success_count > 0:
                    with st.spinner("Generating New Leads cache..."):
                        cache_success, cache_message = generate_new_leads_cache()
                
                # Store cache results in session state
                st.session_state.monday_cache_success = cache_success
                st.session_state.monday_cache_message = cache_message
                
                st.rerun()
            except Exception as e:
                # Store error in session state
                st.session_state.monday_refresh_errors = [f"Unexpected error: {str(e)}"]
                st.session_state.monday_refresh_detailed_results = [f"EXCEPTION - {str(e)}"]
                st.session_state.monday_refresh_success_count = 0
                import traceback
                st.session_state.monday_cache_message = traceback.format_exc()
                st.rerun()
        
        # Display results from session state (persists after rerun) - only show errors
        if st.session_state.monday_refresh_success_count is not None:
            errors = st.session_state.monday_refresh_errors or []
            detailed_results = st.session_state.monday_refresh_detailed_results or []
            
            # Only show cache warning if cache generation failed
            if st.session_state.monday_cache_success is not None and not st.session_state.monday_cache_success:
                st.warning(f"⚠️ Cache generation: {st.session_state.monday_cache_message}")
            
            # Show errors only
            if errors:
                st.markdown(f"""
                <div class="status-error">
                    <h4>⚠️ Some Errors Occurred:</h4>
                    <ul>
                        {''.join([f'<li>{error}</li>' for error in errors])}
                    </ul>
                </div>
                """, unsafe_allow_html=True)
            
            # Show error results from detailed_results
            if detailed_results:
                error_results = [r for r in detailed_results if "ERROR" in r or "EXCEPTION" in r or "WARNING" in r]
                if error_results:
                    for result in error_results:
                        if "ERROR" in result or "EXCEPTION" in result:
                            st.error(f"❌ {result}")
                        elif "WARNING" in result:
                            st.warning(f"⚠️ {result}")
            
            # Show exception traceback if available
            if st.session_state.monday_cache_message and "Traceback" in st.session_state.monday_cache_message:
                st.code(st.session_state.monday_cache_message)
    
    with col2:
        st.markdown("### 📅 Calendly Data")
        st.markdown("Refresh Calendly events data for TEG calls")
        
        if st.button("🔄 Refresh All Calendly Data", type="primary", use_container_width=True):
            with st.spinner("Refreshing Calendly database..."):
                success, message = refresh_calendly_database()
            
            if success:
                st.markdown(f"""
                <div class="status-success">
                    <h4>✅ Calendly Refresh Complete!</h4>
                    <p>{message}</p>
                    <p>Calendly database has been refreshed with latest data.</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class="status-error">
                    <h4>❌ Calendly Refresh Failed:</h4>
                    <p>{message}</p>
                </div>
                """, unsafe_allow_html=True)
            
            st.rerun()
    
    # Database file info
    st.markdown("---")
    st.subheader("📁 Database Files")
    
    if os.path.exists(MONDAY_DB_PATH):
        file_size = os.path.getsize(MONDAY_DB_PATH)
        st.info(f"📊 Monday Database: `{MONDAY_DB_PATH}` ({file_size:,} bytes)")
    
    if os.path.exists(CALENDLY_DB_PATH):
        file_size = os.path.getsize(CALENDLY_DB_PATH)
        st.info(f"📅 Calendly Database: `{CALENDLY_DB_PATH}` ({file_size:,} bytes)")

if __name__ == "__main__":
    main()
