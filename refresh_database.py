"""
Standalone Database Refresh Script
Can be run via cron job to refresh Monday.com and Calendly databases
Also handles QuickBooks refresh token updates
"""
import requests
import sqlite3
import os
import toml
import sys
from datetime import datetime

# Database paths
MONDAY_DB_PATH = "monday_data.db"
CALENDLY_DB_PATH = "calendly_data.db"

def init_databases():
    """Initialize databases if they don't exist"""
    # Initialize Monday database
    conn = sqlite3.connect(MONDAY_DB_PATH)
    cursor = conn.cursor()
    
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
    
    # Initialize Calendly database
    conn = sqlite3.connect(CALENDLY_DB_PATH)
    cursor = conn.cursor()
    
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

def load_config():
    """Load configuration from secrets.toml"""
    secrets_path = os.path.join('.streamlit', 'secrets.toml')
    
    if not os.path.exists(secrets_path):
        print(f"ERROR: secrets.toml not found at {secrets_path}")
        sys.exit(1)
    
    with open(secrets_path, 'r') as f:
        return toml.load(f)

def update_quickbooks_refresh_token(config):
    """Update QuickBooks refresh token if a new one is available"""
    try:
        if 'quickbooks' not in config:
            print("⚠️ No QuickBooks configuration found")
            return False
        
        qb_config = config['quickbooks']
        client_id = qb_config.get('client_id')
        client_secret = qb_config.get('client_secret')
        current_refresh_token = qb_config.get('refresh_token')
        
        if not current_refresh_token:
            print("⚠️ No refresh_token found in secrets.toml")
            return False
        
        # Try to refresh the token
        auth_url = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
        
        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json"
        }
        
        data = {
            "grant_type": "refresh_token",
            "refresh_token": current_refresh_token
        }
        
        response = requests.post(
            auth_url,
            data=data,
            headers=headers,
            auth=(client_id, client_secret),
            timeout=30
        )
        
        if response.status_code == 200:
            auth_response = response.json()
            new_refresh_token = auth_response.get("refresh_token")
            
            if new_refresh_token and new_refresh_token != current_refresh_token:
                # Update the secrets.toml file
                qb_config['refresh_token'] = new_refresh_token
                
                secrets_path = os.path.join('.streamlit', 'secrets.toml')
                with open(secrets_path, 'w') as f:
                    toml.dump(config, f)
                
                print(f"✅ QuickBooks refresh_token updated successfully")
                return True
            else:
                print("ℹ️ Refresh token is up to date")
                return True
        else:
            print(f"⚠️ Could not refresh QuickBooks token: {response.status_code}")
            return False
            
    except Exception as e:
        print(f"⚠️ Error updating QuickBooks token: {str(e)}")
        return False

def refresh_monday_database(config):
    """Refresh Monday.com database"""
    try:
        if 'monday' not in config:
            print("❌ No Monday.com configuration found")
            return False
        
        monday_config = config['monday']
        api_token = monday_config['api_token']
        
        boards_config = [
            (monday_config['sales_board_id'], 'sales_board'),
            (monday_config['new_leads_board_id'], 'new_leads_board'),
            (monday_config['discovery_call_board_id'], 'discovery_call_board'),
            (monday_config['design_review_board_id'], 'design_review_board'),
            (monday_config['ads_board_id'], 'ads_board')
        ]
        
        success_count = 0
        
        for board_id, table_name in boards_config:
            try:
                all_items = []
                cursor = None
                limit = 500 if table_name != 'sales_board' else 200
                page_count = 0
                
                while page_count < 50:
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
                    
                    url = "https://api.monday.com/v2"
                    headers = {
                        "Authorization": api_token,
                        "Content-Type": "application/json",
                    }
                    
                    response = requests.post(url, json={"query": query}, headers=headers, timeout=120)
                    data = response.json()
                    
                    if "errors" in data:
                        break
                    
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
                
                # Save to database
                conn = sqlite3.connect(MONDAY_DB_PATH)
                cursor = conn.cursor()
                
                # Clear existing data
                cursor.execute(f"DELETE FROM {table_name}")
                
                # Insert new data
                for item in all_items:
                    item_id = item.get("id", "")
                    name = item.get("name", "")
                    column_values = str(item.get("column_values", []))
                    
                    cursor.execute(f'''
                        INSERT INTO {table_name} (id, name, board_type, column_values, updated_at)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (item_id, name, table_name, column_values, datetime.now()))
                
                conn.commit()
                conn.close()
                
                print(f"✅ {table_name}: {len(all_items)} items saved")
                success_count += 1
                
            except Exception as e:
                print(f"❌ {table_name}: Error - {str(e)}")
        
        print(f"\n✅ Monday.com refresh complete: {success_count}/5 boards updated")
        return True
        
    except Exception as e:
        print(f"❌ Error refreshing Monday.com database: {str(e)}")
        return False

def refresh_calendly_database(config):
    """Refresh Calendly database"""
    try:
        if 'calendly' not in config:
            print("❌ No Calendly configuration found")
            return False
        
        calendly_config = config['calendly']
        api_key = calendly_config['calendly_api_key']
        
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        
        # Get user info
        user_response = requests.get('https://api.calendly.com/users/me', headers=headers)
        if user_response.status_code != 200:
            print(f"❌ Failed to get Calendly user info: {user_response.status_code}")
            return False
        
        user_data = user_response.json()
        user_uri = user_data['resource']['uri']
        
        # Get event types
        event_types_response = requests.get(f'https://api.calendly.com/event_types?user={user_uri}', headers=headers)
        if event_types_response.status_code != 200:
            print(f"❌ Failed to get Calendly event types: {event_types_response.status_code}")
            return False
        
        event_types_data = event_types_response.json()
        event_types = event_types_data.get('collection', [])
        
        # Find "TEG - Let's Chat" event type
        teg_event_types = []
        for event_type in event_types:
            if event_type.get('name') == "TEG - Let's Chat":
                teg_event_types.append(event_type)
        
        if not teg_event_types:
            print("❌ No 'TEG - Let's Chat' event type found")
            return False
        
        all_events = []
        
        for event_type in teg_event_types:
            event_type_uri = event_type['uri']
            event_type_uuid = event_type_uri.split('/')[-1]
            
            # Fetch events with pagination
            page_count = 0
            next_page_token = None
            
            min_start_time = "2025-01-01T00:00:00.000000Z"
            max_start_time = "2025-10-23T23:59:59.999999Z"
            
            while page_count < 100:
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
                    break
                
                events_data = events_response.json()
                events = events_data.get('collection', [])
                
                if not events:
                    break
                
                all_events.extend(events)
                
                pagination = events_data.get('pagination', {})
                next_page_token = pagination.get('next_page_token')
                
                if not next_page_token:
                    break
        
        # Save to database
        conn = sqlite3.connect(CALENDLY_DB_PATH)
        cursor = conn.cursor()
        
        cursor.execute("DELETE FROM calendly_events")
        
        for event in all_events:
            uri = event.get('uri', '')
            name = event.get('name', '')
            start_time = event.get('start_time', '')
            end_time = event.get('end_time', '')
            status = event.get('status', '')
            event_type = event.get('event_type', '')
            
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
        
        print(f"✅ Calendly refresh complete: {len(all_events)} events saved")
        return True
        
    except Exception as e:
        print(f"❌ Error refreshing Calendly database: {str(e)}")
        return False

def main():
    """Main function"""
    print("=" * 80)
    print(f"DATABASE REFRESH - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)
    
    # Initialize databases
    print("\n🔧 Initializing databases...")
    init_databases()
    
    # Load configuration
    try:
        config = load_config()
    except Exception as e:
        print(f"❌ Error loading configuration: {str(e)}")
        sys.exit(1)
    
    print("\n🔄 Step 1: Updating QuickBooks refresh token...")
    qb_success = update_quickbooks_refresh_token(config)
    
    print("\n🔄 Step 2: Refreshing Monday.com database...")
    monday_success = refresh_monday_database(config)
    
    print("\n🔄 Step 3: Refreshing Calendly database...")
    calendly_success = refresh_calendly_database(config)
    
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"QuickBooks Token: {'✅ Success' if qb_success else '⚠️ Skipped'}")
    print(f"Monday.com DB: {'✅ Success' if monday_success else '❌ Failed'}")
    print(f"Calendly DB: {'✅ Success' if calendly_success else '❌ Failed'}")
    print("=" * 80)
    
    if monday_success and calendly_success:
        print("\n✅ All databases refreshed successfully!")
        sys.exit(0)
    else:
        print("\n⚠️ Some database refreshes failed (check errors above)")
        sys.exit(1)

if __name__ == "__main__":
    main()

