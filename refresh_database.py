"""
Standalone Database Refresh Script
Can be run via cron job to refresh Monday.com and Calendly databases
"""
import requests
import sqlite3
import os
import toml
import sys
import subprocess
import time
import traceback
from datetime import datetime, timedelta

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
            source TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    try:
        cursor.execute("ALTER TABLE calendly_events ADD COLUMN source TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # Column already exists
    conn.commit()
    conn.close()

def load_config():
    """Load configuration from secrets.toml"""
    # Get absolute path to ensure we're reading the right file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    secrets_path = os.path.join(script_dir, '.streamlit', 'secrets.toml')
    
    print(f"🔍 Looking for secrets.toml at: {secrets_path}")
    print(f"🔍 Current working directory: {os.getcwd()}")
    
    if not os.path.exists(secrets_path):
        print(f"❌ ERROR: secrets.toml not found at {secrets_path}")
        print(f"   File exists: {os.path.exists(secrets_path)}")
        print(f"   Absolute path: {os.path.abspath(secrets_path)}")
        sys.exit(1)
    
    print(f"✅ Found secrets.toml")
    
    with open(secrets_path, 'r') as f:
        config = toml.load(f)
    
    print(f"✅ Loaded configuration with sections: {list(config.keys())}")
    return config

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
                    
                    url = "https://api.monday.com/v2"
                    headers = {
                        "Authorization": api_token,
                        "Content-Type": "application/json",
                    }
                    
                    # Retry logic for rate limit errors
                    max_retries = 5
                    retry_count = 0
                    data = None
                    
                    while retry_count < max_retries:
                        try:
                            response = requests.post(url, json={"query": query}, headers=headers, timeout=120)
                            data = response.json()
                            
                            if "errors" in data:
                                errors = data['errors']
                                # Check if it's a rate limit error
                                is_rate_limit = False
                                retry_seconds = 0
                                # Check if it's Monday.com internal server error (transient; retry with backoff)
                                is_internal_error = False
                                
                                for error in errors:
                                    extensions = error.get('extensions', {})
                                    status_code = extensions.get('status_code')
                                    code = extensions.get('code')
                                    msg = (error.get('message') or '').lower()
                                    
                                    if status_code == 429 or 'RATE_LIMIT' in str(code) or 'LIMIT_EXCEEDED' in str(code):
                                        is_rate_limit = True
                                        retry_seconds = max(retry_seconds, extensions.get('retry_in_seconds', 10))
                                    if code == 'INTERNAL_SERVER_ERROR' or 'internal server error' in msg:
                                        is_internal_error = True
                                
                                if is_rate_limit and retry_count < max_retries - 1:
                                    retry_count += 1
                                    wait_time = retry_seconds + (retry_count * 2)  # Add some buffer
                                    print(f"⏳ Rate limit hit for {table_name}, waiting {wait_time}s before retry {retry_count}/{max_retries-1}...")
                                    time.sleep(wait_time)
                                    continue
                                if is_internal_error and retry_count < max_retries - 1:
                                    retry_count += 1
                                    wait_time = 10 + (retry_count * 5)  # 15s, 20s, ...
                                    print(f"⏳ Monday.com internal server error for {table_name}, waiting {wait_time}s before retry {retry_count}/{max_retries-1}...")
                                    time.sleep(wait_time)
                                    continue
                                # Not retryable or max retries reached
                                print(f"❌ GraphQL errors for {table_name}: {errors}")
                                all_items = []
                                break
                            else:
                                # Success - break out of retry loop
                                break
                                
                        except Exception as e:
                            if retry_count < max_retries - 1:
                                retry_count += 1
                                wait_time = (retry_count * 2) + 5  # Exponential backoff
                                print(f"⏳ Request error for {table_name}, waiting {wait_time}s before retry {retry_count}/{max_retries-1}: {str(e)}")
                                time.sleep(wait_time)
                                continue
                            else:
                                print(f"❌ {table_name}: Error after {max_retries} retries - {str(e)}")
                                all_items = []
                                break
                    
                    if data is None or "errors" in data:
                        if retry_count >= max_retries:
                            print(f"❌ {table_name}: Max retries reached, skipping...")
                        all_items = []
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
                    
                    # Small delay between pages to avoid rate limits
                    time.sleep(0.5)
                
                # If no items and we hit an error earlier, skip saving empty set
                if not all_items:
                    print(f"⚠️ {table_name}: No items to save (skipping table write)")
                else:
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
                
                print(f"✅ {table_name}: {len(all_items)} items processed")
                if all_items:
                    success_count += 1
                
            except Exception as e:
                print(f"❌ {table_name}: Error - {str(e)}")
            
            # Add delay between boards to avoid concurrency limits
            if table_name != boards_config[-1][1]:  # Don't delay after last board
                print(f"⏸️  Waiting 2s before next board...")
                time.sleep(2)
        
        print(f"\n✅ Monday.com refresh complete: {success_count}/5 boards updated")
        return True
        
    except Exception as e:
        print(f"❌ Error refreshing Monday.com database: {str(e)}")
        return False

def refresh_calendly_database(config):
    """Refresh Calendly database"""
    try:
        # Debug: Print config structure
        print(f"🔍 Config keys: {list(config.keys())}")
        
        if 'calendly' not in config:
            print("❌ No Calendly configuration found")
            print(f"   Available sections: {list(config.keys())}")
            return False
        
        calendly_config = config['calendly']
        print(f"🔍 Calendly config keys: {list(calendly_config.keys())}")
        
        api_key = calendly_config.get('calendly_api_key') or calendly_config.get('api_key')
        burki_key = calendly_config.get('calendly_burki_api_key')
        if not api_key and not burki_key:
            print("❌ No Calendly API key found (set calendly_api_key or calendly_burki_api_key)")
            return False
        api_key = api_key or burki_key
        # Request by year to avoid API result limit (API returns ~9k events; ascending order so 2026 can be cut off)
        now = datetime.now()
        start_year = now.year - 1
        end_year = now.year + 1
        year_ranges = [
            (y, f"{y}-01-01T00:00:00.000000Z", f"{y}-12-31T23:59:59.999999Z")
            for y in range(start_year, end_year + 1)
        ]
        print(f"   Requesting Calendly events by year: {start_year} to {end_year}")
        
        # Helpers (used by both org path and user path)
        DESIGN_REVIEW_PERSONS = [
            ('anthony-the-evans-group', 'Anthony'),
            ('heather-the-evans-group', 'Heather'),
            ('ian-the-evans-group', 'Ian'),
        ]
        def _person_from_user_resource(user_resource):
            """Infer Anthony/Heather/Ian from User resource (name, slug) - GET /users/{uuid} response."""
            name_lower = (user_resource.get('name') or '').lower()
            slug_lower = (user_resource.get('slug') or '').lower()
            if 'anthony' in name_lower or 'anthony' in slug_lower:
                return 'Anthony'
            if 'heather' in name_lower or 'heather' in slug_lower:
                return 'Heather'
            if ('ian' in name_lower or 'ian' in slug_lower) and 'christian' not in name_lower and 'christian' not in slug_lower:
                return 'Ian'
            return ''
        owner_to_person = {}
        def _person_from_owner(owner_uri):
            """Resolve profile.owner via GET /users/{uuid} and infer person from user name/slug (cached)."""
            if not owner_uri or owner_uri in owner_to_person:
                return owner_to_person.get(owner_uri, '')
            uuid = owner_uri.rstrip('/').split('/')[-1]
            try:
                r = requests.get(f'https://api.calendly.com/users/{uuid}', headers=headers, timeout=10)
                if r.status_code == 200:
                    user_resource = r.json().get('resource', {})
                    owner_to_person[owner_uri] = _person_from_user_resource(user_resource)
                else:
                    owner_to_person[owner_uri] = ''
            except Exception:
                owner_to_person[owner_uri] = ''
            return owner_to_person[owner_uri]
        def _person_from_event_type(et):
            """Infer Anthony/Heather/Ian from event type: scheduling_url, slug, name, profile.name, or GET user by profile.owner."""
            url = (et.get('scheduling_url') or '').lower()
            slug_lower = (et.get('slug') or '').lower()
            name_lower = (et.get('name') or '').lower()
            for url_part, person in DESIGN_REVIEW_PERSONS:
                if url_part in url or url_part.replace('-the-evans-group', '') in name_lower or url_part.replace('-the-evans-group', '') in slug_lower:
                    return person
            profile = et.get('profile') or {}
            profile_name = (profile.get('name') or '').lower()
            if 'anthony' in profile_name:
                return 'Anthony'
            if 'heather' in profile_name:
                return 'Heather'
            if 'ian' in profile_name and 'christian' not in profile_name:
                return 'Ian'
            # Profile.owner is user URI when type=User; when type=Team, owner is team URI (GET /users not applicable)
            owner_uri = profile.get('owner')
            if owner_uri and (profile.get('type') == 'User' or '/users/' in owner_uri):
                return _person_from_owner(owner_uri)
            return ''
        def _person_from_event_memberships(ev):
            """Infer Anthony/Heather/Ian from event.event_memberships (round-robin: who actually handled the call)."""
            memberships = ev.get('event_memberships') or []
            for m in memberships:
                if not isinstance(m, dict):
                    continue
                user_name = (m.get('user_name') or '').lower()
                user_email = (m.get('user_email') or '').lower()
                if 'anthony' in user_name or 'anthony' in user_email:
                    return 'Anthony'
                if 'heather' in user_name or 'heather' in user_email:
                    return 'Heather'
                if ('ian' in user_name or 'ian' in user_email) and 'christian' not in user_name and 'christian' not in user_email:
                    return 'Ian'
                user_uri = m.get('user')
                if user_uri and '/users/' in user_uri:
                    p = _person_from_owner(user_uri)
                    if p:
                        return p
            return ''
        def _teg_relevant_and_source(et):
            """Return (is_teg_relevant, default_source) for an event type."""
            event_name = et.get('name', '')
            scheduling_url = (et.get('scheduling_url') or '').lower()
            name_lower = (event_name or '').lower()
            if event_name and "teg" in name_lower and ("let's chat" in name_lower or "lets chat" in name_lower):
                return True, ''
            if 'teg-introductory-call' in scheduling_url or 'introductory' in name_lower or 'intro call' in name_lower:
                return True, _person_from_event_type(et)
            person = _person_from_event_type(et)
            if person:
                return True, person
            if '/30min' in scheduling_url or '/30-min' in scheduling_url or '30 minute' in name_lower or '30 min' in name_lower:
                return True, 'Design Review'
            return False, ''
        
        all_events = []
        # Optional: try org scope with main key (v2 admin); on timeout/403 use user-scoped only
        use_org_path = False
        org_uri = None
        if api_key and (not burki_key or api_key != burki_key):
            try:
                r_me = requests.get('https://api.calendly.com/users/me',
                                     headers={'Authorization': f'Bearer {api_key}', 'Content-Type': 'application/json'},
                                     timeout=30)
                if r_me.status_code == 200:
                    org_uri = (r_me.json().get('resource') or {}).get('current_organization')
            except Exception:
                pass
        if org_uri and api_key:
            try:
                # Probe one year to confirm org scope works
                y, mn, mx = year_ranges[0]
                r_probe = requests.get('https://api.calendly.com/scheduled_events',
                                       headers={'Authorization': f'Bearer {api_key}', 'Content-Type': 'application/json'},
                                       params={'organization': org_uri, 'min_start_time': mn, 'max_start_time': mx, 'count': 100},
                                       timeout=90)
                if r_probe.status_code == 200:
                    use_org_path = True
            except (requests.exceptions.ReadTimeout, requests.exceptions.RequestException):
                pass
        if use_org_path:
            org_headers = {'Authorization': f'Bearer {api_key}', 'Content-Type': 'application/json'}
            headers = org_headers  # for GET event_types and _person_from_owner
            print("   Using organization scope (admin token)...")
            raw_org_events = []
            for year, min_start_time, max_start_time in year_ranges:
                print(f"   Fetching {year}...")
                page = 0
                next_token = None
                while page < 100:
                    page += 1
                    if page == 1:
                        r = requests.get('https://api.calendly.com/scheduled_events', headers=org_headers,
                                         params={'organization': org_uri, 'min_start_time': min_start_time, 'max_start_time': max_start_time, 'count': 100}, timeout=90)
                    else:
                        r = requests.get('https://api.calendly.com/scheduled_events', headers=org_headers,
                                         params={'organization': org_uri, 'min_start_time': min_start_time, 'max_start_time': max_start_time, 'count': 100, 'page_token': next_token}, timeout=90)
                    if r.status_code != 200:
                        break
                    data = r.json()
                    coll = data.get('collection', [])
                    for item in coll:
                        if isinstance(item, dict) and "resource" in item:
                            raw_org_events.append({**item["resource"], "uri": item.get("uri") or item["resource"].get("uri")})
                        else:
                            raw_org_events.append(item if isinstance(item, dict) else {})
                    next_token = data.get('pagination', {}).get('next_page_token')
                    if not next_token:
                        break
            event_type_cache = {}
            for ev in raw_org_events:
                et_raw = ev.get('event_type')
                et_uri = et_raw if isinstance(et_raw, str) else (et_raw.get('uri') if isinstance(et_raw, dict) else None)
                if et_uri and et_uri not in event_type_cache:
                    uuid = et_uri.rstrip('/').split('/')[-1]
                    try:
                        rr = requests.get(f'https://api.calendly.com/event_types/{uuid}', headers=headers, timeout=10)
                        event_type_cache[et_uri] = rr.json().get('resource', {}) if rr.status_code == 200 else {}
                    except Exception:
                        event_type_cache[et_uri] = {}
                et = event_type_cache.get(et_uri) if et_uri else {}
                relevant, default_source = _teg_relevant_and_source(et) if et else (False, '')
                if not relevant and ev.get('name'):
                    relevant, default_source = _teg_relevant_and_source({'name': ev.get('name'), 'scheduling_url': '', 'slug': ''})
                if relevant:
                    src = _person_from_event_memberships(ev) or default_source
                    all_events.append({**ev, 'source': src})
        else:
            # USER PATH: fetch with one or both keys (Burki + v2) and merge events
            if burki_key and api_key and burki_key != api_key:
                keys_to_fetch = [burki_key, api_key]
                print(f"🔑 Using both Calendly keys (Burki + v2) for user-scoped fetch...")
            else:
                keys_to_fetch = [burki_key or api_key]
                print(f"🔑 Using Calendly API key: {(burki_key or api_key)[:30]}...")
            for key in keys_to_fetch:
                headers = {'Authorization': f'Bearer {key}', 'Content-Type': 'application/json'}
                user_response = requests.get('https://api.calendly.com/users/me', headers=headers, timeout=30)
                if user_response.status_code != 200:
                    print(f"   ⚠️ Skip key ...{key[-8:]}: users/me returned {user_response.status_code}")
                    continue
                user_uri = (user_response.json().get('resource') or {}).get('uri')
                if not user_uri:
                    continue
                resp = requests.get(f'https://api.calendly.com/event_types?user={user_uri}', headers=headers, timeout=30)
                if resp.status_code != 200:
                    print(f"   ⚠️ Skip key ...{key[-8:]}: event_types returned {resp.status_code}")
                    continue
                event_types = resp.json().get('collection', [])
                teg_event_types = []
                for event_type in event_types:
                    event_name = event_type.get('name', '')
                    scheduling_url = (event_type.get('scheduling_url') or '').lower()
                    name_lower = (event_name or '').lower()
                    source = ''
                    if event_name and "teg" in name_lower and ("let's chat" in name_lower or "lets chat" in name_lower):
                        teg_event_types.append((event_type, source))
                    elif 'teg-introductory-call' in scheduling_url or 'introductory' in name_lower or 'intro call' in name_lower:
                        teg_event_types.append((event_type, _person_from_event_type(event_type)))
                    else:
                        person = _person_from_event_type(event_type)
                        if person:
                            teg_event_types.append((event_type, person))
                        elif '/30min' in scheduling_url or '/30-min' in scheduling_url or '30 minute' in name_lower or '30 min' in name_lower:
                            teg_event_types.append((event_type, 'Design Review'))
                if not teg_event_types:
                    continue
                for event_type, source in teg_event_types:
                    event_type_uri = event_type['uri']
                    event_type_uuid = event_type_uri.split('/')[-1]
                    event_name = event_type['name']
                    print(f"   Fetching events for: {event_name}")
                    for year, min_start_time, max_start_time in year_ranges:
                        page_count = 0
                        next_page_token = None
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
                                                          headers=headers, params=params, timeout=60)
                            if events_response.status_code != 200:
                                print(f"   ⚠️ Failed to get events for {event_name} ({year}): {events_response.status_code}")
                                break
                            events_data = events_response.json()
                            raw_collection = events_data.get('collection', [])
                            events = []
                            for item in raw_collection:
                                if isinstance(item, dict) and "resource" in item:
                                    events.append({**item["resource"], "uri": item.get("uri") or item["resource"].get("uri")})
                                else:
                                    events.append(item if isinstance(item, dict) else {})
                            if not events:
                                break
                            for event in events:
                                if not event.get("name") and event_name:
                                    event = {**event, "name": event_name}
                                source_from_membership = _person_from_event_memberships(event)
                                final_source = source_from_membership if source_from_membership else source
                                all_events.append({**event, "source": final_source})
                            next_page_token = events_data.get('pagination', {}).get('next_page_token')
                            if not next_page_token:
                                break
            if not all_events:
                print("❌ No TEG events from any key ('TEG - Let's Chat', TEG Introductory Call, or Design Review 30min)")
                return False
        
        # Deduplicate by URI (same event can be returned under multiple event types)
        seen_uris = set()
        unique_events = []
        for ev in all_events:
            u = (ev.get('uri') or '').strip()
            if u and u not in seen_uris:
                seen_uris.add(u)
                unique_events.append(ev)
        
        # Save to database (same logic as pages/database_refresh.py save_calendly_data_to_db)
        conn = sqlite3.connect(CALENDLY_DB_PATH)
        cursor = conn.cursor()
        try:
            cursor.execute("ALTER TABLE calendly_events ADD COLUMN source TEXT")
            conn.commit()
        except sqlite3.OperationalError:
            pass
        cursor.execute("DELETE FROM calendly_events")
        saved_count = 0
        for event in unique_events:
            if not isinstance(event, dict):
                continue
            uri = event.get('uri') or ''
            name = event.get('name') or ''
            start_time = event.get('start_time') or ''
            end_time = event.get('end_time') or ''
            status = event.get('status') or ''
            raw_event_type = event.get('event_type', '')
            if isinstance(raw_event_type, dict):
                event_type_val = raw_event_type.get('uri') or raw_event_type.get('name') or ''
            else:
                event_type_val = str(raw_event_type) if raw_event_type else ''
            source = event.get('source') or ''
            invitees = event.get('invitees') or []
            invitee_name = invitee_email = ""
            if invitees and isinstance(invitees[0], dict):
                invitee_name = invitees[0].get('name') or ''
                invitee_email = invitees[0].get('email') or ''
            cursor.execute('''
                INSERT OR REPLACE INTO calendly_events 
                (uri, name, start_time, end_time, status, event_type, invitee_name, invitee_email, source, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (uri, name, start_time, end_time, status, event_type_val, invitee_name, invitee_email, source, datetime.now()))
            saved_count += 1
        conn.commit()
        conn.close()
        
        print(f"✅ Calendly refresh complete: {saved_count} events saved (out of {len(unique_events)} unique)")
        return True
        
    except Exception as e:
        print(f"❌ Error refreshing Calendly database: {str(e)}")
        print(traceback.format_exc())
        return False

def generate_new_leads_cache():
    """Generate the new leads month cache by running the cache generation script."""
    try:
        # Get absolute path to script directory to handle cron job working directory issues
        script_dir = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(script_dir, "scripts", "generate_new_leads_month_cache.py")
        if os.path.exists(script_path):
            result = subprocess.run(
                [sys.executable, script_path],
                capture_output=True,
                text=True,
                cwd=script_dir
            )
            if result.returncode == 0:
                print(result.stdout)
                return True
            else:
                print(f"⚠️ Cache generation returned non-zero exit code")
                if result.stderr:
                    print(result.stderr)
                return False
        else:
            print(f"⚠️ Cache script not found at {script_path}")
            print(f"   Script directory: {script_dir}")
            print(f"   Current working directory: {os.getcwd()}")
            return False
    except Exception as e:
        print(f"❌ Error generating cache: {str(e)}")
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
    
    print("\n🔄 Step 1: Refreshing Monday.com database...")
    monday_success = refresh_monday_database(config)
    
    print("\n🔄 Step 2: Refreshing Calendly database...")
    calendly_success = refresh_calendly_database(config)
    
    print("\n🔄 Step 3: Generating New Leads month cache...")
    cache_success = generate_new_leads_cache()
    
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Monday.com DB: {'✅ Success' if monday_success else '❌ Failed'}")
    print(f"Calendly DB: {'✅ Success' if calendly_success else '❌ Failed'}")
    print(f"New Leads Cache: {'✅ Success' if cache_success else '⚠️ Skipped'}")
    print("=" * 80)
    
    if monday_success and calendly_success:
        print("\n✅ All databases refreshed successfully!")
        sys.exit(0)
    else:
        print("\n⚠️ Some database refreshes failed (check errors above)")
        sys.exit(1)

if __name__ == "__main__":
    main()

