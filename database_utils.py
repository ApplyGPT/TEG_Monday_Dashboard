import sqlite3
import pandas as pd
import json
import ast
from datetime import datetime
import os

# Database configuration
DB_PATH = "monday_data.db"

def get_db_connection():
    """Get SQLite database connection"""
    return sqlite3.connect(DB_PATH)

def get_board_data(table_name):
    """Get all data from a specific board table"""
    conn = get_db_connection()
    
    try:
        query = f"SELECT * FROM {table_name} ORDER BY updated_at DESC"
        df = pd.read_sql_query(query, conn)
        
        # Don't parse JSON here - let get_board_data_as_items handle it
        # This avoids the JSON parsing error
        
        return df
    except Exception as e:
        print(f"Error reading from {table_name}: {str(e)}")
        return pd.DataFrame()
    finally:
        conn.close()

def get_board_data_as_items(table_name):
    """Get board data in the same format as Monday.com API (for compatibility)"""
    df = get_board_data(table_name)
    
    if df.empty:
        return []
    
    items = []
    for _, row in df.iterrows():
        try:
            # Parse column_values JSON string back to list
            column_values_str = row['column_values']
            if isinstance(column_values_str, str):
                # Try to parse as JSON first
                try:
                    column_values = json.loads(column_values_str)
                except json.JSONDecodeError:
                    # If JSON parsing fails, try ast.literal_eval for Python dict syntax
                    try:
                        column_values = ast.literal_eval(column_values_str)
                    except (ValueError, SyntaxError):
                        # If both fail, create empty list
                        print(f"Warning: Could not parse column_values for item {row['id']}: {column_values_str[:100]}...")
                        column_values = []
            else:
                column_values = column_values_str if column_values_str else []
            
            item = {
                'id': row['id'],
                'name': row['name'],
                'column_values': column_values
            }
            items.append(item)
        except Exception as e:
            print(f"Error processing item {row.get('id', 'unknown')}: {str(e)}")
            continue
    
    return items

def debug_sales_board():
    """Debug function to see what's in the sales board"""
    items = get_board_data_as_items('sales_board')
    
    print(f"Total items in sales_board: {len(items)}")
    
    # Look for Susan Glenn specifically
    susan_items = []
    for item in items:
        name = item.get("name", "")
        if "susan" in name.lower() or "glenn" in name.lower():
            susan_items.append(item)
    
    print(f"Items containing 'susan' or 'glenn': {len(susan_items)}")
    
    # Show first few items to see the structure
    print("\nFirst 3 items in sales_board:")
    for i, item in enumerate(items[:3]):
        print(f"Item {i+1}: {item.get('name', 'No name')}")
        print(f"  ID: {item.get('id', 'No ID')}")
        print(f"  Columns: {len(item.get('column_values', []))}")
        
        # Show date columns
        for col_val in item.get("column_values", []):
            if col_val.get("type") == "date":
                print(f"    DATE: {col_val.get('text', 'No text')} (ID: {col_val.get('id', 'No ID')})")
        print()
    
    return susan_items

def search_item_by_name(item_name):
    """Search for a specific item across all boards"""
    boards = ['sales_board', 'new_leads_board', 'discovery_call_board', 'design_review_board']
    results = []
    
    for board in boards:
        items = get_board_data_as_items(board)
        
        for item in items:
            item_name_lower = item.get("name", "").lower()
            search_term_lower = item_name.lower()
            
            # Try different search methods
            if (search_term_lower in item_name_lower or 
                item_name_lower in search_term_lower or
                any(word in item_name_lower for word in search_term_lower.split())):
                
                results.append({
                    'board': board,
                    'item': item,
                    'name': item.get("name", ""),
                    'id': item.get("id", ""),
                    'column_values': item.get("column_values", [])
                })
    
    return results

def find_discovery_call_date_columns():
    """Find the specific Discovery Call Date column for each board"""
    # Based on user feedback, we need to be more specific about which column is "Discovery Call Date"
    # Let's use a more conservative approach - only the most likely Discovery Call Date columns
    known_discovery_columns = {
        'sales_board': ['date_mktqwpzz'],  # This one we know works from Susan Glenn
        'new_leads_board': [],  # Need to identify the correct one
        'discovery_call_board': ['date_mktbrpz6'],  # Most likely the main Discovery Call Date
        'design_review_board': ['date_mktqx5xa']  # Madeline Carter example (2025-10-17)
    }
    
    return known_discovery_columns

def get_discovery_call_dates():
    """Get discovery call dates from specific Discovery Call Date columns for 2025 - QUALIFIED ONLY"""
    boards = ['sales_board', 'new_leads_board', 'discovery_call_board', 'design_review_board']
    all_dates = []
    
    # Find Discovery Call Date columns dynamically
    discovery_columns = find_discovery_call_date_columns()
    
    for board in boards:
        items = get_board_data_as_items(board)
        board_discovery_columns = discovery_columns.get(board, [])
        
        for item in items:
            item_name = item.get("name", "")
            
            # Filter out items that start with "No", "Not", or "Spam"
            if item_name.lower().startswith(('no ', 'not ', 'spam', 'no/', 'not/')):
                continue
            
            # For Discovery Call board, also filter out "NOT QUALIFIED" items
            if board == 'discovery_call_board':
                # Check if item name contains "NOT QUALIFIED" or similar
                if any(phrase in item_name.upper() for phrase in ['NOT QUALIFIED', 'NOT QUA', 'NOT QUAL']):
                    continue
                
                # Check qualification status - only count QUALIFIED items
                qualification_status = None
                for col_val in item.get("column_values", []):
                    if col_val.get("id") == "color_mknxk7eq":  # Qualification status column
                        qualification_status = col_val.get("text", "")
                        break
                
                # Only count items that are "Qualified"
                if qualification_status != "Qualified":
                    continue
                
                # Additional filter: exclude items with "Unqualified" status in color_mknx1h9r
                status_mknx1h9r = None
                for col_val in item.get("column_values", []):
                    if col_val.get("id") == "color_mknx1h9r":
                        status_mknx1h9r = col_val.get("text", "")
                        break
                
                # Exclude items that are "Unqualified" in this status column
                if status_mknx1h9r == "Unqualified":
                    continue
                
                # Additional filter: exclude items with "Negotiation" contract status
                contract_status = None
                for col_val in item.get("column_values", []):
                    if col_val.get("id") == "contract_status":
                        contract_status = col_val.get("text", "")
                        break
                
                # Exclude items that are in "Negotiation" status
                if contract_status == "Negotiation":
                    continue
            
            # For Sales board, also filter by qualification status
            if board == 'sales_board':
                # Check qualification status - only count QUALIFIED items
                qualification_status = None
                for col_val in item.get("column_values", []):
                    if col_val.get("id") == "color_mknxg5zf":  # Sales board Rejection Reason column
                        qualification_status = col_val.get("text", "")
                        break
                
                # Only count items that are "Qualified"
                if qualification_status != "Qualified":
                    continue
            
            # For Design Review board, also filter by qualification status
            if board == 'design_review_board':
                # Check qualification status - only count QUALIFIED items
                qualification_status = None
                for col_val in item.get("column_values", []):
                    if col_val.get("id") == "color_mknxrx3c":  # Design Review board Rejection Reason column
                        qualification_status = col_val.get("text", "")
                        break
                
                # Only count items that are "Qualified"
                if qualification_status != "Qualified":
                    continue
            
            for col_val in item.get("column_values", []):
                col_id = col_val.get("id", "")
                col_type = col_val.get("type", "")
                text = col_val.get("text", "")
                
                # Only look at Discovery Call Date columns
                if col_id in board_discovery_columns and col_type == "date" and text and text.strip():
                    try:
                        # Try multiple date parsing methods
                        parsed_date = None
                        
                        # Method 1: Direct pandas parsing
                        parsed_date = pd.to_datetime(text, errors='coerce')
                        
                        # Method 2: Try different date formats if first fails
                        if pd.isna(parsed_date):
                            # Try common date formats
                            date_formats = [
                                '%Y-%m-%d',
                                '%m/%d/%Y',
                                '%d/%m/%Y',
                                '%Y-%m-%d %H:%M:%S',
                                '%m/%d/%Y %H:%M:%S'
                            ]
                            
                            for fmt in date_formats:
                                try:
                                    parsed_date = pd.to_datetime(text, format=fmt)
                                    break
                                except:
                                    continue
                        
                        if not pd.isna(parsed_date) and parsed_date.year == 2025:
                            all_dates.append({
                                'date': parsed_date,
                                'item_name': item_name,
                                'board': board.replace('_board', '').replace('_', ' ').title(),
                                'column_id': col_id,
                                'raw_text': text
                            })
                    except:
                        continue
    
    return all_dates

def get_sales_data():
    """Get sales data in the format expected by sales dashboard with filtering"""
    items = get_board_data_as_items('sales_board')
    
    # Filter out items that start with "No", "Not", or "Spam"
    filtered_items = []
    for item in items:
        item_name = item.get("name", "")
        if not item_name.lower().startswith(('no ', 'not ', 'spam')):
            filtered_items.append(item)
    
    if not filtered_items:
        return {"data": {"boards": [{"items_page": {"items": []}}]}}
    
    return {
        "data": {
            "boards": [{
                "items_page": {
                    "items": filtered_items
                }
            }]
        }
    }

def get_ads_data():
    """Get ads data in the format expected by ads dashboard with filtering"""
    items = get_board_data_as_items('ads_board')
    
    # Filter out items that start with "No", "Not", or "Spam"
    filtered_items = []
    for item in items:
        item_name = item.get("name", "")
        if not item_name.lower().startswith(('no ', 'not ', 'spam')):
            filtered_items.append(item)
    
    if not filtered_items:
        return {"data": {"boards": [{"items_page": {"items": []}}]}}
    
    return {
        "data": {
            "boards": [{
                "items_page": {
                    "items": filtered_items
                }
            }]
        }
    }

def get_new_leads_data():
    """Get new leads data with filtering"""
    items = get_board_data_as_items('new_leads_board')
    
    # Filter out items that start with "No", "Not", or "Spam"
    filtered_items = []
    for item in items:
        item_name = item.get("name", "")
        if not item_name.lower().startswith(('no ', 'not ', 'spam')):
            filtered_items.append(item)
    
    return filtered_items

def get_discovery_call_data():
    """Get discovery call data with filtering"""
    items = get_board_data_as_items('discovery_call_board')
    
    # Filter out items that start with "No", "Not", or "Spam"
    filtered_items = []
    for item in items:
        item_name = item.get("name", "")
        if not item_name.lower().startswith(('no ', 'not ', 'spam')):
            filtered_items.append(item)
    
    return filtered_items

def get_design_review_data():
    """Get design review data with filtering"""
    items = get_board_data_as_items('design_review_board')
    
    # Filter out items that start with "No", "Not", or "Spam"
    filtered_items = []
    for item in items:
        item_name = item.get("name", "")
        if not item_name.lower().startswith(('no ', 'not ', 'spam')):
            filtered_items.append(item)
    
    return filtered_items

def check_database_exists():
    """Check if database exists and has data"""
    if not os.path.exists(DB_PATH):
        return False, "Database file does not exist"
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        # Check if tables exist and have data
        tables = ['sales_board', 'new_leads_board', 'discovery_call_board', 'design_review_board', 'ads_board']
        
        for table in tables:
            cursor.execute(f"SELECT COUNT(*) FROM {table}")
            count = cursor.fetchone()[0]
            if count == 0:
                return False, f"Table {table} is empty"
        
        return True, "Database is ready"
    except Exception as e:
        return False, f"Database error: {str(e)}"
    finally:
        conn.close()

def get_database_info():
    """Get information about the database"""
    if not os.path.exists(DB_PATH):
        return {"exists": False, "size": 0, "tables": {}}
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        tables = ['sales_board', 'new_leads_board', 'discovery_call_board', 'design_review_board', 'ads_board']
        table_info = {}
        
        for table in tables:
            try:
                cursor.execute(f"SELECT COUNT(*) FROM {table}")
                count = cursor.fetchone()[0]
                
                cursor.execute(f"SELECT MAX(updated_at) FROM {table}")
                last_updated = cursor.fetchone()[0]
                
                table_info[table] = {
                    'count': count,
                    'last_updated': last_updated
                }
            except:
                table_info[table] = {
                    'count': 0,
                    'last_updated': 'Never'
                }
        
        file_size = os.path.getsize(DB_PATH)
        
        return {
            "exists": True,
            "size": file_size,
            "tables": table_info
        }
    finally:
        conn.close()
