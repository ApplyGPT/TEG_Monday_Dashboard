"""
Script to find ALL column IDs for "Estevao Cavalcante (TEST)" item in Sales v2 board
WITH DETAILED VALUE EXTRACTION for formula columns
"""
import sys
import os
import json

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from database_utils import get_sales_data

def find_all_columns():
    """Find all column IDs for Estevao Cavalcante (TEST)"""
    
    # Get sales data
    sales_data = get_sales_data()
    
    if not sales_data or "data" not in sales_data or "boards" not in sales_data["data"]:
        print("No sales data found")
        return
    
    items = sales_data["data"]["boards"][0]["items_page"]["items"]
    
    # Find "Estevao Cavalcante (TEST)" (exact match)
    test_item = None
    for item in items:
        if item.get("name", "") == "Estevao Cavalcante (TEST)":
            test_item = item
            break
    
    if not test_item:
        print("No item found matching 'Estevao Cavalcante (TEST)' (exact match)")
        return
    
    print(f"Item Name: {test_item.get('name', '')}")
    print(f"Item ID: {test_item.get('id', '')}\n")
    print("="*100)
    print("ALL COLUMNS FOR THIS ITEM:")
    print("="*100)
    
    # Parse column_values
    column_values = test_item.get("column_values", [])
    if isinstance(column_values, str):
        try:
            column_values = json.loads(column_values)
        except:
            try:
                import ast
                column_values = ast.literal_eval(column_values)
            except:
                column_values = []
    
    # Sort columns by ID for easier reading
    sorted_columns = sorted(column_values, key=lambda x: x.get("id", ""))
    
    print(f"\nTotal columns found: {len(sorted_columns)}\n")
    
    # Focus on formula columns with detailed extraction
    print("="*100)
    print("FORMULA COLUMNS (DETAILED EXTRACTION):")
    print("="*100)
    
    for col_val in sorted_columns:
        col_id = col_val.get("id", "")
        col_type = col_val.get("type", "")
        
        if col_type == "formula":
            text = col_val.get("text", "")
            value = col_val.get("value", "")
            
            print(f"\nColumn ID: {col_id}")
            print(f"  Type: {col_type}")
            print(f"  Text (raw): {repr(text)}")
            print(f"  Value (raw): {repr(value)}")
            
            # Try to extract meaningful value
            extracted_value = None
            
            # Check text field
            if text and text.strip():
                extracted_value = text.strip()
                print(f"  ✓ Value from TEXT: '{extracted_value}'")
            
            # Check value field - try multiple parsing methods
            if not extracted_value and value:
                try:
                    # If value is already a dict
                    if isinstance(value, dict):
                        print(f"  Value is dict: {value}")
                        # Try common keys
                        if "number" in value:
                            extracted_value = str(value["number"])
                            print(f"  ✓ Value from dict['number']: '{extracted_value}'")
                        elif "text" in value:
                            extracted_value = str(value["text"])
                            print(f"  ✓ Value from dict['text']: '{extracted_value}'")
                        elif "value" in value:
                            extracted_value = str(value["value"])
                            print(f"  ✓ Value from dict['value']: '{extracted_value}'")
                        else:
                            print(f"  Dict keys: {list(value.keys())}")
                            extracted_value = str(value)
                    # If value is a string, try to parse as JSON
                    elif isinstance(value, str):
                        try:
                            value_parsed = json.loads(value)
                            if isinstance(value_parsed, dict):
                                print(f"  Parsed JSON dict: {value_parsed}")
                                if "number" in value_parsed:
                                    extracted_value = str(value_parsed["number"])
                                    print(f"  ✓ Value from JSON['number']: '{extracted_value}'")
                                elif "text" in value_parsed:
                                    extracted_value = str(value_parsed["text"])
                                    print(f"  ✓ Value from JSON['text']: '{extracted_value}'")
                                elif "value" in value_parsed:
                                    extracted_value = str(value_parsed["value"])
                                    print(f"  ✓ Value from JSON['value']: '{extracted_value}'")
                                else:
                                    print(f"  JSON keys: {list(value_parsed.keys())}")
                                    extracted_value = str(value_parsed)
                            else:
                                extracted_value = str(value_parsed)
                                print(f"  ✓ Value from parsed JSON: '{extracted_value}'")
                        except json.JSONDecodeError:
                            # Not JSON, just a string
                            if value.strip() and value.strip() != "None":
                                extracted_value = value.strip()
                                print(f"  ✓ Value from string: '{extracted_value}'")
                    else:
                        extracted_value = str(value)
                        print(f"  ✓ Value from other type: '{extracted_value}'")
                except Exception as e:
                    print(f"  ✗ Error extracting value: {e}")
            
            if extracted_value:
                # Clean up monetary values
                if "$" in extracted_value or "," in extracted_value or "." in extracted_value:
                    print(f"  ⭐ FINAL EXTRACTED VALUE: '{extracted_value}' (looks like monetary value!)")
                else:
                    print(f"  ⭐ FINAL EXTRACTED VALUE: '{extracted_value}'")
            else:
                print(f"  ✗ NO VALUE FOUND")
    
    print("\n" + "="*100)
    print("ALL COLUMNS (SUMMARY):")
    print("="*100)
    
    for idx, col_val in enumerate(sorted_columns, 1):
        col_id = col_val.get("id", "")
        col_type = col_val.get("type", "")
        text = str(col_val.get("text", "")).strip()
        value = col_val.get("value", "")
        
        # Truncate long text values for display
        text_display = text[:60] + "..." if len(text) > 60 else text
        value_display = str(value)[:60] + "..." if len(str(value)) > 60 else str(value)
        
        print(f"{idx:3d}. {col_id:25s} | Type: {col_type:15s} | Text: {text_display:30s}")
        # Show additional_info for mirror columns
        if col_type == "mirror":
            add_info = col_val.get("additional_info")
            if add_info:
                try:
                    parsed = json.loads(add_info) if isinstance(add_info, str) else add_info
                except Exception:
                    parsed = add_info
                print(f"     {' ' * 25}| additional_info: {str(parsed)[:80]}")
        
        # Special handling for formula columns
        if col_type == "formula":
            spaces = " " * 25
            print(f"     {spaces} | *** FORMULA COLUMN ***")

if __name__ == "__main__":
    find_all_columns()
