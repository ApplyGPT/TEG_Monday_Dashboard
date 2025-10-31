"""
Test script to verify the fix for "Amount Paid or Contract Value" extraction
"""
import sys
import os
import json
import pandas as pd

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from database_utils import get_sales_data

def test_format_sales_data():
    """Test the format_sales_data logic"""
    
    # Get sales data
    sales_data = get_sales_data()
    
    # Simulate the format_sales_data logic from ads_dashboard.py
    if not sales_data or "data" not in sales_data or "boards" not in sales_data["data"]:
        print("No sales data found")
        return
    
    boards = sales_data["data"]["boards"]
    if not boards or not boards[0].get("items_page"):
        print("No items page found")
        return
    
    items = boards[0]["items_page"]["items"]
    
    # Find "Estevao Cavalcante (TEST)" items
    test_items = [item for item in items if "Estevao Cavalcante (TEST)" in item.get("name", "")]
    
    if not test_items:
        print("No items found matching 'Estevao Cavalcante (TEST)'")
        return
    
    print(f"Found {len(test_items)} item(s)\n")
    
    for item in test_items:
        print(f"Item: {item.get('name', '')}")
        
        # Parse column_values
        column_values = item.get("column_values", [])
        if isinstance(column_values, str):
            try:
                column_values = json.loads(column_values)
            except:
                try:
                    import ast
                    column_values = ast.literal_eval(column_values)
                except:
                    column_values = []
        
        # Apply the NEW logic: prioritize formula, then sum, then fallback
        formula_value = ""
        contract_amt_value = ""
        numbers3_value = ""
        
        for col_val in column_values:
            col_id = col_val.get("id", "")
            text = (col_val.get("text") or "").strip()
            
            # FIRST PRIORITY: Check formula columns
            if not formula_value:
                if col_id in ["formula_mktj2qh2", "formula_mktk2rgx", "formula_mktks5te", 
                             "formula_mktknqy9", "formula_mktkwnyh", "formula_mktq5ahq",
                             "formula_mktt5nty", "formula_mkv0r139"]:
                    if text:
                        formula_value = text
                        print(f"  Formula value found: {text}")
            
            # Collect contract_amt and numbers3
            if col_id == "contract_amt" and text:
                contract_amt_value = text
                print(f"  contract_amt: {text}")
            elif col_id == "numbers3" and text:
                numbers3_value = text
                print(f"  numbers3: {text}")
        
        # Determine final value
        if formula_value:
            final_value = formula_value
            source = "formula"
        elif contract_amt_value and numbers3_value:
            try:
                contract_amt_num = float(str(contract_amt_value).replace('$', '').replace(',', '').strip())
                numbers3_num = float(str(numbers3_value).replace('$', '').replace(',', '').strip())
                final_value = str(contract_amt_num + numbers3_num)
                source = "sum (contract_amt + numbers3)"
            except:
                final_value = contract_amt_value or numbers3_value or ""
                source = "fallback"
        elif contract_amt_value:
            final_value = contract_amt_value
            source = "contract_amt"
        elif numbers3_value:
            final_value = numbers3_value
            source = "numbers3"
        else:
            final_value = ""
            source = "none"
        
        print(f"\n  Final Value: {final_value} (source: {source})")
        print(f"  Expected for 'Estevao Cavalcante (TEST)': 4000")
        if item.get("name", "") == "Estevao Cavalcante (TEST)":
            try:
                final_num = float(str(final_value).replace('$', '').replace(',', '').strip())
                if final_num == 4000:
                    print(f"  ✅ CORRECT!")
                else:
                    print(f"  ❌ WRONG - got {final_num}, expected 4000")
            except:
                print(f"  ⚠️  Could not parse as number")
        print("\n" + "="*80 + "\n")

if __name__ == "__main__":
    test_format_sales_data()
