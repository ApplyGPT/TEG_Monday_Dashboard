"""
Client Contract Search
A read-only search tool to find all contract/payment details for a client across all boards
"""
import streamlit as st
import pandas as pd
from database_utils import get_board_data_as_items

# Page configuration
st.set_page_config(
    page_title="Search Client Contracts",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
<style>
    /* Hide QuickBooks and SignNow pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"] {
        display: none !important;
    }
    
    h1 {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
        margin-top: 1rem;
    }
    
    .search-container {
        background-color: #f8f9fa;
        padding: 1.5rem;
        border-radius: 8px;
        margin-bottom: 2rem;
    }
    
    .results-header {
        font-size: 1.2rem;
        font-weight: bold;
        color: #2c3e50;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Board configuration with display names
BOARDS_CONFIG = {
    'sales_board': 'Sales',
    'new_leads_board': 'New Leads',
    'discovery_call_board': 'Discovery Call',
    'design_review_board': 'Design Review',
    'ads_board': 'Ads'
}

def extract_column_value(item, column_id):
    """Extract a column value from an item by column ID"""
    for col_val in item.get("column_values", []):
        if col_val.get("id") == column_id:
            return col_val.get("text", "") or col_val.get("value", "")
    return ""

def extract_amount_paid_or_contract_value(item):
    """Extract 'Amount Paid or Contract Value' from contract_amt + numbers3 or formula columns"""
    # Try formula columns first (they may contain the calculated value)
    formula_columns = [
        "formula_mktj2qh2",
        "formula_mktk2rgx",
        "formula_mktks5te",
        "formula_mktknqy9",
        "formula_mktkwnyh",
        "formula_mktq5ahq",
        "formula_mktt5nty",
        "formula_mkv0r139"
    ]
    
    for col_id in formula_columns:
        val = extract_column_value(item, col_id)
        if val and val.strip():
            return val
    
    # If no formula value, calculate: contract_amt + numbers3
    contract_amt_str = extract_column_value(item, "contract_amt")
    numbers3_str = extract_column_value(item, "numbers3")
    
    try:
        contract_amt = float(contract_amt_str) if contract_amt_str and contract_amt_str.strip() else 0
        numbers3 = float(numbers3_str) if numbers3_str and numbers3_str.strip() else 0
        
        total = contract_amt + numbers3
        if total > 0:
            return str(int(total)) if total == int(total) else str(total)
    except (ValueError, TypeError):
        pass
    
    # Fallback to contract_amt alone if numbers3 not available
    if contract_amt_str and contract_amt_str.strip():
        return contract_amt_str
    
    return ""

def search_client_contracts(search_term):
    """Search for client contracts across all boards using database only"""
    if not search_term or not search_term.strip():
        return []
    
    search_term_lower = search_term.lower().strip()
    results = []
    
    # Search across all boards
    for board_table, board_display_name in BOARDS_CONFIG.items():
        try:
            items = get_board_data_as_items(board_table)
            
            for item in items:
                item_name = item.get("name", "")
                item_name_lower = item_name.lower()
                
                # Simple contains search (case-insensitive)
                if search_term_lower in item_name_lower:
                    # Extract 'Amount Paid or Contract Value'
                    amount_value = extract_amount_paid_or_contract_value(item)
                    
                    results.append({
                        'item_name': item_name,
                        'board': board_display_name,
                        'amount_paid_or_contract_value': amount_value if amount_value else "N/A"
                    })
        except Exception as e:
            # Continue searching other boards even if one fails
            continue
    
    return results

def main():
    """Main search page"""
    st.markdown('<div><h1>🔍 Search Client Contracts</h1></div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Search container
    search_term = st.text_input(
        "Enter client name:",
        placeholder="Type client name to search...",
        key="client_search",
        help="Search will find all records containing the client name across all boards"
    )
    
    # Perform search
    if search_term:
        with st.spinner("Searching across all boards..."):
            results = search_client_contracts(search_term)
        
        if results:
            st.markdown(f'<div class="results-header">Found {len(results)} matching record(s)</div>', unsafe_allow_html=True)
            
            # Convert to DataFrame for better display
            df = pd.DataFrame(results)
            
            # Display as table
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "item_name": st.column_config.TextColumn(
                        "Item Name",
                        width="medium"
                    ),
                    "board": st.column_config.TextColumn(
                        "Board",
                        width="small"
                    ),
                    "amount_paid_or_contract_value": st.column_config.TextColumn(
                        "Amount Paid or Contract Value",
                        width="medium"
                    )
                }
            )
        else:
            st.info(f"No records found matching '{search_term}'. Try a different search term.")

if __name__ == "__main__":
    main()

