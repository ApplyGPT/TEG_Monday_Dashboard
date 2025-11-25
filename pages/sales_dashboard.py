import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import sys

# Add parent directory to path to import database_utils
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from database_utils import get_sales_data, check_database_exists, get_new_leads_data, get_discovery_call_data, get_design_review_data

# Page configuration
st.set_page_config(
    page_title="Monday.com Sales Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to hide tool pages from sidebar
st.markdown("""
<style>
    /* Hide tool pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="tools"],
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"],
    [data-testid="stSidebarNav"] a[href*="workbook_creator"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_all_leads_for_sales_chart():
    """Get all leads data from all boards for sales chart analysis using correct date fields"""
    import json
    
    all_leads = []
    
    # Get data from all boards using database functions
    boards_data = {
        'New Leads': get_new_leads_data(),
        'Discovery Call': get_discovery_call_data(), 
        'Design Review': get_design_review_data(),
        'Sales': get_sales_data().get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
    }
    
    # Board-specific date column IDs for each stage (based on Monday.com filter results)
    date_columns = {
        'New Leads': 'date_mkwgr4gg',           # Date Created (181 vs 182 expected)
        'Discovery Call': 'date_mktbrpz6',      # Discovery Call Date (144 vs 145 expected)
        'Design Review': 'date3',                # Design Review Date (221 vs 221 expected - PERFECT MATCH!)
        'Sales': 'date_mktqx5me'                # Deck Call Date (390 vs 390 expected - PERFECT MATCH!)
    }
    
    # Process each board's data
    for board_name, items in boards_data.items():
        target_date_column = date_columns.get(board_name)
        
        for item in items:
            # Parse column_values if it's a string
            column_values = item.get("column_values", [])
            if isinstance(column_values, str):
                try:
                    column_values = json.loads(column_values)
                except:
                    column_values = []
            # Look for the specific date column for this board/stage
            stage_date = None
            lead_status = ""
            assigned_person = ""
            for col_val in column_values:
                col_id = col_val.get("id", "")
                col_type = col_val.get("type", "")
                text = (col_val.get("text") or "").strip()
                
                if col_type == "date" and text and col_id == target_date_column and not stage_date:
                    stage_date = text
                
                if board_name == 'Sales':
                    if col_id == "color_mknxd1j2":
                        lead_status = text
                    elif col_id == "color_mkvewcwe":
                        assigned_person = text
            
            # Map board names to chart categories
            if board_name == 'New Leads':
                category = 'New Leads'
            elif board_name == 'Discovery Call':
                category = 'Discovery Call'
            elif board_name == 'Design Review':
                category = 'Design Review Call'
            elif board_name == 'Sales':
                category = 'Deck Call'
            else:
                category = 'Other'
            
            all_leads.append({
                'name': item.get('name', ''),
                'board': board_name,
                'category': category,
                'stage_date': stage_date,  # Changed from 'date_created' to 'stage_date'
                'status': lead_status if board_name == 'Sales' else "",
                'assigned_person': assigned_person if board_name == 'Sales' else ""
            })
    
    return all_leads

# Helper function to format numbers with K format
def format_currency(value):
    """Format currency values with K for thousands"""
    if value >= 1000000:
        return f"${value/1000000:.1f}M"
    elif value >= 1000:
        return f"${value/1000:.1f}K"
    else:
        return f"${value:.0f}"

# Helper function to format numbers with K format (one decimal place)
def format_currency_one_decimal(value):
    """Format currency values with K for thousands - one decimal place"""
    if value >= 1000000:
        return f"${value/1000000:.1f}M"
    elif value >= 1000:
        return f"${value/1000:.1f}K"
    else:
        return f"${value:.1f}"

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
            
        if 'sales_board_id' not in monday_config:
            st.error("Sales board ID not found in secrets.toml. Please add your sales board ID.")
            st.stop()
        
        return {
            'api_token': monday_config['api_token'],
            'sales_board_id': int(monday_config['sales_board_id'])
        }
    except Exception as e:
        st.error(f"Error reading secrets: {str(e)}")
        st.stop()

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_sales_data_from_db():
    """Get sales data from SQLite database"""
    return get_sales_data()

def process_sales_data(data):
    """Convert Monday.com sales data to pandas DataFrame and process it"""
    if not data or "data" not in data or "boards" not in data["data"]:
        return pd.DataFrame()
    
    boards = data["data"]["boards"]
    if not boards:
        return pd.DataFrame()
    
    # Collect items from the board
    all_items = []
    for board in boards:
        if board.get("items_page") and board["items_page"].get("items"):
            all_items.extend(board["items_page"]["items"])
    
    if not all_items:
        return pd.DataFrame()
    
    items = all_items
    
    # Convert to DataFrame
    records = []
    
    
    for item in items:
        record = {
            "Item": item.get("name", ""),
            "Close Date": "",
            "Lead Status": "",
            "Amount Paid or Contract Value": "",
            "Contract Amount": "",
            "Numbers3": "",
            "Assigned Person": "",
            "Client Type": "",
            "Type of Revenue": ""
        }
        
        for col_val in item.get("column_values", []):
            col_id = col_val.get("id", "")
            text = col_val.get("text", "")
            
            if col_id == "color_mknxd1j2":  # Lead Status
                record["Lead Status"] = text if text else ""
            elif col_id == "contract_amt":  # Contract Amount
                record["Contract Amount"] = text if text else ""
            elif col_id == "numbers3":  # Numbers3 column
                record["Numbers3"] = text if text else ""
            elif col_id == "color_mkvewcwe":  # Assigned Person dropdown field (CORRECT ONE)
                record["Assigned Person"] = text if text else ""
            elif col_id == "status_14__1":  # Client Type (CORRECT COLUMN)
                record["Client Type"] = text if text else ""
            elif col_id == "color_mkwp98ks":  # Type of Revenue (CORRECT COLUMN ID)
                record["Type of Revenue"] = text if text else ""
            elif col_id == "date_mktq7npm":  # CORRECT Close Date (Date MK7)
                record["Close Date"] = text if text else ""
            # Try to find the "Amount Paid or Contract Value" formula column
            elif col_id == "formula_mktj2qh2":  # Try first formula column
                record["Amount Paid or Contract Value"] = text if text else ""
            elif col_id == "formula_mktk2rgx":  # Try second formula column
                record["Amount Paid or Contract Value"] = text if text else ""
            elif col_id == "formula_mktks5te":  # Try third formula column
                record["Amount Paid or Contract Value"] = text if text else ""
            elif col_id == "formula_mktknqy9":  # Try fourth formula column
                record["Amount Paid or Contract Value"] = text if text else ""
            elif col_id == "formula_mktkwnyh":  # Try fifth formula column
                record["Amount Paid or Contract Value"] = text if text else ""
            elif col_id == "formula_mktq5ahq":  # Try sixth formula column
                record["Amount Paid or Contract Value"] = text if text else ""
            elif col_id == "formula_mktt5nty":  # Try seventh formula column
                record["Amount Paid or Contract Value"] = text if text else ""
            elif col_id == "formula_mkv0r139":  # Try eighth formula column
                record["Amount Paid or Contract Value"] = text if text else ""
        
        records.append(record)
    
    
    df = pd.DataFrame(records)
    
    # Process monetary values for all columns
    df['Amount Paid or Contract Value'] = pd.to_numeric(df['Amount Paid or Contract Value'], errors='coerce')
    df['Contract Amount'] = pd.to_numeric(df['Contract Amount'].str.replace('$', '').str.replace(',', ''), errors='coerce')
    df['Numbers3'] = pd.to_numeric(df['Numbers3'], errors='coerce')
    
    # Use the CORRECT combination: Contract Amount OR Numbers3 (Amount Paid or Contract Value)
    # This matches the exact $2,013,315 value discovered
    df['Contract Amount'] = pd.to_numeric(df['Contract Amount'], errors='coerce')
    df['Numbers3'] = pd.to_numeric(df['Numbers3'], errors='coerce')
            
    # Use the best available value: Contract Amount if available, otherwise Numbers3
    df['Total Value'] = df['Contract Amount'].fillna(0)
    # If Contract Amount is 0, use Numbers3
    df.loc[df['Total Value'] == 0, 'Total Value'] = df.loc[df['Total Value'] == 0, 'Numbers3']
    
    # Apply the CORRECT filters as discovered:
    # 1. Lead Status = "Closed" OR "Win"
    closed_status_mask = (df['Lead Status'] == 'Closed') | (df['Lead Status'] == 'Win')
            
    # 2. Contract Amount >= 0 OR Numbers3 >= 0 OR both are null (Amount Paid or Contract Value)
    contract_amount_mask = df['Contract Amount'] >= 0
    numbers3_mask = df['Numbers3'] >= 0
    both_null_mask = df['Contract Amount'].isna() & df['Numbers3'].isna()
            
    # Combine filters: (Closed OR Win) AND (Contract Amount >= 0 OR Numbers3 >= 0 OR both are null)
    final_mask = closed_status_mask & (contract_amount_mask | numbers3_mask | both_null_mask)
    df_filtered = df[final_mask].copy()
    
    # Create a separate filter for 2025 data (for KPIs and comparison)
    df['Close Date'] = df['Close Date'].astype(str)
    year_2025_mask = df['Close Date'].str.contains('2025', na=False)
    df_2025 = df_filtered[year_2025_mask].copy()
    
    # Extract year and month for filtered data
    df_filtered['Close Date'] = pd.to_datetime(df_filtered['Close Date'], errors='coerce')
    df_filtered['Year'] = df_filtered['Close Date'].dt.year
    df_filtered['Month'] = df_filtered['Close Date'].dt.month
    df_filtered['Month_Name'] = df_filtered['Close Date'].dt.strftime('%B')
    
    # Show breakdown for 2025
    with_2025 = df_2025[df_2025['Close Date'].astype(str).str.contains('2025', na=False)]
    without_date = df_2025[df_2025['Close Date'].astype(str) == '']
    
    # Show expected values for comparison (2025 specific)
    expected_sum = 2013315
    expected_count = 178
    actual_sum = df_2025['Total Value'].sum()
    actual_count = len(df_2025)
    
    return df_filtered, df_2025

def main():
    """Main application function"""
    # Header
    st.title("📈 Sales Dashboard")
    
    # Check if database exists and has data
    db_exists, db_message = check_database_exists()
    
    if not db_exists:
        st.error(f"❌ Database not ready: {db_message}")
        st.info("💡 Please go to the 'Database Refresh' page to initialize the database with Monday.com data.")
        return
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Settings")
        st.info(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Refresh button
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
    
    # Load and process data from database
    with st.spinner("Loading sales data from database..."):
        data = get_sales_data_from_db()
        
        df_filtered, df_2025 = process_sales_data(data)
    
    if df_filtered.empty:
        st.warning("No closed sales records found. Please check your data and filters.")
        return

    # Current year and month
    current_year = datetime.now().year
    current_month = datetime.now().month
    
    # Filter data for current year and month
    df_current_year = df_filtered[df_filtered['Year'] == current_year]
    df_current_month = df_filtered[(df_filtered['Year'] == current_year) & (df_filtered['Month'] == current_month)]
    
    # Calculate KPIs based on 2025 data (the specific requirement)
    if not df_2025.empty:
        # YTD calculation for 2025 - use 2025 filtered data
        sales_ytd = round(df_2025['Total Value'].sum(), 2)
        
        # MTD calculation for 2025 - current month
        current_month = datetime.now().month
        df_2025['Close Date'] = pd.to_datetime(df_2025['Close Date'], errors='coerce')
        mtd_mask = df_2025['Close Date'].dt.month == current_month
        sales_mtd = df_2025[mtd_mask]['Total Value'].sum()
        
        # Average contract amount for 2025 - calculate from all 2025 records (including NaN/zero values)
        avg_contract = df_2025['Total Value'].sum() / len(df_2025) if len(df_2025) > 0 else 0
    
    else:
        sales_ytd = 0
        sales_mtd = 0
        avg_contract = 0
    
    # Display KPIs in columns at the top with larger numbers and K format
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label="Sales Year-to-Date (YTD)",
            value=f"${sales_ytd:,.2f}",
            delta=None
        )
    
    with col2:
        st.metric(
            label="Sales Month-to-Date (MTD)",
            value=f"${sales_mtd:,.2f}",
            delta=None
        )
    
    with col3:
        st.metric(
            label="Average Contract Amount",
            value=f"${avg_contract:,.2f}",
            delta=None
        )
    
    current_month_name = datetime.now().strftime('%B')
    
    st.subheader(f"Current Month Sales Breakdown by Salesman ({current_month_name} {current_year})")
    df_current_month = df_current_month.copy()
    
    if not df_current_month.empty:
        df_current_month['Contract Amount'] = pd.to_numeric(df_current_month['Contract Amount'], errors='coerce').fillna(0)
        df_current_month['Numbers3'] = pd.to_numeric(df_current_month['Numbers3'], errors='coerce').fillna(0)
        df_current_month['Assigned Person'] = df_current_month['Assigned Person'].fillna('').apply(
            lambda x: x.strip() if isinstance(x, str) else ''
        )
        df_current_month['Assigned Person'] = df_current_month['Assigned Person'].replace('', 'Unassigned')
        
        monthly_salesman = (
            df_current_month
            .groupby('Assigned Person')[['Contract Amount', 'Numbers3']]
            .sum()
            .reset_index()
        )
        monthly_salesman.rename(columns={'Numbers3': 'Amount Paid'}, inplace=True)
        monthly_salesman['Salesman Name'] = monthly_salesman['Assigned Person']
        monthly_salesman['Total'] = monthly_salesman['Contract Amount'] + monthly_salesman['Amount Paid']
        monthly_salesman = monthly_salesman.sort_values('Total', ascending=False)
        
        fig_current_month = go.Figure()
        fig_current_month.add_trace(go.Bar(
            name='Contract Amount',
            x=monthly_salesman['Salesman Name'],
            y=monthly_salesman['Contract Amount'],
            marker_color='#ff7f0e',
            textposition='none'
        ))
        fig_current_month.add_trace(go.Bar(
            name='Amount Paid',
            x=monthly_salesman['Salesman Name'],
            y=monthly_salesman['Amount Paid'],
            marker_color='#1f77b4',
            textposition='none'
        ))
        fig_current_month.add_trace(go.Scatter(
            x=monthly_salesman['Salesman Name'],
            y=monthly_salesman['Total'],
            mode='text',
            text=[format_currency(val) for val in monthly_salesman['Total']],
            textposition='top center',
            textfont=dict(size=14, color='black'),
            showlegend=False,
            hoverinfo='skip'
        ))
        fig_current_month.update_layout(
            barmode='stack',
            xaxis_title='Salesman',
            yaxis_title='Revenue ($)',
            height=500,
            font=dict(size=14),
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.2,
                xanchor="center",
                x=0.5
            )
        )
        st.plotly_chart(fig_current_month, use_container_width=True)
    else:
        st.info(f"No sales data available for {current_month_name} {current_year}.")
    
    # 1. Sales by Month (2025) - Contract Amount vs Amount Paid
    st.subheader(f"Sales by Month (2025)")
    if not df_2025.empty:
        # Extract year and month for 2025 data
        df_2025['Close Date'] = pd.to_datetime(df_2025['Close Date'], errors='coerce')
        df_2025['Year'] = df_2025['Close Date'].dt.year
        df_2025['Month'] = df_2025['Close Date'].dt.month
        df_2025['Month_Name'] = df_2025['Close Date'].dt.strftime('%B')
        
        # Separate contract amounts and amounts paid
        df_2025['Contract Amount'] = pd.to_numeric(df_2025['Contract Amount'], errors='coerce').fillna(0)
        df_2025['Numbers3'] = pd.to_numeric(df_2025['Numbers3'], errors='coerce').fillna(0)
        
        # Group by month and sum both contract amounts and amounts paid
        monthly_contract = df_2025.groupby(['Month', 'Month_Name'])['Contract Amount'].sum().reset_index()
        monthly_paid = df_2025.groupby(['Month', 'Month_Name'])['Numbers3'].sum().reset_index()
        
        # Merge the data
        monthly_sales = monthly_contract.merge(monthly_paid, on=['Month', 'Month_Name'], how='outer')
        monthly_sales = monthly_sales.fillna(0)
        monthly_sales = monthly_sales.sort_values('Month')
        
        # Calculate total for each month (contract amount + amount paid)
        monthly_sales['Total'] = monthly_sales['Contract Amount'] + monthly_sales['Numbers3']
        
        # Create stacked bar chart with two colors
        fig_monthly = go.Figure()
        
        # Add Amount Paid bars
        fig_monthly.add_trace(go.Bar(
            name='Amount Paid',
            x=monthly_sales['Month_Name'],
            y=monthly_sales['Contract Amount'],
            marker_color='#1f77b4',  # Blue color for amount paid
            textposition='inside',  # No text for individual segments
            showlegend=True
        ))
        
        # Add Contract Amount bars
        fig_monthly.add_trace(go.Bar(
            name='Contract Amount',
            x=monthly_sales['Month_Name'],
            y=monthly_sales['Numbers3'],
            marker_color='#ff7f0e',  # Orange color for contract amount
            textposition='inside',  # No text for individual segments
            showlegend=True
        ))
        
        # Add total values on top of the stacked bars
        fig_monthly.add_trace(go.Scatter(
            x=monthly_sales['Month_Name'],
            y=monthly_sales['Total'],
            mode='text',
            text=[format_currency_one_decimal(val) for val in monthly_sales['Total']],
            textposition='top center',
            textfont=dict(size=14, color='black'),
            showlegend=False,
            hoverinfo='skip'
        )) 
        
        fig_monthly.update_layout(
            barmode='stack',
            height=500,
            xaxis_title='Month',
            yaxis_title='Revenue ($)',
            font=dict(size=14),
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.2,
                xanchor="center",
                x=0.5
            )
        )
        
        # Calculate average contract amount per month for display
        monthly_avg = df_2025.groupby(['Month', 'Month_Name'])['Contract Amount'].mean().reset_index()
        monthly_avg = monthly_avg.sort_values('Month')
        
        # Add average contract amount text below each bar
        fig_monthly.add_trace(go.Scatter(
            x=monthly_avg['Month_Name'],
            y=[0] * len(monthly_avg),  # Position at bottom
            mode='text',
            text=[f"Avg. C.A. = ${val:,.0f}" for val in monthly_avg['Contract Amount']],
            textposition='bottom center',
            textfont=dict(size=14, color='black'),
            showlegend=False,
            hoverinfo='skip'
        ))
        
        st.plotly_chart(fig_monthly, use_container_width=True)
    else:
        st.info("No sales data available for 2025.")
    
    # 2. Sales by Year (All Years)
    st.subheader("Sales by Year")
    # Extract year and month for all filtered data
    df_filtered['Close Date'] = pd.to_datetime(df_filtered['Close Date'], errors='coerce')
    df_filtered['Year'] = df_filtered['Close Date'].dt.year
    df_filtered['Month'] = df_filtered['Close Date'].dt.month
    df_filtered['Month_Name'] = df_filtered['Close Date'].dt.strftime('%B')
    
    yearly_sales = df_filtered.groupby('Year')['Total Value'].sum().reset_index()
    yearly_sales = yearly_sales.sort_values('Year')
    
    fig_yearly = px.bar(
        yearly_sales,
        x='Year',
        y='Total Value',
        labels={'Total Value': 'Revenue ($)', 'Year': 'Year'}
    )
    
    # Add numerical amounts above each bar
    fig_yearly.update_traces(
        texttemplate='<b>$%{y:,.2f}</b>',  # Bold text
        textposition='outside',
        textfont=dict(size=16, color='black')  # Larger text
    )
    
    fig_yearly.update_layout(
        height=500, 
        showlegend=False,
        xaxis_title='Year',
        yaxis_title='Revenue ($)',
        font=dict(size=14)  # Larger font for all text
    )
    st.plotly_chart(fig_yearly, use_container_width=True)
    
    # 3. Comparison of Revenue by Year by Month (All Years)
    st.subheader("Comparison of Revenue by Year by Month")
    
    # Create pivot table for grouped bar chart using all filtered data
    monthly_yearly = df_filtered.groupby(['Year', 'Month', 'Month_Name'])['Total Value'].sum().reset_index()
    monthly_yearly = monthly_yearly.sort_values(['Year', 'Month'])
    
    # Create a proper month order for x-axis
    month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                   'July', 'August', 'September', 'October', 'November', 'December']
    
    # Filter to only include months that exist in the data
    available_months = monthly_yearly['Month_Name'].unique()
    month_order_filtered = [month for month in month_order if month in available_months]
    
    # Update the figure to use the proper month order
    fig_grouped = go.Figure()
    
    years = sorted([year for year in monthly_yearly['Year'].unique() if pd.notna(year)])
    # Use stronger colors
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F']
    
    for i, year in enumerate(years):
        year_data = monthly_yearly[monthly_yearly['Year'] == year]
        fig_grouped.add_trace(go.Bar(
            name=str(int(year)),  # Convert to int to remove decimal
            x=year_data['Month_Name'],
            y=year_data['Total Value'],
            marker_color=colors[i],
            text=[format_currency(val) for val in year_data['Total Value']],  # K format
            textposition='outside',
            textfont=dict(size=14, color='black')  # Larger text
        ))
    
    fig_grouped.update_layout(
        barmode='group',
        xaxis_title='Month',
        yaxis_title='Revenue ($)',
        height=500,
        bargap=0.15,
        bargroupgap=0.0,
        font=dict(size=14),  # Larger font for all text
        xaxis=dict(
            categoryorder='array',
            categoryarray=month_order_filtered
        ),
        legend=dict(
            orientation="h",   # horizontal
            yanchor="top",
            y=-0.2,            # below the chart
            xanchor="center",
            x=0.5
        )
    )
    
    st.plotly_chart(fig_grouped, use_container_width=True)
    
    # 4. Comparison of Revenue by Salesman by Month
    st.subheader("Comparison of Revenue by Salesman by Month")
    
    # Year selector for salesman chart - default to 2025
    available_years = sorted([int(year) for year in df_filtered['Year'].unique() if pd.notna(year)])
    default_year_index = available_years.index(2025) if 2025 in available_years else 0
    selected_year_salesman = st.selectbox("Select Year for Salesman Analysis:", available_years, index=default_year_index, key="salesman_year")
    
    df_salesman_year = df_filtered[df_filtered['Year'] == selected_year_salesman]
    
    if not df_salesman_year.empty:
        salesman_monthly = df_salesman_year.groupby(['Month', 'Month_Name', 'Assigned Person'])['Total Value'].sum().reset_index()
        salesman_monthly = salesman_monthly.sort_values('Month')
        
        # Create proper month order for x-axis
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December']
        
        # Filter to only include months that exist in the data
        available_months = salesman_monthly['Month_Name'].unique()
        month_order_filtered = [month for month in month_order if month in available_months]
        
        # Create grouped bar chart by salesman with hardcoded colors
        fig_salesman = go.Figure()
        
        salesmen = sorted(salesman_monthly['Assigned Person'].unique())
        
        # Hardcoded colors for specific salesmen
        salesman_colors = {
            'Jennifer Evans': '#df2f4a',            # Red
            'Gabriela Tamayo': '#a358df',           # Green
            'Anthony Alba': '#579bfc',              # Blue
            'Unassigned': '#96CEB4',                # Green
            'Heather Castagno': '#ffcb00'           # Yellow
        }
        
        # Use strong colors for any other salesmen
        all_colors = ['#DDA0DD', '#98D8C8', '#F7DC6F', '#FF8A80', '#26A69A', '#42A5F5', '#66BB6A', '#FFCA28']
        
        for i, salesman in enumerate(salesmen):
            salesman_data = salesman_monthly[salesman_monthly['Assigned Person'] == salesman]
            
            # Handle empty salesmen
            salesman_name = salesman if salesman and salesman.strip() else 'Unassigned'
            
            # Get color - use hardcoded if available, otherwise cycle through colors
            if salesman_name in salesman_colors:
                color = salesman_colors[salesman_name]
            else:
                color = all_colors[i % len(all_colors)]
            
            fig_salesman.add_trace(go.Bar(
                name=salesman_name,
                x=salesman_data['Month_Name'],
                y=salesman_data['Total Value'],
                marker_color=color,
                text=[format_currency(val) for val in salesman_data['Total Value']],  # K format
                textposition='outside',
                textfont=dict(size=14, color='black')  # Larger text
            ))
        
        fig_salesman.update_layout(
            barmode='group',
            xaxis_title='Month',
            yaxis_title='Revenue ($)',
            height=500,
            bargap=0.15,
            bargroupgap=0.0,
            showlegend=True,
            font=dict(size=14),  # Larger font for all text
            xaxis=dict(
                categoryorder='array',
                categoryarray=month_order_filtered
            ),
            legend=dict(
                orientation="h",   # horizontal
                yanchor="top",
                y=-0.2,            # below the chart
                xanchor="center",
                x=0.5
            )
        )
        st.plotly_chart(fig_salesman, use_container_width=True)
    else:
        st.info(f"No sales data available for {selected_year_salesman}.")
    
    # 5. Comparison of Revenue by Type of Revenue by Month
    st.subheader("Comparison of Revenue by Type of Revenue by Month")
       
    # Year selector for type of revenue chart - default to 2025
    available_years_category = sorted([int(year) for year in df_filtered['Year'].unique() if pd.notna(year)])
    default_year_index_category = available_years_category.index(2025) if 2025 in available_years_category else 0
    selected_year_category = st.selectbox("Select Year for Type of Revenue Analysis:", available_years_category, index=default_year_index_category, key="category_year")
    
    df_category_year = df_filtered[df_filtered['Year'] == selected_year_category]
    
    if not df_category_year.empty:
        category_monthly = df_category_year.groupby(['Month', 'Month_Name', 'Type of Revenue'])['Total Value'].sum().reset_index()
        category_monthly = category_monthly.sort_values('Month')
        
        # Create proper month order for x-axis
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December']
        
        # Filter to only include months that exist in the data
        available_months = category_monthly['Month_Name'].unique()
        month_order_filtered = [month for month in month_order if month in available_months]
        
        # Create grouped bar chart by type of revenue
        fig_category = go.Figure()
        
        categories = sorted(category_monthly['Type of Revenue'].unique())
        
        # Hardcoded colors for specific types of revenue
        category_colors = {
            'New Client': '#00c875',     # Green
            'Add-On': '#ffcb00',         # Yellow
            'Production': '#df2f4a',     # Red
            'OTHER': '#c4c4c4'                # Gray for empty values
        }
        
        # Use strong colors for any other categories
        all_colors = ['#96CEB4', '#FFEAA7', '#98D8C8', '#F7DC6F', '#FF8A80', '#26A69A', '#42A5F5', '#66BB6A', '#FFCA28']
        
        for i, category in enumerate(categories):
            category_data = category_monthly[category_monthly['Type of Revenue'] == category]
            
            # Handle empty categories - ensure we have a proper name for the legend
            category_name = category if category and category.strip() else 'Uncategorized'
            
            # Get color - use hardcoded if available, otherwise cycle through colors
            if category_name in category_colors:
                color = category_colors[category_name]
            else:
                color = all_colors[i % len(all_colors)]
            
            fig_category.add_trace(go.Bar(
                name=category_name,
                x=category_data['Month_Name'],
                y=category_data['Total Value'],
                marker_color=color,
                text=[format_currency(val) for val in category_data['Total Value']],  # K format
                textposition='outside',
                textfont=dict(size=14, color='black')  # Larger text
            ))
        
        fig_category.update_layout(
            barmode='group',
            xaxis_title='Month',
            yaxis_title='Revenue ($)',
            height=500,
            bargap=0.15,
            bargroupgap=0.0,
            showlegend=True,
            font=dict(size=14),  # Larger font for all text
            xaxis=dict(
                categoryorder='array',
                categoryarray=month_order_filtered
            ),
            legend=dict(
                orientation="h",   # horizontal
                yanchor="top",
                y=-0.2,            # below the chart
                xanchor="center",
                x=0.5
            )
        )
        st.plotly_chart(fig_category, use_container_width=True)
    else:
        st.info(f"No sales data available for {selected_year_category}.")
    
    # 6. Sales by Source (Revenue) - Based on UTM Data from Sales Board
    st.subheader("Sales by Source")
    
    @st.cache_data(ttl=300)  # Cache for 5 minutes
    def get_sales_revenue_by_source():
        """Get sales revenue data by source/channel from Sales board for revenue analysis"""
        import json
        
        sales_revenue_data = []
        
        # Get data from Sales board only
        sales_items = get_sales_data_from_db().get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
        
        # Sales board channel column ID - this is the UTM channel column with "Paid search", "Organic search", etc.
        sales_channel_column = 'text_mkrfer1n'
        
        # Process Sales board data
        for item in sales_items:
            # Extract channel information, revenue, and other relevant data
            channel = ""
            close_date = None
            lead_status = ""
            contract_amount = 0
            numbers3_amount = 0
            
            # Parse column_values if it's a string
            column_values = item.get("column_values", [])
            if isinstance(column_values, str):
                try:
                    column_values = json.loads(column_values)
                except:
                    column_values = []
            
            for col_val in column_values:
                col_id = col_val.get("id", "")
                text = (col_val.get("text") or "").strip()
                col_type = col_val.get("type", "")
                
                # Look for the channel column (UTM channel column)
                if col_id == sales_channel_column and text:
                    channel = text
                
                # Look for closed date - use the same column leveraged in process_sales_data
                if col_id == "date_mktq7npm" and text:
                    close_date = text
                
                # Look for lead status
                if col_id == "color_mknxd1j2":  # Lead Status
                    lead_status = text
                
                # Look for revenue amounts
                if col_id == "contract_amt":  # Contract Amount
                    try:
                        contract_amount = float(str(text).replace('$', '').replace(',', '')) if text else 0
                    except:
                        contract_amount = 0
                elif col_id == "numbers3":  # Numbers3 column
                    try:
                        numbers3_amount = float(str(text).replace('$', '').replace(',', '')) if text else 0
                    except:
                        numbers3_amount = 0
                
            # Calculate total revenue (same logic as in process_sales_data)
            total_revenue = contract_amount if contract_amount > 0 else numbers3_amount
            
            # Filter for proper UTM channels (exclude individual names)
            valid_utm_channels = [
                'paid search', 'organic search', 'direct traffic', 'referral', 
                'email marketing', 'social media', 'tradeshow', 'google', 
                'facebook', 'instagram', 'linkedin', 'youtube', 'twitter'
            ]
            
            # Only include items with valid UTM channels and closed/win status and revenue > 0
            if (channel and channel.strip() and 
                close_date and
                lead_status and lead_status.lower() in ['closed', 'win'] and 
                total_revenue > 0 and
                channel.lower() in valid_utm_channels):
                sales_revenue_data.append({
                    'name': item.get('name', ''),
                    'channel': channel,
                    'close_date': close_date,
                    'lead_status': lead_status,
                    'revenue': total_revenue
                })
        
        return sales_revenue_data
    
    # Year selector for sales by source chart - default to 2025
    available_years_source = sorted([int(year) for year in df_filtered['Year'].unique() if pd.notna(year)])
    default_year_index_source = available_years_source.index(2025) if 2025 in available_years_source else 0
    selected_year_source = st.selectbox("Select Year for Sales by Source Analysis:", available_years_source, index=default_year_index_source, key="source_year")
    
    # Get sales revenue data by source
    with st.spinner("Loading Sales by Source data..."):
        sales_revenue_data = get_sales_revenue_by_source()
    
    if sales_revenue_data:
        # Convert to DataFrame
        sales_revenue_df = pd.DataFrame(sales_revenue_data)
        
        # Parse close dates and filter for valid dates
        sales_revenue_df['close_date'] = pd.to_datetime(sales_revenue_df['close_date'], errors='coerce')
        sales_revenue_with_dates = sales_revenue_df.dropna(subset=['close_date'])
        
        # Filter for selected year based on close date
        sales_revenue_with_dates = sales_revenue_with_dates[sales_revenue_with_dates['close_date'].dt.year == selected_year_source]
        
        if not sales_revenue_with_dates.empty:
            # Add month-year column for grouping with proper formatting
            sales_revenue_with_dates['Month Year'] = sales_revenue_with_dates['close_date'].dt.strftime('%B %Y')
            sales_revenue_with_dates['Month'] = sales_revenue_with_dates['close_date'].dt.month
            sales_revenue_with_dates['Month_Name'] = sales_revenue_with_dates['close_date'].dt.strftime('%B')
            
            # Sum revenue by channel and month (instead of counting leads)
            channel_revenue = sales_revenue_with_dates.groupby(['Month', 'Month_Name', 'channel'])['revenue'].sum().reset_index()
            channel_revenue = channel_revenue.sort_values('Month')
            
            # Create proper month order for x-axis
            month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                           'July', 'August', 'September', 'October', 'November', 'December']
            
            # Filter to only include months that exist in the data
            available_months = channel_revenue['Month_Name'].unique()
            month_order_filtered = [month for month in month_order if month in available_months]
            
            # Create grouped bar chart by source/channel
            fig_source = go.Figure()
            
            channels = sorted(channel_revenue['channel'].unique())
            
            # Use strong colors for different channels [[memory:7838657]]
            channel_colors = {
                'Paid search': '#df2f4a',        # Red
                'Organic search': '#00c875',     # Green  
                'Direct': '#4ECDC4',             # Teal
                'Social media': '#ffcb00',       # Yellow
                'Referral': '#579bfc',           # Blue
                'Email': '#a358df'               # Purple
            }
            
            # Use strong colors for any other channels
            all_colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F']
            
            for i, channel in enumerate(channels):
                channel_data = channel_revenue[channel_revenue['channel'] == channel]
                
                # Handle empty channels - ensure we have a proper name for the legend
                channel_name = channel if channel and channel.strip() else 'Unknown'
                
                # Get color - use hardcoded if available, otherwise cycle through colors
                if channel_name in channel_colors:
                    color = channel_colors[channel_name]
                else:
                    color = all_colors[i % len(all_colors)]
                
                fig_source.add_trace(go.Bar(
                    name=channel_name,
                    x=channel_data['Month_Name'],
                    y=channel_data['revenue'],
                    marker_color=color,
                    text=[format_currency(val) for val in channel_data['revenue']],  # K format
                    textposition='outside',
                    textfont=dict(size=14, color='black')  # Larger text
                ))
            
            fig_source.update_layout(
                barmode='group',
                xaxis_title='Month',
                yaxis_title='Revenue ($)',
                height=500,
                bargap=0.15,
                bargroupgap=0.0,
                showlegend=True,
                font=dict(size=14),  # Larger font for all text
                xaxis=dict(
                    categoryorder='array',
                    categoryarray=month_order_filtered
                ),
                legend=dict(
                    orientation="h",   # horizontal
                    yanchor="top",
                    y=-0.2,            # below the chart
                    xanchor="center",
                    x=0.5
                )
            )
            st.plotly_chart(fig_source, use_container_width=True)
        else:
            st.info(f"No sales revenue data with valid dates found for {selected_year_source}.")
    else:
        st.info("No sales revenue data found for source analysis.")
    
    # 7. Number of Deals Closed by Source (Frequency) - Based on UTM Data from Sales Board
    st.subheader("Number of Deals Closed by Source")
    
    # Reuse the same year selector or create a new one (using same selected_year_source)
    # Get sales revenue data by source (same data, but we'll count instead of sum)
    with st.spinner("Loading Number of Deals Closed by Source data..."):
        sales_revenue_data = get_sales_revenue_by_source()
    
    if sales_revenue_data:
        # Convert to DataFrame
        sales_revenue_df = pd.DataFrame(sales_revenue_data)
        
        # Parse close dates and filter for valid dates
        sales_revenue_df['close_date'] = pd.to_datetime(sales_revenue_df['close_date'], errors='coerce')
        sales_revenue_with_dates = sales_revenue_df.dropna(subset=['close_date'])
        
        # Filter for selected year based on close date (use same year as Sales by Source)
        sales_revenue_with_dates = sales_revenue_with_dates[sales_revenue_with_dates['close_date'].dt.year == selected_year_source]
        
        if not sales_revenue_with_dates.empty:
            # Add month columns for grouping
            sales_revenue_with_dates['Month'] = sales_revenue_with_dates['close_date'].dt.month
            sales_revenue_with_dates['Month_Name'] = sales_revenue_with_dates['close_date'].dt.strftime('%B')
            
            # Count deals by channel and month (instead of summing revenue)
            channel_counts = sales_revenue_with_dates.groupby(['Month', 'Month_Name', 'channel']).size().reset_index(name='deal_count')
            channel_counts = channel_counts.sort_values('Month')
            
            # Create proper month order for x-axis
            month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                           'July', 'August', 'September', 'October', 'November', 'December']
            
            # Filter to only include months that exist in the data
            available_months = channel_counts['Month_Name'].unique()
            month_order_filtered = [month for month in month_order if month in available_months]
            
            # Create grouped bar chart by source/channel
            fig_deals_count = go.Figure()
            
            channels = sorted(channel_counts['channel'].unique())
            
            # Use strong colors for different channels (same as Sales by Source) [[memory:7838657]]
            channel_colors = {
                'Paid search': '#df2f4a',        # Red
                'Organic search': '#00c875',     # Green  
                'Direct': '#4ECDC4',             # Teal
                'Social media': '#ffcb00',       # Yellow
                'Referral': '#579bfc',           # Blue
                'Email': '#a358df'               # Purple
            }
            
            # Use strong colors for any other channels
            all_colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F']
            
            for i, channel in enumerate(channels):
                channel_data = channel_counts[channel_counts['channel'] == channel]
                
                # Handle empty channels - ensure we have a proper name for the legend
                channel_name = channel if channel and channel.strip() else 'Unknown'
                
                # Get color - use hardcoded if available, otherwise cycle through colors
                if channel_name in channel_colors:
                    color = channel_colors[channel_name]
                else:
                    color = all_colors[i % len(all_colors)]
                
                fig_deals_count.add_trace(go.Bar(
                    name=channel_name,
                    x=channel_data['Month_Name'],
                    y=channel_data['deal_count'],
                    marker_color=color,
                    text=[str(int(val)) for val in channel_data['deal_count']],  # Show count as integer
                    textposition='outside',
                    textfont=dict(size=14, color='black')  # Larger text
                ))
            
            fig_deals_count.update_layout(
                barmode='group',
                xaxis_title='Month',
                yaxis_title='Number of Deals',
                height=600,
                bargap=0.15,
                bargroupgap=0.0,
                showlegend=True,
                font=dict(size=14),  # Larger font for all text
                xaxis=dict(
                    categoryorder='array',
                    categoryarray=month_order_filtered
                ),
                legend=dict(
                    orientation="h",   # horizontal
                    yanchor="top",
                    y=-0.2,            # below the chart
                    xanchor="center",
                    x=0.5
                )
            )
            st.plotly_chart(fig_deals_count, use_container_width=True)
        else:
            st.info(f"No deals data with valid dates found for {selected_year_source}.")
    else:
        st.info("No deals data found for source analysis.")
    
    st.markdown("---")
    st.subheader("Close Rate by Month - 2025")
    
    with st.spinner("Loading leads data..."):
        all_leads = get_all_leads_for_sales_chart()
    
    leads_with_dates = pd.DataFrame()
    if all_leads:
        leads_df = pd.DataFrame(all_leads)
        leads_df['stage_date'] = pd.to_datetime(leads_df['stage_date'], errors='coerce')
        leads_with_dates = leads_df.dropna(subset=['stage_date'])
    else:
        leads_df = pd.DataFrame()
    
    sales_leads_df = pd.DataFrame()
    if not leads_with_dates.empty:
        sales_leads_df = leads_with_dates[leads_with_dates['board'] == 'Sales'].copy()
    
    if not sales_leads_df.empty:
        sales_leads_df['stage_date'] = pd.to_datetime(sales_leads_df['stage_date'], errors='coerce')
        sales_leads_df = sales_leads_df.dropna(subset=['stage_date'])
        sales_leads_df = sales_leads_df[sales_leads_df['stage_date'].dt.year == 2025]
        
        if not sales_leads_df.empty:
            if 'status' in sales_leads_df.columns:
                lead_status_series = sales_leads_df['status'].fillna('').astype(str)
            else:
                lead_status_series = pd.Series('', index=sales_leads_df.index)
            sales_leads_df['Lead Status'] = lead_status_series
            
            if 'assigned_person' in sales_leads_df.columns:
                assigned_series = sales_leads_df['assigned_person'].fillna('').astype(str)
            else:
                assigned_series = pd.Series('', index=sales_leads_df.index)
            sales_leads_df['Assigned Person'] = assigned_series.replace('', 'Unassigned')
            
            sales_leads_df['Month'] = sales_leads_df['stage_date'].dt.month
            sales_leads_df['Month Label'] = sales_leads_df['stage_date'].dt.strftime('%B')
            
            totals = (
                sales_leads_df
                .groupby(['Month', 'Month Label', 'Assigned Person'])
                .size()
                .reset_index(name='Total Leads')
            )
            
            closed_mask = sales_leads_df['Lead Status'].str.strip().str.lower().isin(['closed', 'win'])
            closed = (
                sales_leads_df[closed_mask]
                .groupby(['Month', 'Month Label', 'Assigned Person'])
                .size()
                .reset_index(name='Closed Leads')
            )
            
            close_rate_df = totals.merge(
                closed,
                on=['Month', 'Month Label', 'Assigned Person'],
                how='left'
            ).fillna({'Closed Leads': 0})
            
            close_rate_df['Closed Leads'] = close_rate_df['Closed Leads'].astype(int)
            close_rate_df['Close Rate'] = close_rate_df['Closed Leads'] / close_rate_df['Total Leads']
            close_rate_df['Close Rate Text'] = close_rate_df['Close Rate'].apply(
                lambda x: f"{x:.1%}" if x > 0 else ""
            )
            
            month_order = ['January', 'February', 'March', 'April', 'May', 'June',
                           'July', 'August', 'September', 'October', 'November', 'December']
            close_rate_df['Month Label'] = pd.Categorical(
                close_rate_df['Month Label'],
                categories=month_order,
                ordered=True
            )
            close_rate_df = close_rate_df.sort_values(['Month Label', 'Assigned Person'])
            
            fig_close_rate = px.bar(
                close_rate_df,
                x='Month Label',
                y='Close Rate',
                color='Assigned Person',
                barmode='group',
                labels={'Month Label': 'Month', 'Close Rate': 'Close Rate'},
                text='Close Rate Text',
                color_discrete_sequence=px.colors.qualitative.Bold
            )
            fig_close_rate.update_layout(
                height=600,
                xaxis_title='Month',
                yaxis_title='Close Rate',
                font=dict(size=14),
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.2,
                    xanchor="center",
                    x=0.5,
                    title=''
                )
            )
            fig_close_rate.update_traces(
                textposition='outside',
                textfont=dict(size=12, color='black')
            )
            fig_close_rate.update_yaxes(tickformat='.0%')
            fig_close_rate.update_xaxes(
                categoryorder='array',
                categoryarray=[m for m in month_order if m in close_rate_df['Month Label'].cat.categories]
            )
            st.plotly_chart(fig_close_rate, use_container_width=True)
        else:
            st.info("No Sales board leads found for 2025 to calculate close rate.")
    else:
        st.info("Close rate data is unavailable. Ensure there are Sales board leads for the selected period.")
    
    # Leads by Category Over Time Chart
    st.markdown("---")
    st.subheader("Leads by Category Over Time - 2025")
    
    # Leads by Category Over Time Chart reuses leads_with_dates (already loaded)
    if not leads_with_dates.empty:
        # Filter for 2025 data only
        leads_with_dates = leads_with_dates[leads_with_dates['stage_date'].dt.year == 2025]
        
        if not leads_with_dates.empty:
            # Add month columns for grouping and ordering
            leads_with_dates['Month'] = leads_with_dates['stage_date'].dt.month
            leads_with_dates['Month_Name'] = leads_with_dates['stage_date'].dt.strftime('%B')
            
            # Count leads by category and month
            category_counts = leads_with_dates.groupby(['Month', 'Month_Name', 'category']).size().reset_index(name='count')
            category_counts = category_counts.sort_values('Month')
            
            # Create pivot table for easier charting
            category_pivot = category_counts.pivot(index='Month_Name', columns='category', values='count').fillna(0)
            
            # Ensure we have the main categories
            main_categories = ['New Leads', 'Discovery Call', 'Design Review Call', 'Deck Call']
            for category in main_categories:
                if category not in category_pivot.columns:
                    category_pivot[category] = 0
            
            # Reorder columns
            category_pivot = category_pivot[main_categories]
            
            # Ensure consistent month order on x-axis
            month_order = ['January', 'February', 'March', 'April', 'May', 'June',
                           'July', 'August', 'September', 'October', 'November', 'December']
            available_months = [month for month in month_order if month in category_pivot.index]
            category_pivot = category_pivot.reindex(available_months).fillna(0)
            
            # Create bar chart using go.Figure for better control
            fig_leads = go.Figure()
            
            # Define colors for each category
            category_colors = {
                'New Leads': '#FF6B6B',           # Red
                'Discovery Call': '#4ECDC4',      # Teal
                'Design Review Call': '#ffcb00',  # Yellow
                'Deck Call': '#00c875'            # Green
            }
            
            # Add traces for each category
            for category in main_categories:
                if category in category_pivot.columns:
                    # Create text labels, but hide zeros
                    text_labels = []
                    for value in category_pivot[category]:
                        if value == 0:
                            text_labels.append("")  # Empty string for zero values
                        else:
                            text_labels.append(str(int(value)))  # Show actual value for non-zeros
                    
                    fig_leads.add_trace(go.Bar(
                        name=category,  # This will show in legend without "category" prefix
                        x=category_pivot.index,
                        y=category_pivot[category],
                        marker_color=category_colors[category],
                        text=text_labels,  # Use custom text labels that hide zeros
                        textposition='outside',
                        textfont=dict(size=12, color='black')
                    ))
            
            # Update layout
            fig_leads.update_layout(
                xaxis_title="Month",
                yaxis_title="Number of Leads",
                barmode='group',  # Side-by-side bars
                height=500,
                showlegend=True,
                xaxis=dict(
                    categoryorder='array',
                    categoryarray=available_months
                ),
                legend=dict(
                    orientation="h",   # horizontal
                    yanchor="top",
                    y=-0.2,            # below the chart
                    xanchor="center",
                    x=0.5
                )
            )
            
            st.plotly_chart(fig_leads, use_container_width=True)
        else:
            st.info("No leads data with valid dates available.")
    else:
        st.info("No leads data available.")

if __name__ == "__main__":
    main()