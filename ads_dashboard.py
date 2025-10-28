import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import sys

# Add parent directory to path to import database_utils
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from database_utils import get_ads_data, get_sales_data, check_database_exists, get_new_leads_data, get_discovery_call_data, get_design_review_data

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

# Page configuration
st.set_page_config(
    page_title="Monday.com Data Viewer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better embedding and responsive design
st.markdown("""
<style>
    .main {
        padding: 1rem;
    }
    .stDataFrame {
        font-size: 12px;
    }
    .stDataFrame > div {
        max-height: 600px;
        overflow-y: auto;
    }
    .stButton > button {
        width: 100%;
        margin-top: 1rem;
    }
    .metric-container {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
    }
    .embed-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
        text-align: center;
    }
    .stMetric {
        background-color: #f8f9fa;
        padding: 0.5rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    /* Hide QuickBooks and SignNow pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"] {
        display: none !important;
    }
    @media (max-width: 768px) {
        .embed-header {
            font-size: 1.2rem;
        }
        .stDataFrame {
            font-size: 10px;
        }
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_ads_data_from_db():
    """Get ads data from SQLite database"""
    return get_ads_data()

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_sales_data_from_db():
    """Get sales data from SQLite database"""
    return get_sales_data()

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_all_leads_for_utm():
    """Get all leads data from all boards using database only for speed"""
    import json
    
    all_leads = []
    
    # Get data from all boards using database functions
    boards_data = {
        'New Leads v2': get_new_leads_data(),
        'Discovery Call v2': get_discovery_call_data(), 
        'Design Review v2': get_design_review_data(),
        'Sales v2': get_sales_data().get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
    }
    
    # Board-specific channel column IDs
    channel_columns = {
        'Sales v2': 'text_mkrfer1n',
        'Design Review v2': 'text_mkrkkpx0',
        'Discovery Call v2': 'text_mkrk2tj8',
        'New Leads v2': 'text_mkref4p0'
    }
    
    # Process each board's data
    for board_name, items in boards_data.items():
        for item in items:
            # Extract channel information
            channel = ""
            date_created = None
            
            target_channel_column = channel_columns.get(board_name)
            
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
                
                # Look for the specific channel column for this board
                if col_id == target_channel_column and text:
                    channel = text
                
                # Look for date created
                if col_type == "date" and text and ("created" in col_id or "date" in col_id):
                    date_created = text
                    break
            
            # Only include items with valid channels (not empty and not placeholder)
            if channel and channel.strip() and channel != "[channel]":
                all_leads.append({
                    'name': item.get('name', ''),
                    'board': board_name,
                    'channel': channel,
                    'date_created': date_created,
                    'channel': channel
                })
    
    return all_leads

def format_ads_data(data):
    """Convert Monday.com ads data to pandas DataFrame"""
    if not data or "data" not in data or "boards" not in data["data"]:
        return pd.DataFrame()
    
    boards = data["data"]["boards"]
    if not boards or not boards[0].get("items_page"):
        return pd.DataFrame()
    
    items_page = boards[0]["items_page"]
    if not items_page or not items_page.get("items"):
        return pd.DataFrame()
    
    items = items_page["items"]
    if not items:
        return pd.DataFrame()
    
    # Convert to DataFrame
    records = []
    for item in items:
        record = {
            "Item": item.get("name", ""),
            "Attribution Date": "",
            "Google Adspend": ""
        }
        
        # Add specific columns we want to display
        for col_val in item.get("column_values", []):
            col_id = col_val.get("id", "")
            text = col_val.get("text") or ""
            value = col_val.get("value", "")
            
            # Map specific column IDs to our desired column names
            if col_id == "name":  # Item name
                record["Item"] = item.get("name", "")
            elif col_id == "date_mkv81p3z":  # Attribution Date
                record["Attribution Date"] = text if text else ""
            elif col_id == "numeric_mkv863mb":  # Google Adspend (the actual column with data)
                record["Google Adspend"] = text if text else ""
        
        records.append(record)
    
    df = pd.DataFrame(records)
    
    # Convert date columns and create month/year column
    df['Attribution Date'] = pd.to_datetime(df['Attribution Date'], errors='coerce')
    
    # Create Month/Year column for x-axis
    df['Month Year'] = df['Attribution Date'].dt.strftime('%B %Y')
    
    # Convert Google Adspend to numeric
    df['Google Adspend'] = pd.to_numeric(df['Google Adspend'], errors='coerce')
    
    # Sort by attribution date
    df = df.sort_values('Attribution Date')
    
    return df

def format_sales_data(data):
    """Convert Monday.com sales data to pandas DataFrame with ROAS filtering"""
    if not data or "data" not in data or "boards" not in data["data"]:
        st.warning("No sales data found in API response")
        return pd.DataFrame()
    
    
    boards = data["data"]["boards"]
    if not boards or not boards[0].get("items_page"):
        st.warning("No items page found in sales board")
        return pd.DataFrame()
    
    items_page = boards[0]["items_page"]
    if not items_page or not items_page.get("items"):
        st.warning("No items found in sales board")
        return pd.DataFrame()
    
    items = items_page["items"]
    if not items:
        st.warning("Empty items list in sales board")
        return pd.DataFrame()
    
    # Convert to DataFrame
    records = []
    for item in items:
        try:
            record = {
                "Item Name": item.get("name", ""),
                "Status": "",
                "Channel": "",
                "Value": "",
                "Date Created": "",
                "Date Closed": "",
                "Assigned Person": ""
            }
            
            # Extract column values
            for col_val in item.get("column_values", []):
                try:
                    col_id = col_val.get("id", "")
                    text = (col_val.get("text") or "").strip()
                    value = col_val.get("value", "")
                    
                    
                    # Map specific column IDs based on the actual Monday.com structure
                    # Status field - using the color column that shows "Sales Qualified"
                    if col_id == "color_mknxd1j2":
                        record["Status"] = text
                    # Channel/Source field - using the correct column for paid search data
                    elif col_id == "text_mkrfer1n":  # This contains "Paid search" data
                        record["Channel"] = text
                    elif col_id == "source":  # Fallback to source column
                        record["Channel"] = text
                    # Amount Paid or Contract Value field (prioritize numbers3 over contract_amt)
                    elif col_id == "numbers3":  # This contains the actual amount paid (10550 for Kimberly)
                        record["Value"] = text
                    elif col_id == "contract_amt":
                        record["Value"] = text if text else record.get("Value", "")
                    elif col_id == "formula_mktj2qh2":  # Try first formula column
                        record["Value"] = text if text else record.get("Value", "")
                    elif col_id == "formula_mktk2rgx":  # Try second formula column
                        record["Value"] = text if text else record.get("Value", "")
                    elif col_id == "formula_mktks5te":  # Try third formula column
                        record["Value"] = text if text else record.get("Value", "")
                    elif col_id == "formula_mktknqy9":  # Try fourth formula column
                        record["Value"] = text if text else record.get("Value", "")
                    elif col_id == "formula_mktkwnyh":  # Try fifth formula column
                        record["Value"] = text if text else record.get("Value", "")
                    elif col_id == "formula_mktq5ahq":  # Try sixth formula column
                        record["Value"] = text if text else record.get("Value", "")
                    elif col_id == "formula_mktt5nty":  # Try seventh formula column
                        record["Value"] = text if text else record.get("Value", "")
                    elif col_id == "formula_mkv0r139":  # Try eighth formula column
                        record["Value"] = text if text else record.get("Value", "")
                    # Date Created field - using date7 column (Jan 14 for Ashley Miles)
                    elif col_id == "date7":
                        record["Date Created"] = text
                    # Date Closed field - using date_mktq7npm column (Aug 20 for Ashley Miles)
                    elif col_id == "date_mktq7npm":
                        record["Date Closed"] = text
                    # Assigned Person field
                    elif col_id == "color_mkvewcwe":
                        record["Assigned Person"] = text
                    # Fallback mappings for other possible columns
                    elif col_id == "status_14__1":  # This shows "OTHER" in your data
                        if record["Channel"] == "":  # Only use if source is empty
                            record["Channel"] = text
                    elif any(word in col_id.lower() for word in ["status", "stage", "state", "phase"]) and record["Status"] == "":
                        record["Status"] = text
                    elif any(word in col_id.lower() for word in ["channel", "source", "utm", "traffic", "medium"]) and record["Channel"] == "":
                        record["Channel"] = text
                    elif any(word in col_id.lower() for word in ["value", "revenue", "amount", "price", "deal", "contract"]) and record["Value"] == "":
                        record["Value"] = text
                except Exception as e:
                    # Skip problematic column values
                    continue
            
            records.append(record)
        except Exception as e:
            # Skip problematic items
            continue
    
    df = pd.DataFrame(records)
    
    # Convert date columns
    df['Date Created'] = pd.to_datetime(df['Date Created'], errors='coerce')
    df['Date Closed'] = pd.to_datetime(df['Date Closed'], errors='coerce')
    
    # Convert Value to numeric (remove $ and commas)
    df['Value'] = df['Value'].astype(str).str.replace('$', '').str.replace(',', '').str.replace(' ', '')
    df['Value'] = pd.to_numeric(df['Value'], errors='coerce')
    
    # Create Month/Year column based on Date Created
    df['Month Year'] = df['Date Created'].dt.strftime('%B %Y')
    df['Year'] = df['Date Created'].dt.year
    
    return df

def filter_roas_data(df, raw_data=None):
    """Filter sales data for ROAS calculation"""
    # Status: ONLY include "Closed" status records
    closed_statuses = ['closed']
    
    # Channel: ONLY include "Paid Search" channel records
    paid_search_channels = ['paid search']
    
    # Use exact matching to accept both "Closed" and "Win" statuses
    df_roas = df[
        (df['Status'].str.lower().isin(['closed', 'win'])) &
        (df['Channel'].str.lower() == 'paid search')
    ].copy()
    
    return df_roas, closed_statuses, paid_search_channels

def calculate_roas(ads_df, sales_df):
    """Calculate ROAS (Return on Ad Spend) by month"""
    if ads_df.empty or sales_df.empty:
        return pd.DataFrame()
    
    # Group ads data by month/year
    ads_monthly = ads_df.groupby('Month Year')['Google Adspend'].sum().reset_index()
    
    # Group sales data by month/year (Date Created)
    sales_monthly = sales_df.groupby('Month Year')['Value'].sum().reset_index()
    
    # Merge the data
    roas_df = pd.merge(ads_monthly, sales_monthly, on='Month Year', how='outer')
    
    # Fill missing values with 0
    roas_df['Google Adspend'] = roas_df['Google Adspend'].fillna(0)
    roas_df['Value'] = roas_df['Value'].fillna(0)
    
    # Calculate ROAS (Revenue / Ad Spend)
    roas_df['ROAS'] = roas_df.apply(
        lambda row: row['Value'] / row['Google Adspend'] if row['Google Adspend'] > 0 else 0,
        axis=1
    )
    
    # Filter for 2025 only
    roas_df = roas_df[roas_df['Month Year'].str.contains('2025', na=False)]
    
    # Sort by date
    roas_df['Date'] = pd.to_datetime(roas_df['Month Year'], errors='coerce')
    roas_df = roas_df.sort_values('Date')
    
    return roas_df

def main():
    """Main application function"""
    # Header
    st.markdown('<div class="embed-header">📊 GOOGLE ADS ATTRIBUTION DASHBOARD</div>', unsafe_allow_html=True)
    
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
    
    # Load data from database
    with st.spinner("Loading data from database..."):
        try:
            ads_data = get_ads_data_from_db()
            sales_data = get_sales_data_from_db()
            ads_df = format_ads_data(ads_data)
            sales_df_raw = format_sales_data(sales_data)
            
            # Filter sales data for ROAS calculation
            sales_df, closed_statuses, paid_search_channels = filter_roas_data(sales_df_raw, sales_data)
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")
            st.info("Please refresh the database using the 'Database Refresh' page")
            return

    # Check if we have data
    if ads_df.empty and sales_df.empty:
        st.warning("No records found in either board. Add some items to Monday.com to see them here.")
        st.info("💡 **Tip**: Make sure your Monday.com boards have items and your API token has the correct permissions.")
    else:
        # ROAS Section
        st.subheader("📈 Return on Ad Spend (ROAS) - 2025")
        
        # Calculate ROAS
        roas_df = calculate_roas(ads_df, sales_df)
        
        # Don't filter - show all months for 2025
        if not roas_df.empty:
            # Create ROAS chart - show all months
            fig_roas = px.bar(
                roas_df,
                x='Month Year',
                y='ROAS',
                title='',
                labels={'ROAS': 'ROAS', 'Month Year': 'Month'},
                range_y=[0, None]  # Start y-axis at 0 to show all values
            )
            
            fig_roas.update_layout(
                xaxis_title="Month",
                yaxis_title="ROAS",
                height=500,
                showlegend=False
            )
            
            # Add vertical line at ROAS = 1 (break-even point)
            fig_roas.add_hline(y=1, line_dash="dash", line_color="red", 
                              annotation_text="Break-even (ROAS = 1)", 
                              annotation_position="top right")
            
            # Add ROAS value labels above each bar (3 decimal places, hide if 0.000)
            fig_roas.update_traces(
                text=[f"<b>{y:.3f}</b>" if y != 0.0 else "" for y in roas_df['ROAS']],
                textposition='outside',
                textfont=dict(size=14, color='black')
            )
            
            # Rotate x-axis labels
            fig_roas.update_xaxes(tickangle=45)
            
            st.plotly_chart(fig_roas, use_container_width=True)
            
            # Profit on Ads 2025 Graph (moved here)
            st.subheader("📈 Profit on Ads 2025")
            
            # Create separate ROAS data for profit chart (include ALL months, not just those with sales)
            roas_df_profit = calculate_roas(ads_df, sales_df)
            
            if not roas_df_profit.empty:
                # Calculate profit (Revenue - Ad Spend) for 2025
                roas_df_profit['Profit'] = roas_df_profit['Value'] - roas_df_profit['Google Adspend']
                
                # Create profit bar chart with solid colors
                fig_profit = go.Figure()
                
                # Add bars with solid red for negative, solid green for positive
                for _, row in roas_df_profit.iterrows():
                    color = 'red' if row['Profit'] < 0 else 'green'
                    fig_profit.add_trace(go.Bar(
                        x=[row['Month Year']],
                        y=[row['Profit']],
                        marker_color=color,
                        showlegend=False
                    ))
                
                # Update layout
                fig_profit.update_layout(
                    xaxis_title="Month",
                    yaxis_title="Profit ($)",
                    height=600,
                    showlegend=False,
                    coloraxis_showscale=False  # Hide color scale
                )
                
                # Rotate x-axis labels
                fig_profit.update_xaxes(tickangle=45)
                
                # Add value labels on bars
                fig_profit.update_traces(
                    texttemplate='<b>$%{y:,.0f}</b>',
                    textposition='outside',
                    textfont=dict(size=14, color='black')
                )
                
                st.plotly_chart(fig_profit, use_container_width=True)
            else:
                st.info("No profit data available for 2025")
            
            # Detailed Sales Table Section
            st.subheader("🔍 Detailed Sales Analysis")
            
            # Month selector for detailed view - only show months with qualifying sales (Closed + Paid Search)
            if not sales_df.empty:
                # Filter sales_df to only include Closed/Win + Paid Search records (same as ROAS)
                # Use exact matching to accept both "Closed" and "Win" statuses
                qualifying_sales = sales_df[
                    (sales_df['Status'].str.lower().isin(['closed', 'win'])) &
                    (sales_df['Channel'].str.lower() == 'paid search')
                ]
                
                # Sort months chronologically and remove NaN values from qualifying sales only
                available_months = sorted([month for month in qualifying_sales['Month Year'].unique() 
                                         if pd.notna(month) and month != 'nan'], 
                                        key=lambda x: pd.to_datetime(x))
                if available_months:
                    selected_month = st.selectbox(
                        "Select Month to View Detailed Sales:",
                        available_months
                    )
                    
                    # Filter sales data for selected month (using same criteria as ROAS calculation)
                    month_sales = qualifying_sales[qualifying_sales['Month Year'] == selected_month]
                    
                    
                    # Display total revenue for selected month
                    month_revenue = month_sales['Value'].sum()
                    st.metric(f"Total Revenue ({selected_month})", f"${month_revenue:,.2f}")
                    
                    if not month_sales.empty:
                        # Prepare data for display
                        display_data = month_sales.copy()
                        
                        # Format Date Created to YYYY-MM-DD
                        display_data['Date Created'] = display_data['Date Created'].dt.strftime('%Y-%m-%d')
                        
                        # Format Date Closed to YYYY-MM-DD (handle NaT values)
                        display_data['Date Closed'] = display_data['Date Closed'].dt.strftime('%Y-%m-%d')
                        display_data['Date Closed'] = display_data['Date Closed'].fillna('')
                        
                        # Format Value with commas for thousands, handle NaN values
                        display_data['Formatted Value'] = display_data['Value'].apply(
                            lambda x: f"${x:,.2f}" if pd.notna(x) and x != 0 else " "
                        )
                        
                        # Add Assigned Person column if it exists
                        if 'Assigned Person' in display_data.columns:
                            display_columns = ['Item Name', 'Formatted Value', 'Assigned Person', 'Date Created', 'Date Closed']
                        else:
                            display_columns = ['Item Name', 'Formatted Value', 'Date Created', 'Date Closed']
                        
                        st.dataframe(
                            display_data[display_columns],
                            width='stretch',
                            hide_index=True,
                            column_config={
                                "Formatted Value": "Revenue ($)",
                                "Date Created": "Date Created",
                                "Date Closed": "Date Closed"
                            }
                        )
                    else:
                        st.info(f"No Closed + Paid Search sales found for {selected_month}")
                else:
                    st.info("No sales data available for detailed analysis")
        
        
        # Original Ad Spend Section
        if not ads_df.empty:
            st.markdown("---")
            
            # Year filter for ads data
            st.subheader("📅 Filter by Year")
            
            # Get unique years from the ads data
            ads_with_dates = ads_df.dropna(subset=['Attribution Date'])
            if not ads_with_dates.empty:
                ads_with_dates['Year'] = ads_with_dates['Attribution Date'].dt.year
                available_years = sorted(ads_with_dates['Year'].unique())
                
                # Add "All Years" option
                year_options = ["All Years"] + [str(year) for year in available_years]
                # Set default to 2025 if available, otherwise "All Years"
                default_index = year_options.index("2025") if "2025" in year_options else 0
                selected_year = st.selectbox("Select Year:", year_options, index=default_index)
                
                # Filter data based on selected year
                if selected_year == "All Years":
                    ads_filtered = ads_with_dates
                    year_label = "All Years"
                else:
                    ads_filtered = ads_with_dates[ads_with_dates['Year'] == int(selected_year)]
                    year_label = selected_year
            else:
                ads_filtered = ads_with_dates
                year_label = "All Years"
                selected_year = "All Years"
            
            # Total Adspend metric
            st.subheader("💰 Total Ad Spend")
            total_adspend = ads_filtered['Google Adspend'].sum()
            st.metric("Total Ad Spend", f"${total_adspend:,.2f}", delta=None)
            
            # Create the bar chart
            st.subheader(f"📊 Adspend by Month - {year_label}")
            
            # Filter out rows with missing data for charting
            ads_chart = ads_filtered.dropna(subset=['Attribution Date', 'Google Adspend'])
            
            if not ads_chart.empty:
                # Create bar chart
                fig = px.bar(
                    ads_chart,
                    x='Month Year',
                    y='Google Adspend',
                    title=f'Google Adspend by Month - {year_label}',
                    labels={'Google Adspend': 'Adspend ($)', 'Month Year': 'Month'},
                    color_discrete_sequence=['#1f77b4']
                )
                
                # Update layout
                fig.update_layout(
                    xaxis_title="Month",
                    yaxis_title="Adspend ($)",
                    height=600,
                    showlegend=False
                )
                
                # Add value labels above each bar
                fig.update_traces(
                    texttemplate='<b>$%{y:,.0f}</b>',
                    textposition='outside',
                    textfont=dict(size=14, color='black')
                )
                
                # Rotate x-axis labels
                fig.update_xaxes(tickangle=45)
                
                # Display the chart
                st.plotly_chart(fig, use_container_width=True)
                
            else:
                st.warning("No ad spend data available for charting.")

    # UTM Data Section at the bottom
    st.markdown("---")
    st.subheader("📊 UTM Data (Leads by Channel - 2025)")
    
    # Get all leads data for UTM analysis
    with st.spinner("Loading UTM data..."):
        all_leads = get_all_leads_for_utm()
    
    if all_leads:
        # Convert to DataFrame
        leads_df = pd.DataFrame(all_leads)
        
        # Parse dates and filter for valid dates
        leads_df['date_created'] = pd.to_datetime(leads_df['date_created'], errors='coerce')
        leads_with_dates = leads_df.dropna(subset=['date_created'])
        
        # Filter for 2025 data only
        leads_with_dates = leads_with_dates[leads_with_dates['date_created'].dt.year == 2025]
        
        if not leads_with_dates.empty:
            # Add month-year column for grouping with proper formatting
            leads_with_dates['Month Year'] = leads_with_dates['date_created'].dt.strftime('%B %Y')
            
            # Count leads by raw channel and month (use channel instead of categorized channel)
            channel_counts = leads_with_dates.groupby(['Month Year', 'channel']).size().reset_index(name='count')
            
            # Create pivot table for easier charting
            channel_pivot = channel_counts.pivot(index='Month Year', columns='channel', values='count').fillna(0)
            
            # Sort the pivot table by month chronologically
            channel_pivot.index = pd.to_datetime(channel_pivot.index)
            channel_pivot = channel_pivot.sort_index()
            channel_pivot.index = channel_pivot.index.strftime('%B %Y')
            
            # Create side-by-side bar chart with dynamic colors
            fig = px.bar(
                channel_pivot,
                title='',
                labels={'value': 'Number of Leads', 'index': 'Month'}
            )
            
            # Update layout
            fig.update_layout(
                xaxis_title="Month",
                yaxis_title="Number of Leads",
                height=500,
                barmode='group',  # Side-by-side bars instead of stacked
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.2,
                    xanchor="center",
                    x=0.5,
                    title_text=""  # Remove legend title
                )
            )
            
            # Remove "channel" prefix from legend entries
            for trace in fig.data:
                if hasattr(trace, 'name') and trace.name:
                    # Remove any "channel" prefix from the trace name
                    trace.name = trace.name.replace('channel', '').strip()
            
            # Rotate x-axis labels
            fig.update_xaxes(tickangle=45)
            
            # Display the chart
            st.plotly_chart(fig, use_container_width=True)
            
        else:
            st.warning("No leads with valid creation dates found for 2025 UTM analysis.")
    else:
        st.warning("No leads data found for UTM analysis.")

    # NEW: Qualified vs Unqualified Breakdown Section
    st.markdown("---")
    st.subheader("🎯 Qualified vs. Unqualified Breakdown by Form Field")
    
    @st.cache_data(ttl=300)
    def get_lead_qualification_data():
        """Extract and process leads for qualification analysis from ALL boards"""
        import json
        
        # Get data from all 4 boards
        new_leads_items = get_new_leads_data()
        discovery_call_items = get_discovery_call_data()
        design_review_items = get_design_review_data()
        sales_items = get_sales_data().get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
        
        print(f"Got {len(new_leads_items)} items from New Leads")
        print(f"Got {len(discovery_call_items)} items from Discovery Call")
        print(f"Got {len(design_review_items)} items from Design Review")
        print(f"Got {len(sales_items)} items from Sales")
        
        # Combine all items from all boards
        all_items = []
        all_items.extend(new_leads_items)
        all_items.extend(discovery_call_items)
        all_items.extend(design_review_items)
        all_items.extend(sales_items)
        
        print(f"Total items from all boards: {len(all_items)}")
        
        # Qualification rule: Unqualified only if status is Disqualified; otherwise Qualified
        
        all_lead_data = []
        
        for item in all_items:
            lead_status = ""
            lead_data = {}
            
            column_values = item.get("column_values", [])
            if isinstance(column_values, str):
                try:
                    column_values = json.loads(column_values)
                except:
                    column_values = []
            
            # Define the "Disqualified" status column for each board type
            # These columns contain the actual Lead Status
            disqualified_status_cols = {
                "status7",  # New Leads
                "color_mknx1h9r",  # Discovery Call
                "color_mknx4zp1",  # Design Review
                "color_mknxd1j2"   # Sales
            }
            
            # Extract ALL column data from the item
            for col_val in column_values:
                col_id = col_val.get("id", "")
                text = (col_val.get("text") or "").strip()
                
                if text:  # Only process columns with values
                    # Get ALL column values - we'll identify form fields by pattern matching values
                    lead_data[col_id] = text
                    
                    # Find Lead Status - check the Disqualified status columns
                    if col_id in disqualified_status_cols and not lead_status:
                        lead_status = text
            
            # New rule: only "Disqualified" is unqualified; anything else (including empty) is qualified
            is_qualified = True if not lead_status else str(lead_status).strip().lower() != "disqualified"
            
            all_lead_data.append({
                'lead_status': lead_status,
                'is_qualified': is_qualified,
                **lead_data
            })
        
        # Debug: Print column IDs we found
        if all_lead_data:
            print(f"Sample lead data keys: {list(all_lead_data[0].keys())[:20]}")
            qualified_count = sum(1 for item in all_lead_data if item['is_qualified'])
            print(f"Qualified: {qualified_count}/{len(all_lead_data)}")
        
        return all_lead_data
    
    # Get qualification data
    with st.spinner("Loading lead qualification data..."):
        qualification_data = get_lead_qualification_data()
    
    if qualification_data:
        df = pd.DataFrame(qualification_data)
        
        # Show available columns for debugging (commented out for production)
        # col_ids = [c for c in df.columns if c not in ['lead_status', 'is_qualified']]
        # st.write(f"Available column IDs ({len(col_ids)}): {col_ids[:10]}...")
        
        # Use explicit column ID mappings based on actual database
        # Each field can have different column IDs on different boards
        field_column_mapping = {
            'CLIENT TYPE?': ['status_1__1', 'status_14__1'],  # status_1__1 for New Leads, status_14__1 for other boards
            'WHAT IS YOUR TIMELINE FOR STARTING?': ['text_mkwf56ca', 'text3__1'],  # text_mkwf56ca for New Leads, text3__1 for other boards
            'WHAT IS YOUR STATUS?': ['text_mkwf2541', 'text_mkwf8r57', 'text37__1'],  # text_mkwf2541 for New Leads, text_mkwf8r57 for Discovery Call, text37__1 for Design Review
            'HOW MANY STYLES DO YOU WANT TO DEVELOP?': ['text_mkwfxk8t', 'text_mkwfs99f', 'text30__1', 'text30__1'],  # text_mkwfxk8t for New Leads, text_mkwfs99f for Discovery, text30__1 for Design Review, text30__1 for Sales
            'WHAT KINDS OF CLOTHING DO YOU WANT TO MAKE?': ['text_mkwfva26', 'text_mkwf8n18', 'text8__1'],  # text_mkwfva26 for New Leads, text_mkwf8n18 for Discovery Call, text8__1 for Design Review and Sales
            'BUDGET FOR DEVELOPMENT (PATTERNS AND SAMPLES)': ['text_mkwfkqex', 'text_mkwf9e6c', 'text7__1']  # text_mkwfkqex for New Leads, text_mkwf9e6c for Discovery Call, text7__1 for Design Review and Sales
        }
        
        # All fields now use lists of possible column IDs
        identified_fields = {}
        for field_name, col_id_list in field_column_mapping.items():
            # Check if any of the columns exist in the DataFrame
            found_cols = [cid for cid in col_id_list if cid in df.columns]
            if found_cols:
                # Store the list of available columns for this field
                identified_fields[field_name] = found_cols[0]  # Use first as primary
                print(f"Mapped '{field_name}' to columns {found_cols} (from {col_id_list})")
            else:
                print(f"WARNING: No columns found for field '{field_name}' (tried {col_id_list})")
        
        if not identified_fields:
            st.error("Could not identify form field columns. Data structure may be different than expected.")
            return
        
        # Create visualizations for each identified field
        for field_name, col_id in identified_fields.items():
            if col_id not in df.columns:
                continue
            
            # Get the list of column IDs for this field
            col_ids_for_field = field_column_mapping.get(field_name, [col_id])
            if not isinstance(col_ids_for_field, list):
                col_ids_for_field = [col_ids_for_field]
            
            # Aggregate values from all possible column IDs for this field
            all_values = []
            available_cols = []
            for col_id_option in col_ids_for_field:
                if col_id_option in df.columns:
                    available_cols.append(col_id_option)
                    valid_data = df[df[col_id_option].notna() & (df[col_id_option] != "")]
                    if not valid_data.empty:
                        all_values.extend(valid_data[col_id_option].tolist())
            
            unique_values = list(set(all_values)) if all_values else []
            
            # Use the first available column ID as primary for filtering
            if available_cols:
                col_id = available_cols[0]
            
            # Filter out invalid values for specific fields
            if 'HOW MANY STYLES' in field_name.upper():
                # Only show valid options
                valid_options = ['LESS THAN 5', '5-10', '11-20', '20+', 'I DON\'T KNOW']
                unique_values = [v for v in unique_values if v.upper() in [o.upper() for o in valid_options]]
            elif field_name.upper().startswith('CLIENT TYPE'):
                # Exclude EXISTING (case insensitive)
                unique_values = [v for v in unique_values if str(v).strip().lower() != 'existing']
            elif 'TIMELINE' in field_name.upper():
                # Only show valid timeline values
                valid_timelines = ['I JUST WANT TO LEARN THE PROCESS', 'READY TO GET STARTED', 'WITHIN THE NEXT 90 DAYS']
                unique_values = [v for v in unique_values if v.upper() in [t.upper() for t in valid_timelines]]
            elif 'BUDGET' in field_name.upper():
                # Only show the five allowed budget values
                allowed_budgets = ['< $5,000', '$5,000 - $10,000', '$10,000 - $20,000', '$20,000 - $50,000', 'other']
                unique_values = [v for v in unique_values if str(v) in allowed_budgets]
            
            if len(unique_values) == 0:
                continue
            
            # Use the field name as the title
            st.markdown(f"### {field_name}")
            
            # Check if this looks like a multi-select field (comma-separated values)
            # CLOTHING is always multi-select, BUDGET is NOT multi-select
            is_clothing_field = 'CLOTHING' in field_name.upper()
            is_budget_field = 'BUDGET' in field_name.upper()
            has_commas = any(',' in str(val) for val in unique_values[:10])
            
            if has_commas and is_clothing_field:
                # Split comma-separated values ONLY for CLOTHING field
                # Check all possible column IDs for this field
                multi_select_data = []
                for _, row in df.iterrows():
                    # Check all column IDs for this field
                    for col_id_option in col_ids_for_field:
                        if col_id_option in df.columns and pd.notna(row[col_id_option]) and row[col_id_option] != "":
                            items = [x.strip() for x in str(row[col_id_option]).split(',')]
                            for item in items:
                                if item:  # Only if not empty after strip
                                    multi_select_data.append({
                                        col_id_option: item,
                                        'is_qualified': row['is_qualified']
                                    })
                
                if multi_select_data:
                    multi_select_df = pd.DataFrame(multi_select_data)
                    # Get unique values from all columns, not just primary
                    all_values = []
                    for col_id_option in col_ids_for_field:
                        if col_id_option in multi_select_df.columns:
                            all_values.extend(multi_select_df[col_id_option].dropna().unique().tolist())
                    unique_values = list(set(all_values)) if all_values else []
                else:
                    unique_values = []
            else:
                multi_select_df = df
            
            # Ensure unique_values is a list
            unique_values = list(unique_values) if not isinstance(unique_values, list) else unique_values
            
            # Filter out invalid values again after multi-select processing
            if 'HOW MANY STYLES' in field_name.upper():
                # Only show valid options
                valid_options = ['LESS THAN 5', '5-10', '11-20', '20+', 'I DON\'T KNOW']
                unique_values = [v for v in unique_values if v.upper() in [o.upper() for o in valid_options]]
            elif 'CLOTHING' in field_name.upper():
                # Only show valid clothing types (must be all uppercase)
                valid_clothing = ['WOMENSWEAR', 'MENSWEAR', 'STREETWEAR', 'ACTIVEWEAR', 'KIDS', 'BRIDAL/COUTURE', 'OTHER']
                unique_values = [v for v in unique_values if v == v.upper() and v in valid_clothing]
            elif 'TIMELINE' in field_name.upper():
                # Only show valid timeline values
                valid_timelines = ['I JUST WANT TO LEARN THE PROCESS', 'READY TO GET STARTED', 'WITHIN THE NEXT 90 DAYS']
                unique_values = [v for v in unique_values if v.upper() in [t.upper() for t in valid_timelines]]
            
            # Create pie charts for each unique value
            num_values = len(unique_values)
            
            if num_values > 0:
                # Create a grid layout (2 columns)
                cols = st.columns(2)
                
                for idx, unique_value in enumerate(unique_values):
                    # Determine which column to use
                    col_idx = idx % 2
                    
                    # Filter data for this value
                    if has_commas and is_clothing_field:
                        # Check all columns in multi_select_df for this field
                        mask = pd.Series([False] * len(multi_select_df))
                        for col_id_option in col_ids_for_field:
                            if col_id_option in multi_select_df.columns:
                                mask = mask | (multi_select_df[col_id_option] == unique_value)
                        filtered_data = multi_select_df[mask]
                    else:
                        # Check all possible column IDs for this field
                        mask = pd.Series([False] * len(df))
                        for col_id_option in col_ids_for_field:
                            if col_id_option in df.columns:
                                mask = mask | (df[col_id_option] == unique_value)
                        filtered_data = df[mask]
                    
                    # Count qualified vs unqualified
                    qualified_count = filtered_data['is_qualified'].sum()
                    unqualified_count = (~filtered_data['is_qualified']).sum()
                    
                    total = qualified_count + unqualified_count
                    
                    if total > 0:
                        with cols[col_idx]:
                            # Create pie chart
                            fig = go.Figure(data=[
                                go.Pie(
                                    labels=['Qualified', 'Unqualified'],
                                    values=[qualified_count, unqualified_count],
                                    marker_colors=['#2ecc71', '#e74c3c'],  # Green for qualified, Red for unqualified
                                    textinfo='label+percent+value',
                                    textfont=dict(size=14, color='white'),
                                    hole=0.3
                                )
                            ])
                            
                            fig.update_layout(
                                title=f"{unique_value}<br>({total} total)",
                                height=500,
                                showlegend=True,
                                legend=dict(
                                    orientation="h",
                                    yanchor="top",
                                    y=-0.15,  # Move legend below the chart
                                    xanchor="center",
                                    x=0.5,
                                    font=dict(size=12)
                                ),
                                margin=dict(b=50, t=80)  # Add bottom margin for legend
                            )
                            
                            st.plotly_chart(fig, use_container_width=True, key=f"{field_name}_{unique_value}_{idx}")
            
            st.markdown("<br>", unsafe_allow_html=True)
    
    else:
        st.warning("No lead qualification data available.")
        st.info("The dataset may be empty or column extraction failed. Check console logs for debug output.")

if __name__ == "__main__":
    main()