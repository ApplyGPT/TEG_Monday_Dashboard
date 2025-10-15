import streamlit as st
import requests
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# Page configuration
st.set_page_config(
    page_title="Monday.com Data Viewer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

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
            
        required_board_ids = ['ads_board_id', 'sales_board_id']
        
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
ADS_BOARD_ID = credentials['ads_board_id']
SALES_BOARD_ID = credentials['sales_board_id']

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
def get_ads_data():
    """Get all items from Ads board with caching"""
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json",
    }
    
    query = f"""
    query {{
        boards(ids: {ADS_BOARD_ID}) {{
            items_page(limit: 100) {{
                items {{
                    id
                    name
                    state
                    created_at
                    updated_at
                    column_values {{
                        id
                        text
                        value
                    }}
                }}
            }}
        }}
    }}
    """
    
    try:
        response = requests.post(url, json={"query": query}, headers=headers, timeout=30)
        
        if response.status_code == 401:
            st.error("401 Unauthorized: Check your API token and permissions.")
            return None
        
        response.raise_for_status()
        return response.json()
    except requests.exceptions.Timeout:
        st.error("Request timed out. Please try again.")
        return None
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching ads data: {str(e)}")
        return None
    except Exception as e:
        st.error(f"Unexpected error: {str(e)}")
        return None

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_sales_data():
    """Get all items from Sales board with caching"""
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json",
    }
    
    query = f"""
    query {{
        boards(ids: {SALES_BOARD_ID}) {{
            items_page(limit: 500) {{
                items {{
                    id
                    name
                    state
                    created_at
                    updated_at
                    column_values {{
                        id
                        text
                        value
                    }}
                }}
            }}
        }}
    }}
    """
    
    try:
        response = requests.post(url, json={"query": query}, headers=headers, timeout=30)
        
        if response.status_code == 401:
            st.error("401 Unauthorized: Check your API token and permissions.")
            return None
        
        response.raise_for_status()
        return response.json()
    except requests.exceptions.Timeout:
        st.error("Request timed out. Please try again.")
        return None
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching sales data: {str(e)}")
        return None
    except Exception as e:
        st.error(f"Unexpected error: {str(e)}")
        return None

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
                    # Contract amount (revenue) field
                    elif col_id == "contract_amt":
                        record["Value"] = text
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
    
    df_roas = df[
        (df['Status'].str.contains('|'.join(closed_statuses), case=False, na=False)) &
        (df['Channel'].str.contains('|'.join(paid_search_channels), case=False, na=False))
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
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Settings")
        st.info(f"Ads Board ID: {ADS_BOARD_ID}")
        st.info(f"Sales Board ID: {SALES_BOARD_ID}")
        st.info(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Refresh button
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
    
    # Load data
    with st.spinner("Loading data from Monday.com..."):
        try:
            ads_data = get_ads_data()
            sales_data = get_sales_data()
            ads_df = format_ads_data(ads_data)
            sales_df_raw = format_sales_data(sales_data)
            
            # Filter sales data for ROAS calculation
            sales_df, closed_statuses, paid_search_channels = filter_roas_data(sales_df_raw, sales_data)
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")
            st.info("Please check your API token and board IDs in secrets.toml")
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
            
            # Month selector for detailed view
            if not sales_df.empty:
                # Sort months chronologically and remove NaN values
                available_months = sorted([month for month in sales_df['Month Year'].unique() 
                                         if pd.notna(month) and month != 'nan'], 
                                        key=lambda x: pd.to_datetime(x))
                if available_months:
                    selected_month = st.selectbox(
                        "Select Month to View Detailed Sales:",
                        available_months
                    )
                    
                    # Filter sales data for selected month (using same criteria as ROAS calculation)
                    month_sales = sales_df[
                        (sales_df['Month Year'] == selected_month) &
                        (sales_df['Status'].str.contains('|'.join(closed_statuses), case=False, na=False)) &
                        (sales_df['Channel'].str.contains('|'.join(paid_search_channels), case=False, na=False))
                    ]
                    
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

if __name__ == "__main__":
    main()