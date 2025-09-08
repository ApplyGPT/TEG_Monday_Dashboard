import streamlit as st
import requests
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# Page configuration
st.set_page_config(
    page_title="Monday.com Sales Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Helper function to format numbers with K format
def format_currency(value):
    """Format currency values with K for thousands"""
    if value >= 1000000:
        return f"${value/1000000:.1f}M"
    elif value >= 1000:
        return f"${value/1000:.1f}K"
    else:
        return f"${value:.0f}"

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

credentials = load_credentials()
API_TOKEN = credentials['api_token']
SALES_BOARD_ID = credentials['sales_board_id']

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_sales_data():
    """Get sales data from Monday.com Sales boards with caching - fetches ALL records using pagination"""
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json",
    }
    
    # Based on the Google Sheets data, we only need the main Sales JEET COPY board
    board_id = SALES_BOARD_ID
    
    all_items = []
    
    with st.spinner(f"🔄 Fetching ALL sales data from Monday.com Sales JEET COPY board..."):
        cursor = None
        limit = 500  # Maximum limit allowed by Monday.com API
        page_count = 0
        
        while True:
            page_count += 1
            
            # GraphQL query to fetch sales data from specific board with pagination
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
            
            try:
                response = requests.post(url, json={"query": query}, headers=headers, timeout=120)  # Increased timeout
                response.raise_for_status()
                data = response.json()
                
                # Check for API errors
                if "errors" in data and data["errors"]:
                    st.error(f"API Error: {data['errors']}")
                    break
                
                if "data" in data and "boards" in data["data"] and data["data"]["boards"]:
                    board_info = data["data"]["boards"][0]
                    items_page = board_info.get("items_page", {})
                    items = items_page.get("items", [])
                    cursor = items_page.get("cursor")
                    
                    all_items.extend(items)
                    
                    # Show progress
                    #st.write(f"📊 Fetched {len(all_items)} records so far...")
                    
                    # If no cursor or fewer items than limit, we've reached the end
                    if not cursor or len(items) < limit:
                        break
                else:
                    st.error("No board data found")
                    break
                    
            except requests.exceptions.Timeout:
                st.error("Request timed out. Please try again.")
                return None
            except requests.exceptions.RequestException as e:
                st.error(f"Error fetching data: {str(e)}")
                return None
            except Exception as e:
                st.error(f"Unexpected error: {str(e)}")
                return None
    
    # Return the data in the expected format
    return {
        "data": {
            "boards": [{
                "items_page": {
                    "items": all_items
                }
            }]
        }
    }

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
            "Client Type": ""
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
    # 1. Lead Status = "Closed"
    closed_status_mask = df['Lead Status'] == 'Closed'
            
    # 2. Contract Amount >= 0 OR Numbers3 >= 0 OR both are null (Amount Paid or Contract Value)
    contract_amount_mask = df['Contract Amount'] >= 0
    numbers3_mask = df['Numbers3'] >= 0
    both_null_mask = df['Contract Amount'].isna() & df['Numbers3'].isna()
            
    # Combine filters: Closed AND (Contract Amount >= 0 OR Numbers3 >= 0 OR both are null)
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
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Settings")
        st.info(f"Fetching from: {SALES_BOARD_ID}")
        st.info(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Refresh button
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
    
    # Load and process data
    with st.spinner("Loading sales data from Monday.com..."):
        data = get_sales_data()
        
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
    
    # 1. Sales by Month (2025)
    st.subheader(f"Sales by Month (2025)")
    if not df_2025.empty:
        # Extract year and month for 2025 data
        df_2025['Close Date'] = pd.to_datetime(df_2025['Close Date'], errors='coerce')
        df_2025['Year'] = df_2025['Close Date'].dt.year
        df_2025['Month'] = df_2025['Close Date'].dt.month
        df_2025['Month_Name'] = df_2025['Close Date'].dt.strftime('%B')
        
        monthly_sales = df_2025.groupby(['Month', 'Month_Name'])['Total Value'].sum().reset_index()
        monthly_sales = monthly_sales.sort_values('Month')
        
        fig_monthly = px.bar(
            monthly_sales,
            x='Month_Name',
            y='Total Value',
            labels={'Total Value': 'Revenue ($)', 'Month_Name': 'Month'}
        )
        
        # Add numerical amounts above each bar
        fig_monthly.update_traces(
            texttemplate='<b>$%{y:,.2f}</b>',  # Bold text
            textposition='outside',
            textfont=dict(size=16, color='black')  # Larger text
        )
        
        fig_monthly.update_layout(
            height=500, 
            showlegend=False,
            xaxis_title='Month',
            yaxis_title='Revenue ($)',
            font=dict(size=14)  # Larger font for all text
        )
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
            'Jeet Sangamnerkar': '#a358df',         # Teal
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
    
    # 5. Comparison of Revenue by Category by Month
    st.subheader("Comparison of Revenue by Category by Month")
    
    # Year selector for category chart - default to 2025
    available_years_category = sorted([int(year) for year in df_filtered['Year'].unique() if pd.notna(year)])
    default_year_index_category = available_years_category.index(2025) if 2025 in available_years_category else 0
    selected_year_category = st.selectbox("Select Year for Category Analysis:", available_years_category, index=default_year_index_category, key="category_year")
    
    df_category_year = df_filtered[df_filtered['Year'] == selected_year_category]
    
    if not df_category_year.empty:
        category_monthly = df_category_year.groupby(['Month', 'Month_Name', 'Client Type'])['Total Value'].sum().reset_index()
        category_monthly = category_monthly.sort_values('Month')
        
        # Create proper month order for x-axis
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December']
        
        # Filter to only include months that exist in the data
        available_months = category_monthly['Month_Name'].unique()
        month_order_filtered = [month for month in month_order if month in available_months]
        
        # Create grouped bar chart by category
        fig_category = go.Figure()
        
        categories = sorted(category_monthly['Client Type'].unique())
        
        # Hardcoded colors for specific categories based on the image
        category_colors = {
            'ESTABLISHED HIGH-END DESIGNER': '#fdab3d',  # Orange
            'EMERGING DESIGNER': '#00c875',               # Green
            'OTHER': '#c4c4c4',                          # Grey
            'SOMEWHERE IN THE MIDDLE': '#579bfc',        # Blue
            'EXISTING': '#df2f4a'                        # Red
        }
        
        # Use strong colors for any other categories
        all_colors = ['#96CEB4', '#FFEAA7', '#98D8C8', '#F7DC6F', '#FF8A80', '#26A69A', '#42A5F5', '#66BB6A', '#FFCA28']
        
        for i, category in enumerate(categories):
            category_data = category_monthly[category_monthly['Client Type'] == category]
            
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
            bargap=0.0,
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

if __name__ == "__main__":
    main()