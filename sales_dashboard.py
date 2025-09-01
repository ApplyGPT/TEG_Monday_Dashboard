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

# Monday.com API settings from credentials.txt
def load_credentials():
    """Load credentials from credentials.txt file"""
    credentials = {}
    try:
        with open('credentials.txt', 'r') as f:
            for line in f:
                line = line.strip()
                if '=' in line and not line.startswith('[') and not line.startswith('#'):
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")  # Remove quotes
                    if key == 'api_token':
                        credentials['api_token'] = value
                    elif key == 'sales_board_id':
                        credentials['sales_board_id'] = int(value)
    except FileNotFoundError:
        st.error("credentials.txt file not found. Please create it with your Monday.com API credentials.")
        st.stop()
    except Exception as e:
        st.error(f"Error reading credentials: {str(e)}")
        st.stop()
    
    return credentials

credentials = load_credentials()
API_TOKEN = credentials['api_token']
SALES_BOARD_ID = credentials['sales_board_id']

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_sales_data():
    """Get sales data from Monday.com board with caching"""
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": API_TOKEN,
        "Content-Type": "application/json",
    }
    
    # GraphQL query to fetch sales data with all required columns
    query = f"""
    query {{
        boards(ids: {SALES_BOARD_ID}) {{
            items_page(limit: 500) {{
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
        response = requests.post(url, json={"query": query}, headers=headers, timeout=60)  # Increased timeout
        response.raise_for_status()
        return response.json()
    except requests.exceptions.Timeout:
        st.error("Request timed out. Please try again.")
        return None
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching data: {str(e)}")
        return None
    except Exception as e:
        st.error(f"Unexpected error: {str(e)}")
        return None

def process_sales_data(data):
    """Convert Monday.com sales data to pandas DataFrame and process it"""
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
            "Close Date": "",
            "Lead Status": "",
            "Amount Paid": "",
            "Contract Value": "",
            "Assigned": "",
            "Client Type": ""
        }
        
        # Extract column values - map by specific column IDs
        for col_val in item.get("column_values", []):
            col_id = col_val.get("id", "")
            col_type = col_val.get("type", "")
            text = col_val.get("text", "")
            value = col_val.get("value", "")
            
            # Store all column data for debugging
            if col_id not in record:
                record[col_id] = text if text else ""
            
            # Map columns by specific column IDs
            if col_id == "color_mknxd1j2":  # Lead Status
                record["Lead Status"] = text if text else ""
            elif col_id == "contract_amt":  # Contract Value
                record["Contract Value"] = text if text else ""
            elif col_id == "person":  # Assigned (Salesman)
                record["Assigned"] = text if text else ""
            elif col_id == "text4__1":  # Client Type (Category)
                record["Client Type"] = text if text else ""
            
            # Store all date columns for later processing
            elif col_id == "date_mktq7npm":
                record["date_mktq7npm"] = text if text else ""
            elif col_id == "date_mktqx5me":
                record["date_mktqx5me"] = text if text else ""
            elif col_id == "date_qualified":
                record["date_qualified"] = text if text else ""
            elif col_id == "date3":
                record["date3"] = text if text else ""
            elif col_id == "date7":
                record["date7"] = text if text else ""
            elif col_id == "contract_invoice_sent":
                record["contract_invoice_sent"] = text if text else ""
        
        # Determine the best close date from multiple date columns
        # Priority order: date_qualified (has 2023), date7 (has 2023), date_mktqx5me (has 2023), date3 (has 2023), date_mktq7npm, contract_invoice_sent
        close_date = ""
        date_priority = ["date_qualified", "date7", "date_mktqx5me", "date3", "date_mktq7npm", "contract_invoice_sent"]
        
        for date_col in date_priority:
            if record.get(date_col, ""):
                close_date = record[date_col]
                break
        
        record["Close Date"] = close_date
        
        records.append(record)
    
    df = pd.DataFrame(records)
    
    # Filter for closed deals only - be more flexible with status matching
    if not df.empty:
        closed_mask = (
            df['Lead Status'].str.contains('Closed', case=False, na=False) |
            df['Lead Status'].str.contains('Won', case=False, na=False) |
            df['Lead Status'].str.contains('Completed', case=False, na=False) |
            df['Lead Status'].str.contains('Done', case=False, na=False) |
            df['Lead Status'].str.contains('Win', case=False, na=False) |
            df['Lead Status'].str.contains('Paid', case=False, na=False) |
            df['Lead Status'].str.contains('Payment Posted', case=False, na=False)
        )
        df = df[closed_mask]
    
    # Process dates
    df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
    df = df.dropna(subset=['Close Date'])
    
    # Extract year and month
    df['Year'] = df['Close Date'].dt.year
    df['Month'] = df['Close Date'].dt.month
    df['Month_Name'] = df['Close Date'].dt.strftime('%B')
    
    # Process monetary values
    df['Amount Paid'] = pd.to_numeric(df['Amount Paid'].str.replace('$', '').str.replace(',', ''), errors='coerce')
    df['Contract Value'] = pd.to_numeric(df['Contract Value'].str.replace('$', '').str.replace(',', ''), errors='coerce')
    
    # Calculate total value
    df['Total Value'] = df['Amount Paid'].fillna(0) + df['Contract Value'].fillna(0)
    
    # Remove rows with no monetary value
    df = df[df['Total Value'] > 0]
    
    return df

def main():
    """Main application function"""
    # Header
    st.title("📈 Sales Dashboard")
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Settings")
        st.info(f"Sales Board ID: {SALES_BOARD_ID}")
        st.info(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Refresh button
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
    
    # Load and process data
    with st.spinner("Loading sales data from Monday.com..."):
        data = get_sales_data()
        

        
        df = process_sales_data(data)
    
    if df.empty:
        st.warning("No closed sales records found. Please check your data and filters.")
        return
    

    
    # Current year and month
    current_year = datetime.now().year
    current_month = datetime.now().month
    
    # Filter data for current year and month
    df_current_year = df[df['Year'] == current_year]
    df_current_month = df[(df['Year'] == current_year) & (df['Month'] == current_month)]
    
    # 1. Sales by Month (Current Year)
    st.subheader(f"Sales by Month ({current_year})")
    if not df_current_year.empty:
        monthly_sales = df_current_year.groupby(['Month', 'Month_Name'])['Total Value'].sum().reset_index()
        monthly_sales = monthly_sales.sort_values('Month')
        
        fig_monthly = px.bar(
            monthly_sales,
            x='Month_Name',
            y='Total Value',
            labels={'Total Value': 'Revenue ($)', 'Month_Name': 'Month'}
        )
        fig_monthly.update_layout(height=500, showlegend=False)
        st.plotly_chart(fig_monthly, use_container_width=True)
    else:
        st.info("No sales data available for the current year.")
    
    # 2. Sales by Year
    st.subheader("Sales by Year")
    yearly_sales = df.groupby('Year')['Total Value'].sum().reset_index()
    yearly_sales = yearly_sales.sort_values('Year')
    
    fig_yearly = px.bar(
        yearly_sales,
        x='Year',
        y='Total Value',
        labels={'Total Value': 'Revenue ($)', 'Year': 'Year'}
    )
    fig_yearly.update_layout(height=500, showlegend=False)
    st.plotly_chart(fig_yearly, use_container_width=True)
    
    # 3. Comparison of Revenue by Year by Month
    st.subheader("Comparison of Revenue by Year by Month")
    
    # Create pivot table for grouped bar chart
    monthly_yearly = df.groupby(['Year', 'Month', 'Month_Name'])['Total Value'].sum().reset_index()
    monthly_yearly = monthly_yearly.sort_values(['Year', 'Month'])
    
    # Create a proper month order for x-axis
    month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                   'July', 'August', 'September', 'October', 'November', 'December']
    
    # Filter to only include months that exist in the data
    available_months = monthly_yearly['Month_Name'].unique()
    month_order_filtered = [month for month in month_order if month in available_months]
    
    # Update the figure to use the proper month order
    fig_grouped = go.Figure()
    
    years = sorted(monthly_yearly['Year'].unique())
    # Use stronger colors
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F']
    
    for i, year in enumerate(years):
        year_data = monthly_yearly[monthly_yearly['Year'] == year]
        fig_grouped.add_trace(go.Bar(
            name=str(year),
            x=year_data['Month_Name'],
            y=year_data['Total Value'],
            marker_color=colors[i]
        ))
    
    fig_grouped.update_layout(
        barmode='group',  # Changed from 'stack' to 'group'
        xaxis_title='Month',
        yaxis_title='Revenue ($)',
        height=500,
        bargap=0.15,  # Small gap between groups of months
        bargroupgap=0.0,  # No gap between bars of the same group
        xaxis=dict(
            categoryorder='array',
            categoryarray=month_order_filtered
        )
    )
    

    st.plotly_chart(fig_grouped, use_container_width=True)
    
    # 4. Comparison of Revenue by Salesman by Month
    st.subheader("Comparison of Revenue by Salesman by Month")
    
    # Year selector for salesman chart
    available_years = sorted(df['Year'].unique())
    selected_year_salesman = st.selectbox("Select Year for Salesman Analysis:", available_years, key="salesman_year")
    
    df_salesman_year = df[df['Year'] == selected_year_salesman]
    
    if not df_salesman_year.empty:
        salesman_monthly = df_salesman_year.groupby(['Month', 'Month_Name', 'Assigned'])['Total Value'].sum().reset_index()
        salesman_monthly = salesman_monthly.sort_values('Month')
        
        # Create proper month order for x-axis
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December']
        
        # Filter to only include months that exist in the data
        available_months = salesman_monthly['Month_Name'].unique()
        month_order_filtered = [month for month in month_order if month in available_months]
        
        # Create stacked bar chart by salesman
        fig_salesman = go.Figure()
        
        salesmen = sorted(salesman_monthly['Assigned'].unique())
        # Use stronger colors
        all_colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F', 
                     '#FF8A80', '#26A69A', '#42A5F5', '#66BB6A', '#FFCA28', '#AB47BC', '#26C6DA', '#FF7043']
        
        for i, salesman in enumerate(salesmen):
            salesman_data = salesman_monthly[salesman_monthly['Assigned'] == salesman]
            # Use modulo to cycle through colors if we have more salesmen than colors
            color = all_colors[i % len(all_colors)]
            # Handle empty salesmen
            salesman_name = salesman if salesman and salesman.strip() else 'Unassigned'
            fig_salesman.add_trace(go.Bar(
                name=salesman_name,
                x=salesman_data['Month_Name'],
                y=salesman_data['Total Value'],
                marker_color=color
            ))
        
        fig_salesman.update_layout(
            barmode='group',  # Changed from 'stack' to 'group'
            xaxis_title='Month',
            yaxis_title='Revenue ($)',
            height=500,
            bargap=0.15,  # Small gap between groups of months
            bargroupgap=0.0,  # No gap between bars of the same group
            showlegend=True,  # Always show legend
            xaxis=dict(
                categoryorder='array',
                categoryarray=month_order_filtered
            )
        )
        st.plotly_chart(fig_salesman, use_container_width=True)
    else:
        st.info(f"No sales data available for {selected_year_salesman}.")
    
    # 5. Comparison of Revenue by Category by Month
    st.subheader("Comparison of Revenue by Category by Month")
    
    # Year selector for category chart
    selected_year_category = st.selectbox("Select Year for Category Analysis:", available_years, key="category_year")
    
    df_category_year = df[df['Year'] == selected_year_category]
    
    if not df_category_year.empty:
        category_monthly = df_category_year.groupby(['Month', 'Month_Name', 'Client Type'])['Total Value'].sum().reset_index()
        category_monthly = category_monthly.sort_values('Month')
        
        # Create proper month order for x-axis
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December']
        
        # Filter to only include months that exist in the data
        available_months = category_monthly['Month_Name'].unique()
        month_order_filtered = [month for month in month_order if month in available_months]
        
        # Create stacked bar chart by category
        fig_category = go.Figure()
        
        categories = sorted(category_monthly['Client Type'].unique())
        # Use stronger colors
        all_colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F', 
                     '#FF8A80', '#26A69A', '#42A5F5', '#66BB6A', '#FFCA28', '#AB47BC', '#26C6DA', '#FF7043']
        
        for i, category in enumerate(categories):
            category_data = category_monthly[category_monthly['Client Type'] == category]
            # Use modulo to cycle through colors if we have more categories than colors
            color = all_colors[i % len(all_colors)]
            # Handle empty categories - ensure we have a proper name for the legend
            category_name = category if category and category.strip() else 'Uncategorized'
            fig_category.add_trace(go.Bar(
                name=category_name,
                x=category_data['Month_Name'],
                y=category_data['Total Value'],
                marker_color=color
            ))
        
        fig_category.update_layout(
            barmode='group',  # Changed from 'stack' to 'group'
            xaxis_title='Month',
            yaxis_title='Revenue ($)',
            height=500,
            bargap=0.15,  # Small gap between groups of months
            bargroupgap=0.0,  # No gap between bars of the same group
            showlegend=True,  # Always show legend
            xaxis=dict(
                categoryorder='array',
                categoryarray=month_order_filtered
            )
        )
        st.plotly_chart(fig_category, use_container_width=True)
    else:
        st.info(f"No sales data available for {selected_year_category}.")

if __name__ == "__main__":
    main()