# TEG Monday Dashboard

A comprehensive multi-page Streamlit dashboard for analyzing Monday.com data across multiple boards.

## Features

### 📊 ADS DASHBOARD (Default Page)
- Google Ads attribution data
- Ad spend tracking by month
- Campaign performance metrics
- Data export functionality
- Year filtering and analysis

### 📈 SALES DASHBOARD (`/sales_dashboard`)
- Sales performance analytics
- Revenue tracking by month/year
- Salesman and category analysis
- YTD and MTD metrics
- Interactive charts and filters

## Setup

1. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure Credentials**
   Create a `.streamlit/secrets.toml` file in the root directory with your API credentials:
   ```toml
   [monday]
   api_token = "your_monday_api_token_here"
   sales_board_id = your_sales_board_id_here
   ads_board_id = your_ads_board_id_here
   
   [openai]
   api_key = "your_openai_api_key_here"
   ``` DASHBOARD (`/sales_dashboard`)
- Sales performance analytics
- Revenue tracking by month/year
- Salesman and category analysis
- YTD and MTD metrics
- Interactive charts and filters

## Setup

1. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure Credentials**
   Create a `.streamlit/secrets.toml` file in the root directory with your API credentials:
   ```toml
   [monday]
   api_token = "your_monday_api_token_here"
   sales_board_id = your_sales_board_id_here
   ads_board_id = your_ads_board_id_here
   
   [openai]
   api_key = "your_openai_api_key_here"
   ```

3. **Run the Application**
   ```bash
   streamlit run ads_dashboard.py
   ```

## Access URLs

Once the application is running, you can access the different dashboards at:

- **ADS DASHBOARD (Default)**: `http://localhost:8501`
- **SALES DASHBOARD**: `http://localhost:8501/sales_dashboard`

## File Structure

```
TEG_Monday_Dashboard/
├── ads_dashboard.py           # ADS DASHBOARD (default page)
├── pages/
│   └── sales_dashboard.py     # SALES DASHBOARD
├── .streamlit/
│   └── secrets.toml          # API credentials (create this)
├── requirements.txt          # Python dependencies
└── README.md                # This file
```

## Navigation

- The ADS DASHBOARD is the default landing page
- Use the navigation sidebar to switch to the SALES DASHBOARD
- Each dashboard has its own refresh button to update data
- All dashboards share the same secrets configuration

## Features

- **Real-time Data**: Connects directly to Monday.com API
- **Caching**: 5-minute cache for better performance
- **Responsive Design**: Works on desktop and mobile
- **Data Export**: Download data as CSV files
- **Interactive Charts**: Built with Plotly for rich visualizations
- **Multi-page Structure**: Clean separation of concerns

## Troubleshooting

1. **API Token Issues**: Make sure your Monday.com API token has the correct permissions
2. **Board Access**: Verify that your API token can access the specified board IDs
3. **Data Loading**: Check the console for any error messages during data fetching
4. **Port Conflicts**: If port 8501 is busy, Streamlit will automatically use the next available port

## Dependencies

- streamlit
- pandas
- plotly
- requests
- datetime

See `secrets.toml` for specific versions.
