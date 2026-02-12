# TEG Monday Dashboard

A comprehensive multi-page Streamlit dashboard for analyzing Monday.com data across multiple boards.

## Overview

This application serves as a central hub for TEG's operations, combining real-time data visualization with automated workflow tools. It integrates data from **Monday.com**, **Calendly**, and **Google Ads** to provide actionable insights for Sales, Marketing, and Operations.

## Key Capabilities

### 📊 Data Visualization & Analytics
- **Multi-Channel Attribution**: Track Google Ads performance and calculate ROAS.
- **Sales Intelligence**: Monitor KPIs, revenue trends, and sales team performance.
- **Lead Tracking**: Visualize lead flow from "New Lead" to "Closed".
- **Call Analytics**: Detailed breakdown of Calendly events for intro calls and design reviews.

### 🛠️ Automation Tools
- **Contract Generation**: Create and send legal agreements via **SignNow**.
- **Deck Creation**: Auto-generate PowerPoint slides for client presentations using **Google Slides/Drive**.
- **Workbook Builder**: Construct detailed Excel workbooks for development packages.
- **Data Sync**: Automated background jobs to keep local databases in sync with external APIs.

## Setup

1. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure Credentials**
   Create a `.streamlit/secrets.toml` file in the root directory with the following structure. Fill in your specific API keys and IDs.

   ```toml
   [monday]
   api_token = "YOUR_MONDAY_API_TOKEN"
   sales_board_id = 123456789
   ads_board_id = 123456789
   new_leads_board_id = 123456789
   discovery_call_board_id = 123456789
   design_review_board_id = 123456789

   [openai]
   api_key = "YOUR_OPENAI_API_KEY"

   [calendly]
   api_key = "YOUR_CALENDLY_API_KEY"
   burki_api_key = "YOUR_BURKI_CALENDLY_KEY" (Optional)

   [signnow]
   client_id = "YOUR_SIGNNOW_CLIENT_ID"
   client_secret = "YOUR_SIGNNOW_CLIENT_SECRET"
   basic_auth_token = "YOUR_BASIC_AUTH_TOKEN"
   
   [google]
   # Service Account Query Parameters (for Sheets/Drive/Slides)
   type = "service_account"
   project_id = "your-project-id"
   private_key_id = "your-private-key-id"
   private_key = "-----BEGIN PRIVATE KEY-----\n..."
   client_email = "your-service-account@..."
   client_id = "..."
   auth_uri = "https://accounts.google.com/o/oauth2/auth"
   token_uri = "https://oauth2.googleapis.com/token"
   auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
   client_x509_cert_url = "..."
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

## Page Breakdown

### Main Dashboard (`ads_dashboard.py`)
- **Description**: The entry point for the application, focusing on Google Ads attribution and ROAS.
- **Key Features**:
  - Displays Google Ad spend and campaign performance.
  - Calculates ROAS (Return on Ad Spend) using sales data.
  - Visualizes data with Plotly charts.
- **Key APIs/IDs**:
  - Monday.com API (`api_token`).
  - Board IDs: `new_leads_board_id`, `sales_board_id`.
- **Dependencies**: `database_utils` (Shared utility for database operations).

### Sales Dashboard (`pages/sales_dashboard.py`)
- **Description**: Comprehensive view of sales team performance.
- **Key Features**:
  - Tracks Year-to-Date (YTD) and Month-to-Date (MTD) sales.
  - Breakdowns by salesman, revenue type, and time periods.
- **Key APIs/IDs**:
  - Board IDs: `sales_board_id`.
- **Dependencies**: `database_utils` (Shared utility for database operations).

### Database Refresh (`pages/database_refresh.py`)
- **Description**: Utility page for syncing external data to local storage.
- **Key Features**:
  - Fetches data from Monday.com and Calendly APIs.
  - Updates local SQLite databases (`monday_data.db` and `calendly_data.db`).
  - Implements retry logic for API stability.
- **Key APIs/IDs**:
  - Monday.com API (`api_token`).
  - Calendly API (`calendly_api_key`).
  - All Board IDs.

### New Leads Check (`pages/new_leads_check.py`)
- **Description**: Monitoring tool for incoming leads.
- **Key Features**:
  - Visualizes new leads with Calendar, Daily, and Weekly views.
  - Checks `monday_data.db` for recent entries.
- **Dependencies**: `database_utils` (Shared utility for database operations).

### Intro Call Dashboard (`pages/intro_call_dashboard.py`)
- **Description**: Analytics for introductory calls.
- **Key Features**:
  - Tracks "Burki" and "Intro Call with TEG" events.
  - Displays call volume by source and time period.
- **Key APIs/IDs**:
  - Calendly API (`calendly_api_key`, `calendly_burki_api_key`).

### Burki Dashboard (`pages/burki_dashboard.py`)
- **Description**: Specialized dashboard for Jamie Burki's calls.
- **Key Features**:
  - Filters for "TEG - Let's Chat" events.
  - Provides focused metrics for a specific team member.
- **Key APIs/IDs**:
  - Calendly API (`calendly_burki_api_key`).

### Design Review Dashboard (`pages/design_review_dashboard.py`)
- **Description**: Analytics for design review meetings.
- **Key Features**:
  - Tracks "TEG Introductory Call" and "Jennifer" events.
  - Uses `calendly_data.db` for sourced data.
- **Dependencies**: `database_utils` (Shared utility for database operations).

### SEO Metrics (`pages/seo_metrics.py`)
- **Description**: Interface for SEO performance reporting.
- **Key Features**:
  - Embeds external Looker Studio reports and Google Sheets.
- **Key APIs/IDs**:
  - Looker Studio Embed URL.

### Tools Hub (`pages/tools.py`)
- **Description**: Central navigation hub for employee productivity tools.
- **Key Features**:
  - Provides quick access to SignNow, Workbook Creator, and Deck Creator tools.

### SignNow Form (`pages/signnow_form.py`)
- **Description**: Contract generation and sending tool.
- **Key Features**:
  - Integration with SignNow API for electronic signatures.
  - Generates contracts for Development and Production.
- **Key APIs/IDs**:
  - SignNow Credentials (`client_id`, `client_secret`, etc.).
- **Dependencies**: `signnow_integration` (Custom module for SignNow API interactions).

### Deck Creator (`pages/deck_creator.py`)
- **Description**: Automated slide deck generator.
- **Key Features**:
  - Creates PowerPoint presentations for various fashion categories (Activewear, Bridal, etc.).
  - Uses Google Drive to fetch image assets.
- **Key APIs/IDs**:
  - Google Drive/Slides API (`google_service_account`).

### Workbook Creator (`pages/workbook_creator.py`)
- **Description**: Excel workbook generator for Development Packages.
- **Key Features**:
  - Populates `Copy of TEG 2025 WORKBOOK TEMPLATES.xlsx` with project data.
  - Exports result to PDF using Google APIs.
- **Key APIs/IDs**:
  - Google Service Account.
- **Dependencies**: `google_sheets_uploader` (Custom module for Google Sheets/Drive uploads).

### A La Carte (`pages/a_la_carte.py`)
- **Description**: Specialized workbook creator for custom/a-la-carte items.
- **Key Features**:
  - Handles flexible pricing and item additions for custom projects.
- **Dependencies**: `google_sheets_uploader` (Custom module for Google Sheets/Drive uploads).

## File Structure

```
TEG_Monday_Dashboard/
├── ads_dashboard.py            # Main application entry point (Ads Dashboard)
├── database_utils.py           # Shared database operations module
├── google_sheets_uploader.py   # Module for Google API interactions
├── signnow_integration.py      # Module for SignNow API interactions
├── pages/
│   ├── sales_dashboard.py      # Sales KPIs and analytics
│   ├── database_refresh.py     # Data sync utility
│   ├── new_leads_check.py      # Lead flow monitoring
│   ├── intro_call_dashboard.py # Calendly analytics
│   ├── burki_dashboard.py      # Individual activity dashboard
│   ├── design_review_dashboard.py # Design review analytics
│   ├── seo_metrics.py          # SEO reporting
│   ├── tools.py                # Tools navigation hub
│   ├── signnow_form.py         # Contract generation tool
│   ├── deck_creator.py         # Presentation generator
│   ├── workbook_creator.py     # Workbook generator
│   └── a_la_carte.py           # Custom workbook tool
├── .streamlit/
│   └── secrets.toml            # API credentials (gitignored)
├── requirements.txt            # Python dependencies
└── README.md                   # This documentation
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

1. **API Token Issues**: Make sure your Monday.com API token has the correct permissions (read/write access to boards).
2. **Board Access**: Verify that your API token can access the specified board IDs. If a board is private, the token must belong to a member or admin of that board.
3. **Database Errors**: If charts are empty, go to the **Database Refresh** page and run a manual refresh. Check the logs for connectivity issues with Monday.com or Calendly.
4. **Google Drive/Slides**: For Deck/Workbook creators, ensure the Google Service Account email has "Editor" access to the specific Drive folders and Template files.
5. **Port Conflicts**: If port 8501 is busy, Streamlit will automatically use the next available port. Check the terminal output for the correct URL.

## Dependencies

- streamlit
- pandas
- plotly
- requests
- datetime
- python-pptx (for Deck Creator)
- google-api-python-client (for Google Drive/Slides/Sheets)
- google-auth-httplib2
- google-auth-oauthlib
- openpyxl (for Workbook Creator)
- Pillow (for image processing)

See `secrets.toml` for specific versions.
