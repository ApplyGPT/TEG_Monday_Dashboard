# Monday.com Streamlit Dashboard

A simple Streamlit dashboard with dummy data designed for embedding in Monday.com.

## Features

- 📊 Interactive charts and visualizations
- 📈 Sales trends and metrics
- 📦 Product performance analysis
- 👥 Team performance tracking
- 📋 Data tables with recent information
- 🎨 Responsive design optimized for embedding

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Running the Dashboard

1. Start the Streamlit app:
```bash
streamlit run app.py
```

2. The dashboard will be available at `http://localhost:8501`

## Embedding in Monday.com

To embed this dashboard in Monday.com:

1. **Deploy the dashboard** to a hosting service (e.g., Streamlit Cloud, Heroku, or your own server)

2. **Get the public URL** of your deployed dashboard

3. **In Monday.com:**
   - Go to your board
   - Add a new column of type "Link"
   - Or use the "Website" widget in a dashboard view
   - Paste your dashboard URL

4. **Alternative embedding methods:**
   - Use an iframe in a Monday.com dashboard
   - Create a custom widget using Monday.com's API

## Dashboard Components

- **KPIs**: Total Sales, Orders, Revenue, and Team Satisfaction
- **Charts**: Sales trends, product performance, team metrics
- **Data Tables**: Recent sales, product, and team data
- **Responsive Design**: Optimized for various screen sizes

## Customization

The dashboard uses dummy data generated randomly. To customize:

- Modify the `generate_dummy_data()` function in `app.py`
- Replace with real data sources
- Adjust charts and metrics as needed
- Update styling in the CSS section

## Dependencies

- streamlit==1.28.1
- pandas==2.1.3
- plotly==5.17.0
- numpy==1.24.3
