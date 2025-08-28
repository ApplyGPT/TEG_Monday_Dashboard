import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from datetime import datetime, timedelta

# Page configuration
st.set_page_config(
    page_title="Monday.com Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better embedding
st.markdown("""
<style>
    .main {
        padding-top: 1rem;
        padding-bottom: 1rem;
    }
    .stApp {
        max-width: 100%;
    }
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.title("📊 Monday.com Dashboard")
st.markdown("---")

# Generate dummy data
@st.cache_data
def generate_dummy_data():
    # Sales data
    dates = pd.date_range(start='2024-01-01', end='2024-12-31', freq='D')
    sales_data = pd.DataFrame({
        'Date': dates,
        'Sales': np.random.normal(1000, 200, len(dates)).cumsum(),
        'Orders': np.random.poisson(50, len(dates)).cumsum(),
        'Revenue': np.random.normal(5000, 1000, len(dates)).cumsum()
    })
    
    # Product performance
    products = ['Product A', 'Product B', 'Product C', 'Product D', 'Product E']
    product_data = pd.DataFrame({
        'Product': products,
        'Sales': np.random.randint(100, 1000, len(products)),
        'Revenue': np.random.randint(5000, 50000, len(products)),
        'Rating': np.random.uniform(3.5, 5.0, len(products))
    })
    
    # Team performance
    team_members = ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve']
    team_data = pd.DataFrame({
        'Member': team_members,
        'Tasks_Completed': np.random.randint(10, 50, len(team_members)),
        'Hours_Worked': np.random.randint(160, 200, len(team_members)),
        'Satisfaction': np.random.uniform(7.0, 10.0, len(team_members))
    })
    
    return sales_data, product_data, team_data

# Load data
sales_data, product_data, team_data = generate_dummy_data()

# Create three columns for KPIs
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        label="Total Sales",
        value=f"${sales_data['Sales'].iloc[-1]:,.0f}",
        delta=f"+{sales_data['Sales'].iloc[-1] - sales_data['Sales'].iloc[-30]:,.0f}"
    )

with col2:
    st.metric(
        label="Total Orders",
        value=f"{sales_data['Orders'].iloc[-1]:,.0f}",
        delta=f"+{sales_data['Orders'].iloc[-1] - sales_data['Orders'].iloc[-30]:,.0f}"
    )

with col3:
    st.metric(
        label="Total Revenue",
        value=f"${sales_data['Revenue'].iloc[-1]:,.0f}",
        delta=f"+{sales_data['Revenue'].iloc[-1] - sales_data['Revenue'].iloc[-30]:,.0f}"
    )

with col4:
    st.metric(
        label="Avg Team Satisfaction",
        value=f"{team_data['Satisfaction'].mean():.1f}/10",
        delta=f"+{team_data['Satisfaction'].mean() - 8.0:.1f}"
    )

st.markdown("---")

# Create two columns for charts
col1, col2 = st.columns(2)

with col1:
    st.subheader("📈 Sales Trend")
    
    # Filter data for last 30 days
    recent_sales = sales_data.tail(30)
    
    fig_sales = px.line(
        recent_sales,
        x='Date',
        y='Sales',
        title='Sales Trend (Last 30 Days)',
        template='plotly_white'
    )
    fig_sales.update_layout(height=300, showlegend=False)
    st.plotly_chart(fig_sales, use_container_width=True)

with col2:
    st.subheader("📦 Product Performance")
    
    fig_products = px.bar(
        product_data,
        x='Product',
        y='Sales',
        title='Sales by Product',
        template='plotly_white',
        color='Sales',
        color_continuous_scale='Blues'
    )
    fig_products.update_layout(height=300, showlegend=False)
    st.plotly_chart(fig_products, use_container_width=True)

# Team performance section
st.markdown("---")
st.subheader("👥 Team Performance")

col1, col2 = st.columns(2)

with col1:
    fig_team_tasks = px.bar(
        team_data,
        x='Member',
        y='Tasks_Completed',
        title='Tasks Completed by Team Member',
        template='plotly_white',
        color='Tasks_Completed',
        color_continuous_scale='Greens'
    )
    fig_team_tasks.update_layout(height=300, showlegend=False)
    st.plotly_chart(fig_team_tasks, use_container_width=True)

with col2:
    fig_team_satisfaction = px.scatter(
        team_data,
        x='Hours_Worked',
        y='Satisfaction',
        size='Tasks_Completed',
        hover_data=['Member'],
        title='Team Satisfaction vs Hours Worked',
        template='plotly_white'
    )
    fig_team_satisfaction.update_layout(height=300)
    st.plotly_chart(fig_team_satisfaction, use_container_width=True)

# Data table section
st.markdown("---")
st.subheader("📋 Recent Data")

tab1, tab2, tab3 = st.tabs(["Sales Data", "Product Data", "Team Data"])

with tab1:
    st.dataframe(sales_data.tail(10), use_container_width=True)

with tab2:
    st.dataframe(product_data, use_container_width=True)

with tab3:
    st.dataframe(team_data, use_container_width=True)

# Footer
st.markdown("---")
st.markdown("*Dashboard created for Monday.com embedding test*")
