import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
import plotly.express as px
import plotly.graph_objects as go
import requests
import json
import os
import sqlite3
import calendar

# Page configuration
st.set_page_config(
    page_title="Burki Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for strong colors
st.markdown("""
<style>
    .main {
        padding: 1rem;
    }
    .stMetric {
        background-color: #f8f9fa;
        padding: 0.5rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
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
    /* Hide QuickBooks and SignNow pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"] {
        display: none !important;
    }
    @media (max-width: 768px) {
        .embed-header {
            font-size: 1.2rem;
        }
    }
</style>
""", unsafe_allow_html=True)

def load_calendly_credentials():
    """Load Calendly credentials from Streamlit secrets"""
    try:
        if 'calendly' not in st.secrets:
            st.error("Calendly configuration not found in secrets.toml. Please check your configuration.")
            st.stop()
        
        calendly_config = st.secrets['calendly']
        
        if 'calendly_api_key' not in calendly_config:
            st.error("Calendly API key not found in secrets.toml. Please add your Calendly API key.")
            st.stop()
            
        return {
            'api_key': calendly_config['calendly_api_key']
        }
    except Exception as e:
        st.error(f"Error reading secrets: {str(e)}")
        st.stop()

def load_calendly_data_from_db():
    """Load Calendly data from SQLite database"""
    CALENDLY_DB_PATH = "calendly_data.db"
    
    try:
        conn = sqlite3.connect(CALENDLY_DB_PATH)
        cursor = conn.cursor()
        
        # Check if table exists and has data
        cursor.execute("SELECT COUNT(*) FROM calendly_events")
        count = cursor.fetchone()[0]
        
        if count == 0:
            conn.close()
            return None, "No Calendly data found in database. Please refresh Calendly data first."
        
        # Get all events
        cursor.execute("""
            SELECT uri, name, start_time, end_time, status, event_type, 
                   invitee_name, invitee_email, updated_at
            FROM calendly_events
            ORDER BY start_time DESC
        """)
        
        rows = cursor.fetchall()
        conn.close()
        
        # Convert to DataFrame
        df = pd.DataFrame(rows, columns=[
            'uri', 'name', 'start_time', 'end_time', 'status', 'event_type',
            'invitee_name', 'invitee_email', 'updated_at'
        ])
        
        # Convert timestamps
        df['start_time'] = pd.to_datetime(df['start_time'])
        df['end_time'] = pd.to_datetime(df['end_time'])
        df['updated_at'] = pd.to_datetime(df['updated_at'])
        
        # Add additional columns for analysis
        df['date'] = df['start_time'].dt.date
        df['month'] = df['start_time'].dt.strftime('%B %Y')
        df['week'] = df['start_time'].dt.isocalendar().week
        df['year'] = df['start_time'].dt.year
        df['day_of_week'] = df['start_time'].dt.strftime('%A')
        df['hour'] = df['start_time'].dt.hour
        
        return df, None
        
    except sqlite3.Error as e:
        return None, f"Database error: {str(e)}"
    except Exception as e:
        return None, f"Error loading Calendly data: {str(e)}"

def get_calendly_data():
    """Get Calendly data for Jamie Burki's TEG events"""
    credentials = load_calendly_credentials()
    api_key = credentials['api_key']
    
    headers = {
        'Authorization': f'Bearer {api_key}',
        'Content-Type': 'application/json'
    }
    
    # Get user info
    user_response = requests.get('https://api.calendly.com/users/me', headers=headers)
    if user_response.status_code != 200:
        st.error(f"Failed to get user info: {user_response.status_code}")
        return {'events': [], 'debug_info': {'error': f'User API error: {user_response.status_code}'}}
    
    user_data = user_response.json()
    user_uri = user_data['resource']['uri']
    user_name = user_data['resource']['name']
    
    # Get event types
    event_types_response = requests.get('https://api.calendly.com/event_types', 
                                      headers=headers,
                                      params={
                                          'user': user_uri,
                                          'count': 100,
                                          'active': True
                                      })
    
    if event_types_response.status_code != 200:
        st.error(f"Failed to get event types: {event_types_response.status_code}")
        return {'events': [], 'debug_info': {'error': f'Event types API error: {event_types_response.status_code}'}}
    
    event_types_data = event_types_response.json()
    event_types = event_types_data.get('collection', [])
    
    # Find TEG-related events
    teg_events = []
    for event_type in event_types:
        name = event_type.get('name', '').lower()
        if 'teg' in name:
            teg_events.append(event_type)
    
    if not teg_events:
        st.warning("❌ No TEG-related events found!")
        return {'events': [], 'debug_info': {'available_event_types': [et.get('name') for et in event_types]}}
    
    # Use the first TEG event found
    selected_event = teg_events[0]
    event_uri = selected_event.get('uri')
    event_name = selected_event.get('name')
    
    # Get scheduled events with pagination to get ALL events
    min_start_time = "2025-01-01T00:00:00.000000Z"
    max_start_time = "2025-10-23T23:59:59.999999Z"
    
    all_events = []
    next_page_token = None
    page_count = 0
    
    while True:
        page_count += 1
        params = {
            'user': user_uri,
            'min_start_time': min_start_time,
            'max_start_time': max_start_time,
            'count': 100
        }
        
        if next_page_token:
            params['page_token'] = next_page_token
        
        events_response = requests.get('https://api.calendly.com/scheduled_events', 
                                     headers=headers,
                                     params=params)
        
        if events_response.status_code != 200:
            st.error(f"Failed to get scheduled events: {events_response.status_code}")
            return {'events': [], 'debug_info': {'error': f'Scheduled events API error: {events_response.status_code}'}}
        
        events_data = events_response.json()
        page_events = events_data.get('collection', [])
        all_events.extend(page_events)
        
        # Check for next page
        pagination = events_data.get('pagination', {})
        next_page_token = pagination.get('next_page_token')
        
        if not next_page_token:
            break
        
        # Safety limit to prevent infinite loops
        if page_count > 50:
            break
    
    # Filter events by selected event type
    filtered_events = []
    for event in all_events:
        if event.get('event_type') == event_uri:
            filtered_events.append(event)
    
    return {
        'events': filtered_events,
        'debug_info': {
            'event_name': event_name,
            'event_uri': event_uri,
            'total_events': len(all_events),
            'filtered_events': len(filtered_events),
            'user_name': user_name
        }
    }

def format_calendly_data(events):
    """Format Calendly events data for analysis"""
    if not events:
        return pd.DataFrame()
    
    records = []
    for event in events:
        start_time = event.get('start_time')
        end_time = event.get('end_time')
        status = event.get('status', '')
        
        # Get invitee information
        invitees = event.get('invitees', [])
        invitee_name = ""
        if invitees:
            invitee = invitees[0]
            invitee_name = invitee.get('name', '')
        
        if start_time:
            start_dt = pd.to_datetime(start_time)
            records.append({
                'event_name': event.get('name', ''),
                'start_time': start_dt,
                'end_time': pd.to_datetime(end_time) if end_time else None,
                'status': status,
                'invitee_name': invitee_name,
                'date': start_dt.date(),
                'month': start_dt.strftime('%B %Y'),
                'week': start_dt.isocalendar()[1],
                'year': start_dt.year,
                'day_of_week': start_dt.strftime('%A'),
                'hour': start_dt.hour,
                'event_type': event.get('event_type', ''),
                'uri': event.get('uri', '')
            })
    
    return pd.DataFrame(records)

def create_daily_chart(df):
    """Create daily calls chart"""
    if df.empty:
        return None
    
    daily_counts = df.groupby('date').size().reset_index(name='count')
    daily_counts = daily_counts.sort_values('date')
    
    fig = px.bar(
        daily_counts,
        x='date',
        y='count',
        title='Calls by Day',
        labels={'count': 'Number of Calls', 'date': 'Date'}
    )
    
    fig.update_layout(
        xaxis_title="Date",
        yaxis_title="Number of Calls",
        height=600,
        showlegend=False
    )
    
    fig.update_traces(
        marker_color='#1f77b4',
        text=daily_counts['count'],
        textposition='outside'
    )
    
    return fig

def create_weekly_chart(df):
    """Create weekly calls chart"""
    if df.empty:
        return None
    
    # Create proper week labels with date ranges
    df_copy = df.copy()
    df_copy['week_start'] = df_copy['start_time'].dt.to_period('W').dt.start_time
    df_copy['week_end'] = df_copy['week_start'] + pd.Timedelta(days=6)
    
    weekly_counts = df_copy.groupby(['week_start', 'week_end']).size().reset_index(name='count')
    
    # Create readable week labels like "Oct 12 - Oct 18"
    weekly_counts['week_label'] = weekly_counts.apply(
        lambda row: f"{row['week_start'].strftime('%b %d')} - {row['week_end'].strftime('%b %d')}", 
        axis=1
    )
    
    # Sort chronologically
    weekly_counts = weekly_counts.sort_values('week_start')
    
    fig = px.bar(
        weekly_counts,
        x='week_label',
        y='count',
        title='Calls by Week',
        labels={'count': 'Number of Calls', 'week_label': 'Week'}
    )
    
    fig.update_layout(
        xaxis_title="Week",
        yaxis_title="Number of Calls",
        height=600,
        showlegend=False
    )
    
    fig.update_traces(
        marker_color='#2E8B57',
        text=weekly_counts['count'],
        textposition='outside'
    )
    
    fig.update_xaxes(tickangle=45)
    
    return fig

def create_monthly_chart(df):
    """Create monthly calls chart"""
    if df.empty:
        return None
    
    # Create proper month sorting by using datetime
    df_copy = df.copy()
    df_copy['month_sort'] = df_copy['start_time'].dt.to_period('M')
    
    monthly_counts = df_copy.groupby(['month_sort', 'month']).size().reset_index(name='count')
    
    # Sort chronologically by month_sort
    monthly_counts = monthly_counts.sort_values('month_sort')
    
    fig = px.bar(
        monthly_counts,
        x='month',
        y='count',
        title='Calls by Month',
        labels={'count': 'Number of Calls', 'month': 'Month'}
    )
    
    fig.update_layout(
        xaxis_title="Month",
        yaxis_title="Number of Calls",
        height=600,
        showlegend=False
    )
    
    fig.update_traces(
        marker_color='#FF6B6B',
        text=monthly_counts['count'],
        textposition='outside'
    )
    
    fig.update_xaxes(tickangle=45)
    
    return fig

def create_monthly_calendar_view(df, selected_month):
    """Create a monthly calendar view showing call counts by day"""
    if df.empty:
        return None
    
    # Filter data for the selected month
    df_filtered = df[
        (df['start_time'].dt.year == selected_month.year) & 
        (df['start_time'].dt.month == selected_month.month)
    ].copy()
    
    if df_filtered.empty:
        return None
    
    # Count calls for each day
    daily_counts = df_filtered.groupby(df_filtered['start_time'].dt.date).size().to_dict()
    
    return daily_counts

def display_calendar_grid(daily_counts, selected_month):
    """Display the calendar grid with call counts"""
    if not daily_counts:
        st.info("No calls scheduled for this month.")
        return
    
    month_name = selected_month.strftime('%B %Y')
    
    # Create calendar grid
    cal = calendar.monthcalendar(selected_month.year, selected_month.month)
    col_names = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    
    # Header row
    cols = st.columns(7)
    for i, col_name in enumerate(col_names):
        with cols[i]:
            st.markdown(f"**{col_name}**")
    
    # Calendar rows
    for week in cal:
        cols = st.columns(7)
        for i, day in enumerate(week):
            with cols[i]:
                if day == 0:
                    st.write("")  # Empty cell for days not in this month
                else:
                    current_date = selected_month.replace(day=day)
                    call_count = daily_counts.get(current_date.date(), 0)
                    
                    if call_count > 0:
                        bg_color = "#2E8B57"  # Sea Green
                        text_color = "#FFFFFF"
                        
                        st.markdown(f"""
                        <div style="
                            background-color: {bg_color};
                            color: {text_color};
                            padding: 8px;
                            border-radius: 8px;
                            text-align: center;
                            margin: 2px;
                            font-weight: bold;
                        ">
                            <strong>{day}</strong><br>
                            <small>{call_count} call{'s' if call_count != 1 else ''}</small>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div style="
                            background-color: #f0f0f0;
                            color: #666666;
                            padding: 8px;
                            border-radius: 8px;
                            text-align: center;
                            margin: 2px;
                        ">
                            <strong>{day}</strong><br>
                            <small>0 calls</small>
                        </div>
                        """, unsafe_allow_html=True)

def main():
    """Main application function"""
    # Header
    st.markdown('<div class="embed-header">📊 BURKI DASHBOARD</div>', unsafe_allow_html=True)
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Settings")
        st.info(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Refresh button
        if st.button("🔄 Refresh Data"):
            st.rerun()
    
    # Load Calendly data from database
    with st.spinner("Loading Calendly data from database..."):
        try:
            df, error = load_calendly_data_from_db()
            
            if error:
                st.error(f"Error loading data: {error}")
                st.info("💡 **Tip:** Go to the Database Refresh page and click 'Refresh All Calendly Data' to populate the database.")
                return
            
            if df is None or df.empty:
                st.warning("No Calendly data found in database.")
                st.info("💡 **Tip:** Go to the Database Refresh page and click 'Refresh All Calendly Data' to populate the database.")
                return
            
            # Skip event information section
            
            # Filter for 2025 data
            df_2025 = df[df['year'] == 2025].copy()
            
            if df_2025.empty:
                st.warning("No events found for 2025.")
                return
            
            # Charts section
            st.markdown("---")
            
            # Create tabs for different views
            tab1, tab2, tab3 = st.tabs(["📅 Daily View", "📊 Weekly View", "📊 Monthly View"])
            
            with tab1:
                # Add month selector for the calendar view
                
                # Get available months from data
                if not df_2025.empty:
                    available_months = df_2025['start_time'].dt.to_period('M').unique()
                    month_strings = [str(month) for month in sorted(available_months)]
                    
                    # Create month selector
                    selected_month_str = st.selectbox(
                        "Select Month:",
                        options=month_strings,
                        index=len(month_strings) - 1 if month_strings else 0,  # Default to most recent month
                        help="Select a month to view calls in calendar format"
                    )
                    
                    # Parse selected month
                    if selected_month_str:
                        selected_month = pd.to_datetime(selected_month_str).to_pydatetime().replace(day=1)
                        
                        # Create and display calendar view
                        daily_counts = create_monthly_calendar_view(df_2025, selected_month)
                        if daily_counts:
                            display_calendar_grid(daily_counts, selected_month)
                        else:
                            st.info(f"No calls scheduled for {selected_month.strftime('%B %Y')}")
            
            with tab2:
                weekly_fig = create_weekly_chart(df_2025)
                if weekly_fig:
                    st.plotly_chart(weekly_fig, use_container_width=True)
                else:
                    st.info("No weekly data available")
            
            with tab3:
                monthly_fig = create_monthly_chart(df_2025)
                if monthly_fig:
                    st.plotly_chart(monthly_fig, use_container_width=True)
                else:
                    st.info("No monthly data available")
            
            # Skip detailed data table
                
        except Exception as e:
            st.error(f"Error loading Calendly data: {str(e)}")
            st.info("Please check your Calendly API key in secrets.toml")

if __name__ == "__main__":
    main()
