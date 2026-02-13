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
import pytz

# California timezone for displaying dates (user's timezone)
CALIFORNIA_TZ = pytz.timezone('America/Los_Angeles')

# Get current year dynamically
CURRENT_YEAR = datetime.now().year

# Intro Call Dashboard: Two Calendly links (for reference, filtering is dynamic)
INTRO_CALL_LINKS = {
    "Burki": "https://calendly.com/jamie-the-evans-group/teg-lets-chat",
    "Intro Call with TEG": "https://calendly.com/d/dv7-542-3nm/intro-call-with-teg",
}

def get_color_palette(sources):
    """Generate a color palette for the given sources dynamically."""
    base_colors = ["#1f77b4", "#2ca02c", "#ff7f0e", "#d62728", "#9467bd", "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"]
    return {source: base_colors[i % len(base_colors)] for i, source in enumerate(sorted(sources))}

# Page configuration
st.set_page_config(
    page_title="TEG Introductory Call Dashboard",
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
    /* Hide tool pages from sidebar */
    [data-testid="stSidebarNav"] a[href*="tools"],
    [data-testid="stSidebarNav"] a[href*="signnow_form"],
    [data-testid="stSidebarNav"] a[href*="workbook_creator"],
    [data-testid="stSidebarNav"] a[href*="deck_creator"],
    [data-testid="stSidebarNav"] a[href*="a_la_carte"] {
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
    """Load Calendly data from SQLite database, filtered to Burki Calls and Intro Call with TEG events.
    Events are identified by:
    - Burki: event name containing "TEG" and "Let's Chat" or "Lets Chat"
    - Intro Call with TEG: event name containing 'introductory' or 'intro call' or scheduling URL contains 'intro-call-with-teg'
    """
    CALENDLY_DB_PATH = "calendly_data.db"
    
    try:
        conn = sqlite3.connect(CALENDLY_DB_PATH)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(calendly_events)")
        columns = [row[1] for row in cursor.fetchall()]
        has_source = "source" in columns
        
        cursor.execute("SELECT COUNT(*) FROM calendly_events")
        count = cursor.fetchone()[0]
        if count == 0:
            conn.close()
            return None, "No Calendly data found in database. Please refresh Calendly data first."
        
        if has_source:
            cursor.execute("""
                SELECT uri, name, start_time, end_time, status, event_type,
                       invitee_name, invitee_email, source, updated_at
                FROM calendly_events
                ORDER BY start_time DESC
            """)
            df = pd.DataFrame(cursor.fetchall(), columns=[
                'uri', 'name', 'start_time', 'end_time', 'status', 'event_type',
                'invitee_name', 'invitee_email', 'source', 'updated_at'
            ])
        else:
            cursor.execute("""
                SELECT uri, name, start_time, end_time, status, event_type,
                       invitee_name, invitee_email, updated_at
                FROM calendly_events
                ORDER BY start_time DESC
            """)
            df = pd.DataFrame(cursor.fetchall(), columns=[
                'uri', 'name', 'start_time', 'end_time', 'status', 'event_type',
                'invitee_name', 'invitee_email', 'updated_at'
            ])
            df["source"] = ""
        conn.close()
        
        # Filter to Intro Call events
        # Based on CSV export Event Type Name: "TEG - Let's Chat" and "*Intro call with TEG*"
        # Note: Database may store names with or without asterisks, so we check both
        if not df.empty and 'name' in df.columns:
            name_str = df['name'].astype(str)
            name_lower = name_str.str.lower()
            
            # Match "TEG - Let's Chat" exactly (case-insensitive, handles asterisks)
            # CSV: "TEG - Let's Chat"
            is_teg_lets_chat = (
                name_lower.str.contains("teg", na=False) &
                (name_lower.str.contains("let's chat", na=False) | name_lower.str.contains("lets chat", na=False))
            )
            
            # Match "*Intro call with TEG*" exactly (case-insensitive, handles asterisks)
            # CSV: "*Intro call with TEG*"
            # Must contain "intro call" but NOT "introductory" (to exclude "*TEG Introductory Call*")
            # Exclude "TEG Intro Call" (without asterisks, different event type)
            is_intro_call_with_teg = (
                name_lower.str.contains("intro call", na=False) &
                name_lower.str.contains("teg", na=False) &
                ~name_lower.str.contains("introductory", na=False) &
                ~(name_lower == "teg intro call")  # Exclude exact match "TEG Intro Call"
            )
            
            # Create mask for filtering
            mask = is_teg_lets_chat | is_intro_call_with_teg
            df = df[mask].copy()
            
            # Filter to only active events (exclude canceled)
            if 'status' in df.columns:
                df = df[df['status'].str.lower() == 'active'].copy()
            
            # Preserve existing source if present (person names like Ian, Anthony, Burki, etc.)
            df["source"] = df["source"].fillna("").astype(str).str.strip()
            # Remove generic labels - they should be replaced with person names from database
            generic_labels = ["Intro Call with TEG", "Other", "TEG - Let's Chat", "*Intro call with TEG*"]
            df.loc[df["source"].isin(generic_labels), "source"] = ""
            # Only set "Burki" for TEG - Let's Chat events if source is truly empty
            # For "*Intro call with TEG*" events, preserve person names (like Ian) - don't overwrite
            empty_source = (df["source"] == "")
            # Apply is_teg_lets_chat mask to the filtered dataframe
            df_teg_lets_chat_mask = df['name'].astype(str).str.lower().str.contains("teg", na=False) & (
                df['name'].astype(str).str.lower().str.contains("let's chat", na=False) | 
                df['name'].astype(str).str.lower().str.contains("lets chat", na=False)
            )
            df.loc[df_teg_lets_chat_mask & empty_source, "source"] = "Burki"
            # If source is still empty, leave it empty (don't set to "Other" - let it show as empty rather than wrong)
            # This helps identify events that need database refresh
        
        if df.empty:
            return None, "No Burki Calls or Intro Call with TEG events in database. Refresh Calendly data from the Database Refresh page."
        
        # Convert timestamps (Calendly API uses UTC) and add analysis columns
        # Convert to California timezone before extracting date (user's timezone)
        df['start_time'] = pd.to_datetime(df['start_time'], utc=True)
        df['end_time'] = pd.to_datetime(df['end_time'], utc=True)
        df['updated_at'] = pd.to_datetime(df['updated_at'], utc=True)
        df['start_time_local'] = df['start_time'].dt.tz_convert(CALIFORNIA_TZ)
        df['date'] = df['start_time_local'].dt.date
        df['month'] = df['start_time_local'].dt.strftime('%B %Y')
        df['week'] = df['start_time_local'].dt.isocalendar().week
        df['year'] = df['start_time_local'].dt.year
        df['day_of_week'] = df['start_time_local'].dt.strftime('%A')
        df['hour'] = df['start_time_local'].dt.hour
        return df, None
        
    except sqlite3.Error as e:
        return None, f"Database error: {str(e)}"
    except Exception as e:
        return None, f"Error loading Calendly data: {str(e)}"

def get_calendly_data():
    """Get Calendly data for TEG Introductory Call events (scheduling link: TEG_INTRO_CALL_SCHEDULING_URL)"""
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
    
    # Get scheduled events with pagination to get ALL events for current year
    min_start_time = f"{CURRENT_YEAR}-01-01T00:00:00.000000Z"
    max_start_time = f"{CURRENT_YEAR}-12-31T23:59:59.999999Z"
    
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

def create_two_week_daily_chart(df, start_date, end_date):
    """Create daily calls chart for the selected date range (all days in range, including 0 calls)."""
    if start_date is None or end_date is None:
        return None

    # Count calls for each day in the (already date-filtered) df
    daily_counts = df.groupby('date').size().reset_index(name='count')

    # Create a complete date range for the selected range (including days with 0 calls)
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    complete_dates = pd.DataFrame({'date': [d.date() for d in date_range]})

    # Merge with actual counts
    daily_counts = complete_dates.merge(daily_counts, on='date', how='left')
    daily_counts['count'] = daily_counts['count'].fillna(0).astype(int)
    daily_counts = daily_counts.sort_values('date')

    # Convert date to datetime for better Plotly handling
    daily_counts['date_datetime'] = pd.to_datetime(daily_counts['date'])

    date_range_label = f"{start_date.strftime('%b %d, %Y')} – {end_date.strftime('%b %d, %Y')}"

    # Create bar chart
    fig = px.bar(
        daily_counts,
        x='date_datetime',
        y='count',
        title=f'Calls by Day ({date_range_label})',
        labels={'count': 'Number of Calls', 'date_datetime': 'Date'},
        text='count'
    )
    
    # Customize x-axis to show dates in "Oct 29" format
    fig.update_xaxes(
        tickformat='%b %d',
        dtick=86400000  # One day in milliseconds
    )
    
    fig.update_layout(
        xaxis_title="Date",
        yaxis_title="Number of Calls",
        height=600,
        showlegend=False,
        xaxis=dict(
            tickangle=45
        )
    )
    
    fig.update_traces(
        marker_color='#1f77b4',
        textposition='outside'
    )
    
    return fig


def create_stacked_daily_chart(df, start_date, end_date):
    """Stacked bar: each day on x-axis, count by Source (dynamic)."""
    if start_date is None or end_date is None:
        return None
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    all_dates = [d.date() for d in date_range]
    if df.empty:
        counts = pd.DataFrame(columns=['date', 'source', 'count'])
        sources = []
    else:
        counts = df.groupby(['date', 'source']).size().reset_index(name='count')
        # Filter out empty sources - only include sources that have at least one event
        sources = sorted([s for s in df['source'].unique() if s and str(s).strip()])
    colors = get_color_palette(sources)
    rows = []
    for d in all_dates:
        for source in sources:
            n = counts[(counts['date'] == d) & (counts['source'] == source)]['count'].sum() if not counts.empty else 0
            rows.append({'date': pd.Timestamp(d), 'Source': source, 'count': int(n)})
    long = pd.DataFrame(rows)
    fig = px.bar(
        long, x='date', y='count', color='Source',
        color_discrete_map=colors,
        barmode='stack',
        title='Calls by Day (by source)',
        labels={'count': 'Number of Calls', 'date': 'Date'},
        category_orders={'Source': sources}
    )
    fig.update_layout(xaxis_title="Date", yaxis_title="Number of Calls", height=500, xaxis_tickangle=45)
    fig.update_traces(textposition='inside', texttemplate='%{y}')
    return fig


def create_stacked_weekly_chart(df):
    """Stacked bar: week on x-axis, count by Source (dynamic)."""
    if df.empty:
        return None
    df_copy = df.copy()
    df_copy['week_start'] = df_copy['start_time_local'].dt.to_period('W').dt.start_time
    df_copy['week_label'] = df_copy['week_start'].apply(
        lambda ts: f"{ts.strftime('%b %d')} - {(ts + pd.Timedelta(days=6)).strftime('%b %d')}"
    )
    counts = df_copy.groupby(['week_start', 'week_label', 'source']).size().reset_index(name='count')
    counts = counts.sort_values('week_start')
    long = counts.rename(columns={'source': 'Source'})
    # Filter out empty sources - only include sources that have at least one event
    sources = sorted([s for s in df['source'].unique() if s and str(s).strip()])
    colors = get_color_palette(sources)
    weeks = long[['week_start', 'week_label']].drop_duplicates()
    rows = []
    for _, r in weeks.iterrows():
        for source in sources:
            n = long[(long['week_start'] == r['week_start']) & (long['Source'] == source)]['count'].sum()
            rows.append({'week_start': r['week_start'], 'week_label': r['week_label'], 'Source': source, 'count': int(n)})
    long = pd.DataFrame(rows)
    fig = px.bar(
        long, x='week_label', y='count', color='Source',
        color_discrete_map=colors,
        barmode='stack',
        title='Calls by Week (by source)',
        labels={'count': 'Number of Calls', 'week_label': 'Week'},
        category_orders={'Source': sources}
    )
    fig.update_layout(xaxis_title="Week", yaxis_title="Number of Calls", height=500, xaxis_tickangle=45)
    fig.update_traces(textposition='inside', texttemplate='%{y}')
    return fig


def create_stacked_monthly_chart(df):
    """Stacked bar: month on x-axis, count by Source (dynamic)."""
    if df.empty:
        return None
    counts = df.groupby(['month', 'source']).size().reset_index(name='count')
    counts = counts.rename(columns={'source': 'Source'})
    # Filter out empty sources - only include sources that have at least one event
    sources = sorted([s for s in df['source'].unique() if s and str(s).strip()])
    colors = get_color_palette(sources)
    months = counts['month'].unique()
    rows = []
    for m in months:
        for source in sources:
            n = counts[(counts['month'] == m) & (counts['Source'] == source)]['count'].sum()
            rows.append({'month': m, 'Source': source, 'count': int(n)})
    counts = pd.DataFrame(rows)
    counts['month_dt'] = pd.to_datetime(counts['month'])
    counts = counts.sort_values('month_dt')
    fig = px.bar(
        counts, x='month', y='count', color='Source',
        color_discrete_map=colors,
        barmode='stack',
        title='Calls by Month (by source)',
        labels={'count': 'Number of Calls', 'month': 'Month'},
        category_orders={'Source': sources}
    )
    fig.update_layout(xaxis_title="Month", yaxis_title="Number of Calls", height=500, xaxis_tickangle=45)
    fig.update_traces(textposition='inside', texttemplate='%{y}')
    return fig

def create_weekly_chart(df):
    """Create weekly calls chart"""
    if df.empty:
        return None
    
    # Create proper week labels with date ranges
    df_copy = df.copy()
    df_copy['week_start'] = df_copy['start_time_local'].dt.to_period('W').dt.start_time
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
    df_copy['month_sort'] = df_copy['start_time_local'].dt.to_period('M')
    
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
    
    # Filter data for the selected month (using California timezone)
    df_filtered = df[
        (df['start_time_local'].dt.year == selected_month.year) & 
        (df['start_time_local'].dt.month == selected_month.month)
    ].copy()
    
    if df_filtered.empty:
        return None
    
    # Count calls for each day (using California timezone)
    daily_counts = df_filtered.groupby(df_filtered['start_time_local'].dt.date).size().to_dict()
    
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
    st.markdown('<div class="embed-header">📊 INTRO CALL DASHBOARD</div>', unsafe_allow_html=True)
    
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
            
            # Date range filter at the very top - applies to all charts (same UX as ads_dashboard)
            # Use form so the page only reruns when user clicks Apply (not on every date change)
            st.subheader("📅 Date Range")
            if "intro_call_start_date" not in st.session_state:
                st.session_state.intro_call_start_date = date(CURRENT_YEAR, 1, 1)
            if "intro_call_end_date" not in st.session_state:
                st.session_state.intro_call_end_date = date.today()
            with st.form(key="intro_call_date_range_form", clear_on_submit=False):
                date_col1, date_col2, date_col3 = st.columns([1, 1, 1])
                with date_col1:
                    start_input = st.date_input(
                        "Start Date",
                        value=st.session_state.intro_call_start_date,
                        help="Start date for all metrics and charts",
                        key="intro_call_start_date_input",
                    )
                with date_col2:
                    end_input = st.date_input(
                        "End Date",
                        value=st.session_state.intro_call_end_date,
                        help="End date for all metrics and charts",
                        key="intro_call_end_date_input",
                    )
                with date_col3:
                    st.markdown("<div style='margin-top: 14px; padding-top: 14px'></div>", unsafe_allow_html=True)
                    apply_clicked = st.form_submit_button("Apply Date Range Filters")
                if apply_clicked:
                    if start_input > end_input:
                        st.session_state.intro_call_start_date = end_input
                        st.session_state.intro_call_end_date = end_input
                    else:
                        st.session_state.intro_call_start_date = start_input
                        st.session_state.intro_call_end_date = end_input
                    st.rerun()
            start_date = st.session_state.intro_call_start_date
            end_date = st.session_state.intro_call_end_date
            if start_date > end_date:
                end_date = start_date

            # Filter data by selected date range
            df_filtered = df[(df["date"] >= start_date) & (df["date"] <= end_date)].copy()

            if df_filtered.empty:
                st.warning("No events found for the selected date range.")
                return

            # Charts section (stacked by person)
            st.markdown("---")

            # Create tabs for different views (stacked by person)
            tab1, tab2, tab3 = st.tabs(["📅 Daily View", "📊 Weekly View", "📊 Monthly View"])

            with tab1:
                stacked_daily = create_stacked_daily_chart(df_filtered, start_date, end_date)
                if stacked_daily:
                    st.plotly_chart(stacked_daily, use_container_width=True)
                else:
                    st.info("No calls in the selected date range.")

            with tab2:
                stacked_weekly = create_stacked_weekly_chart(df_filtered)
                if stacked_weekly:
                    st.plotly_chart(stacked_weekly, use_container_width=True)
                else:
                    st.info("No weekly data available")

            with tab3:
                stacked_monthly = create_stacked_monthly_chart(df_filtered)
                if stacked_monthly:
                    st.plotly_chart(stacked_monthly, use_container_width=True)
                else:
                    st.info("No monthly data available")
            
            # Skip detailed data table
                
        except Exception as e:
            st.error(f"Error loading Calendly data: {str(e)}")
            st.info("Please check your Calendly API key in secrets.toml")

if __name__ == "__main__":
    main()
