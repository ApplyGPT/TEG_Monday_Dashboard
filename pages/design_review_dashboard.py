import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
import os

# Get current year dynamically
CURRENT_YEAR = datetime.now().year

# Design Review Call: Two Calendly links (for reference, filtering is dynamic)
DESIGN_REVIEW_LINKS = {
    "TEG Introductory Call": "https://calendly.com/d/ctc8-ndq-rjz/teg-introductory-call",
    "Jennifer": "https://calendly.com/jennifer-teg/30minutegooglemeet",
}

def get_color_palette(sources):
    """Generate a color palette for the given sources dynamically."""
    base_colors = ["#1f77b4", "#2ca02c", "#ff7f0e", "#d62728", "#9467bd", "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"]
    return {source: base_colors[i % len(base_colors)] for i, source in enumerate(sorted(sources))}

# Page configuration
st.set_page_config(
    page_title="Design Review Call Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
<style>
    .main { padding: 1rem; }
    .stMetric { background-color: #f8f9fa; padding: 0.5rem; border-radius: 0.5rem; border-left: 4px solid #1f77b4; }
    .embed-header { font-size: 1.5rem; font-weight: bold; color: #1f77b4; margin-bottom: 1rem; text-align: center; }
    [data-testid="stSidebarNav"] a[href*="tools"], [data-testid="stSidebarNav"] a[href*="signnow_form"],
    [data-testid="stSidebarNav"] a[href*="workbook_creator"], [data-testid="stSidebarNav"] a[href*="deck_creator"],
    [data-testid="stSidebarNav"] a[href*="a_la_carte"] { display: none !important; }
</style>
""", unsafe_allow_html=True)


def load_design_review_data_from_db():
    """Load Calendly data for Design Review: TEG Introductory Call and Jennifer events.
    Events are identified by:
    - TEG Introductory Call: scheduling URL contains 'ctc8-ndq-rjz/teg-introductory-call' or name contains 'introductory'
    - Jennifer: scheduling URL contains 'jennifer-teg/30minutegooglemeet' or name contains 'jennifer'
    Preserves existing source values from database (e.g., person names like Anthony, Heather, Ian).
    """
    CALENDLY_DB_PATH = "calendly_data.db"
    if not os.path.exists(CALENDLY_DB_PATH):
        return None, "Calendly database not found. Refresh Calendly data from the Database Refresh page."
    try:
        conn = sqlite3.connect(CALENDLY_DB_PATH)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(calendly_events)")
        columns = [row[1] for row in cursor.fetchall()]
        if "source" not in columns:
            conn.close()
            return None, "Design Review 'source' column not in database. Refresh Calendly data from the Database Refresh page to include Design Review links."
        cursor.execute("""
            SELECT uri, name, start_time, end_time, status, event_type, invitee_name, invitee_email, source, updated_at
            FROM calendly_events
            ORDER BY start_time DESC
        """)
        rows = cursor.fetchall()
        conn.close()
        if not rows:
            return pd.DataFrame(), None
        df = pd.DataFrame(rows, columns=[
            'uri', 'name', 'start_time', 'end_time', 'status', 'event_type',
            'invitee_name', 'invitee_email', 'source', 'updated_at'
        ])
        # Filter to TEG Introductory Call and Jennifer events
        if not df.empty:
            name_lower = df['name'].astype(str).str.lower()
            event_type_lower = df['event_type'].astype(str).str.lower()
            
            # TEG Introductory Call: URL contains 'ctc8-ndq-rjz/teg-introductory-call' or name contains 'introductory'
            is_teg_intro = (
                event_type_lower.str.contains('ctc8-ndq-rjz/teg-introductory-call', na=False) |
                name_lower.str.contains('introductory', na=False)
            )
            
            # Jennifer: URL contains 'jennifer-teg/30minutegooglemeet' or name contains 'jennifer'
            is_jennifer = (
                event_type_lower.str.contains('jennifer-teg/30minutegooglemeet', na=False) |
                name_lower.str.contains('jennifer', na=False)
            )
            
            # Create mask for filtering
            mask = is_teg_intro | is_jennifer
            df = df[mask].copy()
            
            # Preserve existing source if present, otherwise set based on pattern
            df["source"] = df["source"].fillna("").astype(str).str.strip()
            # Only set source if it's empty or "Other" or "Design Review" (generic)
            empty_or_generic = (df["source"] == "") | (df["source"] == "Other") | (df["source"] == "Design Review")
            df.loc[is_teg_intro[mask] & empty_or_generic, "source"] = "TEG Introductory Call"
            df.loc[is_jennifer[mask] & empty_or_generic, "source"] = "Jennifer"
            # If source is still empty after pattern matching, set to "Other"
            df.loc[df["source"] == "", "source"] = "Other"
        
        if df.empty:
            return pd.DataFrame(), None
        # Calendly API uses UTC; parse as UTC so date doesn't shift by timezone
        df['start_time'] = pd.to_datetime(df['start_time'], utc=True)
        df['end_time'] = pd.to_datetime(df['end_time'], utc=True)
        df['updated_at'] = pd.to_datetime(df['updated_at'], utc=True)
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
        return None, f"Error loading data: {str(e)}"


def create_stacked_daily_chart(df, start_date, end_date):
    """Stacked bar: each day on x-axis, count by Source (dynamic)."""
    if start_date is None or end_date is None:
        return None
    # Full date range
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    all_dates = [d.date() for d in date_range]
    # Count by date and source (df may be empty)
    if df.empty:
        counts = pd.DataFrame(columns=['date', 'source', 'count'])
        sources = []
    else:
        counts = df.groupby(['date', 'source']).size().reset_index(name='count')
        sources = sorted(df['source'].unique())
    colors = get_color_palette(sources)
    # One row per (date, Source) with count (0 if missing) so all sources always appear
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
        title='Calls by Day',
        labels={'count': 'Number of Calls', 'date': 'Date'},
        category_orders={'Source': sources}
    )
    fig.update_layout(xaxis_title="Date", yaxis_title="Number of Calls", height=500, xaxis_tickangle=45)
    fig.update_traces(textposition='inside', texttemplate='%{y}')
    return fig


def create_stacked_weekly_chart(df):
    """Stacked bar: week range on x-axis, count by Source (dynamic)."""
    if df.empty:
        return None
    df_copy = df.copy()
    df_copy['week_start'] = df_copy['start_time'].dt.to_period('W').dt.start_time
    df_copy['week_label'] = df_copy['week_start'].apply(
        lambda ts: f"{ts.strftime('%b %d')} - {(ts + pd.Timedelta(days=6)).strftime('%b %d')}"
    )
    counts = df_copy.groupby(['week_start', 'week_label', 'source']).size().reset_index(name='count')
    counts = counts.sort_values('week_start')
    long = counts.rename(columns={'source': 'Source'})
    sources = sorted(df['source'].unique())
    colors = get_color_palette(sources)
    # Ensure all sources appear in legend (add 0 rows for missing sources per week)
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
        title='Calls by Week',
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
    sources = sorted(df['source'].unique())
    colors = get_color_palette(sources)
    # Ensure all sources per month (fill 0) so all appear in legend
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
        title='Calls by Month',
        labels={'count': 'Number of Calls', 'month': 'Month'},
        category_orders={'Source': sources}
    )
    fig.update_layout(xaxis_title="Month", yaxis_title="Number of Calls", height=500, xaxis_tickangle=45)
    fig.update_traces(textposition='inside', texttemplate='%{y}')
    return fig


def main():
    st.markdown('<div class="embed-header">📊 DESIGN REVIEW CALL DASHBOARD</div>', unsafe_allow_html=True)
    with st.sidebar:
        st.header("⚙️ Settings")
        st.info(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if st.button("🔄 Refresh Data"):
            st.rerun()

    with st.spinner("Loading Design Review data..."):
        df, error = load_design_review_data_from_db()
    if error:
        st.error(f"Error loading data: {error}")
        st.info("💡 Go to the Database Refresh page and click 'Refresh All Calendly Data' to include Design Review links (TEG Introductory Call, Jennifer).")
        return
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        st.warning("No Design Review events in database. Refresh Calendly data and ensure the Design Review links (TEG Introductory Call, Jennifer) are under your Calendly account.")
        return

    # Date range (form so page only reruns on Apply)
    st.subheader("📅 Date Range")
    if "design_review_start_date" not in st.session_state:
        st.session_state.design_review_start_date = date(CURRENT_YEAR, 1, 1)
    if "design_review_end_date" not in st.session_state:
        st.session_state.design_review_end_date = date.today()
    with st.form(key="design_review_date_form", clear_on_submit=False):
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            start_in = st.date_input("Start Date", value=st.session_state.design_review_start_date, key="dr_start")
        with c2:
            end_in = st.date_input("End Date", value=st.session_state.design_review_end_date, key="dr_end")
        with c3:
            st.markdown("<div style='margin-top: 14px; padding-top: 14px'></div>", unsafe_allow_html=True)
            apply_btn = st.form_submit_button("Apply Date Range Filters")
        if apply_btn:
            if start_in > end_in:
                st.session_state.design_review_start_date = end_in
                st.session_state.design_review_end_date = end_in
            else:
                st.session_state.design_review_start_date = start_in
                st.session_state.design_review_end_date = end_in
            st.rerun()
    start_date = st.session_state.design_review_start_date
    end_date = st.session_state.design_review_end_date
    if start_date > end_date:
        end_date = start_date

    df_filtered = df[(df["date"] >= start_date) & (df["date"] <= end_date)].copy()
    if df_filtered.empty:
        st.warning("No events in the selected date range.")
        return

    # Tabs: Daily, Weekly, Monthly (stacked bar each)
    st.markdown("---")
    tab1, tab2, tab3 = st.tabs(["📅 Daily View", "📊 Weekly View", "📊 Monthly View"])
    with tab1:
        fig_d = create_stacked_daily_chart(df_filtered, start_date, end_date)
        if fig_d:
            st.plotly_chart(fig_d, use_container_width=True)
        else:
            st.info("No daily data for the selected range.")
    with tab2:
        fig_w = create_stacked_weekly_chart(df_filtered)
        if fig_w:
            st.plotly_chart(fig_w, use_container_width=True)
        else:
            st.info("No weekly data available.")
    with tab3:
        fig_m = create_stacked_monthly_chart(df_filtered)
        if fig_m:
            st.plotly_chart(fig_m, use_container_width=True)
        else:
            st.info("No monthly data available.")


if __name__ == "__main__":
    main()
