import streamlit as st
import pandas as pd
from datetime import datetime, date
import calendar
import sys, os
import json

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from database_utils import (
    check_database_exists,
    get_new_leads_data,
    get_discovery_call_data,
    get_design_review_data,
    get_sales_data,
)

st.set_page_config(
    page_title="New Leads Check",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
/* Hide QuickBooks and SignNow pages from sidebar */
[data-testid="stSidebarNav"] a[href*="quickbooks_form"],
[data-testid="stSidebarNav"] a[href*="signnow_form"] {
    display: none !important;
}

.embed-header {
    font-size: 1.5rem;
    font-weight: bold;
    color: #1f77b4;
    margin-bottom: 1rem;
    text-align: center;
}

.stDataFrame {
    font-size: 12px;
}

.stDataFrame > div {
    max-height: 600px;
    overflow-y: auto;
}
</style>
""",
    unsafe_allow_html=True,
)

# ----------------------
# Data functions
# ----------------------
@st.cache_data(ttl=600, show_spinner=False)
def get_all_leads_data_from_db():
    """Load all leads data from local SQLite database (cached)."""
    # Keep imports of board fetchers centralized above for faster reloads
    boards = {
        "New Leads v2": get_new_leads_data(),
        "Discovery Call v2": get_discovery_call_data(),
        "Design Review v2": get_design_review_data(),
        "Sales v2": (
            get_sales_data()
            .get("data", {})
            .get("boards", [{}])[0]
            .get("items_page", {})
            .get("items", [])
        ),
    }

    # Flatten structure and tag board name
    return [
        {**item, "board_name": board_name}
        for board_name, items in boards.items()
        for item in items
    ]


def _current_month_bounds(today: date):
    month_start = today.replace(day=1)
    return month_start, today


def _is_current_month(selected_date: date) -> bool:
    today = date.today()
    return selected_date.year == today.year and selected_date.month == today.month


def _cache_file_path() -> str:
    # Store cache under inputs for fast local read
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "inputs", "new_leads_current_month.json")


@st.cache_data(ttl=300, show_spinner=False)
def try_load_cached_current_month_df(_cache_path: str) -> tuple[pd.DataFrame, dict]:
    """Load precomputed current-month leads dataframe and daily counts from JSON cache if available.
    
    Args:
        _cache_path: Cache file path (prefixed with _ to prevent Streamlit from treating it as a parameter)
    
    Returns:
        tuple: (dataframe, daily_counts_dict) or (empty_df, empty_dict) if cache not available
    """
    try:
        if os.path.exists(_cache_path):
            with open(_cache_path, "r", encoding="utf-8") as f:
                cache_data = json.load(f)
            
            # Handle both old and new cache formats
            if isinstance(cache_data, dict) and "records" in cache_data:
                # New format with pre-calculated daily counts
                records = cache_data["records"]
                daily_counts = cache_data.get("daily_counts", {})
                if records:
                    df = pd.DataFrame.from_records(records)
                    # Ensure date columns are proper types
                    if "Effective Date" in df.columns:
                        df["Effective Date"] = pd.to_datetime(df["Effective Date"], errors="coerce")
                    if "Effective Date Date" in df.columns:
                        df["Effective Date Date"] = pd.to_datetime(df["Effective Date Date"], errors="coerce").dt.date
                    return df, daily_counts
            elif isinstance(cache_data, list):
                # Old format - just records list
                if cache_data:
                    df = pd.DataFrame.from_records(cache_data)
                    if "Effective Date" in df.columns:
                        df["Effective Date"] = pd.to_datetime(df["Effective Date"], errors="coerce")
                    if "Effective Date Date" in df.columns:
                        df["Effective Date Date"] = pd.to_datetime(df["Effective Date Date"], errors="coerce").dt.date
                    return df, {}
    except Exception:
        pass
    return pd.DataFrame(), {}


@st.cache_data(ttl=600)
def format_leads_data(leads_data):
    if not leads_data:
        return pd.DataFrame()

    df = pd.DataFrame(
        [
            {
                "Item Name": i.get("name", ""),
                "Current Board": i.get("board_name", ""),
                "Created At": i.get("created_at", ""),
                "Date Created (Custom)": next(
                    (
                        c.get("text")
                        for c in (i.get("column_values") or [])
                        if (
                            c.get("type") == "date"
                            and c.get("text")
                            and "new lead form fill date"
                            not in (c.get("id") or "").lower()
                        )
                    ),
                    None,
                ),
            }
            for i in leads_data
        ]
    )

    df["Effective Date"] = pd.to_datetime(df["Date Created (Custom)"], errors="coerce")
    mask = df["Effective Date"].isna()
    if mask.any():
        df.loc[mask, "Effective Date"] = pd.to_datetime(
            df.loc[mask, "Created At"], errors="coerce"
        )

    df["Effective Date Date"] = df["Effective Date"].dt.date
    return df


def filter_leads_by_date(df, selected_date):
    if df.empty:
        return df
    if isinstance(selected_date, str):
        selected_date = pd.to_datetime(selected_date).date()
    return df[df["Effective Date Date"] == selected_date].copy()


def get_daily_counts(df, selected_date, precalculated_daily_counts=None):
    """Get daily counts for the selected date's month.
    
    Args:
        df: DataFrame with leads data
        selected_date: Selected date
        precalculated_daily_counts: Pre-calculated daily counts dict from cache (optional)
    
    Returns:
        pd.Series with daily counts
    """
    # If pre-calculated counts are available from cache, prefer them
    if precalculated_daily_counts:
        # Filter to only dates up to selected_date
        filtered_counts = {
            date_str: count
            for date_str, count in precalculated_daily_counts.items()
            if pd.to_datetime(date_str).date() <= selected_date
        }
        # Convert to Series with proper date index
        if filtered_counts:
            dates = [pd.to_datetime(k).date() for k in filtered_counts.keys()]
            counts = pd.Series(list(filtered_counts.values()), index=dates).sort_index()
            return counts
    
    # Fallback to calculating from dataframe
    if df.empty:
        return pd.Series(dtype=int)  # explicit dtype to avoid warnings

    month_start = selected_date.replace(day=1)
    mask = (df["Effective Date Date"] >= month_start) & (
        df["Effective Date Date"] <= selected_date
    )
    subset = df.loc[mask, "Effective Date Date"]
    # value_counts returns dtype int by default; convert index to date objects if needed
    counts = subset.value_counts().sort_index()
    # Ensure index objects are plain date if they are Timestamp
    counts.index = [d.date() if hasattr(d, "date") else d for d in counts.index]
    return counts


def display_calendar_html(daily_counts, selected_date):
    if daily_counts.empty:
        st.info("No leads this month.")
        return

    first_weekday, days_in_month = calendar.monthrange(
        selected_date.year, selected_date.month
    )
    html = "<table style='width:100%; border-collapse:collapse; text-align:center;'>"
    html += (
        "<tr>"
        + "".join(f"<th>{d}</th>" for d in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"])
        + "</tr><tr>"
    )

    # Leading empty cells
    for _ in range(first_weekday):
        html += "<td></td>"

    # Fill the days (only up to selected_date.day)
    for day in range(1, selected_date.day + 1):
        date_obj = selected_date.replace(day=day)
        count = int(daily_counts.get(date_obj, 0))
        color = "#b3e6b3"
        html += (
            f"<td style='padding:8px; border:1px solid #ccc; background:{color};'>"
            f"<b>{day}</b><br><small>{count} leads</small></td>"
        )
        if (first_weekday + day) % 7 == 0:
            html += "</tr><tr>"

    html += "</tr></table>"
    st.markdown(html, unsafe_allow_html=True)


# ----------------------
# Main UI
# ----------------------
def main():
    st.markdown('<div class="embed-header">🔍 NEW LEADS CHECK</div>', unsafe_allow_html=True)

    db_exists, db_message = check_database_exists()
    if not db_exists:
        st.error(f"❌ Database not ready: {db_message}")
        st.info(
            "💡 Please go to the 'Database Refresh' page to initialize the database with Monday.com data."
        )
        return

    # Sidebar
    with st.sidebar:
        st.header("⚙️ Settings")
        refresh = st.button("🔄 Refresh Data")
        st.info(f"Last Updated: {datetime.now():%Y-%m-%d %H:%M:%S}")

    # If user requested refresh, clear cache BEFORE calling cached function
    if refresh:
        st.cache_data.clear()

    # Date filter stays as-is (confirmed requirement)
    st.subheader("📅 Select Date")
    selected_date = st.date_input(
        "Choose a date to view leads created on that day:",
        value=date.today(),
        help="Select the date to filter leads by their creation date",
    )

    # Try cached JSON first (fast path), regardless of selected month
    df = pd.DataFrame()
    daily_counts_cache = {}
    cache_path = _cache_file_path()
    df, daily_counts_cache = try_load_cached_current_month_df(cache_path)

    # Fallback to database if cache not present or if viewing past months
    if df.empty:
        with st.spinner("Loading leads data from database..."):
            leads_data = get_all_leads_data_from_db()
            df = format_leads_data(leads_data)

    if df.empty:
        st.warning(
            "No leads data found. Please refresh the database from the 'Database Refresh' page."
        )
        return

    st.subheader("Monthly Calendar View")
    # Use pre-calculated daily counts if available for near-instantaneous load
    daily_counts = get_daily_counts(df, selected_date, daily_counts_cache)
    display_calendar_html(daily_counts, selected_date)
    st.markdown("---")

    filtered_df = filter_leads_by_date(df, selected_date)
    if filtered_df.empty:
        st.info(f"No leads were created on {selected_date:%B %d, %Y}.")
    else:
        st.metric("Leads Found", len(filtered_df))
        st.dataframe(
            filtered_df[["Item Name", "Current Board"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Item Name": "Item Name",
                "Current Board": "Current Board",
            },
        )


if __name__ == "__main__":
    main()
