import streamlit as st
import pandas as pd
from datetime import datetime, date
import calendar
import sys, os

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

@st.cache_data(ttl=600, show_spinner=False)
def get_all_leads_data_from_db(force_refresh=False):
    """Load all leads data from local SQLite database (cached)."""
    if force_refresh:
        st.cache_data.clear()

    boards = {
        "New Leads v2": get_new_leads_data(),
        "Discovery Call v2": get_discovery_call_data(),
        "Design Review v2": get_design_review_data(),
        "Sales v2": get_sales_data()
        .get("data", {})
        .get("boards", [{}])[0]
        .get("items_page", {})
        .get("items", []),
    }

    # Flatten structure and tag board name
    return [
        {**item, "board_name": board_name}
        for board_name, items in boards.items()
        for item in items
    ]


@st.cache_data(ttl=600)
def format_leads_data(leads_data):
    """Convert Monday.com leads JSON into DataFrame (fast & vectorized)."""
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

    # Parse datetime efficiently
    df["Effective Date"] = pd.to_datetime(df["Date Created (Custom)"], errors="coerce")
    mask = df["Effective Date"].isna()
    if mask.any():
        df.loc[mask, "Effective Date"] = pd.to_datetime(
            df.loc[mask, "Created At"], errors="coerce"
        )

    df["Effective Date Date"] = df["Effective Date"].dt.date
    return df


def filter_leads_by_date(df, selected_date):
    """Filter leads DataFrame by selected date."""
    if df.empty:
        return df
    if isinstance(selected_date, str):
        selected_date = pd.to_datetime(selected_date).date()
    return df[df["Effective Date Date"] == selected_date].copy()


def get_daily_counts(df, selected_date):
    """Return Series of daily lead counts for the month up to selected date."""
    if df.empty:
        return pd.Series(dtype=int)

    month_start = selected_date.replace(day=1)
    mask = (df["Effective Date Date"] >= month_start) & (
        df["Effective Date Date"] <= selected_date
    )
    subset = df.loc[mask, "Effective Date Date"]
    return subset.value_counts().sort_index()


def display_calendar_html(daily_counts, selected_date):
    """Render a lightweight HTML calendar instead of heavy Streamlit columns."""
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

    # Fill the days
    for day in range(1, selected_date.day + 1):
        date_obj = selected_date.replace(day=day)
        count = int(daily_counts.get(date_obj, 0))
        color = ("#b3e6b3")
        html += (
            f"<td style='padding:8px; border:1px solid #ccc; background:{color};'>"
            f"<b>{day}</b><br><small>{count} leads</small></td>"
        )
        if (first_weekday + day) % 7 == 0:
            html += "</tr><tr>"

    html += "</tr></table>"
    st.markdown(html, unsafe_allow_html=True)


def main():
    st.markdown('<div class="embed-header">🔍 NEW LEADS CHECK</div>', unsafe_allow_html=True)

    # Check database
    db_exists, db_message = check_database_exists()
    if not db_exists:
        st.error(f"❌ Database not ready: {db_message}")
        st.info(
            "💡 Please go to the 'Database Refresh' page to initialize the database with Monday.com data."
        )
        return

    # Sidebar controls
    with st.sidebar:
        st.header("⚙️ Settings")
        refresh = st.button("🔄 Refresh Data")
        st.info(f"Last Updated: {datetime.now():%Y-%m-%d %H:%M:%S}")

    # Load data
    with st.spinner("Loading leads data from database..."):
        leads_data = get_all_leads_data_from_db(force_refresh=refresh)
        df = format_leads_data(leads_data)

    if df.empty:
        st.warning(
            "No leads data found in database. Please refresh the database from the 'Database Refresh' page."
        )
        return

    # Date input
    st.subheader("📅 Select Date")
    selected_date = st.date_input(
        "Choose a date to view leads created on that day:",
        value=date.today(),
        help="Select the date to filter leads by their creation date",
    )

    # Monthly view
    st.subheader("Monthly Calendar View")
    daily_counts = get_daily_counts(df, selected_date)
    display_calendar_html(daily_counts, selected_date)
    st.markdown("---")

    # Filtered data
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
