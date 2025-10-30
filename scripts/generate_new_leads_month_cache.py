import os
import json
from datetime import date
import pandas as pd

import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from database_utils import (
    get_new_leads_data,
    get_discovery_call_data,
    get_design_review_data,
    get_sales_data,
)


def _cache_file_path() -> str:
    return os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "inputs",
        "new_leads_current_month.json",
    )


def _format_leads_data(leads_data):
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


def main():
    today = date.today()
    month_start = today.replace(day=1)

    # Build combined leads like the page does
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

    leads_data = [
        {**item, "board_name": board_name}
        for board_name, items in boards.items()
        for item in items
    ]

    df = _format_leads_data(leads_data)
    if df.empty:
        print("No data to cache.")
        return

    # Filter to current month
    df_current = df[(df["Effective Date Date"] >= month_start) & (df["Effective Date Date"] <= today)].copy()

    # Serialize to JSON records for fast load
    records = df_current.to_dict(orient="records")
    cache_path = _cache_file_path()
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False)

    print(f"Wrote {len(records)} records to {cache_path}")


if __name__ == "__main__":
    main()


