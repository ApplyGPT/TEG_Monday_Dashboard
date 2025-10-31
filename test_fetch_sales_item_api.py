"""
Direct Monday.com API fetch for a single Sales item to inspect raw column_values
Reads credentials and board id from .streamlit/secrets.toml
"""
import os
import json
import toml
import requests


def load_secrets():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    secrets_path = os.path.join(script_dir, '.streamlit', 'secrets.toml')
    with open(secrets_path, 'r', encoding='utf-8') as f:
        return toml.load(f)


def fetch_sales_board_items(api_token: str, sales_board_id: int):
    url = "https://api.monday.com/v2"
    headers = {"Authorization": api_token, "Content-Type": "application/json"}
    # Fetch in one page (limit 200) for simplicity
    query = f"""
    query {{
      boards(ids: [{sales_board_id}]) {{
        items_page(limit: 200) {{
          items {{
            id
            name
            column_values {{
              id
              type
              text
              value
            }}
          }}
        }}
      }}
    }}
    """
    resp = requests.post(url, json={"query": query}, headers=headers, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    return data.get("data", {}).get("boards", [{}])[0].get("items_page", {}).get("items", [])


def fetch_items_by_ids(api_token: str, board_id: int, item_ids):
    url = "https://api.monday.com/v2"
    headers = {"Authorization": api_token, "Content-Type": "application/json"}
    ids_list = ",".join(str(x) for x in item_ids)
    query = f"""
    query {{
      boards(ids: [{board_id}]) {{
        items_page(query_params: {{ ids: [{ids_list}] }}) {{
          items {{
            id
            name
            column_values {{ id type text value }}
          }}
        }}
      }}
    }}
    """
    resp = requests.post(url, json={"query": query}, headers=headers, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    return data.get("data", {}).get("boards", [{}])[0].get("items_page", {}).get("items", [])


def main():
    secrets = load_secrets()
    monday = secrets.get("monday", {})
    api_token = monday.get("api_token")
    sales_board_id = int(monday.get("sales_board_id"))

    items = fetch_sales_board_items(api_token, sales_board_id)
    target_name = "Estevao Cavalcante (TEST)"
    targets = [it for it in items if it.get("name") == target_name]

    if not targets:
        print("Item not found:", target_name)
        print(f"Fetched {len(items)} items from sales board")
        return

    item = targets[0]
    print("Item:", item.get("name"), "ID:", item.get("id"))
    print("Total columns:", len(item.get("column_values", [])))
    connect_json = None
    for cv in item.get("column_values", []):
        cid = cv.get("id")
        ctype = cv.get("type")
        text = cv.get("text")
        value = cv.get("value")
        disp_text = (text or "")
        value_short = (value[:100] + "...") if isinstance(value, str) and len(value) > 100 else value
        print(f"- {cid:20s} | {ctype:10s} | text='{disp_text}' | value='{value_short}'")
        if cid.startswith("connect_boards") and value and not connect_json:
            try:
                connect_json = json.loads(value) if isinstance(value, str) else value
            except Exception:
                connect_json = None

    # If we have connected items, fetch their values to emulate Mirror
    if connect_json and isinstance(connect_json, dict):
        linked = connect_json.get("linkedPulseIds") or []
        linked_ids = [int(x.get("linkedPulseId")) for x in linked if x.get("linkedPulseId")]
        board_ids = connect_json.get("boardIds") or []
        linked_board_id = int(board_ids[0]) if board_ids else board_id
        if linked_ids:
            print("\nFollowing connect_boards → fetching linked items:")
            linked_items = fetch_items_by_ids(api_token, linked_board_id, linked_ids)
            for li in linked_items:
                print("Linked:", li.get("name"), "ID:", li.get("id"))
                # Print key monetary columns and any formula
                for cv in li.get("column_values", []):
                    cid = cv.get("id")
                    if any(k in cid for k in ["numbers3", "contract_amt", "formula", "numeric", "numbers"]):
                        print("  ", cid, "=>", cv.get("text"), cv.get("value"))


if __name__ == "__main__":
    main()


