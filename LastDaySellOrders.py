"""Find existing single-leg SELLs (including children) to replace, or create missing SELLs."""

import os
from collections import Counter
import pandas as pd
import schwabdev
from schwab_tokens import TOKENS_DB
from dotenv import load_dotenv
from datetime import datetime, timedelta, timezone

# 1. Load hidden variables from the .env file
load_dotenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")

if not all([client_id, client_secret, callback_url]):
    raise ValueError("Error: One or more sensitive keys failed to load from the .env file.")

# 2. Initialize the Schwab client
client = schwabdev.Client(client_id, client_secret, callback_url, tokens_db=TOKENS_DB)

# 3. Retrieve linked accounts and extract the account hash
linked_accounts = client.linked_accounts().json()
if isinstance(linked_accounts, list) and len(linked_accounts) > 0:
    account_hash = linked_accounts[0].get('hashValue')
elif isinstance(linked_accounts, dict):
    account_securities = linked_accounts.get('securitiesAccountList', [])
    account_hash = account_securities[0].get('hashValue') if account_securities else linked_accounts.get('hashValue')
else:
    account_hash = None

if not account_hash:
    raise SystemExit("Error: Could not retrieve a valid account hash.")

# 4. Fetch the last 30 days, including queued orders, not just WORKING orders.
print("Fetching orders entered in the last 30 days from the first linked account...")

now_utc = datetime.now(timezone.utc)
from_time = (now_utc - timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
to_time = now_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")

# Keep the request within Schwab's order result limit; abort below if saturated.
max_results = 3000
orders_resp = client.account_orders(
    accountHash=account_hash,
    fromEnteredTime=from_time,
    toEnteredTime=to_time,
    maxResults=max_results,
)

if orders_resp.status_code != 200:
    raise SystemExit(
        f"Error fetching orders. Status: {orders_resp.status_code}. "
        f"Server response: {orders_resp.text}"
    )

orders = orders_resp.json()
if not isinstance(orders, list) or any(not isinstance(order, dict) for order in orders):
    raise SystemExit("Error: Expected a list of orders from Schwab; batch aborted.")
if len(orders) >= max_results:
    raise SystemExit(
        "Error: Order query reached its result limit; results may be incomplete. "
        "Batch aborted. Query smaller date windows before replacing orders."
    )

terminal_statuses = {"FILLED", "CANCELED", "REJECTED", "EXPIRED", "REPLACED"}
pending_changes = {"PENDING_CANCEL", "PENDING_REPLACE"}
all_orders = []
pending_orders = list(orders)
while pending_orders:
    order = pending_orders.pop()
    all_orders.append(order)
    pending_orders.extend(order.get("childOrderStrategies", []))

status_counts = Counter(order.get("status", "UNKNOWN") for order in all_orders)
print(f"Retrieved {len(orders)} orders. Status counts: {dict(status_counts)}")
# Replace individual leaf orders, never flatten a parent or multi-leg strategy.
open_orders = [
    order for order in all_orders
    if order.get("status")
    and order["status"] not in terminal_statuses | pending_changes | {"UNKNOWN"}
    and order.get("orderStrategyType") == "SINGLE"
    and len(order.get("orderLegCollection", [])) == 1
    and not order.get("childOrderStrategies")
    and order.get("orderId") is not None
]
print(f"Found {len(open_orders)} existing single-leg orders eligible for replacement attempts.")
attempted_order_ids = set()
attempted_sells = set()

# 5. Load and iterate through your Excel file
excel_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "LastDaySellOrders.xlsx")
try:
    df = pd.read_excel(excel_file)
    print(f"Successfully loaded {len(df)} lines from {excel_file}.\n")
except Exception as e:
    raise SystemExit(f"Error reading Excel file: {e}")

# 6. Replace an eligible order, or place a SELL when no unresolved SELL exists.
for index, row in df.iterrows():
    symbol = str(row['symbol']).strip().upper()
    instruction = str(row['instruction']).strip().upper()  # "BUY" or "SELL"
    quantity = int(row['quantity'])
    order_type = str(row['orderType']).strip().upper()
    duration = str(row['duration']).strip().upper()
    session_type = str(row.get('session', 'NORMAL')).strip().upper()

    if session_type == "SEEMLESS":
        session_type = "SEAMLESS"

    print(f"Processing row {index + 1}: Searching for open {instruction} order on {symbol}...")

    if instruction == "SELL" and symbol in attempted_sells:
        print(f"Skipped row {index + 1}: A SELL submission was already attempted for {symbol} in this batch.")
        print("-" * 40)
        continue

    # --- AUTO-LOOKUP LOGIC ---
    matched_order_ids = []
    for order in open_orders:
        if instruction != "SELL" and order.get("status") not in {"WORKING", "QUEUED"}:
            continue
        # Navigate Schwab's standard JSON order leg structure
        legs = order.get('orderLegCollection', [])
        if legs:
            leg = legs[0]
            leg_symbol = leg.get('instrument', {}).get('symbol', '').upper()
            leg_instruction = leg.get('instruction', '').upper()
            
            # Check if this active order matches our Excel symbol and buy/sell intent
            if leg_symbol == symbol and leg_instruction == instruction:
                order_id = str(order['orderId'])
                if order_id not in attempted_order_ids:
                    matched_order_ids.append(order_id)
    matched_order_ids = list(dict.fromkeys(matched_order_ids))

    if not matched_order_ids:
        if instruction != "SELL":
            print(f"Skipped row {index + 1}: No eligible {instruction} order found for {symbol}; only missing SELL orders are created.")
            print("-" * 40)
            continue

        unresolved_sells = []
        for order in all_orders:
            if order.get("status") in terminal_statuses:
                continue
            for leg in order.get("orderLegCollection", []):
                if (
                    leg.get("instrument", {}).get("symbol", "").upper() == symbol
                    and leg.get("instruction", "").upper() == "SELL"
                ):
                    unresolved_sells.append(
                        f"ID={order.get('orderId', 'missing')}, "
                        f"status={order.get('status', 'UNKNOWN')}, "
                        f"strategy={order.get('orderStrategyType', 'missing')}"
                    )
                    break

        if unresolved_sells:
            print(f"Skipped row {index + 1}: Cannot safely replace existing SELL for {symbol}: {'; '.join(unresolved_sells)}. Pending changes, missing IDs/status, and parent/multi-leg strategies require review. No new SELL placed.")
            print("-" * 40)
            continue

    if len(matched_order_ids) > 1:
        print(f"Skipped row {index + 1}: Multiple eligible {instruction} orders found for {symbol}: {', '.join(matched_order_ids)}. Select an order explicitly rather than replacing an arbitrary match.")
        print("-" * 40)
        continue

    matched_order_id = matched_order_ids[0] if matched_order_ids else None
    if matched_order_id:
        print(f"Found matching Order ID {matched_order_id}. Formatting replacement...")
    else:
        print(f"No active SELL found for {symbol}. Formatting a new SELL from the Excel row...")

    # Safety Check for Market Orders in Extended Hours
    if session_type in ["SEAMLESS", "SEPARATE"] and order_type == "MARKET":
        print(f"⚠️ Skipping row {index + 1}: Extended session '{session_type}' does not allow MARKET orders.")
        print("-" * 40)
        continue

    # Build the modification payload format
    order_payload = {
        "orderType": order_type,
        "session": session_type,        
        "duration": duration,
        "orderStrategyType": "SINGLE",
        "orderLegCollection": [
            {
                "instruction": instruction,
                "quantity": quantity,
                "instrument": {
                    "symbol": symbol,
                    "assetType": "EQUITY"
                }
            }
        ]
    }

    if order_type == "LIMIT":
        order_payload["price"] = f"{float(row['price']):.2f}"

    # Mark attempts before sending: a timeout does not prove submission failed.
    if instruction == "SELL":
        attempted_sells.add(symbol)
    if matched_order_id:
        attempted_order_ids.add(matched_order_id)
    action = f"replace order {matched_order_id}" if matched_order_id else f"place new SELL for {symbol}"
    try:
        if matched_order_id:
            resp = client.replace_order(account_hash, matched_order_id, order_payload)
        else:
            resp = client.place_order(account_hash, order_payload)
        
        if resp.status_code == 201:
            location_header = resp.headers.get('location', '')
            new_order_id = location_header.split('/')[-1] if location_header else "Unknown ID"
            print(f"Success: {action}. New Order ID: {new_order_id}")
        else:
            print(f"Failed to {action}. Status Code: {resp.status_code}")
            print(f"Response: {resp.text}")
            
    except Exception as e:
        print(f"Error attempting to {action}: {e}. Submission outcome may be unknown; verify with Schwab before retrying.")
        
    print("-" * 40)

print("Batch SELL placement/replacement processing complete.")
