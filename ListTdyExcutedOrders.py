import os
import pandas as pd
import schwabdev
from schwab_tokens import TOKENS_DB
from dotenv import load_dotenv
from datetime import datetime, timezone, timedelta

# 1. Load hidden variables from the .env file
load_dotenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")

if not all([client_id, client_secret, callback_url]):
    raise ValueError("Error: One or more sensitive keys failed to load from the .env file.")

# 2. Initialize the Schwab client
try:
    client = schwabdev.Client(client_id, client_secret, callback_url, tokens_db=TOKENS_DB)
except Exception as e:
    raise SystemExit(f"💥 Critical Error: Failed to initialize Schwab Client. Details: {e}")

# 3. Retrieve linked accounts and extract the account hash
try:
    linked_accounts = client.linked_accounts().json()
    if isinstance(linked_accounts, list) and len(linked_accounts) > 0:
        account_hash = linked_accounts[0].get('hashValue')
    elif isinstance(linked_accounts, dict):
        account_securities = linked_accounts.get('securitiesAccountList', [])
        account_hash = account_securities[0].get('hashValue') if account_securities else linked_accounts.get('hashValue')
    else:
        account_hash = None
except Exception as e:
    raise SystemExit(f"💥 Critical Error: Could not connect to your Schwab account to get linked accounts. Details: {e}")

if not account_hash:
    raise SystemExit("Error: Could not retrieve a valid account hash.")

# 4. Helper function to check if a date is a NYSE Stock Market Holiday
def is_market_holiday(dt):
    month, day, year = dt.month, dt.day, dt.year
    weekday = dt.weekday() # 0=Monday, 3=Thursday, 4=Friday, etc.
    dom_count = (day - 1) // 7 + 1

    # New Year's Day (Jan 1 or Jan 2 if Jan 1 is Sunday)
    if month == 1 and (day == 1 or (day == 2 and weekday == 0)):
        return True
    # Martin Luther King Jr. Day (Third Monday in January)
    if month == 1 and weekday == 0 and dom_count == 3:
        return True
    # Washington's Birthday / Presidents' Day (Third Monday in February)
    if month == 2 and weekday == 0 and dom_count == 3:
        return True
    # Memorial Day (Last Monday in May)
    if month == 5 and weekday == 0 and (day + 7 > 31):
        return True
    # Juneteenth (June 19, or observed Mon/Fri if weekend)
    if month == 6 and (day == 19 or (day == 20 and weekday == 0) or (day == 18 and weekday == 4)):
        return True
    # Independence Day (July 4, or observed Mon/Fri if weekend)
    if month == 7 and (day == 4 or (day == 5 and weekday == 0) or (day == 3 and weekday == 4)):
        return True
    # Labor Day (First Monday in September)
    if month == 9 and weekday == 0 and dom_count == 1:
        return True
    # Thanksgiving Day (Fourth Thursday in November)
    if month == 11 and weekday == 3 and dom_count == 4: 
        return True
    # Christmas Day (Dec 25, or observed Mon/Fri if weekend)
    if month == 12 and (day == 25 or (day == 26 and weekday == 0) or (day == 24 and weekday == 4)):
        return True
    return False

# 5. Auto-calculate the boundaries for the LAST ACTIVE TRADING DAY
target_day = datetime.now(timezone.utc)
while True:
    if target_day.weekday() in [5, 6] or is_market_holiday(target_day):
        target_day -= timedelta(days=1)
    else:
        break

from_time = target_day.strftime("%Y-%m-%dT00:00:00.000Z")
to_time = target_day.strftime("%Y-%m-%dT23:59:59.000Z")
print(f"Targeting trading day: {target_day.strftime('%A, %b %d, %Y')}")
print("Fetching filled orders from Schwab...")

# 6. Query filled statuses for the selected target day (Accepting BUY & SELL)
try:
    resp = client.account_orders(
        accountHash=account_hash,
        fromEnteredTime=from_time,
        toEnteredTime=to_time,
        status="FILLED"
    )
    if resp.status_code == 200:
        orders = resp.json()
        extracted_data = []
        for o in orders:
            legs = o.get('orderLegCollection', [])
            leg = legs[0] if legs else {}
            
            # Read instruction safely and normalize to uppercase
            instruction = str(leg.get("instruction", "")).upper()
            
            # Skip empty records or unrelated instructions
            if not instruction:
                continue
                
            # Grouping instructions to explicitly catch all Buy vs Sell types
            if "BUY" in instruction:
                order_action = "BUY"
            elif "SELL" in instruction:
                order_action = "SELL"
            else:
                order_action = instruction # Keeps SHORT, COVER, etc., if applicable

            extracted_data.append({
                "orderId": o.get("orderId"),
                "symbol": leg.get("instrument", {}).get("symbol"),
                "instruction": order_action,          # Displays clean 'BUY' or 'SELL'
                "rawInstruction": instruction,        # Keeps original Schwab string for debug
                "quantity": leg.get("quantity"),
                "orderType": o.get("orderType"),
                "price": o.get("price"),
                "duration": o.get("duration"),
                "session": o.get("session"),
                "status": o.get("status"),
                "enteredTime": o.get("enteredTime")
            })
            
        if extracted_data:
            df = pd.DataFrame(extracted_data)
            
            # Parse string timestamps into datetime objects
            df['enteredTime'] = pd.to_datetime(df['enteredTime'])
            
            # Sort descending by transaction time (Latest first)
            df = df.sort_values(by='enteredTime', ascending=False)
            
            # Remove timezone info to satisfy Excel engine requirements
            df['enteredTime'] = df['enteredTime'].dt.tz_localize(None)
            
            output_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "ListTdyExecutedOrders.xlsx")
            df.to_excel(output_file, index=False)
            print(f"✅ Success! Exported {len(df)} FILLED orders (BUY & SELL) to {output_file}")
        else:
            print("ℹ️ No filled BUY or SELL orders found matching that specific calendar day window.")
    else:
        print(f"⚠️ Warning: Server responded with status code {resp.status_code}")
        print(f"Response details: {resp.text}")
except Exception as e:
    print(f"💥 Critical error fetching or exporting orders: {e}")
