import os
import pandas as pd
import schwabdev
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
client = schwabdev.Client(client_id, client_secret, callback_url)

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

# 4. Helper function to check if a date is a NYSE Stock Market Holiday
def is_market_holiday(dt):
    month, day, year = dt.month, dt.day, dt.year
    weekday = dt.weekday()  # 0=Monday, 3=Thursday, 4=Friday
    dom_count = (day - 1) // 7 + 1

    if month == 1 and (day == 1 or (day == 2 and weekday == 0)):
        return True
    if month == 1 and weekday == 0 and dom_count == 3:
        return True
    if month == 2 and weekday == 0 and dom_count == 3:
        return True
    if month == 5 and weekday == 0 and (day + 7 > 31):
        return True
    if month == 6 and (day == 19 or (day == 20 and weekday == 0) or (day == 18 and weekday == 4)):
        return True
    if month == 7 and (day == 4 or (day == 5 and weekday == 0) or (day == 3 and weekday == 4)):
        return True
    if month == 9 and weekday == 0 and dom_count == 1:
        return True
    if month == 11 and weekday == 3 and dom_count == 4:
        return True
    if month == 12 and (day == 25 or (day == 26 and weekday == 0) or (day == 24 and weekday == 4)):
        return True

    return False

# 5. Calculate the boundaries for the PREVIOUS LAST TRADING DAY
target_day = datetime.now(timezone.utc)

# Step 1: Find the most recent active trading day
while True:
    if target_day.weekday() in [5, 6] or is_market_holiday(target_day):
        target_day -= timedelta(days=1)
    else:
        break

# Step 2: Push back by 1 day and skip weekends/holidays again to find the PREVIOUS trading day
target_day -= timedelta(days=1)
while True:
    if target_day.weekday() in [5, 6] or is_market_holiday(target_day):
        target_day -= timedelta(days=1)
    else:
        break

# Format the time window specifically for Schwab's API requirements
from_time = target_day.strftime("%Y-%m-%dT00:00:00.000Z")
to_time = target_day.strftime("%Y-%m-%dT23:59:59.000Z")

print(f"Targeting previous trading day: {target_day.strftime('%A, %b %d, %Y')}")
print("Fetching filled orders from Schwab...")

# 6. Query filled statuses for the selected target day (Filter "BUY" locally)
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
            leg = legs[0] if legs else {}   # Safely pull the first leg item dictionary
            
            # 💡 Local Filter: Only proceed if the instruction is an exact match for BUY
            if leg.get("instruction") != "BUY":
                continue
            
            extracted_data.append({
                "orderId": o.get("orderId"),
                "symbol": leg.get("instrument", {}).get("symbol"),
                "instruction": leg.get("instruction"),
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
            output_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "schwab_orders_previous_trading_day.xlsx")
            df.to_excel(output_file, index=False)
            print(f"✅ Success! Exported {len(df)} FILLED BUY orders to {output_file}")
        else:
            print("ℹ️ No filled BUY orders found matching that specific calendar day window.")
            
    else:
        print(f"⚠️ Warning: Server responded with status code {resp.status_code}")
        print(f"Response details: {resp.text}")

except Exception as e:
    print(f"💥 Critical error fetching or exporting orders: {e}")
