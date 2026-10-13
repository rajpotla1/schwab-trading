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
try:
    client = schwabdev.Client(client_id, client_secret, callback_url)
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

# 4. Auto-calculate the boundaries for the time window
# Look back up to 30 days to catch active working orders still on the books
now_utc = datetime.now(timezone.utc)
from_time = (now_utc - timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
to_time = now_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")

print("Fetching active open orders from Schwab...")

# 5. Query WORKING statuses and filter for BUY orders
try:
    resp = client.account_orders(
        accountHash=account_hash,
        fromEnteredTime=from_time,
        toEnteredTime=to_time,
        status="WORKING"
    )
    
    if resp.status_code == 200:
        orders = resp.json()
        extracted_data = []
        
        for o in orders:
            legs = o.get('orderLegCollection', [])
            leg = legs[0] if legs else {}
            
            instruction = str(leg.get("instruction", "")).upper()
            
            # Local Filter: Only process orders containing standard "BUY" instructions
            if "BUY" not in instruction:
                continue
                
            extracted_data.append({
                "symbol": leg.get("instrument", {}).get("symbol"),
                "enteredTime": o.get("enteredTime")
            })
            
        if extracted_data:
            df = pd.DataFrame(extracted_data)
            df['enteredTime'] = pd.to_datetime(df['enteredTime'])
            # Sort descending by transaction time (Latest first)
            df = df.sort_values(by='enteredTime', ascending=False)
            
            print(f"\n✅ Found {len(df)} Open Working BUY Orders:\n")
            # Output only the symbol column without the index row numbers
            print(df[['symbol']].to_string(index=False))
        else:
            print("ℹ️ No active open BUY orders found on the account profile.")
    else:
        print(f"⚠️ Warning: Server responded with status code {resp.status_code}")
        print(f"Response details: {resp.text}")
        
except Exception as e:
    print(f"💥 Critical error fetching open orders: {e}")
