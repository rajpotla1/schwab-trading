import os
import pandas as pd
import schwabdev
from dotenv import load_dotenv

# 1. Load configuration and initialize the Schwab client
load_dotenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")

if not all([client_id, client_secret, callback_url]):
    raise ValueError("Error: Sensitive keys failed to load from the .env file.")

client = schwabdev.Client(client_id, client_secret, callback_url)

# 2. Retrieve linked accounts and extract the account hash
linked_accounts = client.linked_accounts().json()
if isinstance(linked_accounts, list) and len(linked_accounts) > 0:
    account_hash = linked_accounts[0].get("hashValue")
elif isinstance(linked_accounts, dict):
    account_securities = linked_accounts.get("securitiesAccountList", [])
    account_hash = (
        account_securities[0].get("hashValue")
        if account_securities
        else linked_accounts.get("hashValue")
    )
else:
    account_hash = None

if not account_hash:
    raise SystemExit("Error: Could not retrieve a valid account hash.")

print("Connecting to Schwab to fetch account positions...")

try:
    # 3. Pull Account Details with the positions parameter enabled
    response = client.account_details(accountHash=account_hash, fields="positions")

    if response.status_code == 200:
        account_data = response.json()
        securities_account = account_data.get("securitiesAccount", {})
        positions = securities_account.get("positions", [])

        if not positions:
            print("⚠️ No open positions found in this account.")
        else:
            print(f"✅ Successfully retrieved {len(positions)} open positions.")

        # 4. Parse positions into a structured list of dictionaries
        parsed_positions = []
        for pos in positions:
            quantity = pos.get("longQuantity", 0) or pos.get("shortQuantity", 0)
            average_price = pos.get("averagePrice", 0.0)
            market_value = pos.get("marketValue", 0.0)

            instrument = pos.get("instrument", {})
            symbol = instrument.get("symbol", "N/A")

            # Fallback check to ensure the full name is captured
            description = (
                instrument.get("description") 
                or instrument.get("name") 
                or instrument.get("fundamental", {}).get("companyName")
                or "N/A"
            )

            parsed_positions.append(
                {
                    "Symbol": symbol,
                    "Company Name": description,  # Clearly named column
                    "Quantity": quantity,
                    "Avg Price": average_price,
                    "Current Market Value": market_value,
                }
            )


        # 5. Convert to a DataFrame and export to Excel
        df = pd.DataFrame(parsed_positions)

        excel_filename = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "trades.xlsx")
        df.to_excel(excel_filename, index=False)
        print(f"📊 Successfully saved results to {excel_filename}")

    else:
        print(
            f"❌ Failed to fetch account positions. Status Code: {response.status_code}"
        )
        print(f"Response: {response.text}")

except Exception as e:
    print(f"💥 Error reaching Schwab API account details endpoint: {e}")

