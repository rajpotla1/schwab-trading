import sys
import os
from datetime import datetime, date
import pytz
from dotenv import load_dotenv
from schwab import auth, client
from schwab.orders.generic import OrderBuilder
from schwab.orders.common import OrderType, Duration, Session



# 1. Load hidden variables from the .env file
load_dotenv()
client_id = os.getenv("SCHWAB_CLIENT_ID")
client_secret = os.getenv("SCHWAB_CLIENT_SECRET")
callback_url = os.getenv("SCHWAB_CALLBACK_URL")

if not all([client_id, client_secret, callback_url]):
    raise ValueError("Error: One or more sensitive keys failed to load from the .env file.")

# --- ADDITIONAL SETUP ---
TOKEN_PATH = 'token.json'
ACCOUNT_ID = os.getenv("SCHWAB_ACCOUNT_ID", "YOUR_SCHWAB_ACCOUNT_ID")

# --- SAFETY CLOCK GUARDRAIL ---
def verify_market_cutoff():
    """
    Checks the current US/Eastern time. Aborts if it's past the 3:45 PM EST safety buffer
    to ensure orders aren't queued up for the next trading day.
    """
    est_timezone = pytz.timezone('US/Eastern')
    now_est = datetime.now(est_timezone)
    
    cutoff_hour = 15
    cutoff_minute = 45
    
    print(f"Current System Time (EST): {now_est.strftime('%Y-%m-%d %I:%M:%S %p')}")
    
    # Explicitly defining weekend days (5 = Saturday, 6 = Sunday)
    weekend_days = [5, 6]
    if now_est.weekday() in weekend_days:
        print("CRITICAL EXCEPTION: Today is a weekend. Schwab mutual funds do not price or settle on weekends.")
        return False
        
    if now_est.hour > cutoff_hour or (now_est.hour == cutoff_hour and now_est.minute >= cutoff_minute):
        print(f"CRITICAL EXCEPTION: Execution halted. Current time is past the 3:45 PM EST safety buffer.")
        print("Placing a mutual fund liquidation now would mistakenly queue the transaction for TOMORROW's close.")
        return False
        
    print("Market time window verified: Safe to proceed with intraday settlement calculations.")
    return True

def get_total_buying_cost_today(c, account_id):
    """
    Fetches today's orders and calculates total cost of completed or filled stock/ETF buy trades.
    """
    total_cost = 0.0
    
    # 1. Get the current date in the EST timezone (matching your guardrail logic)
    est_timezone = pytz.timezone('US/Eastern')
    now_est = datetime.now(est_timezone)
    
    # 2. Create native datetime objects for the absolute start and end of the day in EST
    start_dt = est_timezone.localize(datetime(now_est.year, now_est.month, now_est.day, 0, 0, 0))
    end_dt = est_timezone.localize(datetime(now_est.year, now_est.month, now_est.day, 23, 59, 59))
    
    # FIX: Pass the native datetime objects directly to the client wrapper
    response = c.get_orders_for_account(
        account_id, 
        from_entered_datetime=start_dt, 
        to_entered_datetime=end_dt
    )
    
    if response.status_code != 200:
        print(f"Error fetching orders: {response.text}")
        return 0.0
        
    orders = response.json()
    valid_order_statuses = ["FILLED", "PARTIALLY_FILLED"]
    
    for order in orders:
        status = order.get('status', '')
        if status in valid_order_statuses:
            for activity in order.get('orderActivityCollection', []):
                if activity.get('activityType') == 'EXECUTION':
                    for leg in order.get('orderLegCollection', []):
                        if leg.get('instruction') == 'BUY' and leg.get('assetType') in ['EQUITY']:
                            for execution in activity.get('executionLegs', []):
                                price = execution.get('price', 0.0)
                                quantity = execution.get('quantity', 0.0)
                                total_cost += (price * quantity)
                                
    return round(total_cost, 2)

def place_swvxx_sell_order(c, account_id, amount):
    """
    Builds and places a mutual fund liquidation order for SWVXX for a specific dollar amount.
    """
    if amount <= 0:
        print("No buying cost detected today. Skipping SWVXX liquidation.")
        return

    # Using a raw JSON payload structure accepted directly by Schwab's API
    # This bypasses the OrderBuilder structure limitations entirely
    mutual_fund_order_payload = {
        "orderType": "MARKET",
        "session": "NORMAL",
        "duration": "GOOD_TILL_CANCEL",
        "orderStrategyType": "SINGLE",
        "orderLegCollection": [
            {
                "instruction": "SELL",
                "quantity": amount,  # Schwab API uses 'quantity' for dollar value when assetType is MUTUAL_FUND
                "instrument": {
                    "symbol": "SWVXX",
                    "assetType": "MUTUAL_FUND"
                }
            }
        ]
    }
    
    # Transmit the raw dictionary directly to the place_order endpoint
    response = c.place_order(account_id, mutual_fund_order_payload)
    
    http_success_codes = [200, 201, 202]
    if response.status_code in http_success_codes:
        print(f"SUCCESS: Placed sell order for ${amount} of SWVXX.")
    else:
        print(f"FAILED: SWVXX sell order rejected: {response.text}")


if __name__ == "__main__":
    if not verify_market_cutoff():
        print("Script execution terminated automatically.")
        sys.exit(1)
        
    # --- FIXED MULTI-STAGE AUTHENTICATION WORKAROUND ---
    c = None
    try:
        # Stage 1: Try loading the existing unexpired token cleanly without triggering the server check
        if os.path.exists(TOKEN_PATH):
            c = auth.client_from_token_file(TOKEN_PATH, client_id, client_secret)
            print("Successfully authenticated using cached token file.")
    except Exception as token_err:
        print(f"Cached token bypass unavailable or expired, falling back to full login: {token_err}")

 


    # Stage 2: If no token exists or it has fully expired, initialize the safe browser flow
    if c is None:
        try:
            print("Initializing interactive OAuth browser authentication flow...")
            #c = auth.client_from_login_flow(client_id, client_secret, callback_url, TOKEN_PATH)
            c = auth.client_from_manual_flow(client_id, client_secret, callback_url, TOKEN_PATH)
        except Exception as flow_err:
            print(f"Authentication totally failed: {flow_err}")
            sys.exit(1)
        
    # Execute core script actions
    buying_cost = get_total_buying_cost_today(c, ACCOUNT_ID)
    print(f"Total equity buying cost detected for today ({date.today()}): ${buying_cost}")
    
    place_swvxx_sell_order(c, ACCOUNT_ID, buying_cost)
