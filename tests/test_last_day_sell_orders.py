import contextlib
from datetime import datetime
import io
import os
from pathlib import Path
import runpy
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch


SCRIPT = Path(__file__).resolve().parents[1] / "LastDaySellOrders.py"


def make_order(status="WORKING", order_id=123, symbol="TMO", instruction="SELL"):
    return {
        "orderId": order_id,
        "status": status,
        "orderStrategyType": "SINGLE",
        "orderLegCollection": [
            {"instruction": instruction, "instrument": {"symbol": symbol}}
        ],
    }


class LastDaySellOrdersTests(unittest.TestCase):
    def run_script(self, orders, *, status_code=200, rows=None, place_status=201,
                   place_error=None):
        client = Mock()
        client.linked_accounts.return_value.json.return_value = [{"hashValue": "test"}]
        response = client.account_orders.return_value
        response.status_code = status_code
        response.json.return_value = orders
        response.text = "Test API error"
        client.replace_order.return_value.status_code = 201
        client.replace_order.return_value.headers = {"location": "/orders/456"}
        client.place_order.return_value.status_code = place_status
        client.place_order.return_value.headers = {"location": "/orders/789"}
        client.place_order.return_value.text = "Test placement error"
        client.place_order.side_effect = place_error
        if rows is None:
            rows = [{
                "symbol": "TMO", "instruction": "SELL", "quantity": 2,
                "orderType": "LIMIT", "duration": "GOOD_TILL_CANCEL",
                "session": "NORMAL", "price": 600,
            }]
        class Frame:
            def __len__(self):
                return len(rows)

            def iterrows(self):
                return enumerate(rows)

        modules = {
            "schwabdev": SimpleNamespace(Client=Mock(return_value=client)),
            "schwab_tokens": SimpleNamespace(TOKENS_DB="unused"),
            "dotenv": SimpleNamespace(load_dotenv=lambda: None),
            "pandas": SimpleNamespace(read_excel=Mock(return_value=Frame())),
        }
        output = io.StringIO()
        error = None
        with patch.dict("sys.modules", modules), patch.dict(os.environ, {
            "SCHWAB_CLIENT_ID": "test", "SCHWAB_CLIENT_SECRET": "test",
            "SCHWAB_CALLBACK_URL": "https://localhost",
        }), contextlib.redirect_stdout(output):
            try:
                runpy.run_path(str(SCRIPT), run_name="__main__")
            except SystemExit as exc:
                error = str(exc)
        return client, output.getvalue(), error

    def test_queued_order_is_replaced(self):
        client, _, error = self.run_script([make_order(status="QUEUED")])
        self.assertIsNone(error)
        client.replace_order.assert_called_once()
        self.assertEqual(client.replace_order.call_args.args[:2], ("test", "123"))
        client.place_order.assert_not_called()

    def test_search_covers_only_last_30_days_and_all_statuses(self):
        client, _, _ = self.run_script([])
        query = client.account_orders.call_args.kwargs
        start = datetime.fromisoformat(query["fromEnteredTime"])
        end = datetime.fromisoformat(query["toEnteredTime"])
        self.assertEqual((end - start).total_seconds(), 30 * 24 * 60 * 60)
        self.assertNotIn("status", query)
        self.assertEqual(query["maxResults"], 3000)

    def test_working_order_payload_is_preserved(self):
        client, _, error = self.run_script([make_order()])
        self.assertIsNone(error)
        payload = client.replace_order.call_args.args[2]
        self.assertEqual(payload["price"], "600.00")
        self.assertEqual(payload["orderLegCollection"][0]["quantity"], 2)
        self.assertEqual(payload["orderStrategyType"], "SINGLE")

    def test_terminal_and_pending_orders_are_not_replaced(self):
        for status in ("FILLED", "CANCELED", "REJECTED", "EXPIRED",
                       "PENDING_CANCEL", "PENDING_REPLACE"):
            with self.subTest(status=status):
                client, _, _ = self.run_script([make_order(status=status)])
                client.replace_order.assert_not_called()

    def test_api_error_aborts_instead_of_reporting_missing_orders(self):
        client, output, error = self.run_script([], status_code=401)
        self.assertIsNotNone(error)
        self.assertIn("401", error)
        self.assertNotIn("Skipped row", output)
        client.replace_order.assert_not_called()
        client.place_order.assert_not_called()

    def test_invalid_response_aborts(self):
        client, _, error = self.run_script({"error": "invalid response"})
        self.assertIsNotNone(error)
        client.replace_order.assert_not_called()
        client.place_order.assert_not_called()

    def test_ambiguous_orders_are_not_replaced(self):
        client, output, _ = self.run_script([make_order(), make_order(order_id=124)])
        client.replace_order.assert_not_called()
        self.assertIn("multiple", output.lower())
        client.place_order.assert_not_called()

    def test_order_without_id_is_not_replaced(self):
        order = make_order()
        del order["orderId"]
        client, _, _ = self.run_script([order])
        client.replace_order.assert_not_called()
        client.place_order.assert_not_called()

    def test_other_symbol_and_instruction_are_not_replaced(self):
        client, _, _ = self.run_script([
            make_order(symbol="MSFT"), make_order(instruction="BUY")
        ])
        client.replace_order.assert_not_called()
        client.place_order.assert_called_once()

    def test_truncated_response_aborts(self):
        client, _, error = self.run_script([make_order()] * 3000)
        self.assertIsNotNone(error)
        self.assertIn("results may be incomplete", error)
        client.replace_order.assert_not_called()
        client.place_order.assert_not_called()

    def test_conditional_and_multi_leg_orders_are_not_replaced(self):
        parent = make_order()
        parent["childOrderStrategies"] = [make_order(order_id=124, symbol="MSFT")]
        multi_leg = make_order()
        multi_leg["orderLegCollection"].append(
            {"instruction": "BUY", "instrument": {"symbol": "MSFT"}}
        )
        for order in (parent, multi_leg, {
            **make_order(), "orderStrategyType": "OCO"
        }):
            with self.subTest(order=order):
                client, _, _ = self.run_script([order])
                client.replace_order.assert_not_called()
                client.place_order.assert_not_called()

    def test_duplicate_rows_do_not_reuse_the_original_order_id(self):
        row = {
            "symbol": "TMO", "instruction": "SELL", "quantity": 2,
            "orderType": "LIMIT", "duration": "GOOD_TILL_CANCEL", "price": 600,
        }
        client, _, _ = self.run_script([make_order()], rows=[row, row])
        client.replace_order.assert_called_once()
        client.place_order.assert_not_called()

    def test_extended_hours_market_order_is_still_skipped(self):
        row = {
            "symbol": "TMO", "instruction": "SELL", "quantity": 2,
            "orderType": "MARKET", "duration": "DAY", "session": "SEAMLESS",
        }
        client, _, _ = self.run_script([make_order()], rows=[row])
        client.replace_order.assert_not_called()
        client.place_order.assert_not_called()

    def test_missing_sell_places_exact_excel_payload(self):
        client, output, error = self.run_script([])
        self.assertIsNone(error)
        client.replace_order.assert_not_called()
        client.place_order.assert_called_once_with("test", {
            "orderType": "LIMIT", "session": "NORMAL",
            "duration": "GOOD_TILL_CANCEL", "orderStrategyType": "SINGLE",
            "price": "600.00",
            "orderLegCollection": [{
                "instruction": "SELL", "quantity": 2,
                "instrument": {"symbol": "TMO", "assetType": "EQUITY"},
            }],
        })
        self.assertIn("789", output)

    def test_terminal_sell_allows_new_sell(self):
        for status in ("FILLED", "CANCELED", "EXPIRED", "REJECTED", "REPLACED"):
            with self.subTest(status=status):
                client, _, _ = self.run_script([make_order(status=status)])
                client.place_order.assert_called_once()
                client.replace_order.assert_not_called()

    def test_unresolved_sell_blocks_new_sell(self):
        for status in ("PENDING_CANCEL", "PENDING_REPLACE", "UNKNOWN"):
            with self.subTest(status=status):
                client, _, _ = self.run_script([make_order(status=status)])
                client.place_order.assert_not_called()
                client.replace_order.assert_not_called()

    def test_nested_sell_is_replaced_by_child_id(self):
        parent = make_order(symbol="MSFT")
        parent["childOrderStrategies"] = [
            make_order(status="AWAITING_PARENT_ORDER", order_id=124)
        ]
        client, _, _ = self.run_script([parent])
        client.place_order.assert_not_called()
        client.replace_order.assert_called_once()
        self.assertEqual(client.replace_order.call_args.args[:2], ("test", "124"))

    def test_existing_sell_in_other_live_statuses_is_replaced(self):
        for status in ("AWAITING_PARENT_ORDER", "AWAITING_STOP_CONDITION",
                       "ACCEPTED", "PENDING_ACTIVATION", "AWAITING_MANUAL_REVIEW"):
            with self.subTest(status=status):
                client, _, _ = self.run_script([make_order(status=status)])
                client.place_order.assert_not_called()
                client.replace_order.assert_called_once()

    def test_multiple_nested_sells_are_not_arbitrarily_replaced(self):
        parent = {
            "orderStrategyType": "OCO",
            "childOrderStrategies": [
                make_order(order_id=124), make_order(order_id=125)
            ],
        }
        client, output, _ = self.run_script([parent])
        self.assertIn("Multiple", output)
        client.replace_order.assert_not_called()
        client.place_order.assert_not_called()

    def test_blocked_sell_reports_order_details(self):
        client, output, _ = self.run_script([make_order(status="PENDING_CANCEL")])
        self.assertIn("ID=123", output)
        self.assertIn("status=PENDING_CANCEL", output)
        client.replace_order.assert_not_called()
        client.place_order.assert_not_called()

    def test_missing_buy_is_not_created(self):
        row = {
            "symbol": "TMO", "instruction": "BUY", "quantity": 2,
            "orderType": "LIMIT", "duration": "DAY", "price": 600,
        }
        client, _, _ = self.run_script([], rows=[row])
        client.place_order.assert_not_called()
        client.replace_order.assert_not_called()

    def test_duplicate_missing_sells_are_not_submitted_twice(self):
        row = {
            "symbol": "TMO", "instruction": "SELL", "quantity": 2,
            "orderType": "LIMIT", "duration": "DAY", "price": 600,
        }
        for options in ({}, {"place_status": 400},
                        {"place_error": TimeoutError("timed out")}):
            with self.subTest(options=options):
                client, output, _ = self.run_script([], rows=[row, row], **options)
                client.place_order.assert_called_once()
                client.replace_order.assert_not_called()
                if options:
                    self.assertNotIn("Success:", output)

    def test_missing_market_sell_obeys_extended_hours_guard(self):
        row = {
            "symbol": "TMO", "instruction": "SELL", "quantity": 2,
            "orderType": "MARKET", "duration": "DAY", "session": "SEAMLESS",
        }
        client, _, _ = self.run_script([], rows=[row])
        client.place_order.assert_not_called()


if __name__ == "__main__":
    unittest.main()
