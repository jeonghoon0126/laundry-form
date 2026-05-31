import importlib.util
from pathlib import Path
import unittest


def load_generate_invoices():
    module_path = Path(__file__).resolve().parents[1] / "scripts" / "generate_invoices.py"
    spec = importlib.util.spec_from_file_location("generate_invoices", module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ProfitSummaryTests(unittest.TestCase):
    def setUp(self):
        self.gi = load_generate_invoices()

    def test_profit_summary_uses_revenue_based_labor_and_net_vat(self):
        summary = self.gi.calculate_profit_summary(6_948_700)

        self.assertEqual(summary["revenue"], 6_948_700)
        self.assertEqual(summary["labor_cost"], 2_124_443)
        self.assertEqual(summary["logistics_cost"], 1_080_000)
        self.assertEqual(summary["rent_utility"], 770_000)
        self.assertEqual(summary["electricity"], 250_000)
        self.assertEqual(summary["water"], 100_000)
        self.assertEqual(summary["insurance"], 60_000)
        self.assertEqual(summary["supplies_cost"], 277_948)
        self.assertEqual(summary["withholding_tax"], 105_747)
        self.assertEqual(summary["vat"], 218_056)
        self.assertEqual(summary["total_cost"], 4_986_194)
        self.assertEqual(summary["operating_profit"], 1_962_506)
        self.assertAlmostEqual(summary["operating_margin"], 0.2824, places=4)

    def test_profit_summary_text_is_ready_for_email(self):
        summary = self.gi.calculate_profit_summary(6_948_700)
        text = self.gi.format_profit_summary_text(summary)

        self.assertIn("매출: 6,948,700원", text)
        self.assertIn("총 지출: 4,986,194원", text)
        self.assertIn("영업이익: 1,962,506원", text)
        self.assertIn("영업이익률: 28.2%", text)
        self.assertIn("인건비: 2,124,443원", text)

    def test_profit_summary_for_may_baseline_revenue(self):
        summary = self.gi.calculate_profit_summary(7_457_500)

        self.assertEqual(summary["labor_cost"], 2_280_000)
        self.assertEqual(summary["electricity"], 250_000)
        self.assertEqual(summary["supplies_cost"], 298_300)
        self.assertEqual(summary["withholding_tax"], 110_880)
        self.assertEqual(summary["vat"], 250_832)
        self.assertEqual(summary["total_cost"], 5_200_012)
        self.assertEqual(summary["operating_profit"], 2_257_488)
        self.assertAlmostEqual(summary["operating_margin"], 0.3027, places=4)

    def test_profit_sheet_update_writes_final_profit_columns(self):
        captured = {}

        self.gi.get_sheets_token = lambda: "token"
        self.gi._sheets_get = lambda spreadsheet_id, token, range_str: [
            ["정산월"],
            [],
            ["2025-12"],
            ["2026-1"],
            ["2026-2"],
            ["2026-3"],
            ["2026-4"],
        ]

        def fake_batch_update(spreadsheet_id, token, data):
            captured["data"] = data

        self.gi._sheets_batch_update = fake_batch_update
        self.gi.format_profit_sheet_row = lambda token, target_row: captured.setdefault("formatted_row", target_row)

        self.assertTrue(self.gi.update_profit_sheet(2026, 4, 6_948_700))

        values_by_range = {
            item["range"]: item["values"][0][0]
            for item in captured["data"]
        }
        self.assertEqual(values_by_range["영업이익계산!E7"], 2_124_443)
        self.assertEqual(values_by_range["영업이익계산!N7"], 105_747)
        self.assertEqual(values_by_range["영업이익계산!P7"], 4_986_194)
        self.assertEqual(values_by_range["영업이익계산!Q7"], 1_962_506)
        self.assertAlmostEqual(values_by_range["영업이익계산!R7"], 0.2824, places=4)
        self.assertEqual(captured["formatted_row"], 7)

    def test_gangnam_location_uses_kops_business_and_default_prices(self):
        reg_no, name, owner = self.gi.BUSINESS_MAP["강남구 봉은사로37길 8"]

        self.assertEqual(reg_no, "767-87-02214")
        self.assertEqual(name, "주식회사 콥스")
        self.assertEqual(owner, "남택호")
        self.assertEqual(self.gi.get_location_prices("강남구 봉은사로37길 8"), self.gi.PRICES)

    def test_janghanpyeong_settlement_ends_from_2026_06_01(self):
        location = "동대문구 장한로26나길 21"

        self.assertTrue(self.gi.is_settlement_location_active(location, self.gi.date(2026, 5, 31)))
        self.assertFalse(self.gi.is_settlement_location_active(location, self.gi.date(2026, 6, 1)))


if __name__ == "__main__":
    unittest.main()
