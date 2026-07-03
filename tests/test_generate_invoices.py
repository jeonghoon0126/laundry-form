import importlib.util
import os
from pathlib import Path
import sys
import unittest
from unittest.mock import patch


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
        self.assertEqual(summary["supplies_cost"], 447_251)
        self.assertEqual(summary["withholding_tax"], 105_746)
        self.assertEqual(summary["withholding_national_tax"], 96_133)
        self.assertEqual(summary["withholding_local_tax"], 9_613)
        self.assertEqual(summary["output_vat"], 631_700)
        self.assertEqual(summary["input_vat_credit"], 147_934)
        self.assertEqual(summary["vat"], 483_766)
        self.assertEqual(summary["total_cost"], 5_421_206)
        self.assertEqual(summary["operating_profit"], 1_527_494)
        self.assertAlmostEqual(summary["operating_margin"], 0.2198, places=4)

    def test_profit_summary_text_is_ready_for_email(self):
        summary = self.gi.calculate_profit_summary(6_948_700)
        close = self.gi.calculate_monthly_close_summary(6_948_700)
        text = self.gi.format_profit_summary_text(summary, close)

        self.assertIn("수입", text)
        self.assertIn("정산 매출: 6,948,700원", text)
        self.assertIn("지출", text)
        self.assertIn("운영비 합계: 4,831,694원", text)
        self.assertIn("세금/적립 합계: 726,048원", text)
        self.assertIn("영업이익률: 22.0%", text)
        self.assertIn("인건비: 2,124,443원", text)
        self.assertIn("매입세액공제 항목", text)
        self.assertIn("월세+관리비 매입세액: 70,000원", text)
        self.assertIn("소모품 매입세액: 40,660원", text)
        self.assertIn("부가세 계산", text)
        self.assertIn("매출 부가세: 631,700원", text)
        self.assertIn("매입세액공제: 147,934원", text)
        self.assertIn("납부 예상 부가세: 483,766원", text)
        self.assertIn("종소세 적립 기준", text)
        self.assertIn("월 적립액: 136,536원", text)
        self.assertIn("종소세 적립 후 순수익: 1,390,958원", text)

    def test_profit_summary_html_groups_email_breakdown_sections(self):
        summary = self.gi.calculate_profit_summary(7_457_500)
        close = self.gi.calculate_monthly_close_summary(7_457_500)
        html = self.gi.format_profit_summary_html(summary, close)

        self.assertIn("구분", html)
        self.assertIn("항목", html)
        self.assertIn("금액", html)
        self.assertIn("설명", html)
        self.assertIn("수입", html)
        self.assertIn("지출 - 운영비", html)
        self.assertIn("지출 - 세금/적립", html)
        self.assertIn("매입세액공제 항목", html)
        self.assertIn("부가세 계산", html)
        self.assertIn("종소세 적립 기준", html)
        self.assertIn("최종 순수익", html)
        self.assertIn("정산 매출", html)
        self.assertIn("월세+관리비 매입세액", html)
        self.assertIn("납부 예상 부가세", html)
        self.assertIn("종소세 적립 후 순수익", html)

    def test_monthly_close_text_shows_cash_itemized_when_inputs_exist(self):
        close = self.gi.calculate_monthly_close_summary(
            7_457_500,
            kops_receivable_amount=2_409_500,
            current_bank_balance=470_000,
            late_receipt_amount=2_409_500,
            payroll_to_pay=2_280_000,
            owner_draw=1_000_000,
            income_tax_reserve_rate=0.10,
        )
        text = self.gi.format_monthly_close_text(close)

        self.assertIn("월마감 현금 기준", text)
        self.assertIn("현재 통장 잔액: 470,000원", text)
        self.assertIn("콥스 입금 후 현금: 2,879,500원", text)
        self.assertIn("인건비 지급 후 현금: 599,500원", text)
        self.assertIn("이미 빼간 영업이익: 1,000,000원", text)
        self.assertIn("추가로 남길 수 있는 이익: 619,618원", text)

    def test_monthly_close_html_shows_cash_itemized_when_inputs_exist(self):
        close = self.gi.calculate_monthly_close_summary(
            7_457_500,
            kops_receivable_amount=2_409_500,
            current_bank_balance=470_000,
            late_receipt_amount=2_409_500,
            payroll_to_pay=2_280_000,
            owner_draw=1_000_000,
            income_tax_reserve_rate=0.10,
        )
        html = self.gi.format_monthly_close_html(close)

        self.assertIn("월마감 현금 기준", html)
        self.assertIn("현재 통장 잔액", html)
        self.assertIn("콥스 입금 후 현금", html)
        self.assertIn("인건비 지급 후 현금", html)
        self.assertIn("이미 빼간 영업이익", html)
        self.assertIn("추가로 남길 수 있는 이익", html)

    def test_profit_summary_for_may_baseline_revenue(self):
        summary = self.gi.calculate_profit_summary(7_457_500)

        self.assertEqual(summary["labor_cost"], 2_280_000)
        self.assertEqual(summary["electricity"], 250_000)
        self.assertEqual(summary["supplies_cost"], 480_000)
        self.assertEqual(summary["withholding_tax"], 110_880)
        self.assertEqual(summary["withholding_national_tax"], 100_800)
        self.assertEqual(summary["withholding_local_tax"], 10_080)
        self.assertEqual(summary["output_vat"], 677_955)
        self.assertEqual(summary["input_vat_credit"], 150_911)
        self.assertEqual(summary["vat"], 527_044)
        self.assertEqual(summary["total_cost"], 5_657_924)
        self.assertEqual(summary["operating_profit"], 1_799_576)
        self.assertAlmostEqual(summary["operating_margin"], 0.2413, places=4)

    def test_monthly_close_separates_profit_from_cash_timing(self):
        close = self.gi.calculate_monthly_close_summary(
            7_457_500,
            kops_receivable_amount=2_409_500,
            current_bank_balance=470_000,
            late_receipt_amount=2_409_500,
            payroll_to_pay=2_280_000,
            owner_draw=1_000_000,
            income_tax_reserve_rate=0.10,
        )

        self.assertEqual(close["profit_summary"]["operating_profit"], 1_799_576)
        self.assertEqual(close["income_tax_reserve"], 179_958)
        self.assertEqual(close["net_profit_after_income_tax_reserve"], 1_619_618)
        self.assertEqual(close["remaining_profit_after_owner_draw"], 619_618)
        self.assertEqual(close["cash_timing"]["month_end_collection"], 5_048_000)
        self.assertEqual(close["cash_timing"]["delayed_kops_collection"], 2_409_500)
        self.assertEqual(close["cash_timing"]["cash_after_late_receipt"], 2_879_500)
        self.assertEqual(close["cash_timing"]["cash_after_payroll"], 599_500)
        self.assertEqual(close["cash_timing"]["logistics_payments"], [
            {"day": 15, "amount": 540_000},
            {"day": 30, "amount": 540_000},
        ])

    def test_monthly_close_uses_annualized_monthly_profit_for_income_tax_reserve(self):
        close = self.gi.calculate_monthly_close_summary(8_723_000)

        self.assertEqual(close["profit_summary"]["operating_profit"], 2_476_310)
        self.assertEqual(close["income_tax_reserve_detail"]["annual_tax_base"], 29_715_720)
        self.assertEqual(close["income_tax_reserve_detail"]["income_tax_rate"], 0.15)
        self.assertEqual(close["income_tax_reserve_detail"]["progressive_deduction"], 1_260_000)
        self.assertEqual(close["income_tax_reserve_detail"]["annual_income_tax"], 3_197_358)
        self.assertEqual(close["income_tax_reserve_detail"]["annual_local_income_tax"], 319_736)
        self.assertEqual(close["income_tax_reserve_detail"]["annual_income_tax_total"], 3_517_094)
        self.assertEqual(close["income_tax_reserve"], 293_091)
        self.assertEqual(close["net_profit_after_income_tax_reserve"], 2_183_219)

    def test_business_totals_find_kops_receivable_for_monthly_close(self):
        business_data = {
            self.gi.KOPS_REG_NO: {
                "locations": {
                    "서대문구 연희로4길 25-7": {
                        "blanket": 10,
                        "mat": 5,
                        "pillow_cover": 20,
                        "towel": 30,
                        "body_towel": 0,
                        "pillow_fill": 0,
                        "cotton_blanket": 0,
                    }
                },
                "extra_items": [],
            },
            "554-88-03481": {
                "locations": {
                    "중구 장충단로 225": {
                        "blanket": 7,
                        "mat": 3,
                        "pillow_cover": 12,
                        "towel": 18,
                        "body_towel": 4,
                        "pillow_fill": 0,
                        "cotton_blanket": 0,
                    }
                },
                "extra_items": [{"name": "adjustment", "qty": 1, "price": 10_000, "amount": 10_000}],
            },
        }

        kops_total = self.gi.calculate_kops_receivable_amount(business_data)
        total = self.gi.calculate_total_amount_from_business_data(business_data)

        self.assertEqual(kops_total, 60_000)
        self.assertEqual(total, 113_400)

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
        self.assertEqual(values_by_range["영업이익계산!N7"], 105_746)
        self.assertEqual(values_by_range["영업이익계산!O7"], 483_766)
        self.assertEqual(values_by_range["영업이익계산!P7"], 5_421_206)
        self.assertEqual(values_by_range["영업이익계산!Q7"], 1_527_494)
        self.assertAlmostEqual(values_by_range["영업이익계산!R7"], 0.2198, places=4)
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

    def test_resolve_invoice_sheet_name_allows_spacing_differences(self):
        sheet_titles = [
            "invoice(거래명세서)_강남구 봉은사로 37길 8",
            "invoice(거래명세서)_장충동 메종 드 브릭",
        ]

        self.assertEqual(
            self.gi.resolve_invoice_sheet_name("강남구 봉은사로37길 8", sheet_titles),
            "invoice(거래명세서)_강남구 봉은사로 37길 8",
        )
        self.assertEqual(
            self.gi.resolve_invoice_sheet_name("중구 장충단로 225", sheet_titles),
            "invoice(거래명세서)_장충동 메종 드 브릭",
        )

    def test_invoice_sheets_only_mode_skips_email_generation(self):
        calls = {}
        rows = [
            (
                self.gi.date(2026, 6, 1),
                "강남구 봉은사로37길 8",
                1,
                2,
                3,
                4,
                0,
                0,
                0,
            )
        ]

        self.gi.get_monthly_data = lambda year, month: rows
        self.gi.update_invoice_sheets = lambda year, month, row_data: calls.setdefault(
            "invoice_sheets", (year, month, row_data)
        ) or True

        def fail_if_called(*args, **kwargs):
            raise AssertionError("email/PDF generation should not run in invoice_sheets_only mode")

        self.gi.generate_pdf = fail_if_called
        self.gi.generate_excel = fail_if_called
        self.gi.send_report_email = fail_if_called
        self.gi.send_email = fail_if_called
        self.gi.update_profit_sheet = fail_if_called

        with patch.dict(os.environ, {"INVOICE_JOB_MODE": "invoice_sheets_only"}), \
                patch.object(sys, "argv", ["generate_invoices.py", "2026", "6"]):
            self.gi.main()

        self.assertEqual(calls["invoice_sheets"], (2026, 6, rows))


if __name__ == "__main__":
    unittest.main()
