import importlib.util
from pathlib import Path
import unittest


def load_weekly_report():
    module_path = Path(__file__).resolve().parents[1] / "scripts" / "send_weekly_revenue_report.py"
    spec = importlib.util.spec_from_file_location("send_weekly_revenue_report", module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class WeeklyRevenueReportTests(unittest.TestCase):
    def setUp(self):
        self.report = load_weekly_report()

    def test_previous_week_window_uses_last_completed_monday_to_sunday(self):
        start, end = self.report.previous_week_window(self.report.date(2026, 6, 8))

        self.assertEqual(start, self.report.date(2026, 6, 1))
        self.assertEqual(end, self.report.date(2026, 6, 8))

    def test_build_summary_groups_revenue_by_day_and_location(self):
        rows = [
            (self.report.date(2026, 6, 1), "강남구 봉은사로37길 8", 2, 1, 0, 4, 0, 0, 0),
            (self.report.date(2026, 6, 2), "중구 장충단로 225", 0, 0, 0, 0, 3, 0, 0),
            (self.report.date(2026, 6, 2), "동대문구 장한로26나길 21", 10, 0, 0, 0, 0, 0, 0),
            (self.report.date(2026, 6, 7), "동대문구 회기로 189", 0, 0, 0, 0, 0, 0, 0),
        ]

        summary = self.report.build_weekly_summary(rows, self.report.date(2026, 6, 1), self.report.date(2026, 6, 8))

        self.assertEqual(summary["total"], 13_300)
        self.assertEqual(summary["locations"]["강남구 봉은사로37길 8"], 10_000)
        self.assertEqual(summary["locations"]["중구 장충단로 225"], 3_300)
        self.assertNotIn("동대문구 회기로 189", summary["locations"])
        self.assertNotIn("동대문구 장한로26나길 21", summary["locations"])
        self.assertEqual(summary["daily"][self.report.date(2026, 6, 1)], 10_000)
        self.assertEqual(summary["daily"][self.report.date(2026, 6, 2)], 3_300)
        self.assertEqual(summary["record_count"], 2)

    def test_email_defaults_to_requested_recipient(self):
        self.assertEqual(self.report.DEFAULT_RECIPIENT, "kham0126@gmail.com")

    def test_html_contains_weekly_total_and_location_rows(self):
        summary = {
            "start": self.report.date(2026, 6, 1),
            "end_exclusive": self.report.date(2026, 6, 8),
            "total": 13_300,
            "record_count": 2,
            "locations": {
                "강남구 봉은사로37길 8": 10_000,
                "중구 장충단로 225": 3_300,
            },
            "daily": {
                self.report.date(2026, 6, 1): 10_000,
                self.report.date(2026, 6, 2): 3_300,
            },
        }

        html = self.report.build_weekly_email_html(summary)

        self.assertIn("06/01~06/07", html)
        self.assertIn("13,300원", html)
        self.assertIn("강남구 봉은사로37길 8", html)
        self.assertIn("10,000원", html)


if __name__ == "__main__":
    unittest.main()
