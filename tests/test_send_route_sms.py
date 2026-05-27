import importlib.util
from datetime import date
from pathlib import Path
import unittest


def load_send_route_sms():
    module_path = Path(__file__).resolve().parents[1] / "scripts" / "send_route_sms.py"
    spec = importlib.util.spec_from_file_location("send_route_sms", module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class SendRouteSmsTests(unittest.TestCase):
    def setUp(self):
        self.sms = load_send_route_sms()

    def test_gangnam_route_starts_on_2026_05_28_with_bupyeong_roundtrip_order(self):
        route = self.sms.get_route(date(2026, 5, 28))

        self.assertEqual(route, [
            "봉은사로37길 8",
            "가락로28길 3-10",
            "능동로 165-1",
            "장한로26나길 21",
            "회기로 189",
            "고산자로 508-3",
            "장충단로 225",
            "연희로4길 25-7",
        ])

    def test_gangnam_route_keeps_wangsanro_biweekly_monday_slot(self):
        route = self.sms.get_route(date(2026, 6, 1))

        self.assertEqual(route, [
            "봉은사로37길 8",
            "가락로28길 3-10",
            "능동로 165-1",
            "왕산로 200, 1004호",
            "회기로 189",
            "고산자로 508-3",
            "장충단로 225",
            "연희로4길 25-7",
        ])

    def test_route_before_gangnam_start_date_keeps_existing_order(self):
        route = self.sms.get_route(date(2026, 5, 25))

        self.assertEqual(route, [
            "연희로4길 25-7",
            "장충단로 225",
            "고산자로 508-3",
            "회기로 189",
            "능동로 165-1",
            "가락로28길 3-10",
        ])

    def test_gangnam_message_marks_missing_detail(self):
        subject, body = self.sms.build_message(date(2026, 5, 28), self.sms.get_route(date(2026, 5, 28)))

        self.assertEqual(subject, "5/28(목) 동선")
        self.assertIn("① 강남 | 신규 숙소", body)
        self.assertIn("봉은사로37길 8", body)
        self.assertIn("상세주소/출입정보 확인 필요", body)


if __name__ == "__main__":
    unittest.main()
