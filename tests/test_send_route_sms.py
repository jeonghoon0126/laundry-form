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

    def test_janghanpyeong_route_ends_from_2026_06_01(self):
        route = self.sms.get_route(date(2026, 6, 4))

        self.assertEqual(route, [
            "봉은사로37길 8",
            "가락로28길 3-10",
            "능동로 165-1",
            "회기로 189",
            "고산자로 508-3",
            "장충단로 225",
            "연희로4길 25-7",
        ])

    def test_janghanpyeong_next_note_stops_after_end_date(self):
        _, body = self.sms.build_message(date(2026, 6, 1), self.sms.get_route(date(2026, 6, 1)))

        self.assertNotIn("장한평", body)

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

    def test_gangnam_message_includes_access_detail(self):
        subject, body = self.sms.build_message(date(2026, 5, 28), self.sms.get_route(date(2026, 5, 28)))

        self.assertEqual(subject, "5/28(목) 동선")
        self.assertIn("① 강남 | 신규 숙소", body)
        self.assertIn("서울 강남구 봉은사로37길 8", body)
        self.assertIn("건물출입: 종버튼 +2580", body)
        self.assertIn("5층 엘리베이터 옆 수납창고 자물쇠 000*", body)
        self.assertIn("주차장 협소: 건물 앞 정차 권장", body)

    def test_wangsanro_returns_to_biweekly_monday_cycle_in_august(self):
        july_last_thursday = self.sms.get_route(date(2026, 7, 30))
        august_first_monday = self.sms.get_route(date(2026, 8, 3))
        august_second_monday = self.sms.get_route(date(2026, 8, 10))
        august_third_monday = self.sms.get_route(date(2026, 8, 17))
        august_fourth_monday = self.sms.get_route(date(2026, 8, 24))

        self.assertIn("왕산로 200, 1004호", july_last_thursday)
        self.assertNotIn("왕산로 200, 1004호", august_first_monday)
        self.assertIn("왕산로 200, 1004호", august_second_monday)
        self.assertNotIn("왕산로 200, 1004호", august_third_monday)
        self.assertIn("왕산로 200, 1004호", august_fourth_monday)

    def test_itaewon_is_added_to_august_route_after_jangchung(self):
        route = self.sms.get_route(date(2026, 8, 3))

        self.assertEqual(route, [
            "봉은사로37길 8",
            "가락로28길 3-10",
            "능동로 165-1",
            "회기로 189",
            "고산자로 508-3",
            "장충단로 225",
            "회나무로 50",
            "연희로4길 25-7",
        ])

    def test_itaewon_message_includes_address_and_elevator_note(self):
        route = self.sms.get_route(date(2026, 8, 3))
        _, body = self.sms.build_message(date(2026, 8, 3), route)

        self.assertIn("이태원 | 이태원 숙소", body)
        self.assertIn("서울특별시 용산구 회나무로 50 (이태원동)", body)
        self.assertIn("엘리베이터 있음", body)
        self.assertIn("5층 엘리베이터 진입 후 반층 위 렉 설치 예정", body)


if __name__ == "__main__":
    unittest.main()
