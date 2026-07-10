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

    def test_gangnam_message_includes_eonju_access_detail(self):
        subject, body = self.sms.build_message(date(2026, 5, 28), self.sms.get_route(date(2026, 5, 28)))

        self.assertEqual(subject, "5/28(목) 동선")
        self.assertIn("① 강남 | 강남 언주로 숙소", body)
        self.assertIn("서울 강남구 봉은사로37길 8", body)
        self.assertIn("건물 앞 정차 권장", body)
        self.assertIn("건물출입 종버튼 +2580", body)
        self.assertIn("엘리베이터 이동", body)
        self.assertIn("5층 엘리베이터 옆 수납창고", body)
        self.assertIn("자물쇠 000*", body)
        self.assertNotIn("상세주소/출입정보 확인 필요", body)

    def test_wangsanro_is_added_on_july_monday_and_thursday_only(self):
        july_thursday = self.sms.get_route(date(2026, 7, 2))
        july_monday = self.sms.get_route(date(2026, 7, 6))
        august_monday = self.sms.get_route(date(2026, 8, 3))

        self.assertIn("왕산로 200, 1004호", july_thursday)
        self.assertIn("왕산로 200, 1004호", july_monday)
        self.assertNotIn("왕산로 200, 1004호", august_monday)

    def test_wangsanro_july_override_keeps_last_thursday_and_august_note(self):
        july_last_thursday = self.sms.get_route(date(2026, 7, 30))
        _, august_body = self.sms.build_message(date(2026, 8, 3), self.sms.get_route(date(2026, 8, 3)))

        self.assertIn("왕산로 200, 1004호", july_last_thursday)
        self.assertIn("청량리는 다음 일정 8/10(월)", august_body)

    def test_owner_sms_failure_does_not_retry_driver_sms(self):
        calls = []

        def fake_send_single(api_key, api_secret, sender, to, text, msg_type="LMS", subject=""):
            calls.append(to)
            if to == "owner":
                raise RuntimeError("owner failed")

        self.sms._send_single = fake_send_single
        old_env = dict(self.sms.os.environ)
        try:
            self.sms.os.environ.update({
                "SOLAPI_API_KEY": "key",
                "SOLAPI_API_SECRET": "secret",
                "SOLAPI_SENDER": "sender",
                "RECIPIENT_PHONE": "driver",
                "OWNER_PHONE": "owner",
            })

            self.sms.send_sms(("subject", "body"), 1)
        finally:
            self.sms.os.environ.clear()
            self.sms.os.environ.update(old_env)

        self.assertEqual(calls, ["driver", "owner"])

    def test_driver_sms_failure_still_fails_workflow(self):
        def fake_send_single(api_key, api_secret, sender, to, text, msg_type="LMS", subject=""):
            raise RuntimeError("driver failed")

        self.sms._send_single = fake_send_single
        old_env = dict(self.sms.os.environ)
        try:
            self.sms.os.environ.update({
                "SOLAPI_API_KEY": "key",
                "SOLAPI_API_SECRET": "secret",
                "SOLAPI_SENDER": "sender",
                "RECIPIENT_PHONE": "driver",
                "OWNER_PHONE": "owner",
            })

            with self.assertRaises(RuntimeError):
                self.sms.send_sms(("subject", "body"), 1)
        finally:
            self.sms.os.environ.clear()
            self.sms.os.environ.update(old_env)


if __name__ == "__main__":
    unittest.main()
