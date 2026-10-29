import importlib.util
import io
from contextlib import redirect_stdout
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

    def test_driver_pickup_request_is_in_each_scheduled_message(self):
        request = "기사님, 수거 동선에 안내된 숙소는 빠짐없이 방문해 주시고, 검은 가방도 모두 수거해 주시면 감사하겠습니다."
        for day in (date(2026, 11, 2), date(2026, 11, 5), date(2026, 11, 7)):
            with self.subTest(day=day):
                route = self.sms.get_route(day)
                _, body = self.sms.build_message(day, route)
                self.assertEqual(body.count(request), 1)
                self.assertIn("안전 운전하시고 감사합니다!", body)

    def test_saturday_route_contains_only_jangchung(self):
        self.assertEqual(self.sms.get_route(date(2026, 8, 19)), [])
        self.assertEqual(
            self.sms.get_route(date(2026, 8, 22)),
            ["장충단로 225"],
        )
        self.assertEqual(
            self.sms.get_route(date(2026, 8, 29)),
            ["장충단로 225"],
        )

    def test_saturday_message_contains_no_other_route_notes(self):
        route_date = date(2026, 8, 22)
        subject, body = self.sms.build_message(route_date, self.sms.get_route(route_date))

        self.assertEqual(subject, "8/22(토) 동선")
        self.assertIn("① 장충동 | 메종드브릭", body)
        self.assertIn("장충단로 225", body)
        self.assertNotIn("청량리", body)
        self.assertNotIn("↓", body)

    def test_workflow_schedules_each_monday_thursday_and_saturday_slot(self):
        workflow_path = Path(__file__).resolve().parents[1] / ".github" / "workflows" / "send-route-sms.yml"
        workflow = workflow_path.read_text(encoding="utf-8")

        self.assertEqual(workflow.count("* * 1,4,6"), 4)
        self.assertNotIn("* * 1,3,4", workflow)

    def test_main_previews_saturday_without_sending_in_dry_run(self):
        old_env = dict(self.sms.os.environ)
        try:
            self.sms.os.environ.update({
                "TEST_DATE": "2026-08-22",
                "DRY_RUN": "true",
            })
            self.sms.send_sms = lambda *_args: self.fail("dry-run must not send SMS")
            output = io.StringIO()

            with redirect_stdout(output):
                self.sms.main()
        finally:
            self.sms.os.environ.clear()
            self.sms.os.environ.update(old_env)

        self.assertIn("8/22(토) 동선", output.getvalue())
        self.assertIn("① 장충동 | 메종드브릭", output.getvalue())
        self.assertIn("[DRY_RUN] 실제 발송 안 함", output.getvalue())

    def test_main_skips_days_outside_monday_thursday_and_saturday(self):
        old_env = dict(self.sms.os.environ)
        try:
            self.sms.os.environ["TEST_DATE"] = "2026-08-19"
            self.sms.send_sms = lambda *_args: self.fail("Wednesday must not send SMS")
            output = io.StringIO()

            with redirect_stdout(output):
                self.sms.main()
        finally:
            self.sms.os.environ.clear()
            self.sms.os.environ.update(old_env)

        self.assertIn("[SKIP] 수요일은 발송 대상 아님", output.getvalue())

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

    def test_wangsanro_keeps_july_monday_thursday_schedule_from_august(self):
        july_thursday = self.sms.get_route(date(2026, 7, 2))
        july_monday = self.sms.get_route(date(2026, 7, 6))
        august_monday = self.sms.get_route(date(2026, 8, 3))
        august_thursday = self.sms.get_route(date(2026, 8, 6))
        _, august_body = self.sms.build_message(date(2026, 8, 3), august_monday)

        self.assertIn("왕산로 200, 1004호", july_thursday)
        self.assertIn("왕산로 200, 1004호", july_monday)
        self.assertIn("왕산로 200, 1004호", august_monday)
        self.assertIn("왕산로 200, 1004호", august_thursday)
        self.assertNotIn("청량리는 다음 일정", august_body)

    def test_wangsanro_switches_to_weekly_monday_from_november_2(self):
        for day, included in [(date(2026, 10, 29), True),
                              (date(2026, 11, 2), True),
                              (date(2026, 11, 5), False),
                              (date(2026, 11, 9), True),
                              (date(2026, 11, 12), False)]:
            with self.subTest(day=day):
                route = self.sms.get_route(day)
                self.assertEqual(self.sms.WANGSANRO_KEY in route, included)

    def test_thursday_notes_point_to_the_immediately_following_monday(self):
        for day, expected in [(date(2026, 11, 5), "청량리는 다음 일정 11/9(월)"),
                              (date(2026, 11, 12), "청량리는 다음 일정 11/16(월)")]:
            with self.subTest(day=day):
                _, body = self.sms.build_message(day, self.sms.get_route(day))
                self.assertIn(expected, body)

    def test_november_change_preserves_other_stops_and_order(self):
        expected = ["봉은사로37길 8", "가락로28길 3-10", "능동로 165-1",
                    "회기로 189", "고산자로 508-3", "장충단로 225",
                    "회나무로 50", "연희로4길 25-7", "통일로 863-10"]
        for day in [date(2026, 10, 29), date(2026, 11, 2), date(2026, 11, 5),
                    date(2026, 11, 9), date(2026, 11, 12)]:
            with self.subTest(day=day):
                route = self.sms.get_route(day)
                self.assertEqual([key for key in route if key != self.sms.WANGSANRO_KEY], expected)
                if self.sms.WANGSANRO_KEY in route:
                    self.assertEqual(route[3], self.sms.WANGSANRO_KEY)
                    self.assertNotIn("청량리는 다음 일정", self.sms.build_message(day, route)[1])
        self.assertEqual(self.sms.get_route(date(2026, 11, 3)), [])
        self.assertEqual(self.sms.get_route(date(2026, 11, 7)), [self.sms.JANGCHUNG_KEY])

    def test_wangsanro_biweekly_schedule_is_preserved_before_july_override(self):
        june_first_monday = self.sms.get_route(date(2026, 6, 1))
        june_second_monday = self.sms.get_route(date(2026, 6, 8))

        self.assertIn("왕산로 200, 1004호", june_first_monday)
        self.assertNotIn("왕산로 200, 1004호", june_second_monday)

    def test_itaewon_route_starts_on_2026_08_03(self):
        july_23_route = self.sms.get_route(date(2026, 7, 23))
        july_30_route = self.sms.get_route(date(2026, 7, 30))
        august_3_route = self.sms.get_route(date(2026, 8, 3))

        self.assertNotIn("회나무로 50", july_23_route)
        self.assertNotIn("회나무로 50", july_30_route)
        self.assertIn("회나무로 50", august_3_route)

    def test_itaewon_is_added_after_jangchung_from_august(self):
        route = self.sms.get_route(date(2026, 8, 3))

        self.assertEqual(route, [
            "봉은사로37길 8",
            "가락로28길 3-10",
            "능동로 165-1",
            "왕산로 200, 1004호",
            "회기로 189",
            "고산자로 508-3",
            "장충단로 225",
            "회나무로 50",
            "연희로4길 25-7",
        ])

    def test_itaewon_message_includes_access_detail(self):
        route = self.sms.get_route(date(2026, 8, 3))
        _, body = self.sms.build_message(date(2026, 8, 3), route)

        self.assertIn("이태원 | 이태원 숙소", body)
        self.assertIn("서울특별시 용산구 회나무로 50 (이태원동)", body)
        self.assertIn("5층 엘베 내려 반층위 옥상문앞", body)
        self.assertIn("공동현관 비밀번호: [🗝️열쇠] + 3571 + [🔔종]", body)
        self.assertNotIn("렉 설치 예정", body)

    def test_itaewon_is_present_twice_weekly_from_august_3(self):
        for route_date in [
            date(2026, 8, 3),
            date(2026, 8, 6),
            date(2026, 8, 10),
            date(2026, 8, 13),
        ]:
            with self.subTest(route_date=route_date):
                self.assertIn("회나무로 50", self.sms.get_route(route_date))

        self.assertEqual(self.sms.get_route(date(2026, 8, 4)), [])

    def test_eunpyeong_is_added_after_yeonnam_from_september_7(self):
        self.assertNotIn("통일로 863-10", self.sms.get_route(date(2026, 9, 3)))

        for route_date in [date(2026, 9, 7), date(2026, 9, 10)]:
            with self.subTest(route_date=route_date):
                route = self.sms.get_route(route_date)
                self.assertEqual(route[-2:], ["연희로4길 25-7", "통일로 863-10"])

    def test_eunpyeong_message_includes_ground_floor_storage_and_no_elevator(self):
        route = self.sms.get_route(date(2026, 9, 7))
        _, body = self.sms.build_message(date(2026, 9, 7), route)

        self.assertIn("은평 | 은평 숙소", body)
        self.assertIn("서울 은평구 통일로 863-10", body)
        self.assertTrue("정문현관 - 5052*" in body, "은평 정문 출입안내 누락")
        self.assertIn("엘리베이터 없음", body)
        self.assertIn("1층 세탁물 보관", body)

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
