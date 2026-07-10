import importlib.util
from datetime import date
from pathlib import Path
import unittest


def load_dispatch_route_sms():
    module_path = Path(__file__).resolve().parents[1] / "scripts" / "dispatch_route_sms.py"
    spec = importlib.util.spec_from_file_location("dispatch_route_sms", module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class DispatchRouteSmsTests(unittest.TestCase):
    def setUp(self):
        self.dispatcher = load_dispatch_route_sms()

    def test_find_successful_send_run_ignores_dry_run_success(self):
        runs = [
            {
                "id": 1,
                "conclusion": "success",
                "event": "workflow_dispatch",
                "run_started_at": "2026-07-13T01:00:00Z",
                "display_title": "캐리 동선 문자 발송 today dry-run",
            },
        ]
        self.dispatcher.fetch_workflow_runs = lambda limit, env=None: runs

        existing = self.dispatcher.find_successful_send_run(
            current_day=date(2026, 7, 13),
            target_day=date(2026, 7, 13),
        )

        self.assertIsNone(existing)

    def test_find_successful_send_run_matches_manual_target_date_title(self):
        runs = [
            {
                "id": 2,
                "conclusion": "success",
                "event": "workflow_dispatch",
                "run_started_at": "2026-07-10T01:00:00Z",
                "display_title": "캐리 동선 문자 발송 2026-07-09 send",
            },
        ]
        self.dispatcher.fetch_workflow_runs = lambda limit, env=None: runs

        existing = self.dispatcher.find_successful_send_run(
            current_day=date(2026, 7, 10),
            target_day=date(2026, 7, 9),
        )

        self.assertEqual(existing["id"], 2)


if __name__ == "__main__":
    unittest.main()
