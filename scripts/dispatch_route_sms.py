#!/usr/bin/env python3
"""
Route SMS workflow helper.

- Local launchd: dispatch the GitHub workflow at the exact local time.
- GitHub Actions: skip duplicate same-day sends when a successful run already exists.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

KST = ZoneInfo("Asia/Seoul")
REPO = "jeonghoon0126/laundry-form"
WORKFLOW = "send-route-sms.yml"
TARGET_WEEKDAYS = {0, 3}  # Monday, Thursday
DEFAULT_GH = Path.home() / ".local/bin/gh"


def gh_path() -> str:
    if DEFAULT_GH.exists():
        return str(DEFAULT_GH)

    found = shutil.which("gh")
    if found:
        return found

    raise FileNotFoundError("gh CLI not found")


def run_gh(*args: str, capture_output: bool = False, env: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    command = [gh_path(), *args]
    return subprocess.run(
        command,
        check=True,
        text=True,
        capture_output=capture_output,
        env=env,
    )


def now_kst() -> datetime:
    return datetime.now(KST)


def today_kst() -> date:
    return now_kst().date()


def parse_run_date(timestamp: str) -> date:
    return datetime.fromisoformat(timestamp.replace("Z", "+00:00")).astimezone(KST).date()


def fetch_workflow_runs(limit: int, env: dict[str, str] | None = None) -> list[dict[str, Any]]:
    result = run_gh(
        "api",
        f"repos/{REPO}/actions/workflows/{WORKFLOW}/runs?per_page={limit}",
        capture_output=True,
        env=env,
    )
    payload = json.loads(result.stdout)
    return payload.get("workflow_runs", [])


def find_successful_run_today(
    *,
    current_day: date,
    exclude_run_id: int | None = None,
    limit: int = 20,
    env: dict[str, str] | None = None,
) -> dict[str, Any] | None:
    for run in fetch_workflow_runs(limit=limit, env=env):
        if exclude_run_id and int(run["id"]) == exclude_run_id:
            continue
        if run.get("conclusion") != "success":
            continue
        if run.get("event") not in {"schedule", "workflow_dispatch"}:
            continue
        if parse_run_date(run["run_started_at"]) != current_day:
            continue
        return run
    return None


def write_github_output(name: str, value: str) -> None:
    output_path = os.environ.get("GITHUB_OUTPUT")
    if not output_path:
        return

    with open(output_path, "a", encoding="utf-8") as handle:
        handle.write(f"{name}={value}\n")


def check_only(limit: int) -> int:
    env = os.environ.copy()
    current_run_id = int(env.get("GITHUB_RUN_ID", "0"))
    current_day = today_kst()
    existing = find_successful_run_today(
        current_day=current_day,
        exclude_run_id=current_run_id or None,
        limit=limit,
        env=env,
    )

    if existing:
        existing_id = str(existing["id"])
        existing_started = existing["run_started_at"]
        write_github_output("should_send", "false")
        write_github_output("existing_run_id", existing_id)
        print(
            f"[SKIP] already sent today via run {existing_id} "
            f"({existing.get('event')} / {existing_started})"
        )
        return 0

    write_github_output("should_send", "true")
    write_github_output("existing_run_id", "")
    print("[OK] no successful same-day run found")
    return 0


def dispatch(limit: int, allow_offday: bool) -> int:
    current = now_kst()
    if current.weekday() not in TARGET_WEEKDAYS and not allow_offday:
        print(f"[SKIP] {current.date()} is not a send day")
        return 0

    existing = find_successful_run_today(current_day=current.date(), limit=limit)
    if existing:
        print(
            f"[SKIP] already sent today via run {existing['id']} "
            f"({existing.get('event')} / {existing['run_started_at']})"
        )
        return 0

    run_gh(
        "workflow",
        "run",
        WORKFLOW,
        "--repo",
        REPO,
        "--ref",
        "main",
        "-f",
        "dry_run=false",
    )
    print(f"[OK] workflow dispatched at {current.isoformat(timespec='seconds')}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check-only", action="store_true")
    parser.add_argument("--allow-offday", action="store_true")
    parser.add_argument("--limit", type=int, default=20)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.check_only:
        return check_only(limit=args.limit)

    return dispatch(limit=args.limit, allow_offday=args.allow_offday)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # pragma: no cover - operational helper
        print(f"[ERROR] {exc}", file=sys.stderr)
        raise
