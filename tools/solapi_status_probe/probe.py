"""One-shot SOLAPI account status probe. Default is I/O-free dry-run.

No send/upload/payment endpoint exists. Execute mode is for a future explicitly
approved GitHub Actions runtime only. No send/upload/payment path exists.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
import hashlib
import hmac
import http.client
import json
import math
import re
import secrets
import ssl
import sys
from typing import Callable, Mapping

HOST = "api.solapi.com"
ORIGIN = "https://" + HOST
REPOSITORY = "jeonghoon0126/laundry-form"
REF = "refs/heads/main"
WORKFLOW_REF = REPOSITORY + "/.github/workflows/solapi-status-readonly.yml@" + REF
ACK = "APPROVED_READ_ONLY_SOLAPI_STATUS"
MAX_RESPONSE_BYTES = 256_000
TIMEOUT_SECONDS = 10


class Endpoint(Enum):
    BALANCE = "/cash/v1/balance"
    SENDERS = "/senderid/v1/numbers/active"
    PRICING = "/pricing/v1/messaging?countryId=82&serviceMethod=MT&withoutProfit=false"


LABELS = {Endpoint.BALANCE: "balance_and_recharge", Endpoint.SENDERS: "active_sender_listing",
          Endpoint.PRICING: "account_mms_pricing"}


class ProbeError(ValueError):
    """Fixed error codes only. Never include request, response, env or exceptions."""


@dataclass(frozen=True)
class Credentials:
    key: str = field(repr=False)
    secret: str = field(repr=False)


@dataclass(frozen=True)
class Response:
    status: int
    body: bytes = field(repr=False)
    content_type: str = field(default="application/json", repr=False)


def default_report(mode="dry_run", code="NO_ACCOUNT_READS_PERFORMED") -> dict:
    return {"mode": mode, "code": code, "scope": "ACCOUNT_READS_ONLY",
            "checks": {LABELS[item]: {"api_read": "not_run", "response_fields": "unverified"} for item in Endpoint},
            "specific_sender_verified": False, "approved_account_match": "unverified",
            "tax_basis": "unverified", "budget_verified": False,
            "send_authorized": False, "spending_authorized": False,
            "send_performed": False, "upload_performed": False,
            "account_changed": False}


def check_runtime(env: Mapping[str, str]) -> bool:
    """An accident guard, NOT evidence that the user actually granted approval."""
    expected = {"GITHUB_ACTIONS": "true", "GITHUB_EVENT_NAME": "workflow_dispatch",
                "GITHUB_REPOSITORY": REPOSITORY, "GITHUB_REF": REF,
                "GITHUB_WORKFLOW_REF": WORKFLOW_REF, "SOLAPI_STATUS_READ_ACK": ACK}
    if any(env.get(key) != value for key, value in expected.items()):
        return False
    revision = env.get("SOLAPI_STATUS_APPROVED_REVISION", "")
    return (type(revision) is str and re.fullmatch(r"[0-9a-f]{40}", revision) is not None
            and env.get("GITHUB_SHA") == revision)


def credential_values(env: Mapping[str, str]) -> Credentials:
    # This function is reachable only after explicit mode and all runtime guards.
    key, secret = env.get("SOLAPI_API_KEY", ""), env.get("SOLAPI_API_SECRET", "")
    if (type(key) is not str or not re.fullmatch(r"[A-Za-z0-9_-]{1,200}", key)
            or type(secret) is not str or not 1 <= len(secret) <= 1000
            or any(ord(c) < 32 or ord(c) > 126 for c in secret)):
        raise ProbeError("CREDENTIALS_UNAVAILABLE_OR_INVALID")
    return Credentials(key, secret)


def sign(credentials: Credentials, now: datetime, salt: str) -> str:
    if (type(credentials) is not Credentials or not isinstance(now, datetime)
            or now.tzinfo is None or now.utcoffset() is None
            or type(salt) is not str or not re.fullmatch(r"[a-f0-9]{32}", salt)):
        raise ProbeError("SIGNING_INPUT_INVALID")
    stamp = now.astimezone(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    signature = hmac.new(credentials.secret.encode("utf-8"),
                         (stamp + salt).encode("utf-8"), hashlib.sha256).hexdigest()
    return f"HMAC-SHA256 apiKey={credentials.key}, date={stamp}, salt={salt}, signature={signature}"


def verified_tls_context():
    # Unlike create_default_context(), direct SSLContext construction does not
    # honor SSLKEYLOGFILE. Never create a TLS key-log file in the future runner.
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    context.check_hostname = True
    context.verify_mode = ssl.CERT_REQUIRED
    context.load_default_certs(ssl.Purpose.SERVER_AUTH)
    if context.keylog_filename is not None:
        raise ProbeError("TLS_KEY_LOGGING_FORBIDDEN")
    return context


class HTTPSReadTransport:
    """GET-only, exact official host/path, TLS verification, no proxy/redirect/retry.

    Constructor performs no I/O. No arbitrary URL, port, method, body, redirect
    target or connection factory is accepted from workflow inputs.
    """
    def get(self, endpoint: Endpoint, authorization: str) -> Response:
        if type(endpoint) is not Endpoint:
            raise ProbeError("READ_ENDPOINT_NOT_ALLOWED")
        if (type(authorization) is not str or "\r" in authorization or "\n" in authorization
                or not authorization.startswith("HMAC-SHA256 ")):
            raise ProbeError("AUTHORIZATION_HEADER_INVALID")
        connection = None
        try:
            context = verified_tls_context()
            connection = http.client.HTTPSConnection(HOST, port=443, timeout=TIMEOUT_SECONDS, context=context)
            connection.set_debuglevel(0)
            connection.request("GET", endpoint.value, body=None,
                               headers={"Authorization": authorization, "Accept": "application/json"})
            response = connection.getresponse()
            # Do not follow redirects or consume server error bodies.
            if not 200 <= response.status < 300:
                return Response(response.status, b"")
            body = response.read(MAX_RESPONSE_BYTES + 1)
            if len(body) > MAX_RESPONSE_BYTES:
                raise ProbeError("RESPONSE_LIMIT_EXCEEDED")
            return Response(response.status, body, response.getheader("Content-Type", ""))
        except Exception:
            # Fixed category; never serialize underlying exceptions/headers/body.
            raise ProbeError("READ_TRANSPORT_UNCONFIRMED") from None
        finally:
            if connection is not None:
                try:
                    connection.close()
                except Exception:
                    pass


def parse_json(response: Response):
    if (type(response) is not Response or type(response.body) is not bytes
            or len(response.body) > MAX_RESPONSE_BYTES
            or type(response.content_type) is not str
            or response.content_type.split(";", 1)[0].strip().lower() != "application/json"):
        return None
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError("duplicate")
            result[key] = value
        return result
    def invalid(_):
        raise ValueError("nonfinite")
    try:
        return json.loads(response.body, object_pairs_hook=unique, parse_constant=invalid)
    except (ValueError, UnicodeError, RecursionError):
        return None


def number(value) -> bool:
    return type(value) in (int, float) and math.isfinite(value) and value >= 0


def timestamp_or_null(value) -> bool:
    if value is None:
        return True  # Valid schema, not proof that a specific sender is usable.
    if type(value) is not str:
        return False
    try:
        moment = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return moment.tzinfo is not None and moment.utcoffset() is not None
    except (ValueError, OverflowError):
        return False


def verifiable_fields(endpoint: Endpoint, data) -> bool:
    """Schema/presence only. Do NOT output or infer the actual private setting."""
    if endpoint is Endpoint.BALANCE:
        if type(data) is not dict:
            return False
        fields = ("balance", "balanceOnly", "deposit", "point", "minimumCash", "rechargeTo")
        return (type(data.get("accountId")) is str and bool(data["accountId"])
                and all(number(data.get(name)) for name in fields)
                and type(data.get("autoRecharge")) is int and data["autoRecharge"] >= 0)
    if endpoint is Endpoint.PRICING:
        return (type(data) is dict and data.get("countryId") == "82"
                and data.get("serviceMethod") == "MT" and number(data.get("mms")) and data["mms"] > 0)
    if endpoint is Endpoint.SENDERS:
        if type(data) is not list or len(data) > 1000:
            return False
        return all(type(item) is dict and type(item.get("phoneNumber")) is str
                   and re.fullmatch(r"[0-9]{8,12}", item["phoneNumber"]) is not None
                   and item.get("status") == "ACTIVE" and "expireAt" in item
                   and timestamp_or_null(item["expireAt"]) for item in data)
    return False


def classify(endpoint: Endpoint, response: Response) -> dict:
    result = {"api_read": "unconfirmed", "response_fields": "unverified"}
    if type(response) is not Response or type(response.status) is not int:
        return result
    if response.status in (401, 403):
        result["api_read"] = "permission_or_authentication_denied"
    elif 200 <= response.status < 300:
        data = parse_json(response)
        # An HTTP success without the expected response isn't authentication or
        # configuration proof. Never echo an errorMessage returned in JSON.
        if data is not None and not (type(data) is dict and "errorCode" in data):
            result["api_read"] = "succeeded"
            try:
                valid = verifiable_fields(endpoint, data)
            except (TypeError, ValueError, OverflowError, RecursionError):
                valid = False
            if valid:
                result["response_fields"] = "verifiable"
    return result


def run(*, execute_read_only=False, environment_provider: Callable | None = None,
        transport_factory: Callable | None = None, now_provider: Callable | None = None,
        salt_provider: Callable | None = None) -> dict:
    if execute_read_only is not True:
        return default_report()  # Before environment/credentials/clock/TLS/network.
    report = default_report("read_only_diagnostic", "ACTIVATION_GUARD_BLOCKED")
    try:
        if environment_provider is None:
            import os
            env = os.environ
        else:
            env = environment_provider()
        if not check_runtime(env):
            return report
        credentials = credential_values(env)
        transport = transport_factory() if transport_factory else HTTPSReadTransport()
        now_provider = now_provider or (lambda: datetime.now(timezone.utc))
        salt_provider = salt_provider or (lambda: secrets.token_hex(16))
        # No endpoint repeat. Any failure or unverifiable schema stops this
        # diagnostic without expanding into further requests.
        used_salts = set()
        for endpoint in Endpoint:
            try:
                salt = salt_provider()
                if type(salt) is not str or salt in used_salts:
                    raise ProbeError("FRESH_SALT_REQUIRED")
                used_salts.add(salt)
                auth = sign(credentials, now_provider(), salt)
                response = transport.get(endpoint, auth)
                check = classify(endpoint, response)
            except Exception:
                check = {"api_read": "unconfirmed", "response_fields": "unverified"}
            report["checks"][LABELS[endpoint]] = check
            if check["api_read"] != "succeeded" or check["response_fields"] != "verifiable":
                break
        report["code"] = "READ_ONLY_OBSERVATION_NOT_SEND_AUTHORIZATION"
        return report
    except Exception:
        report["code"] = "READ_ONLY_PREREQUISITES_UNVERIFIED"
        return report


def main(argv=None) -> int:
    # A fixed one-flag interface avoids argparse/gettext locale/environment
    # reads in the no-argument path and never echoes untrusted arguments.
    arguments = sys.argv[1:] if argv is None else argv
    if arguments == []:
        report = run()
    elif arguments == ["--execute-read-only"]:
        report = run(execute_read_only=True)
    elif arguments == ["--help"]:
        print("Default: offline dry-run. --execute-read-only: future approved GitHub runtime only.")
        return 0
    else:
        report = default_report(code="INVALID_ARGUMENTS_OR_UNVERIFIED_RUNTIME")
    print(json.dumps(report, ensure_ascii=True, sort_keys=True))
    if report["mode"] == "dry_run":
        return 0 if report["code"] == "NO_ACCOUNT_READS_PERFORMED" else 2
    return 0 if all(v["api_read"] == "succeeded" and v["response_fields"] == "verifiable"
                    for v in report["checks"].values()) else 2


if __name__ == "__main__":
    raise SystemExit(main())
