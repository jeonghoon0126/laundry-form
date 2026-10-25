"""Fake environment and mock-only tests. Never contact GitHub or SOLAPI."""
from datetime import datetime, timezone
import io
import json
from pathlib import Path
import unittest
from unittest.mock import patch
import probe as p

NOW = datetime(2026, 10, 25, 12, 0, tzinfo=timezone.utc)
KEY, SECRET = "FAKE_API_KEY_ONLY", "FAKE_SECRET_ONLY"
PRIVATE_PHONE = "01000000000"
PRIVATE_BALANCE = 9876543
REVISION = "a" * 40


def environment():
    return {"SOLAPI_STATUS_READ_TARGET": "standard_status", "GITHUB_ACTIONS": "true", "GITHUB_EVENT_NAME": "workflow_dispatch",
        "GITHUB_REPOSITORY": p.REPOSITORY, "GITHUB_REF": p.REF,
        "GITHUB_WORKFLOW_REF": p.WORKFLOW_REF, "GITHUB_SHA": REVISION,
        "SOLAPI_STATUS_READ_ACK": p.ACK, "SOLAPI_STATUS_APPROVED_REVISION": REVISION,
        "SOLAPI_API_KEY": KEY, "SOLAPI_API_SECRET": SECRET}


def response(body, status=200):
    return p.Response(status, json.dumps(body).encode())


def balance():
    return {"accountId": "PRIVATE_FAKE_ACCOUNT", "balance": PRIVATE_BALANCE,
            "balanceOnly": PRIVATE_BALANCE, "deposit": 0, "point": 0,
            "minimumCash": 12345, "rechargeTo": 54321, "autoRecharge": 1}


def senders():
    return [{"phoneNumber": PRIVATE_PHONE, "status": "ACTIVE", "expireAt": "2026-12-01T00:00:00Z"}]


def pricing():
    return {"countryId": "82", "serviceMethod": "MT", "mms": 123}


class MockTransport:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []
        self.headers = []

    def get(self, endpoint, authorization):
        self.calls.append(endpoint)
        self.headers.append(authorization)
        result = self.responses.pop(0)
        if isinstance(result, BaseException):
            raise result
        return result


class ProbeTests(unittest.TestCase):
    def setUp(self):
        self.counter = 0
        guard = patch("socket.socket", side_effect=AssertionError("real network forbidden in tests"))
        self.socket_mock = guard.start()
        self.addCleanup(guard.stop)

    def tearDown(self):
        self.assertEqual(self.socket_mock.call_count, 0, "tests attempted a real socket")

    def salt(self):
        self.counter += 1
        return f"{self.counter:032x}"

    def execute(self, values=None, env=None):
        mock = MockTransport(values if values is not None else [response(balance()), response(senders()), response(pricing())])
        result = p.run(execute_read_only=True, environment_provider=lambda: environment() if env is None else env,
                       transport_factory=lambda: mock, now_provider=lambda: NOW, salt_provider=self.salt)
        return result, mock

    def test_default_dry_run_zero_io(self):
        def explode():
            raise AssertionError("provider invoked")
        class NoEnvironment(dict):
            def get(self, *args, **kwargs):
                raise AssertionError("environment accessed")
            def __getitem__(self, key):
                raise AssertionError("environment accessed")
        with patch("builtins.open", side_effect=AssertionError("file I/O")), \
             patch("pathlib.Path.open", side_effect=AssertionError("path I/O")), \
             patch("os.getenv", side_effect=AssertionError("env read")), \
             patch("os.environ", NoEnvironment()), \
             patch("socket.socket", side_effect=AssertionError("network")), \
             patch("http.client.HTTPSConnection", side_effect=AssertionError("HTTP")), \
             patch("ssl.create_default_context", side_effect=AssertionError("TLS I/O")):
            result = p.run(environment_provider=explode, transport_factory=explode, now_provider=explode, salt_provider=explode)
            self.assertEqual(result, p.default_report())
            with patch("sys.stdout", new_callable=io.StringIO) as output:
                self.assertEqual(p.main([]), 0)
                self.assertIn("NO_ACCOUNT_READS_PERFORMED", output.getvalue())

    def test_only_literal_true_selects_execute(self):
        for value in (False, None, "true", 1):
            result = p.run(execute_read_only=value, environment_provider=lambda: self.fail("env called"))
            self.assertEqual(result["mode"], "dry_run")

    def test_every_runtime_guard_before_secret_read(self):
        class GuardMapping(dict):
            def get(self, key, default=None):
                if key in ("SOLAPI_API_KEY", "SOLAPI_API_SECRET"):
                    raise AssertionError("must not read secrets")
                return super().get(key, default)
        for field in ("GITHUB_ACTIONS", "GITHUB_EVENT_NAME", "GITHUB_REPOSITORY", "GITHUB_REF", "GITHUB_WORKFLOW_REF",
                      "GITHUB_SHA", "SOLAPI_STATUS_READ_ACK", "SOLAPI_STATUS_APPROVED_REVISION", "SOLAPI_STATUS_READ_TARGET"):
            with self.subTest(field=field):
                env = GuardMapping(environment())
                env[field] = "wrong"
                result, mock = self.execute([], env)
                self.assertEqual(result["code"], "ACTIVATION_GUARD_BLOCKED")
                self.assertEqual(mock.calls, [])

    def test_missing_or_invalid_credentials_no_network(self):
        for field in ("SOLAPI_API_KEY", "SOLAPI_API_SECRET"):
            for bad in ("", None, "x\r\nleak", 1):
                env = environment()
                env[field] = bad
                result, mock = self.execute([], env)
                self.assertEqual(result["code"], "READ_ONLY_PREREQUISITES_UNVERIFIED")
                self.assertEqual(mock.calls, [])

    def test_all_reads_success_without_authorization_claim(self):
        result, mock = self.execute()
        self.assertEqual(mock.calls, list(p.Endpoint))
        for value in result["checks"].values():
            self.assertEqual(value, {"api_read": "succeeded", "response_fields": "verifiable"})
        for key in ("specific_sender_verified", "budget_verified", "send_authorized", "spending_authorized",
                    "send_performed", "upload_performed", "account_changed"):
            self.assertIs(result[key], False)
        self.assertEqual(result["tax_basis"], "unverified")
        self.assertEqual(result["approved_account_match"], "unverified")
        self.assertEqual(len(set(mock.headers)), 3)

    def test_no_sensitive_values_in_any_report(self):
        result, _ = self.execute()
        text = json.dumps(result) + repr(p.Credentials(KEY, SECRET)) + repr(response(balance()))
        for private in (KEY, SECRET, PRIVATE_PHONE, str(PRIVATE_BALANCE), "PRIVATE_FAKE_ACCOUNT", "12345", "54321"):
            self.assertNotIn(private, text)
        self.assertNotIn("autoRecharge", text)
        self.assertNotIn("Authorization", text)

    def test_on_off_recharge_same_public_shape(self):
        on, _ = self.execute()
        off_balance = balance()
        off_balance["autoRecharge"] = 0
        off, _ = self.execute([response(off_balance), response(senders()), response(pricing())])
        self.assertEqual(on, off)

    def test_empty_sender_list_does_not_claim_sender_verified(self):
        result, _ = self.execute([response(balance()), response([]), response(pricing())])
        self.assertEqual(result["checks"]["active_sender_listing"]["response_fields"], "verifiable")
        self.assertIs(result["specific_sender_verified"], False)

    def test_optional_sender_expiry_does_not_infer_validity(self):
        for sender in ({"phoneNumber": PRIVATE_PHONE, "status": "ACTIVE"},
                       {"phoneNumber": PRIVATE_PHONE, "status": "ACTIVE", "expireAt": None}):
            result, mock = self.execute([response(balance()), response([sender]), response(pricing())])
            self.assertEqual(mock.calls, list(p.Endpoint))
            self.assertEqual(result["checks"]["active_sender_listing"]["response_fields"], "verifiable")
            self.assertIs(result["specific_sender_verified"], False)
            self.assertIs(result["send_authorized"], False)

    def test_unknown_or_malformed_fields_remain_unverified(self):
        for endpoint, bad_values in ((p.Endpoint.BALANCE, ({}, {**balance(), "autoRecharge": "0"}, {**balance(), "balance": True})),
                                    (p.Endpoint.SENDERS, ({}, [None], [{**senders()[0], "status": "PENDING"}], [{**senders()[0], "expireAt": "not date"}])),
                                    (p.Endpoint.PRICING, ({}, {**pricing(), "mms": 0}, {**pricing(), "mms": "123"}, {**pricing(), "countryId": "1"}))):
            for data in bad_values:
                self.assertEqual(p.classify(endpoint, response(data))["response_fields"], "unverified")

    def test_malformed_http_json_and_provider_errors_redacted(self):
        for body in (b"not json", b'{"a":1,"a":2}', b'{"balance":NaN}', b"x" * (p.MAX_RESPONSE_BYTES + 1), b"\xff"):
            out = p.classify(p.Endpoint.BALANCE, p.Response(200, body))
            self.assertEqual(out, {"api_read": "unconfirmed", "response_fields": "unverified"})
        out = p.classify(p.Endpoint.BALANCE, response({"errorCode": "SENSITIVE_FAKE_ERROR", "errorMessage": PRIVATE_PHONE}))
        self.assertEqual(out["api_read"], "unconfirmed")
        self.assertNotIn(PRIVATE_PHONE, json.dumps(out))
        for value in (None, "200", True):
            self.assertEqual(p.classify(p.Endpoint.BALANCE, p.Response(value, b"{}"))["api_read"], "unconfirmed")

    def test_auth_denied_stops_without_retry(self):
        for code in (401, 403):
            result, mock = self.execute([response({"errorMessage": SECRET}, code)])
            self.assertEqual(len(mock.calls), 1)
            self.assertEqual(result["checks"]["balance_and_recharge"]["api_read"], "permission_or_authentication_denied")
            self.assertEqual(result["checks"]["active_sender_listing"]["api_read"], "not_run")
            self.assertNotIn(SECRET, json.dumps(result))

    def test_timeout_no_retry_and_no_raw_exception(self):
        result, mock = self.execute([TimeoutError(SECRET), response(senders()), response(pricing())])
        self.assertEqual(mock.calls, [p.Endpoint.BALANCE])
        self.assertEqual(result["checks"]["balance_and_recharge"]["api_read"], "unconfirmed")
        self.assertNotIn(SECRET, json.dumps(result))

    def test_non_2xx_never_followed_or_interpreted_as_success(self):
        for code in (301, 302, 307, 308, 408, 429, 500, 503):
            result = p.classify(p.Endpoint.BALANCE, response(balance(), code))
            self.assertEqual(result["api_read"], "unconfirmed")

    def test_error_redirect_or_bad_schema_never_expands_requests(self):
        bad_responses = [response({"errorMessage": SECRET}, code) for code in (401, 429, 301, 302, 307, 308, 500)]
        bad_responses += [response({}), p.Response(200, b"broken"), response({**balance(), "autoRecharge": "unknown"})]
        for bad in bad_responses:
            with self.subTest(status=bad.status):
                result, mock = self.execute([bad, response(senders()), response(pricing())])
                self.assertEqual(mock.calls, [p.Endpoint.BALANCE])
                self.assertEqual(result["checks"]["active_sender_listing"]["api_read"], "not_run")
                self.assertEqual(len(mock.responses), 2)
                self.assertNotIn(SECRET, json.dumps(result))

    def test_reused_salt_prevents_subsequent_calls(self):
        mock = MockTransport([response(balance())])
        result = p.run(execute_read_only=True, environment_provider=environment,
            transport_factory=lambda: mock, now_provider=lambda: NOW, salt_provider=lambda: "0" * 32)
        self.assertEqual(mock.calls, [p.Endpoint.BALANCE])
        self.assertEqual(result["checks"]["active_sender_listing"]["api_read"], "unconfirmed")

    def test_hmac_known_fake_vector(self):
        header = p.sign(p.Credentials("dummy_key", "dummy_secret"), NOW, "0" * 32)
        self.assertEqual(header, "HMAC-SHA256 apiKey=dummy_key, date=2026-10-25T12:00:00Z, salt=00000000000000000000000000000000, signature=44fa15ceec81cfe7f03b4c874fe39892aa0deaa9f9562c5ddb7b1ff1c4cfb459")

    def test_cli_invalid_input_is_redacted(self):
        with patch("sys.stdout", new_callable=io.StringIO) as output:
            self.assertEqual(p.main(["--" + PRIVATE_PHONE]), 2)
            self.assertNotIn(PRIVATE_PHONE, output.getvalue())

    def test_tls_verification_preserved_and_keylogging_never_enabled(self):
        class FakeContext:
            keylog_filename = None
            def load_default_certs(self, purpose):
                self.purpose = purpose
        context = FakeContext()
        with patch("probe.ssl.SSLContext", return_value=context) as constructor, \
             patch("probe.ssl.create_default_context", side_effect=AssertionError("would honor SSLKEYLOGFILE")), \
             patch.dict("os.environ", {"SSLKEYLOGFILE": "DO-NOT-CREATE-FAKE-KEYLOG"}, clear=True):
            self.assertIs(p.verified_tls_context(), context)
            constructor.assert_called_once_with(p.ssl.PROTOCOL_TLS_CLIENT)
        self.assertIs(context.check_hostname, True)
        self.assertEqual(context.verify_mode, p.ssl.CERT_REQUIRED)
        self.assertEqual(context.purpose, p.ssl.Purpose.SERVER_AUTH)
        self.assertIsNone(context.keylog_filename)

    def test_transport_fixed_get_no_body_no_proxy_no_redirect(self):
        class FakeHTTP:
            status = 200
            def read(self, amount):
                self.amount = amount
                return b"{}"
            def getheader(self, name, default):
                return "application/json"
        class FakeConnection:
            def __init__(self):
                self.calls = []
                self.closed = False
                self.raw = FakeHTTP()
            def set_debuglevel(self, level):
                self.debug = level
            def request(self, method, path, body, headers):
                self.calls.append((method, path, body, headers))
            def getresponse(self):
                return self.raw
            def close(self):
                self.closed = True
        connection = FakeConnection()
        with patch("probe.verified_tls_context", return_value="FAKE_VERIFIED_CONTEXT"), \
             patch("probe.http.client.HTTPSConnection", return_value=connection) as factory:
            p.HTTPSReadTransport().get(p.Endpoint.BALANCE, "HMAC-SHA256 FAKE_HEADER")
            factory.assert_called_once_with("api.solapi.com", port=443, timeout=10, context="FAKE_VERIFIED_CONTEXT")
        self.assertEqual(connection.calls[0][:3], ("GET", "/cash/v1/balance", None))
        self.assertEqual(connection.debug, 0)
        self.assertTrue(connection.closed)
        self.assertEqual(connection.raw.amount, p.MAX_RESPONSE_BYTES + 1)

    def test_transport_rejects_arbitrary_endpoint_before_io(self):
        with patch("probe.verified_tls_context", side_effect=AssertionError("I/O")):
            for endpoint in ("https://evil.invalid", "/messages/v4/send-many/detail", "/cash/v1/balance", "POST"):
                with self.assertRaisesRegex(p.ProbeError, "READ_ENDPOINT_NOT_ALLOWED"):
                    p.HTTPSReadTransport().get(endpoint, "HMAC-SHA256 FAKE_HEADER")

    def test_transport_errors_sanitized_and_redirect_body_not_read(self):
        class Fake:
            status = 307
            def set_debuglevel(self, _): pass
            def request(self, *args, **kwargs): pass
            def getresponse(self): return self
            def read(self, _): raise AssertionError("must not read redirect body")
            def close(self): pass
        with patch("probe.verified_tls_context", return_value=object()), \
             patch("probe.http.client.HTTPSConnection", return_value=Fake()):
            result = p.HTTPSReadTransport().get(p.Endpoint.BALANCE, "HMAC-SHA256 FAKE_HEADER")
            self.assertEqual(result.status, 307)
            self.assertEqual(result.body, b"")
        with patch("probe.verified_tls_context", side_effect=OSError(SECRET)):
            with self.assertRaises(p.ProbeError) as caught:
                p.HTTPSReadTransport().get(p.Endpoint.BALANCE, "HMAC-SHA256 FAKE_HEADER")
            self.assertNotIn(SECRET, str(caught.exception))

    def diagnostic(self, values=None, env=None):
        env = environment() if env is None else env
        env['SOLAPI_STATUS_READ_TARGET'] = p.ReadTarget.SENDER_SCHEMA.value
        mock = MockTransport(values if values is not None else [response(senders())])
        result = p.run(execute_read_only=True, read_target=p.ReadTarget.SENDER_SCHEMA,
                       environment_provider=lambda: env, transport_factory=lambda: mock,
                       now_provider=lambda: NOW, salt_provider=self.salt)
        return result, mock

    def test_sender_diagnostic_exactly_one_get_and_no_scope_expansion(self):
        for data in (senders(), [], {}):
            result, mock = self.diagnostic([response(data)])
            self.assertEqual(mock.calls, [p.Endpoint.SENDERS])
            self.assertEqual(result['read_target'], 'sender_schema_diagnostic')
            for name in ('balance_and_recharge', 'account_mms_pricing'):
                self.assertEqual(result['checks'][name]['api_read'], 'not_run')
            for name in ('specific_sender_verified', 'budget_verified', 'send_authorized',
                         'spending_authorized', 'send_performed', 'upload_performed', 'account_changed'):
                self.assertIs(result[name], False)

    def test_sender_schema_reasons_are_only_fixed_enums(self):
        item = senders()[0]
        cases = [({}, 'root_type_unsupported'), ([None] * 1001, 'list_limit_exceeded'),
                 ([None], 'item_type_unsupported'), ([{}], 'missing_required_field'),
                 ([{**item, 'phoneNumber': None}], 'phone_type_unexpected'),
                 ([{**item, 'phoneNumber': 'fake-not-a-number'}], 'phone_format_unexpected'),
                 ([{**item, 'status': 'FAKE_OTHER_STATUS'}], 'status_value_unexpected'),
                 ([{**item, 'expireAt': 'FAKE_NOT_A_DATE'}], 'expiry_format_unexpected'),
                 ([], 'verified'), (senders(), 'verified')]
        for data, expected in cases:
            self.assertEqual(p.sender_schema_reason(data).value, expected)
            result, _ = self.diagnostic([response(data)])
            self.assertEqual(result['checks']['active_sender_listing']['schema_reason'], expected)

    def test_sender_acceptance_conditions_match_previous_validator(self):
        import itertools
        def previous(data):
            if type(data) is not list or len(data) > 1000:
                return False
            return all(type(item) is dict and type(item.get('phoneNumber')) is str
                       and p.re.fullmatch(r'[0-9]{8,12}', item['phoneNumber']) is not None
                       and item.get('status') == 'ACTIVE'
                       and ('expireAt' not in item or p.timestamp_or_null(item['expireAt']))
                       for item in data)
        missing = object()
        cases = [None, {}, [], [None], [None] * 1001]
        for phone, status, expiry in itertools.product(
                (missing, None, 0, '', '00000000', PRIVATE_PHONE, '0' * 13, '000-0000-0000'),
                (missing, None, 'ACTIVE', 'FAKE_OTHER', 0, {}),
                (missing, None, 0, '', '2027-01-01T00:00:00Z', '2027-01-01T00:00:00')):
            item = {}
            for name, value in (('phoneNumber', phone), ('status', status), ('expireAt', expiry)):
                if value is not missing:
                    item[name] = value
            cases.append([item])
        for data in cases:
            self.assertEqual(p.sender_schema_reason(data) is p.SchemaReason.VERIFIED, previous(data))

    def test_diagnostic_scope_guard_precedes_credentials(self):
        class GuardMapping(dict):
            def get(self, key, default=None):
                if key in ('SOLAPI_API_KEY', 'SOLAPI_API_SECRET'):
                    self.fail_if_secret_read = True
                    raise AssertionError('secret read forbidden')
                return super().get(key, default)
        for target, mismatch in ((p.ReadTarget.SENDER_SCHEMA, 'standard_status'),
                                 (p.ReadTarget.STANDARD, 'sender_schema_diagnostic'),
                                 (p.ReadTarget.SENDER_SCHEMA, None)):
            env = GuardMapping(environment())
            env['SOLAPI_STATUS_READ_TARGET'] = mismatch
            result = p.run(execute_read_only=True, read_target=target, environment_provider=lambda: env,
                           transport_factory=lambda: self.fail('transport constructed'))
            self.assertEqual(result['code'], 'ACTIVATION_GUARD_BLOCKED')
            self.assertFalse(getattr(env, 'fail_if_secret_read', False))

    def test_diagnostic_cannot_select_execution_without_literal_execute(self):
        for execute in (False, None, 1, 'true'):
            result = p.run(execute_read_only=execute, read_target=p.ReadTarget.SENDER_SCHEMA,
                           environment_provider=lambda: self.fail('environment touched'))
            self.assertEqual(result, p.default_report())
        for target in (None, 'sender_schema_diagnostic', 1):
            result = p.run(execute_read_only=True, read_target=target,
                           environment_provider=lambda: self.fail('environment touched'))
            self.assertEqual(result['code'], 'ACTIVATION_GUARD_BLOCKED')

    def test_diagnostic_errors_never_retry_or_call_other_endpoints(self):
        cases = [response({}, n) for n in (301, 307, 401, 403, 429, 500)]
        cases += [p.Response(200, b'not-json'), response({'errorCode': SECRET}), TimeoutError(SECRET)]
        for bad in cases:
            result, mock = self.diagnostic([bad])
            self.assertEqual(mock.calls, [p.Endpoint.SENDERS])
            check = result['checks']['active_sender_listing']
            self.assertEqual(check['schema_reason'], 'not_evaluated')
            self.assertNotEqual(check['api_read'], 'succeeded')
            self.assertEqual(result['checks']['account_mms_pricing']['api_read'], 'not_run')
            self.assertNotIn(SECRET, json.dumps(result))

    def test_diagnostic_never_emits_unknown_keys_values_or_exception_repr(self):
        private = ['PRIVATE_ARBITRARY_SERVER_KEY', 'PRIVATE_SERVER_VALUE', SECRET, KEY,
                   PRIVATE_PHONE, str(PRIVATE_BALANCE), 'PRIVATE_HEADER_SENTINEL']
        cases = [{private[0]: private[1]},
                 [{**senders()[0], 'status': private[1], private[0]: private[6]}],
                 [{**senders()[0], 'phoneNumber': private[1]}],
                 [{**senders()[0], 'expireAt': private[6]}]]
        for data in cases:
            result, _ = self.diagnostic([response(data)])
            output = json.dumps(result)
            for value in private:
                self.assertNotIn(value, output)
            self.assertNotIn('phoneNumber', output)
            self.assertNotIn('expireAt', output)
            self.assertNotIn('Authorization', output)
        with patch('probe.timestamp_or_null', side_effect=ValueError(SECRET)):
            result, _ = self.diagnostic()
            self.assertEqual(result['checks']['active_sender_listing']['schema_reason'], 'validation_unconfirmed')
            self.assertNotIn(SECRET, json.dumps(result))

    def test_diagnostic_unexpected_reason_return_cannot_leak(self):
        with patch('probe.sender_schema_reason', return_value=SECRET):
            result, _ = self.diagnostic()
            self.assertEqual(result['checks']['active_sender_listing']['schema_reason'], 'validation_unconfirmed')
            self.assertNotIn(SECRET, json.dumps(result))

    def test_diagnostic_reason_not_added_to_standard_output(self):
        result, _ = self.execute()
        self.assertNotIn('schema_reason', json.dumps(result))
        self.assertNotIn('read_target', result)

    def test_diagnostic_cli_exit_depends_on_sender_only(self):
        result, _ = self.diagnostic()
        with patch('probe.run', return_value=result) as run, patch('sys.stdout', new_callable=io.StringIO):
            self.assertEqual(p.main(['--execute-sender-schema-diagnostic']), 0)
            run.assert_called_once_with(execute_read_only=True, read_target=p.ReadTarget.SENDER_SCHEMA)
        result, _ = self.diagnostic([response({})])
        with patch('probe.run', return_value=result), patch('sys.stdout', new_callable=io.StringIO):
            self.assertEqual(p.main(['--execute-sender-schema-diagnostic']), 2)

    def test_workflow_diagnostic_scope_is_fixed_choice_and_guarded(self):
        import yaml
        workflow = yaml.safe_load((Path(__file__).resolve().parents[2] / '.github/workflows/solapi-status-readonly.yml').read_text())
        target = workflow['on']['workflow_dispatch']['inputs']['read_target']
        self.assertEqual(target['options'], ['standard_status', 'sender_schema_diagnostic'])
        self.assertEqual(target['default'], 'standard_status')
        job = workflow['jobs']['read_only']
        self.assertIn("inputs.read_target == 'sender_schema_diagnostic'", job['if'])
        step = job['steps'][1]
        self.assertEqual(step['env']['SOLAPI_STATUS_READ_TARGET'], '${{ inputs.read_target }}')
        self.assertIn('--execute-sender-schema-diagnostic', step['run'])
        self.assertNotIn('${{', step['run'])
        self.assertIn('*) exit 2', step['run'])

    def test_workflow_manual_only_no_secrets_or_checkout_in_dry_run(self):
        import yaml  # Already-installed test utility, not a runtime requirement.
        workflow = yaml.safe_load((Path(__file__).resolve().parents[2] / ".github/workflows/solapi-status-readonly.yml").read_text())
        self.assertEqual(set(workflow["on"]), {"workflow_dispatch"})
        inputs = workflow["on"]["workflow_dispatch"]["inputs"]
        self.assertIs(inputs["execute_read_only"]["default"], False)
        self.assertEqual(inputs["authorization_ack"]["default"], "NOT_APPROVED")
        self.assertEqual(workflow["permissions"], {})
        dry = workflow["jobs"]["dry_run"]
        self.assertEqual(dry["permissions"], {})
        dry_text = json.dumps(dry)
        for blocked in ("secrets.", "uses", "checkout", "curl", "requests", "http.client"):
            self.assertNotIn(blocked, dry_text)
        execute = workflow["jobs"]["read_only"]
        self.assertEqual(execute["permissions"], {"contents": "read"})
        self.assertFalse(execute["steps"][0]["with"]["persist-credentials"])
        self.assertRegex(execute["steps"][0]["uses"], r"actions/checkout@[0-9a-f]{40}$")
        self.assertEqual(len(execute["steps"]), 2)
        self.assertEqual(execute["steps"][1]["env"]["SOLAPI_API_KEY"], "${{ secrets.SOLAPI_API_KEY }}")
        self.assertNotIn("id-token", json.dumps(workflow))
        self.assertNotIn("upload-artifact", json.dumps(workflow))
        self.assertIn("inputs.approved_revision == github.sha", execute["if"])


if __name__ == "__main__":
    unittest.main()
