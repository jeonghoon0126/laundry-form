import importlib.util
import os
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]

def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, ROOT / path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

class DatabaseConfigurationTests(unittest.TestCase):
    def test_absent_secret_does_not_create_fallback_or_connect(self):
        with patch.dict(os.environ, {}, clear=True):
            gi = load_module('generate_invoices', 'scripts/generate_invoices.py')
        self.assertEqual(gi.SUPABASE_URI, '')
        with patch.object(gi.psycopg2, 'connect') as connect:
            with self.assertRaisesRegex(RuntimeError, 'SUPABASE_URI must be configured'):
                gi.get_monthly_data(2026, 10)
            connect.assert_not_called()

    def test_whitespace_secret_fails_closed(self):
        with patch.dict(os.environ, {'SUPABASE_URI': ' \n\t '}):
            gi = load_module('generate_invoices', 'scripts/generate_invoices.py')
        with self.assertRaises(RuntimeError):
            gi.require_supabase_uri()

    def test_configured_secret_is_trimmed_and_passed_to_driver(self):
        fake_uri = 'postgresql://unit-test.invalid/example'
        with patch.dict(os.environ, {'SUPABASE_URI': f' {fake_uri}\n'}):
            gi = load_module('generate_invoices', 'scripts/generate_invoices.py')
        with patch.object(gi.psycopg2, 'connect') as connect:
            connect.return_value.cursor.return_value.fetchall.return_value = []
            self.assertEqual(gi.get_monthly_data(2026, 10), [])
            connect.assert_called_once_with(fake_uri)

    def test_weekly_report_reuses_fail_closed_validation(self):
        with patch.dict(os.environ, {}, clear=True):
            gi = load_module('generate_invoices', 'scripts/generate_invoices.py')
            with patch.dict(sys.modules, {'generate_invoices': gi}):
                report = load_module('send_weekly_revenue_report', 'scripts/send_weekly_revenue_report.py')
        with patch.object(report.psycopg2, 'connect') as connect:
            with self.assertRaisesRegex(RuntimeError, 'SUPABASE_URI must be configured'):
                report.get_weekly_rows(report.date(2026, 10, 1), report.date(2026, 10, 8))
            connect.assert_not_called()

    def test_config_error_does_not_contain_secret(self):
        gi = load_module('generate_invoices', 'scripts/generate_invoices.py')
        with patch.object(gi, 'SUPABASE_URI', ''):
            with self.assertRaises(RuntimeError) as error:
                gi.require_supabase_uri()
        self.assertNotIn('postgresql://', str(error.exception))

if __name__ == '__main__':
    unittest.main()
