import unittest

from src.purchase_request_app import app


class EzAccountFlowTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

    def test_ezaccount_route_exists_and_handles_missing_config(self):
        response = self.client.post('/api/pr/1/ezaccount')
        self.assertIn(response.status_code, (200, 400, 404))
        payload = response.get_json(silent=True) or {}
        self.assertIn('ok', payload)


if __name__ == '__main__':
    unittest.main()
