import os
import unittest

try:
    from fastapi.testclient import TestClient
    HAS_FASTAPI = True
except ImportError:  # pragma: no cover - beror på miljön
    HAS_FASTAPI = False


@unittest.skipUnless(HAS_FASTAPI, "fastapi/httpx saknas i miljön")
class CurrentUserEndpointTests(unittest.TestCase):
    """Endpointen /me för raden "Inloggad som: …"."""

    @classmethod
    def setUpClass(cls):
        # main.py använder relativa sökvägar (static, templates) från repots rot
        cls._cwd = os.getcwd()
        os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from app import main
        cls.client = TestClient(main.app)

    @classmethod
    def tearDownClass(cls):
        os.chdir(cls._cwd)

    def get_user(self, header=None):
        headers = {"X-MS-CLIENT-PRINCIPAL-NAME": header} if header is not None else {}
        response = self.client.get("/me", headers=headers)
        self.assertEqual(response.status_code, 200)
        return response.json()["user"]

    def test_no_user_without_azure_login(self):
        # Lokalt (ingen Azure-inloggning framför tjänsten) visas ingen rad
        self.assertIsNone(self.get_user())

    def test_user_name_from_azure_header_is_decoded(self):
        self.assertEqual(self.get_user("gran.rorstrom%40iaf.se"), "gran.rorstrom@iaf.se")
        self.assertEqual(self.get_user("G%C3%B6ran+R%C3%B6rstr%C3%B6m"), "Göran Rörström")

    def test_blank_or_overlong_values(self):
        self.assertIsNone(self.get_user("   "))
        self.assertEqual(len(self.get_user("a" * 500)), 200)


if __name__ == "__main__":
    unittest.main()
