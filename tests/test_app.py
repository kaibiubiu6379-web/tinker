import io
import re
import unittest
from pathlib import Path

from openpyxl import load_workbook

from app import create_app
from domain_to_excel import parse_text


SAMPLE_TEXT = """DEMO-000
正常example-h5.com
ns-1.example.com
DEMO-000-H5
2027-02-14 17:59:32(156天)

DEMO-000
正常example-web.com
ns-2.example.com
DEMO-000-WEB
2027-04-25 16:46:19(226天)
"""

UNKNOWN_CATEGORY_TEXT = """DEMO-SEO 备案域名
正常seo-example.com
ns-1.example.com
备案域名已退回
2027-02-10 09:34:21(151天)
使用中
"""


class AppTestCase(unittest.TestCase):
    def setUp(self):
        app = create_app(
            {
                "TESTING": True,
                "SECRET_KEY": "test-secret",
                "APP_USERNAME": "tester",
                "APP_PASSWORD": "secret",
            }
        )
        self.client = app.test_client()

    def login(self):
        response = self.client.post(
            "/api/login", json={"username": "tester", "password": "secret"}
        )
        self.assertEqual(response.status_code, 200)
        with self.client.session_transaction() as session:
            return session["csrf_token"]

    def test_protected_page_redirects_to_login(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.location.endswith("/login"))

    def test_invalid_login_is_rejected(self):
        response = self.client.post(
            "/api/login", json={"username": "tester", "password": "wrong"}
        )
        self.assertEqual(response.status_code, 401)

    def test_pasted_text_returns_excel(self):
        csrf_token = self.login()
        response = self.client.post(
            "/api/convert",
            data={"text": SAMPLE_TEXT, "title": "测试域名"},
            headers={"X-CSRF-Token": csrf_token},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["X-Record-Count"], "2")
        self.assertRegex(
            response.headers["Content-Disposition"],
            re.compile(r"filename\*=UTF-8''%E6%B5%8B%E8%AF%95%E5%9F%9F%E5%90%8D-\d{8}-\d{6}\.xlsx"),
        )
        workbook = load_workbook(io.BytesIO(response.data), read_only=True)
        sheet = workbook.active
        self.assertEqual(sheet["A2"].value, "WEB")
        self.assertEqual(sheet["C2"].value, "H5")
        self.assertEqual(sheet["A3"].value, "example-web.com")
        self.assertEqual(sheet["C3"].value, "example-h5.com")

    def test_file_upload_returns_excel(self):
        csrf_token = self.login()
        response = self.client.post(
            "/api/convert",
            data={
                "file": (io.BytesIO(SAMPLE_TEXT.encode("utf-8")), "domains.txt"),
                "title": "上传测试",
            },
            headers={"X-CSRF-Token": csrf_token},
        )
        self.assertEqual(response.status_code, 200)

    def test_convert_rejects_missing_csrf_token(self):
        self.login()
        response = self.client.post("/api/convert", data={"text": SAMPLE_TEXT})
        self.assertEqual(response.status_code, 403)

    def test_unknown_category_uses_first_field_of_header(self):
        records = parse_text(UNKNOWN_CATEGORY_TEXT)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["category"], "DEMO-SEO")


if __name__ == "__main__":
    unittest.main()
