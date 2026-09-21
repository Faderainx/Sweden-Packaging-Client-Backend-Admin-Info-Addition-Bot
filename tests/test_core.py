import json
import sys
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parents[1]))
from main import Customer, _parse_mail_timestamp, extract_invoice_email, extract_verification_code, load_config, resolve_mail_route


class CoreTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = load_config(Path(__file__).parents[1] / "config.json")

    def test_route_exact_domains(self):
        self.assertEqual(resolve_mail_route("x@erp-aid.example.invalid", self.config["mail_routes"])["key"], "aid")
        self.assertEqual(resolve_mail_route("x@example.invalid", self.config["mail_routes"])["key"], "er-helper")
        self.assertEqual(resolve_mail_route("x@erp-helper.example.invalid", self.config["mail_routes"])["key"], "er-helper")
        self.assertEqual(resolve_mail_route("x@163.com", self.config["mail_routes"])["key"], "163")
        self.assertEqual(resolve_mail_route("x@foo.ecopv0316.example.invalid", self.config["mail_routes"])["key"], "aliyun")

    def test_unknown_route_is_none(self):
        self.assertIsNone(resolve_mail_route("x@unknown.example", self.config["mail_routes"]))

    def test_invoice_email_extraction(self):
        self.assertEqual(extract_invoice_email("Invoice email\ninvoice@example.com\nEdit invoice email"), "invoice@example.com")
        self.assertEqual(extract_invoice_email("Faktura e-post\nsupport@example.com\nÄndra faktura e-post"), "support@example.com")

    def test_chinese_customer_headers(self):
        customer = Customer.from_row(2, {
            "客户中文名称": "测试客户",
            "邮箱": "test@eu-erp.example.invalid",
            "邮箱密码": "secret",
        })
        self.assertEqual(customer.name, "测试客户")
        self.assertEqual(customer.portal_email, "test@eu-erp.example.invalid")
        self.assertEqual(customer.portal_password, "secret")
        self.assertEqual(customer.mail_email, "test@eu-erp.example.invalid")
        self.assertEqual(customer.mail_password, "secret")

    def test_verification_code_extraction_uses_label(self):
        self.assertEqual(extract_verification_code("账户验证码：\n31354256\n如果没有请求获取验证码"), "31354256")
        self.assertIsNone(extract_verification_code("今天的订单号是 31354256"))

    def test_mail_timestamp_labels(self):
        reference = datetime(2026, 9, 17, 17, 20)
        self.assertEqual(_parse_mail_timestamp("NPA kundportal\n今天 17:18\n账户验证码", reference), datetime(2026, 9, 17, 17, 18))
        self.assertEqual(_parse_mail_timestamp("NPA kundportal\nToday 17:18\n账户验证码", reference), datetime(2026, 9, 17, 17, 18))
        self.assertEqual(_parse_mail_timestamp("NPA kundportal\nYesterday 23:05\n账户验证码", reference), datetime(2026, 9, 16, 23, 5))
        self.assertEqual(_parse_mail_timestamp("NPA kundportal\nTue 18:47\n账户验证码", reference), datetime(2026, 9, 15, 18, 47))
        self.assertEqual(_parse_mail_timestamp("NPA kundportal\n2026-09-16 23:05\n账户验证码", reference), datetime(2026, 9, 16, 23, 5))
        self.assertEqual(_parse_mail_timestamp("NPA kundportal\n9月17日 17:19\n账户验证码", reference), datetime(2026, 9, 17, 17, 19))
        self.assertEqual(_parse_mail_timestamp("时间：2026年9月18日 09:36 (星期五)", reference), datetime(2026, 9, 18, 9, 36))


if __name__ == "__main__":
    unittest.main()
