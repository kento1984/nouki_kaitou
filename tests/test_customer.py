"""customer.py のユニットテスト"""

from nouki_kaitou.customer import (
    check_customer_master,
    convert_day_name_to_number,
    get_customer_delivery_days,
    get_email_addresses,
    get_retention_days,
    is_route_delivery,
)
from nouki_kaitou.models import CacheStore


# テスト用モックワークシート
class MockWorksheet:
    def __init__(self, data: list[list]):
        self._data = data

    def iter_rows(self, min_row=1, max_col=None, values_only=False):
        for row in self._data[min_row - 1:]:
            if max_col is not None:
                yield tuple(row[:max_col])
            else:
                yield tuple(row)


# ============================================
# convert_day_name_to_number
# ============================================
class TestConvertDayNameToNumber:
    def test_single_char(self):
        assert convert_day_name_to_number("月") == 2
        assert convert_day_name_to_number("火") == 3
        assert convert_day_name_to_number("水") == 4
        assert convert_day_name_to_number("木") == 5
        assert convert_day_name_to_number("金") == 6
        assert convert_day_name_to_number("土") == 7
        assert convert_day_name_to_number("日") == 1

    def test_with_suffix(self):
        assert convert_day_name_to_number("月曜") == 2
        assert convert_day_name_to_number("月曜日") == 2

    def test_invalid(self):
        assert convert_day_name_to_number("abc") == 0
        assert convert_day_name_to_number("") == 0

    def test_with_whitespace(self):
        assert convert_day_name_to_number("  月  ") == 2


# ============================================
# get_customer_delivery_days
# ============================================
class TestGetCustomerDeliveryDays:
    def test_found(self):
        cache = CacheStore()
        cache.cust_days = {
            "岡安産業（株）　千葉営業所": [2, 4, 6],
            "共同ガス（株）　本社": [],
        }
        assert get_customer_delivery_days("岡安産業（株）　千葉営業所", cache) == [2, 4, 6]

    def test_empty_days(self):
        cache = CacheStore()
        cache.cust_days = {"顧客A": []}
        assert get_customer_delivery_days("顧客A", cache) == []

    def test_not_found(self):
        cache = CacheStore()
        assert get_customer_delivery_days("存在しない顧客", cache) == []


# ============================================
# get_retention_days
# ============================================
class TestGetRetentionDays:
    def test_found(self):
        cache = CacheStore()
        cache.cust_retention = {"顧客A": 2, "顧客B": 0}
        assert get_retention_days("顧客A", cache) == 2

    def test_zero(self):
        cache = CacheStore()
        cache.cust_retention = {"顧客B": 0}
        assert get_retention_days("顧客B", cache) == 0

    def test_not_found(self):
        cache = CacheStore()
        assert get_retention_days("存在しない", cache) == 0


# ============================================
# is_route_delivery
# ============================================
class TestIsRouteDelivery:
    def test_true(self):
        cache = CacheStore()
        cache.cust_route = {"共同ガス（株）　本社": True}
        assert is_route_delivery("共同ガス（株）　本社", cache) is True

    def test_false(self):
        cache = CacheStore()
        cache.cust_route = {"岡安産業（株）　千葉営業所": False}
        assert is_route_delivery("岡安産業（株）　千葉営業所", cache) is False

    def test_not_found(self):
        cache = CacheStore()
        assert is_route_delivery("存在しない", cache) is False


# ============================================
# get_email_addresses
# ============================================
class TestGetEmailAddresses:
    def test_found(self):
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便", "mail1", "mail2"],
            ["顧客A", "月水金", 2, "", "a@test.com", "b@test.com"],
            ["顧客B", "", 0, "", "c@test.com", ""],
        ]
        ws = MockWorksheet(data)
        assert get_email_addresses("顧客A", ws) == "a@test.com; b@test.com"
        assert get_email_addresses("顧客B", ws) == "c@test.com"

    def test_not_found(self):
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便", "mail1"],
            ["顧客A", "", 0, "", "a@test.com"],
        ]
        ws = MockWorksheet(data)
        assert get_email_addresses("存在しない", ws) == ""

    def test_no_email(self):
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便"],
            ["顧客A", "", 0, ""],
        ]
        ws = MockWorksheet(data)
        assert get_email_addresses("顧客A", ws) == ""

    def test_none_worksheet(self):
        assert get_email_addresses("顧客A", None) == ""


# ============================================
# check_customer_master
# ============================================
class TestCheckCustomerMaster:
    def test_all_registered(self):
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便", "mail1"],
            ["顧客A", "", 0, "", "a@test.com"],
            ["顧客B", "", 0, "", "b@test.com"],
        ]
        ws = MockWorksheet(data)
        result = check_customer_master(["顧客A", "顧客B"], ws)
        assert result == ""

    def test_missing_email(self):
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便"],
            ["顧客A", "", 0, ""],  # メールアドレスなし
            ["顧客B", "", 0, ""],
        ]
        ws = MockWorksheet(data)
        result = check_customer_master(["顧客A", "顧客B"], ws)
        assert "・顧客A" in result
        assert "・顧客B" in result

    def test_missing_customer(self):
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便", "mail1"],
            ["顧客A", "", 0, "", "a@test.com"],
        ]
        ws = MockWorksheet(data)
        result = check_customer_master(["顧客A", "顧客C"], ws)
        assert "・顧客A" not in result
        assert "・顧客C" in result

    def test_none_worksheet(self):
        result = check_customer_master(["顧客A"], None)
        assert "・顧客A" in result
