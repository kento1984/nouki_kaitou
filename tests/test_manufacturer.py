"""manufacturer.py のユニットテスト"""

from nouki_kaitou.manufacturer import (
    get_delivery_days_to_add,
    get_manufacturer_name,
)
from nouki_kaitou.models import CacheStore


class TestGetManufacturerName:
    def test_found(self):
        cache = CacheStore()
        cache.mfg_name = {"D01": "（株）ダイヘン", "K01": "コベルコ溶接テクノ"}
        assert get_manufacturer_name("D01", cache) == "（株）ダイヘン"

    def test_not_found(self):
        cache = CacheStore()
        cache.mfg_name = {"D01": "ダイヘン"}
        assert get_manufacturer_name("X99", cache) == ""

    def test_empty_code(self):
        cache = CacheStore()
        cache.mfg_name = {"D01": "ダイヘン"}
        assert get_manufacturer_name("", cache) == ""

    def test_whitespace_code(self):
        cache = CacheStore()
        cache.mfg_name = {"D01": "ダイヘン"}
        assert get_manufacturer_name("  D01  ", cache) == "ダイヘン"


class TestGetDeliveryDaysToAdd:
    def test_found(self):
        cache = CacheStore()
        cache.mfg_days = {"D01": 2, "K01": 3}
        assert get_delivery_days_to_add("D01", cache) == 2
        assert get_delivery_days_to_add("K01", cache) == 3

    def test_not_found_returns_default(self):
        cache = CacheStore()
        assert get_delivery_days_to_add("X99", cache) == 2

    def test_empty_code(self):
        cache = CacheStore()
        assert get_delivery_days_to_add("", cache) == 2
