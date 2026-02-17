"""confirming.py のユニットテスト"""

import datetime

from nouki_kaitou.confirming import (
    get_confirmed_delivery_date,
    get_confirming_status,
)
from nouki_kaitou.models import CacheStore


class TestGetConfirmedDeliveryDate:
    def test_found_with_date(self):
        cache = CacheStore()
        cache.confirm = {
            "GL2Z444462|10": ("未", "未処理", None),
            "GL2F445985|10": ("済", "回答待ち", datetime.date(2026, 3, 15)),
        }
        result = get_confirmed_delivery_date("GL2F445985", "10", cache)
        assert result == datetime.date(2026, 3, 15)

    def test_found_no_date(self):
        cache = CacheStore()
        cache.confirm = {
            "GL2Z444462|10": ("未", "未処理", None),
        }
        result = get_confirmed_delivery_date("GL2Z444462", "10", cache)
        assert result is None

    def test_not_found(self):
        cache = CacheStore()
        result = get_confirmed_delivery_date("NOTEXIST", "10", cache)
        assert result is None

    def test_empty_order(self):
        cache = CacheStore()
        assert get_confirmed_delivery_date("", "10", cache) is None

    def test_empty_detail(self):
        cache = CacheStore()
        assert get_confirmed_delivery_date("GL2Z444462", "", cache) is None


class TestGetConfirmingStatus:
    def test_found(self):
        cache = CacheStore()
        cache.confirm = {
            "GL2Z444462|10": ("未", "未処理", None),
            "GL2F445985|10": ("済", "分納", datetime.date(2026, 3, 15)),
        }
        assert get_confirming_status("GL2Z444462", "10", cache) == "未処理"
        assert get_confirming_status("GL2F445985", "10", cache) == "分納"

    def test_not_found(self):
        cache = CacheStore()
        assert get_confirming_status("NOTEXIST", "10", cache) == ""

    def test_empty_order(self):
        cache = CacheStore()
        assert get_confirming_status("", "10", cache) == ""

    def test_empty_detail(self):
        cache = CacheStore()
        assert get_confirming_status("GL2Z444462", "", cache) == ""
