"""stockout.py のユニットテスト"""

from nouki_kaitou.models import CacheStore
from nouki_kaitou.stockout import (
    extract_approx_delivery,
    get_storage_place_from_same_order,
    remove_stockout_text,
)


# ============================================
# extract_approx_delivery
# ============================================
class TestExtractApproxDelivery:
    def test_joujun(self):
        """上旬パターン"""
        assert extract_approx_delivery("欠品中 3月上旬予定") == "3月上旬入荷予定"

    def test_chuujun(self):
        """中旬パターン"""
        assert extract_approx_delivery("欠品中 3月中旬予定") == "3月中旬入荷予定"

    def test_gejun(self):
        """下旬パターン"""
        assert extract_approx_delivery("欠品中 3月下旬予定") == "3月下旬入荷予定"

    def test_slash_goro(self):
        """「○/○頃」パターン"""
        assert extract_approx_delivery("欠品中 3/15頃") == "3月15日頃入荷予定"

    def test_month_day_goro(self):
        """「○月○日頃」パターン"""
        assert extract_approx_delivery("欠品中 3月15日頃") == "3月15日頃入荷予定"

    def test_gatsu_matsu(self):
        """「○月末」パターン"""
        assert extract_approx_delivery("欠品中 3月末予定") == "3月末入荷予定"

    def test_no_stockout(self):
        """欠品中がない"""
        assert extract_approx_delivery("通常コメント") == ""

    def test_stockout_only(self):
        """欠品中のみ（アバウト納期なし）"""
        assert extract_approx_delivery("欠品中") == ""

    def test_empty(self):
        assert extract_approx_delivery("") == ""

    def test_full_width_digits(self):
        """全角数字を含む場合"""
        assert extract_approx_delivery("欠品中　３月上旬予定") == "3月上旬入荷予定"

    def test_with_prefix(self):
        """欠品中の前にテキストがある場合（VBA InStr互換: 部分一致）"""
        assert extract_approx_delivery("一部欠品中 3月上旬予定") == "3月上旬入荷予定"

    def test_stockout_with_space(self):
        """欠品中の後にスペース"""
        assert extract_approx_delivery("欠品中　3月中旬") == "3月中旬入荷予定"

    def test_invalid_month_day(self):
        """無効な月日"""
        assert extract_approx_delivery("欠品中 13/32頃") == ""


# ============================================
# remove_stockout_text
# ============================================
class TestRemoveStockoutText:
    def test_stockout_with_approx(self):
        """欠品中 + アバウト納期 → 全除去"""
        assert remove_stockout_text("欠品中 3月上旬予定") == ""

    def test_stockout_only(self):
        """欠品中のみ"""
        assert remove_stockout_text("欠品中") == ""

    def test_prefix_text(self):
        """前にテキストあり"""
        result = remove_stockout_text("テスト 欠品中 3月上旬予定")
        assert result == "テスト"

    def test_suffix_text(self):
        """後ろにテキストあり"""
        result = remove_stockout_text("欠品中 3月上旬予定 備考")
        assert "備考" in result

    def test_no_stockout(self):
        """欠品中なし → そのまま"""
        assert remove_stockout_text("通常コメント") == "通常コメント"

    def test_empty(self):
        assert remove_stockout_text("") == ""

    def test_stockout_with_goro(self):
        """「頃」パターン"""
        assert remove_stockout_text("欠品中 3/15頃") == ""

    def test_stockout_with_matsu(self):
        """「月末」パターン"""
        assert remove_stockout_text("欠品中 3月末") == ""


# ============================================
# get_storage_place_from_same_order
# ============================================
class TestGetStoragePlaceFromSameOrder:
    def test_found(self):
        cache = CacheStore()
        cache.storage = {
            "GL2Z444369": "転送中（直送用）",
            "GL2C444510": "関東商品センター",
        }
        result = get_storage_place_from_same_order("GL2Z444369", cache)
        assert result == "転送中（直送用）"

    def test_not_found(self):
        cache = CacheStore()
        assert get_storage_place_from_same_order("NOTEXIST", cache) == ""

    def test_empty_order(self):
        cache = CacheStore()
        assert get_storage_place_from_same_order("", cache) == ""
