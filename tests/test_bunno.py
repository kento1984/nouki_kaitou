"""bunno.py のユニットテスト

分納処理の全10関数を網羅的にテスト。
"""

import datetime

import pytest

from nouki_kaitou.bunno import (
    calculate_bunno_date,
    extract_bunno_info,
    extract_location_from_token,
    has_bunno_kakuninchu,
    has_bunno_mitei,
    is_date_token,
    normalize_bunno_date,
    remove_bunno_text,
    split_qty_and_date,
    starts_with_number,
)
from nouki_kaitou.models import BunnoEntry, CacheStore


# ============================================
# StartsWithNumber
# ============================================
class TestStartsWithNumber:
    def test_digit_start(self):
        assert starts_with_number("100個") is True

    def test_alpha_start(self):
        assert starts_with_number("ABC") is False

    def test_japanese_start(self):
        assert starts_with_number("未定") is False

    def test_empty(self):
        assert starts_with_number("") is False

    def test_single_digit(self):
        assert starts_with_number("5") is True

    def test_zero_start(self):
        assert starts_with_number("0.5kg") is True


# ============================================
# IsDateToken
# ============================================
class TestIsDateToken:
    def test_mitei(self):
        assert is_date_token("未定") is True

    def test_slash_date(self):
        assert is_date_token("12/20") is True

    def test_fullwidth_slash(self):
        assert is_date_token("１２／２０") is True

    def test_month_period(self):
        assert is_date_token("3月上旬予定") is True

    def test_keppin(self):
        assert is_date_token("欠品中") is True

    def test_kakuninchu(self):
        assert is_date_token("確認中") is True

    def test_quantity(self):
        assert is_date_token("100個") is False

    def test_location(self):
        assert is_date_token("東京") is False

    def test_empty(self):
        assert is_date_token("") is False


# ============================================
# NormalizeBunnoDate
# ============================================
class TestNormalizeBunnoDate:
    def test_remove_shukka(self):
        assert normalize_bunno_date("12/22出荷") == "12/22"

    def test_remove_chaku(self):
        assert normalize_bunno_date("12/22着") == "12/22"

    def test_keppin_to_mitei(self):
        assert normalize_bunno_date("欠品中納期確認中") == "未定"

    def test_kakuninchu_to_mitei(self):
        assert normalize_bunno_date("確認中") == "未定"

    def test_yotei_without_jun(self):
        """「予定」は○旬以外では除去"""
        assert normalize_bunno_date("12/22予定") == "12/22"

    def test_yotei_with_jun(self):
        """「○旬予定」はそのまま保持"""
        assert normalize_bunno_date("3月上旬予定") == "3月上旬予定"

    def test_plain_date(self):
        assert normalize_bunno_date("12/22") == "12/22"

    def test_chuujun_yotei(self):
        assert normalize_bunno_date("1月中旬予定") == "1月中旬予定"

    def test_gejun_yotei(self):
        assert normalize_bunno_date("2月下旬予定") == "2月下旬予定"

    def test_trim(self):
        assert normalize_bunno_date("  12/22  ") == "12/22"

    def test_shukka_yotei(self):
        """「出荷予定」→「出荷」も「予定」も除去 → 空 or 日付"""
        assert normalize_bunno_date("12/22出荷予定") == "12/22"


# ============================================
# ExtractLocationFromToken
# ============================================
class TestExtractLocationFromToken:
    def test_fullwidth_brackets(self):
        token, location = extract_location_from_token("700m（滋賀）")
        assert token == "700m"
        assert location == "滋賀"

    def test_halfwidth_brackets(self):
        token, location = extract_location_from_token("300個(東京)")
        assert token == "300個"
        assert location == "東京"

    def test_no_brackets(self):
        token, location = extract_location_from_token("500本")
        assert token == "500本"
        assert location == ""

    def test_only_brackets(self):
        token, location = extract_location_from_token("（大阪）")
        assert token == ""
        assert location == "大阪"

    def test_nested_text(self):
        """括弧内に文字列がある場合"""
        token, location = extract_location_from_token("100個（ABC工場）")
        assert token == "100個"
        assert location == "ABC工場"


# ============================================
# SplitQtyAndDate
# ============================================
class TestSplitQtyAndDate:
    def test_split_ko_date(self):
        result = split_qty_and_date("1個12/19")
        assert result == ("1個", "12/19")

    def test_split_hon_date(self):
        result = split_qty_and_date("5本12/20")
        assert result == ("5本", "12/20")

    def test_split_m_date(self):
        result = split_qty_and_date("700m12/17")
        assert result == ("700m", "12/17")

    def test_split_kg_date(self):
        result = split_qty_and_date("10kg3/5")
        assert result == ("10kg", "3/5")

    def test_split_mitei(self):
        result = split_qty_and_date("3個未定")
        assert result == ("3個", "未定")

    def test_no_split_quantity_only(self):
        result = split_qty_and_date("100個")
        assert result is None

    def test_no_split_no_unit(self):
        result = split_qty_and_date("12/19")
        assert result is None

    def test_set_unit(self):
        result = split_qty_and_date("2セット1/15")
        assert result == ("2セット", "1/15")

    def test_no_split_unit_then_text(self):
        """単位の後が日付でもなく「未定」でもなければ分割しない"""
        result = split_qty_and_date("5個東京")
        assert result is None

    def test_split_fukuro_date(self):
        """袋単位で分割"""
        result = split_qty_and_date("10袋12/25")
        assert result == ("10袋", "12/25")

    def test_split_maki_date(self):
        """巻単位で分割"""
        result = split_qty_and_date("3巻1/8")
        assert result == ("3巻", "1/8")

    def test_split_cho_date(self):
        """丁単位で分割"""
        result = split_qty_and_date("5丁2/14")
        assert result == ("5丁", "2/14")

    def test_split_kumi_date(self):
        """組単位で分割"""
        result = split_qty_and_date("2組3/20")
        assert result == ("2組", "3/20")

    def test_split_fukuro_mitei(self):
        """袋単位で未定と分割"""
        result = split_qty_and_date("5袋未定")
        assert result == ("5袋", "未定")


# ============================================
# ExtractBunnoInfo
# ============================================
class TestExtractBunnoInfo:
    def test_basic(self):
        result = extract_bunno_info("分納:700m 12/17 滋賀、300m 12/17 東京")
        assert len(result) == 2
        assert result[0].quantity == "700m"
        assert result[0].date_str == "12/17"
        assert result[0].location == "滋賀"
        assert result[1].quantity == "300m"
        assert result[1].date_str == "12/17"
        assert result[1].location == "東京"

    def test_fullwidth_colon(self):
        result = extract_bunno_info("分納：50個 1/10,30個 未定")
        assert len(result) == 2
        assert result[0].quantity == "50個"
        assert result[0].date_str == "1/10"
        assert result[1].quantity == "30個"
        assert result[1].date_str == "未定"

    def test_no_bunno(self):
        result = extract_bunno_info("通常コメント")
        assert len(result) == 0

    def test_empty(self):
        result = extract_bunno_info("")
        assert len(result) == 0

    def test_qty_date_combined(self):
        """数量と日付が連結しているパターン"""
        result = extract_bunno_info("分納:1個12/19,2個12/20")
        assert len(result) == 2
        assert result[0].quantity == "1個"
        assert result[0].date_str == "12/19"
        assert result[1].quantity == "2個"
        assert result[1].date_str == "12/20"

    def test_with_brackets_location(self):
        """括弧内に場所があるパターン"""
        result = extract_bunno_info("分納:50個（貴社） 12/20,30個（工場） 1/15")
        assert len(result) == 2
        assert result[0].location == "貴社"
        assert result[1].location == "工場"

    def test_no_date_defaults_mitei(self):
        """日付がなければ未定"""
        result = extract_bunno_info("分納:100個 東京")
        assert len(result) == 1
        assert result[0].date_str == "未定"
        assert result[0].location == "東京"

    def test_fullwidth_numbers(self):
        """全角数字は半角に変換される"""
        result = extract_bunno_info("分納:５０個 １２／２０")
        assert len(result) == 1
        assert result[0].quantity == "50個"
        assert result[0].date_str == "12/20"

    def test_truncate_at_double_space(self):
        """2スペース以降は無視"""
        result = extract_bunno_info("分納:50個 12/20  欠品中ABC")
        assert len(result) == 1
        assert result[0].date_str == "12/20"

    def test_keppin_date_normalized(self):
        """欠品中は未定に正規化"""
        result = extract_bunno_info("分納:100個 欠品中")
        assert len(result) == 1
        assert result[0].date_str == "未定"

    def test_jun_yotei_preserved(self):
        """○旬予定はそのまま保持"""
        result = extract_bunno_info("分納:200個 3月上旬予定")
        assert len(result) == 1
        assert result[0].date_str == "3月上旬予定"

    def test_invalid_qty_skipped(self):
        """数値として無効な数量はスキップ"""
        result = extract_bunno_info("分納:ABC 12/20")
        assert len(result) == 0

    def test_multiple_comma_separated(self):
        """カンマ区切り複数"""
        result = extract_bunno_info("分納:10個 1/5,20個 1/10,30個 未定")
        assert len(result) == 3

    def test_fukuro_unit(self):
        """袋単位の分納"""
        result = extract_bunno_info("分納:10袋 1/20,5袋 2/5")
        assert len(result) == 2
        assert result[0].quantity == "10袋"
        assert result[1].quantity == "5袋"

    def test_maki_unit(self):
        """巻単位の分納"""
        result = extract_bunno_info("分納:3巻 12/25")
        assert len(result) == 1
        assert result[0].quantity == "3巻"

    def test_cho_unit(self):
        """丁単位の分納"""
        result = extract_bunno_info("分納:2丁 1/15 東京")
        assert len(result) == 1
        assert result[0].quantity == "2丁"
        assert result[0].location == "東京"

    def test_kumi_unit(self):
        """組単位の分納"""
        result = extract_bunno_info("分納:1組 3/10")
        assert len(result) == 1
        assert result[0].quantity == "1組"

    def test_mixed_new_units(self):
        """新単位の混在"""
        result = extract_bunno_info("分納:5袋 1/10,2巻 1/15,3丁 未定")
        assert len(result) == 3
        assert result[0].quantity == "5袋"
        assert result[1].quantity == "2巻"
        assert result[2].quantity == "3丁"
        assert result[2].date_str == "未定"


# ============================================
# RemoveBunnoText
# ============================================
class TestRemoveBunnoText:
    def test_remove_all(self):
        assert remove_bunno_text("分納:100個 12/20") == ""

    def test_remove_with_prefix(self):
        result = remove_bunno_text("ABC 分納:100個 12/20")
        assert result == "ABC "

    def test_remove_with_suffix(self):
        """2スペースで区切られている場合"""
        result = remove_bunno_text("分納:100個 12/20  残りコメント")
        assert result == "  残りコメント"

    def test_fullwidth_colon(self):
        result = remove_bunno_text("分納：50個 1/10")
        assert result == ""

    def test_no_bunno(self):
        assert remove_bunno_text("通常テキスト") == "通常テキスト"

    def test_empty(self):
        assert remove_bunno_text("") == ""


# ============================================
# HasBunnoMitei
# ============================================
class TestHasBunnoMitei:
    def test_has_mitei(self):
        entries = [BunnoEntry("100個", "未定", "")]
        assert has_bunno_mitei(entries) is True

    def test_no_mitei(self):
        entries = [BunnoEntry("100個", "12/20", "")]
        assert has_bunno_mitei(entries) is False

    def test_has_keppin(self):
        entries = [BunnoEntry("50個", "欠品中", "")]
        assert has_bunno_mitei(entries) is True

    def test_has_kakuninchu(self):
        entries = [BunnoEntry("50個", "確認中", "")]
        assert has_bunno_mitei(entries) is True

    def test_has_yotei(self):
        entries = [BunnoEntry("50個", "3月上旬予定", "")]
        assert has_bunno_mitei(entries) is True

    def test_empty_list(self):
        assert has_bunno_mitei([]) is False

    def test_mixed_confirmed_and_mitei(self):
        """確定済みと未定が混在"""
        entries = [
            BunnoEntry("50個", "12/20", ""),
            BunnoEntry("30個", "未定", ""),
        ]
        assert has_bunno_mitei(entries) is True

    def test_mitei_with_confirmed_date(self):
        """確認中一覧に確定日がある場合は未定とみなさない"""
        cache = CacheStore()
        cache.confirm["123|10"] = ("未", "分納", datetime.date(2026, 3, 15))
        entries = [BunnoEntry("100個", "未定", "")]
        assert has_bunno_mitei(
            entries, cache=cache,
            order_number="123", detail_number="10"
        ) is False

    def test_mitei_without_confirmed_date(self):
        """確認中一覧に確定日がない場合は未定"""
        cache = CacheStore()
        cache.confirm["123|10"] = ("未", "分納", None)
        entries = [BunnoEntry("100個", "未定", "")]
        assert has_bunno_mitei(
            entries, cache=cache,
            order_number="123", detail_number="10"
        ) is True


# ============================================
# HasBunnoKakuninchu
# ============================================
class TestHasBunnoKakuninchu:
    def test_has_kakuninchu(self):
        detail = [["50個", "未定", "", "確認中"]]
        assert has_bunno_kakuninchu(detail) is True

    def test_no_kakuninchu(self):
        detail = [["50個", "12/20", "", "12月20日出荷予定"]]
        assert has_bunno_kakuninchu(detail) is False

    def test_jun_yotei(self):
        """○旬予定は未確定"""
        detail = [["50個", "未定", "", "3月上旬予定"]]
        assert has_bunno_kakuninchu(detail) is True

    def test_shukka_yotei_excluded(self):
        """出荷予定は確定扱い"""
        detail = [["50個", "12/20", "", "12月20日出荷予定"]]
        assert has_bunno_kakuninchu(detail) is False

    def test_haitatsu_yotei_excluded(self):
        """配達予定は確定扱い"""
        detail = [["50個", "12/20", "", "12月22日配達予定"]]
        assert has_bunno_kakuninchu(detail) is False

    def test_empty(self):
        assert has_bunno_kakuninchu([]) is False

    def test_short_item(self):
        """4要素未満のアイテム"""
        detail = [["50個", "12/20", ""]]
        assert has_bunno_kakuninchu(detail) is False


# ============================================
# CalculateBunnoDate
# ============================================
class TestCalculateBunnoDate:
    """CalculateBunnoDateのテスト。today引数でテスト日付を固定。"""

    @pytest.fixture
    def today(self):
        return datetime.date(2026, 2, 16)

    def test_mitei_no_cache(self, today):
        """未定でキャッシュなし → 確認中"""
        result = calculate_bunno_date("未定", False, 1, today=today)
        assert result == "確認中"

    def test_mitei_with_confirmed_ship_rule(self, today):
        """未定 + 確定日あり + 直送ルール"""
        cache = CacheStore()
        cache.confirm["123|10"] = ("未", "分納", datetime.date(2026, 3, 10))
        result = calculate_bunno_date(
            "未定", is_ship_rule=True, days_to_add=1,
            cache=cache, order_number="123", detail_number="10",
            today=today,
        )
        assert result == "3月10日出荷予定"

    def test_mitei_with_confirmed_past_ship(self, today):
        """未定 + 確定日が過去 + 直送ルール → 出荷済み"""
        cache = CacheStore()
        cache.confirm["123|10"] = ("未", "分納", datetime.date(2026, 1, 15))
        result = calculate_bunno_date(
            "未定", is_ship_rule=True, days_to_add=1,
            cache=cache, order_number="123", detail_number="10",
            today=today,
        )
        assert result == "1月15日出荷済み"

    def test_mitei_with_confirmed_delivery(self, today):
        """未定 + 確定日あり + 自社便（配達ルール）"""
        cache = CacheStore()
        cache.confirm["123|10"] = ("未", "分納", datetime.date(2026, 3, 10))
        result = calculate_bunno_date(
            "未定", is_ship_rule=False, days_to_add=1,
            cache=cache, order_number="123", detail_number="10",
            today=today,
        )
        # 3/10 + 1営業日 = 3/11(水)
        assert result == "3月11日配達予定"

    def test_mitei_with_confirmed_rosenbin(self, today):
        """未定 + 確定日あり + 路線便"""
        cache = CacheStore()
        cache.confirm["123|10"] = ("未", "分納", datetime.date(2026, 3, 10))
        result = calculate_bunno_date(
            "未定", is_ship_rule=False, days_to_add=2,
            cache=cache, order_number="123", detail_number="10",
            is_rosenbin=True, today=today,
        )
        # 路線便: 3/10 + max(2-1, 0) = 1営業日 = 3/11(水)
        assert result == "3月11日出荷予定"

    def test_jun_yotei_no_cache(self, today):
        """○旬予定でキャッシュなし → そのまま"""
        result = calculate_bunno_date("3月上旬予定", False, 1, today=today)
        assert result == "3月上旬予定"

    def test_jun_yotei_with_confirmed(self, today):
        """○旬予定 + 確定日あり → 計算結果"""
        cache = CacheStore()
        cache.confirm["123|10"] = ("未", "分納", datetime.date(2026, 3, 5))
        result = calculate_bunno_date(
            "3月上旬予定", is_ship_rule=True, days_to_add=1,
            cache=cache, order_number="123", detail_number="10",
            today=today,
        )
        assert result == "3月5日出荷予定"

    def test_date_format_ship_future(self, today):
        """M/D形式 + 直送 + 未来 → 出荷予定"""
        result = calculate_bunno_date(
            "3/10", is_ship_rule=True, days_to_add=0,
            today=today,
        )
        assert result == "3月10日出荷予定"

    def test_date_format_ship_past(self, today):
        """M/D形式 + 直送 + 過去 → 出荷済み"""
        result = calculate_bunno_date(
            "1/15", is_ship_rule=True, days_to_add=0,
            today=today,
        )
        assert result == "1月15日出荷済み"

    def test_date_format_delivery_future(self, today):
        """M/D形式 + 自社便 + 未来 → 配達予定"""
        result = calculate_bunno_date(
            "3/10", is_ship_rule=False, days_to_add=1,
            today=today,
        )
        # 3/10 + 1営業日 = 3/11
        assert result == "3月11日配達予定"

    def test_date_format_rosenbin(self, today):
        """M/D形式 + 路線便"""
        result = calculate_bunno_date(
            "3/10", is_ship_rule=False, days_to_add=2,
            is_rosenbin=True, today=today,
        )
        # 路線便: 3/10 + max(2-1, 0) = 1営業日 = 3/11
        assert result == "3月11日出荷予定"

    def test_invalid_date(self, today):
        """不正な日付 → 確認中"""
        result = calculate_bunno_date("ABC", False, 1, today=today)
        assert result == "確認中"

    def test_invalid_month(self, today):
        """月が範囲外"""
        result = calculate_bunno_date("13/15", False, 1, today=today)
        assert result == "確認中"

    def test_past_date_next_year(self, today):
        """180日以上過去の日付 → 翌年"""
        # 2026/2/16基準で、7/15は未来なので7/15として処理
        result = calculate_bunno_date(
            "7/15", is_ship_rule=True, days_to_add=0,
            today=today,
        )
        assert result == "7月15日出荷予定"

    def test_rosenbin_zero_add(self, today):
        """路線便でdaysToAdd=0のとき、max(0-1, 0)=0"""
        result = calculate_bunno_date(
            "3/10", is_ship_rule=False, days_to_add=0,
            is_rosenbin=True, today=today,
        )
        # max(0-1, 0) = 0 → 日付そのまま
        assert result == "3月10日出荷予定"

    def test_no_slash(self, today):
        """スラッシュなし → 確認中"""
        result = calculate_bunno_date("ABC", False, 1, today=today)
        assert result == "確認中"
