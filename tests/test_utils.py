"""utils.py のユニットテスト"""

import datetime

import pytest

from nouki_kaitou.utils import (
    build_report_filename,
    build_sheet_name,
    convert_to_half_width,
    extract_date_from_string,
    format_date_japanese,
    is_december_31,
    is_numeric_char,
    normalize_name_for_comparison,
    parse_date,
    parse_time,
    to_circled_number,
    to_half_width_num,
)


# ============================================
# convert_to_half_width
# ============================================
class TestConvertToHalfWidth:
    def test_full_width_digits(self):
        assert convert_to_half_width("０１２３４５６７８９") == "0123456789"

    def test_full_width_slash(self):
        assert convert_to_half_width("３／１０") == "3/10"

    def test_full_width_colon(self):
        assert convert_to_half_width("１５：００") == "15:00"

    def test_mixed(self):
        assert convert_to_half_width("３月１０日") == "3月10日"

    def test_already_half_width(self):
        assert convert_to_half_width("12/31") == "12/31"

    def test_empty(self):
        assert convert_to_half_width("") == ""

    def test_none(self):
        assert convert_to_half_width(None) == ""

    def test_japanese_text_unchanged(self):
        assert convert_to_half_width("確認中") == "確認中"


# ============================================
# to_half_width_num
# ============================================
class TestToHalfWidthNum:
    def test_full_width_zero(self):
        assert to_half_width_num("０") == "0"

    def test_full_width_nine(self):
        assert to_half_width_num("９") == "9"

    def test_half_width(self):
        assert to_half_width_num("5") == "5"

    def test_non_digit(self):
        assert to_half_width_num("あ") == "あ"


# ============================================
# is_numeric_char
# ============================================
class TestIsNumericChar:
    def test_half_width(self):
        assert is_numeric_char("5") is True

    def test_full_width(self):
        assert is_numeric_char("５") is True

    def test_letter(self):
        assert is_numeric_char("A") is False

    def test_empty(self):
        assert is_numeric_char("") is False


# ============================================
# to_circled_number
# ============================================
class TestToCircledNumber:
    def test_one(self):
        assert to_circled_number(1) == "①"

    def test_ten(self):
        assert to_circled_number(10) == "⑩"

    def test_over_ten(self):
        assert to_circled_number(11) == "11"

    def test_zero(self):
        assert to_circled_number(0) == "0"


# ============================================
# extract_date_from_string
# ============================================
class TestExtractDateFromString:
    def test_standard_format(self):
        today = datetime.date.today()
        result = extract_date_from_string("3月15日出荷予定")
        assert result is not None
        assert result.month == 3
        assert result.day == 15

    def test_with_prefix(self):
        result = extract_date_from_string("1月8日配達済み")
        assert result is not None
        assert result.month == 1
        assert result.day == 8

    def test_no_date(self):
        assert extract_date_from_string("確認中") is None

    def test_empty(self):
        assert extract_date_from_string("") is None

    def test_none(self):
        assert extract_date_from_string(None) is None

    def test_invalid_month(self):
        assert extract_date_from_string("13月1日") is None

    def test_invalid_day(self):
        assert extract_date_from_string("2月32日") is None


# ============================================
# parse_date
# ============================================
class TestParseDate:
    def test_datetime_object(self):
        dt = datetime.datetime(2026, 1, 15, 10, 30)
        assert parse_date(dt) == datetime.date(2026, 1, 15)

    def test_date_object(self):
        d = datetime.date(2026, 3, 10)
        assert parse_date(d) == d

    def test_slash_string(self):
        assert parse_date("2026/1/15") == datetime.date(2026, 1, 15)

    def test_slash_string_zero_padded(self):
        assert parse_date("2026/01/05") == datetime.date(2026, 1, 5)

    def test_none(self):
        assert parse_date(None) is None

    def test_empty_string(self):
        assert parse_date("") is None

    def test_zero(self):
        assert parse_date(0) is None


# ============================================
# parse_time
# ============================================
class TestParseTime:
    def test_standard(self):
        assert parse_time("8:54:58") == (8, 54)

    def test_two_digit_hour(self):
        assert parse_time("15:30:00") == (15, 30)

    def test_none(self):
        assert parse_time(None) is None

    def test_empty(self):
        assert parse_time("") is None


# ============================================
# is_december_31
# ============================================
class TestIsDecember31:
    def test_true(self):
        assert is_december_31(datetime.date(2026, 12, 31)) is True

    def test_false(self):
        assert is_december_31(datetime.date(2026, 3, 15)) is False

    def test_none(self):
        assert is_december_31(None) is False


# ============================================
# format_date_japanese
# ============================================
class TestFormatDateJapanese:
    def test_standard(self):
        assert format_date_japanese(datetime.date(2026, 3, 15)) == "3月15日"

    def test_single_digit(self):
        assert format_date_japanese(datetime.date(2026, 1, 5)) == "1月5日"


# ============================================
# build_report_filename
# ============================================
class TestBuildReportFilename:
    def test_normal(self):
        dt = datetime.datetime(2026, 2, 14)
        result = build_report_filename("岡安産業（株）　千葉営業所", dt)
        assert result == "納期回答書_岡安産業（株）　千葉営業所様_20260214.xlsx"

    def test_with_rep(self):
        dt = datetime.datetime(2026, 2, 14)
        result = build_report_filename("マツモト産業（株）", dt, rep_name="柏原")
        assert result == "納期回答書_マツモト産業（株）様_柏原様_20260214.xlsx"

    def test_single_order_number(self):
        dt = datetime.datetime(2026, 2, 14)
        result = build_report_filename(
            "テスト顧客", dt, order_numbers=["GL2C444510"]
        )
        assert result == "納期回答書_テスト顧客様_GL2C444510_20260214.xlsx"

    def test_multiple_order_numbers(self):
        dt = datetime.datetime(2026, 2, 14)
        result = build_report_filename(
            "テスト顧客", dt, order_numbers=["GL2C444510", "GL2C444511"]
        )
        assert result == "納期回答書_テスト顧客様_複数注番_20260214.xlsx"


# ============================================
# build_sheet_name
# ============================================
class TestBuildSheetName:
    def test_normal(self):
        assert build_sheet_name("テスト顧客") == "テスト顧客様"

    def test_with_rep(self):
        assert build_sheet_name("テスト顧客", "柏原") == "テスト顧客_柏原様"

    def test_truncation(self):
        """31文字を超える場合は切り詰める"""
        long_name = "あ" * 30
        result = build_sheet_name(long_name)
        assert len(result) <= 31


# ============================================
# normalize_name_for_comparison
# ============================================
class TestNormalizeNameForComparison:
    def test_fullwidth_parens(self):
        """全角括弧→半角括弧"""
        assert normalize_name_for_comparison("（有）三橋機工") == "(有)三橋機工"

    def test_halfwidth_parens_unchanged(self):
        """半角括弧はそのまま"""
        assert normalize_name_for_comparison("(有)三橋機工") == "(有)三橋機工"

    def test_fullwidth_space(self):
        """全角スペース→半角スペース"""
        assert normalize_name_for_comparison("本多酸素（株）　八潮営業所") == "本多酸素(株) 八潮営業所"

    def test_strip(self):
        """前後の空白を除去"""
        assert normalize_name_for_comparison("  テスト商事  ") == "テスト商事"

    def test_same_after_normalize(self):
        """全角/半角の揺れがある同一顧客が一致する"""
        a = normalize_name_for_comparison("（有）三橋機工")
        b = normalize_name_for_comparison("(有)三橋機工")
        assert a == b

    def test_kabu_fullwidth_vs_halfwidth(self):
        """（株）vs (株)"""
        a = normalize_name_for_comparison("テスト（株）")
        b = normalize_name_for_comparison("テスト(株)")
        assert a == b

    def test_different_names(self):
        """異なる顧客名は異なる"""
        a = normalize_name_for_comparison("テスト商事（株）")
        b = normalize_name_for_comparison("別会社（有）")
        assert a != b

    def test_fullwidth_dash(self):
        """全角ハイフン→半角ハイフン"""
        assert normalize_name_for_comparison("テスト－商事") == "テスト-商事"

    def test_fullwidth_digits(self):
        """全角数字→半角数字"""
        assert normalize_name_for_comparison("第１工場") == "第1工場"

    def test_fullwidth_alpha(self):
        """全角英字→半角英字"""
        assert normalize_name_for_comparison("ＫＧＫサービス") == "KGKサービス"
