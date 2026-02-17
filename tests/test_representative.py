"""representative.py のユニットテスト"""

from nouki_kaitou.representative import (
    contains_rep,
    get_rep_email_addresses,
    get_rep_list,
    is_split_by_rep,
    parse_rep_names,
    should_include_for_rep,
)


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
# parse_rep_names
# ============================================
class TestParseRepNames:
    def test_comma_separated(self):
        assert parse_rep_names("田中、鈴木") == ["田中", "鈴木"]

    def test_half_comma(self):
        assert parse_rep_names("田中,鈴木") == ["田中", "鈴木"]

    def test_dot_separated(self):
        assert parse_rep_names("田中・鈴木") == ["田中", "鈴木"]

    def test_sama_separated(self):
        """「様」が区切り文字として扱われる"""
        assert parse_rep_names("柏原様首藤様") == ["柏原", "首藤"]

    def test_single_name(self):
        assert parse_rep_names("柏原") == ["柏原"]

    def test_empty(self):
        assert parse_rep_names("") == []

    def test_whitespace(self):
        assert parse_rep_names("  ") == []

    def test_mixed(self):
        assert parse_rep_names("柏原様、首藤") == ["柏原", "首藤"]

    def test_trailing_sama(self):
        """末尾の様でも空文字は除外される"""
        assert parse_rep_names("柏原様") == ["柏原"]


# ============================================
# contains_rep
# ============================================
class TestContainsRep:
    def test_found(self):
        assert contains_rep(["柏原", "首藤", "入山"], "首藤") is True

    def test_not_found(self):
        assert contains_rep(["柏原", "首藤"], "田中") is False

    def test_empty_list(self):
        assert contains_rep([], "柏原") is False


# ============================================
# is_split_by_rep
# ============================================
class TestIsSplitByRep:
    def test_found(self):
        data = [
            ["顧客名", "担当者名", "mail1"],
            ["マツモト産業（株）", "柏原", "kashi@test.com"],
            ["マツモト産業（株）", "首藤", "shuto@test.com"],
        ]
        ws = MockWorksheet(data)
        assert is_split_by_rep("マツモト産業（株）", ws) is True

    def test_not_found(self):
        data = [
            ["顧客名", "担当者名", "mail1"],
            ["マツモト産業（株）", "柏原", "kashi@test.com"],
        ]
        ws = MockWorksheet(data)
        assert is_split_by_rep("存在しない顧客", ws) is False

    def test_none_ws(self):
        assert is_split_by_rep("任意", None) is False


# ============================================
# get_rep_list
# ============================================
class TestGetRepList:
    def test_normal(self):
        data = [
            ["顧客名", "担当者名", "mail1"],
            ["マツモト産業（株）", "柏原", "kashi@test.com"],
            ["マツモト産業（株）", "首藤", "shuto@test.com"],
            ["マツモト産業（株）", "入山", "iriyama@test.com"],
        ]
        ws = MockWorksheet(data)
        result = get_rep_list("マツモト産業（株）", ws)
        assert result == ["柏原", "首藤", "入山"]

    def test_remove_sama(self):
        """末尾の「様」を除去"""
        data = [
            ["顧客名", "担当者名", "mail1"],
            ["テスト", "柏原様", "test@test.com"],
        ]
        ws = MockWorksheet(data)
        result = get_rep_list("テスト", ws)
        assert result == ["柏原"]

    def test_no_duplicates(self):
        """重複排除"""
        data = [
            ["顧客名", "担当者名", "mail1"],
            ["テスト", "柏原", "test1@test.com"],
            ["テスト", "柏原", "test2@test.com"],
        ]
        ws = MockWorksheet(data)
        result = get_rep_list("テスト", ws)
        assert result == ["柏原"]

    def test_not_found(self):
        data = [
            ["顧客名", "担当者名"],
            ["テスト", "柏原"],
        ]
        ws = MockWorksheet(data)
        assert get_rep_list("存在しない", ws) == []

    def test_none_ws(self):
        assert get_rep_list("任意", None) == []


# ============================================
# get_rep_email_addresses
# ============================================
class TestGetRepEmailAddresses:
    def test_found(self):
        data = [
            ["顧客名", "担当者名", "mail1", "mail2"],
            ["テスト", "柏原", "kashi1@test.com", "kashi2@test.com"],
            ["テスト", "首藤", "shuto@test.com", ""],
        ]
        ws = MockWorksheet(data)
        result = get_rep_email_addresses("テスト", "柏原", ws)
        assert result == "kashi1@test.com; kashi2@test.com"

    def test_with_sama(self):
        """マスターの担当者名に「様」がついている場合"""
        data = [
            ["顧客名", "担当者名", "mail1"],
            ["テスト", "柏原様", "kashi@test.com"],
        ]
        ws = MockWorksheet(data)
        result = get_rep_email_addresses("テスト", "柏原", ws)
        assert result == "kashi@test.com"

    def test_not_found(self):
        data = [
            ["顧客名", "担当者名", "mail1"],
            ["テスト", "柏原", "kashi@test.com"],
        ]
        ws = MockWorksheet(data)
        assert get_rep_email_addresses("テスト", "田中", ws) == ""

    def test_none_ws(self):
        assert get_rep_email_addresses("テスト", "柏原", None) == ""


# ============================================
# should_include_for_rep
# ============================================
class TestShouldIncludeForRep:
    def test_no_filter(self):
        """rep_name空 → 常にTrue"""
        assert should_include_for_rep("柏原", "", []) is True
        assert should_include_for_rep("", "", []) is True

    def test_specific_rep_found(self):
        """特定担当者がセルに含まれている"""
        assert should_include_for_rep("柏原、首藤", "柏原", []) is True

    def test_specific_rep_not_found(self):
        """特定担当者がセルに含まれていない"""
        assert should_include_for_rep("柏原、首藤", "入山", []) is False

    def test_other_empty_cell(self):
        """__OTHER__: 空欄 → True"""
        assert should_include_for_rep("", "__OTHER__", ["柏原"]) is True

    def test_other_unregistered(self):
        """__OTHER__: 未登録担当者がいる → True"""
        assert should_include_for_rep(
            "田中、鈴木", "__OTHER__", ["柏原", "首藤"]
        ) is True

    def test_other_all_registered(self):
        """__OTHER__: 全員登録済み → False"""
        assert should_include_for_rep(
            "柏原、首藤", "__OTHER__", ["柏原", "首藤", "入山"]
        ) is False

    def test_sama_in_cell(self):
        """セル値に「様」が含まれる場合のパース"""
        assert should_include_for_rep("柏原様", "柏原", []) is True
