"""config.py のユニットテスト"""

import datetime

import pytest

from nouki_kaitou.config import (
    get_branch_settings,
    load_branch_settings,
    load_holidays,
)
from nouki_kaitou.models import BranchSettings


# ============================================
# モックWorkbook/Worksheetクラス
# ============================================
class MockWorksheet:
    """openpyxlのWorksheetをモックする簡易クラス"""

    def __init__(self, data: list[list]):
        self._data = data

    def iter_rows(self, min_row=1, max_col=None, values_only=False):
        for row in self._data[min_row - 1:]:
            if max_col is not None:
                yield tuple(row[:max_col])
            else:
                yield tuple(row)


class MockWorkbook:
    """openpyxlのWorkbookをモックする簡易クラス"""

    def __init__(self, sheets: dict[str, MockWorksheet]):
        self._sheets = sheets

    def __getitem__(self, key: str) -> MockWorksheet:
        if key not in self._sheets:
            raise KeyError(key)
        return self._sheets[key]


# ============================================
# load_holidays
# ============================================
class TestLoadHolidays:
    def test_normal(self):
        """祝日と特別締切時間の読み込み"""
        data = [
            ["日付", "説明", "締切時間"],  # ヘッダー
            [datetime.date(2026, 1, 1), "元旦", None],       # 祝日（C列空）
            [datetime.date(2026, 1, 13), "成人の日", None],   # 祝日
            [datetime.date(2026, 12, 30), "年末", 12],        # 特別締切時間12時
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"特別日カレンダー": ws})

        result = load_holidays(wb)

        assert datetime.date(2026, 1, 1) in result
        assert result[datetime.date(2026, 1, 1)] is None      # 祝日
        assert datetime.date(2026, 1, 13) in result
        assert result[datetime.date(2026, 1, 13)] is None
        assert result[datetime.date(2026, 12, 30)] == 12       # 特別締切12時

    def test_empty_calendar(self):
        """空のカレンダー"""
        data = [["日付", "説明", "締切時間"]]  # ヘッダーのみ
        ws = MockWorksheet(data)
        wb = MockWorkbook({"特別日カレンダー": ws})

        result = load_holidays(wb)
        assert result == {}

    def test_missing_sheet(self):
        """シートが存在しない場合"""
        wb = MockWorkbook({})  # 空のWorkbook
        result = load_holidays(wb)
        assert result == {}

    def test_invalid_date(self):
        """無効な日付はスキップ"""
        data = [
            ["日付", "説明", "締切時間"],
            ["invalid", "テスト", None],     # 無効日付
            [None, "テスト2", None],          # None
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"特別日カレンダー": ws})

        result = load_holidays(wb)
        assert result == {}

    def test_cutoff_value_non_numeric(self):
        """C列が数値以外の場合はNone（祝日扱い）"""
        data = [
            ["日付", "説明", "締切時間"],
            [datetime.date(2026, 3, 21), "春分の日", "abc"],  # 数値変換失敗
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"特別日カレンダー": ws})

        result = load_holidays(wb)
        assert result[datetime.date(2026, 3, 21)] is None


# ============================================
# load_branch_settings
# ============================================
class TestLoadBranchSettings:
    def _make_source_data(self, order_num: str = "GL2Z444369") -> list[list[str]]:
        """注番を含む最小限のsource_dataを構築"""
        rows = [[""] * 5 for _ in range(6)]  # 行0-5
        # 行6: データ行（注番=0列目とする）
        rows.append([order_num, "", "顧客名"])
        return rows

    def test_normal(self):
        """正常な営業所設定の読み込み"""
        branch_data = [
            ["コード", "営業所名", "締切", "センター", "メール", "開始日"],
            ["GL", "京葉営業所", 15, "関東商品センター", "keiyou@mac-exe.co.jp", datetime.date(2026, 1, 6)],
        ]
        branch_ws = MockWorksheet(branch_data)
        wb = MockWorkbook({"営業所設定": branch_ws})

        cols = {"受発注伝票": 0}
        source = self._make_source_data("GL2Z444369")

        settings = load_branch_settings(wb, source, cols)

        assert settings.name == "京葉営業所"
        assert settings.default_cutoff == 15
        assert settings.base_center == "関東商品センター"
        assert settings.shared_email == "keiyou@mac-exe.co.jp"
        assert "京葉営業所" in settings.signature

    def test_no_matching_branch(self):
        """営業所コード不一致→デフォルト設定"""
        branch_data = [
            ["コード", "営業所名", "締切", "センター", "メール", "開始日"],
            ["TK", "東京営業所", 14, "東京センター", "tokyo@mac-exe.co.jp", None],
        ]
        branch_ws = MockWorksheet(branch_data)
        wb = MockWorkbook({"営業所設定": branch_ws})

        cols = {"受発注伝票": 0}
        source = self._make_source_data("GL2Z444369")

        settings = load_branch_settings(wb, source, cols)
        assert settings.name == ""  # デフォルト

    def test_no_branch_sheet(self):
        """営業所設定シートなし→デフォルト"""
        wb = MockWorkbook({})
        cols = {"受発注伝票": 0}
        source = self._make_source_data("GL2Z444369")

        settings = load_branch_settings(wb, source, cols)
        assert settings.name == ""

    def test_no_alpha_order_num(self):
        """英字プレフィックスの注番がない場合→デフォルト"""
        branch_data = [
            ["コード", "営業所名", "締切", "センター", "メール", "開始日"],
            ["GL", "京葉営業所", 15, "関東商品センター", "", None],
        ]
        branch_ws = MockWorksheet(branch_data)
        wb = MockWorkbook({"営業所設定": branch_ws})

        cols = {"受発注伝票": 0}
        source = self._make_source_data("12345678")  # 英字プレフィックスなし

        settings = load_branch_settings(wb, source, cols)
        assert settings.name == ""

    def test_no_order_col(self):
        """受発注伝票列がない場合→デフォルト"""
        wb = MockWorkbook({})
        cols: dict[str, int] = {}
        source = self._make_source_data()

        settings = load_branch_settings(wb, source, cols)
        assert settings.name == ""

    def test_remarks_mode_renrakujikou(self):
        """G列「連絡事項」→ remarks_mode=external"""
        branch_data = [
            ["コード", "営業所名", "締切", "センター", "メール", "開始日", "注番列の表示"],
            ["GL", "松本営業所", 15, "関東商品センター", "", None, "連絡事項"],
        ]
        branch_ws = MockWorksheet(branch_data)
        wb = MockWorkbook({"営業所設定": branch_ws})
        cols = {"受発注伝票": 0}
        source = self._make_source_data("GL2Z444369")

        settings = load_branch_settings(wb, source, cols)
        assert settings.remarks_mode == "external"

    def test_remarks_mode_empty(self):
        """G列が空欄 → remarks_mode=detail（デフォルト）"""
        branch_data = [
            ["コード", "営業所名", "締切", "センター", "メール", "開始日", "注番列の表示"],
            ["GL", "京葉営業所", 15, "関東商品センター", "", None, ""],
        ]
        branch_ws = MockWorksheet(branch_data)
        wb = MockWorkbook({"営業所設定": branch_ws})
        cols = {"受発注伝票": 0}
        source = self._make_source_data("GL2Z444369")

        settings = load_branch_settings(wb, source, cols)
        assert settings.remarks_mode == "detail"

    def test_remarks_mode_no_g_column(self):
        """G列がない旧フォーマット → remarks_mode=detail（デフォルト）"""
        branch_data = [
            ["コード", "営業所名", "締切", "センター", "メール", "開始日"],
            ["GL", "京葉営業所", 15, "関東商品センター", "", None],
        ]
        branch_ws = MockWorksheet(branch_data)
        wb = MockWorkbook({"営業所設定": branch_ws})
        cols = {"受発注伝票": 0}
        source = self._make_source_data("GL2Z444369")

        settings = load_branch_settings(wb, source, cols)
        assert settings.remarks_mode == "detail"

    def test_remarks_mode_unknown_value(self):
        """G列に不明な値 → remarks_mode=detail（デフォルト）"""
        branch_data = [
            ["コード", "営業所名", "締切", "センター", "メール", "開始日", "注番列の表示"],
            ["GL", "京葉営業所", 15, "関東商品センター", "", None, "不明"],
        ]
        branch_ws = MockWorksheet(branch_data)
        wb = MockWorkbook({"営業所設定": branch_ws})
        cols = {"受発注伝票": 0}
        source = self._make_source_data("GL2Z444369")

        settings = load_branch_settings(wb, source, cols)
        assert settings.remarks_mode == "detail"


# ============================================
# get_branch_settings
# ============================================
class TestGetBranchSettings:
    def test_normal(self):
        """営業所名・締切時間・署名の取得"""
        branch = BranchSettings(
            name="京葉営業所",
            default_cutoff=15,
            signature="マツモト産業\n京葉営業所",
        )
        name, cutoff, sig = get_branch_settings(branch)
        assert name == "京葉営業所"
        assert cutoff == 15
        assert "京葉営業所" in sig

    def test_special_cutoff_override(self):
        """特別締切時間で上書き"""
        branch = BranchSettings(
            name="京葉営業所",
            default_cutoff=15,
            signature="マツモト産業\n京葉営業所",
        )
        holidays = {
            datetime.date(2026, 12, 30): 12,  # 特別締切12時
        }
        name, cutoff, sig = get_branch_settings(
            branch, holidays, target_date=datetime.date(2026, 12, 30)
        )
        assert cutoff == 12  # 12時に上書きされている

    def test_holiday_no_cutoff(self):
        """祝日（締切時間なし）→デフォルト締切のまま"""
        branch = BranchSettings(
            name="京葉営業所",
            default_cutoff=15,
        )
        holidays = {
            datetime.date(2026, 1, 1): None,  # 祝日（締切時間なし）
        }
        name, cutoff, sig = get_branch_settings(
            branch, holidays, target_date=datetime.date(2026, 1, 1)
        )
        assert cutoff == 15  # 変更されない

    def test_no_holidays(self):
        """祝日辞書なし→デフォルト"""
        branch = BranchSettings(default_cutoff=15)
        name, cutoff, sig = get_branch_settings(branch, None)
        assert cutoff == 15

    def test_non_holiday_date(self):
        """祝日でない日→デフォルト"""
        branch = BranchSettings(default_cutoff=15)
        holidays = {
            datetime.date(2026, 1, 1): None,
        }
        name, cutoff, sig = get_branch_settings(
            branch, holidays, target_date=datetime.date(2026, 1, 5)
        )
        assert cutoff == 15
