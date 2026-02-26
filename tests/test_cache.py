"""cache.py のユニットテスト"""

import datetime
from typing import Optional

import pytest

from nouki_kaitou.cache import (
    _detect_customer_master_format,
    _parse_weekday_string,
    build_all_caches,
    build_confirming_cache,
    build_customer_cache,
    build_manufacturer_cache,
    build_pattern_cache,
    build_storage_cache,
)
from nouki_kaitou.models import CacheStore, DeliveryPattern


# ============================================
# モックWorksheet/Workbookクラス
# ============================================
class MockWorksheet:
    """openpyxlのWorksheetをモックする簡易クラス"""

    def __init__(self, data: list[list]):
        self._data = data

    def iter_rows(self, min_row=1, max_row=None, max_col=None, values_only=False):
        end = max_row if max_row is not None else len(self._data)
        for row in self._data[min_row - 1:end]:
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
# _parse_weekday_string
# ============================================
class TestParseWeekdayString:
    def test_normal(self):
        assert _parse_weekday_string("月水金") == [2, 4, 6]

    def test_single(self):
        assert _parse_weekday_string("火") == [3]

    def test_all_days(self):
        assert _parse_weekday_string("日月火水木金土") == [1, 2, 3, 4, 5, 6, 7]

    def test_empty(self):
        assert _parse_weekday_string("") == []

    def test_non_weekday(self):
        assert _parse_weekday_string("ABC") == []

    def test_mixed(self):
        """曜日以外の文字が混じっている場合"""
        assert _parse_weekday_string("月・水・金") == [2, 4, 6]


# ============================================
# build_manufacturer_cache
# ============================================
class TestBuildManufacturerCache:
    def test_normal(self):
        """メーカー名と加算日数の読み込み"""
        data = [
            ["品目Group", "メーカー名", "加算日数"],  # ヘッダー
            ["D01", "（株）ダイヘン", 2],
            ["K01", "コベルコ溶接テクノ", 3],
            ["A01", "アズワン", None],  # 日数なし→デフォルト2
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"メーカー一覧": ws})

        mfg_name, mfg_days = build_manufacturer_cache(wb)

        assert mfg_name["D01"] == "（株）ダイヘン"
        assert mfg_name["K01"] == "コベルコ溶接テクノ"
        assert mfg_days["D01"] == 2
        assert mfg_days["K01"] == 3
        assert mfg_days["A01"] == 2  # デフォルト値

    def test_missing_sheet(self):
        """シートが存在しない場合"""
        wb = MockWorkbook({})
        mfg_name, mfg_days = build_manufacturer_cache(wb)
        assert mfg_name == {}
        assert mfg_days == {}

    def test_empty_key_skipped(self):
        """キーが空の行はスキップ"""
        data = [
            ["品目Group", "メーカー名", "加算日数"],
            ["", "空キー", 2],
            [None, "Noneキー", 2],
            ["D01", "（株）ダイヘン", 2],
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"メーカー一覧": ws})

        mfg_name, mfg_days = build_manufacturer_cache(wb)
        assert len(mfg_name) == 1
        assert "D01" in mfg_name

    def test_duplicate_key(self):
        """重複キーは最初の値を使用"""
        data = [
            ["品目Group", "メーカー名", "加算日数"],
            ["D01", "ダイヘン1", 2],
            ["D01", "ダイヘン2", 5],  # 重複、スキップされる
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"メーカー一覧": ws})

        mfg_name, mfg_days = build_manufacturer_cache(wb)
        assert mfg_name["D01"] == "ダイヘン1"
        assert mfg_days["D01"] == 2

    def test_invalid_days_value(self):
        """加算日数が不正な値→デフォルト2"""
        data = [
            ["品目Group", "メーカー名", "加算日数"],
            ["D01", "ダイヘン", "abc"],
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"メーカー一覧": ws})

        _, mfg_days = build_manufacturer_cache(wb)
        assert mfg_days["D01"] == 2


# ============================================
# build_customer_cache
# ============================================
class TestBuildCustomerCache:
    def test_normal(self):
        """顧客マスターの読み込み"""
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便"],  # ヘッダー
            ["岡安産業（株）　千葉営業所", "月水金", 2, ""],
            ["共同ガス（株）　本社", "", 0, "○"],   # 路線便あり
            ["コイケ酸商（株）　白井営業所", "火木", None, ""],
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"顧客マスター": ws})

        cust_days, cust_retention, cust_route, cust_pattern = build_customer_cache(wb)

        # 配送曜日
        assert cust_days["岡安産業（株）　千葉営業所"] == [2, 4, 6]  # 月水金
        assert cust_days["共同ガス（株）　本社"] == []  # 空
        assert cust_days["コイケ酸商（株）　白井営業所"] == [3, 5]  # 火木

        # 保持日数
        assert cust_retention["岡安産業（株）　千葉営業所"] == 2
        assert cust_retention["共同ガス（株）　本社"] == 0
        assert cust_retention["コイケ酸商（株）　白井営業所"] == 0

        # 路線便
        assert cust_route["岡安産業（株）　千葉営業所"] is False
        assert cust_route["共同ガス（株）　本社"] is True
        assert cust_route["コイケ酸商（株）　白井営業所"] is False

        # パターン（旧フォーマットなので空）
        assert cust_pattern == {}

    def test_missing_sheet(self):
        """シートが存在しない場合"""
        wb = MockWorkbook({})
        cust_days, cust_retention, cust_route, cust_pattern = build_customer_cache(wb)
        assert cust_days == {}
        assert cust_pattern == {}

    def test_empty_customer_skipped(self):
        """顧客名が空の行はスキップ"""
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便"],
            ["", "", 0, ""],
            ["テスト顧客", "月", 1, ""],
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"顧客マスター": ws})

        cust_days, _, _, _ = build_customer_cache(wb)
        assert len(cust_days) == 1
        assert "テスト顧客" in cust_days

    def test_duplicate_customer(self):
        """重複顧客名は最初の値を使用"""
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便"],
            ["テスト顧客", "月水金", 2, ""],
            ["テスト顧客", "火木", 5, "○"],  # 重複、スキップ
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"顧客マスター": ws})

        cust_days, cust_retention, cust_route, _ = build_customer_cache(wb)
        assert cust_days["テスト顧客"] == [2, 4, 6]  # 最初の値
        assert cust_retention["テスト顧客"] == 2
        assert cust_route["テスト顧客"] is False

    def test_new_format_with_pattern(self):
        """新フォーマット（E列=配送パターン）の読み込み"""
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便", "配送パターン"],
            ["顧客A", "月水金", 0, "", "近隣2便"],
            ["顧客B", "", 0, "○", "遠方午前"],
            ["顧客C", "", 0, "", ""],  # パターンなし
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"顧客マスター": ws})

        _, _, _, cust_pattern = build_customer_cache(wb)
        assert cust_pattern["顧客A"] == "近隣2便"
        assert cust_pattern["顧客B"] == "遠方午前"
        assert "顧客C" not in cust_pattern


# ============================================
# build_confirming_cache
# ============================================
class TestBuildConfirmingCache:
    def test_normal(self):
        """確認中テーブルの読み込み"""
        data = [
            # ヘッダー（11列）
            ["送付日時", "受注日", "顧客名", "受発注伝票", "明細",
             "メーカー名", "品名", "問合せ状況", "ステータス", "受注納期", "送付者"],
            # データ行
            ["2026-02-15", "2026-01-06", "コイケ酸商", "GL2Z444462", "10",
             "タケダ機械", "修理品", "未", "未処理", None, "boxeo"],
            ["2026-02-15", "2026-01-28", "コイケ酸商", "GL2F445985", "10",
             "育良精機", "ライトアーク", "済", "回答待ち", datetime.date(2026, 3, 15), "boxeo"],
        ]
        ws = MockWorksheet(data)

        cache = build_confirming_cache(ws)

        assert "GL2Z444462|10" in cache
        status1 = cache["GL2Z444462|10"]
        assert status1[0] == "未"         # 問合せ状況
        assert status1[1] == "未処理"     # ステータス
        assert status1[2] is None         # 受注納期

        assert "GL2F445985|10" in cache
        status2 = cache["GL2F445985|10"]
        assert status2[0] == "済"
        assert status2[1] == "回答待ち"
        assert status2[2] == datetime.date(2026, 3, 15)

    def test_none_worksheet(self):
        """Noneシート"""
        cache = build_confirming_cache(None)
        assert cache == {}

    def test_empty_table(self):
        """ヘッダーのみ"""
        data = [
            ["送付日時", "受注日", "顧客名", "受発注伝票", "明細",
             "メーカー名", "品名", "問合せ状況", "ステータス", "受注納期", "送付者"],
        ]
        ws = MockWorksheet(data)
        cache = build_confirming_cache(ws)
        assert cache == {}

    def test_duplicate_key(self):
        """重複キーは最初の値を使用"""
        data = [
            ["送付日時", "受注日", "顧客名", "受発注伝票", "明細",
             "メーカー名", "品名", "問合せ状況", "ステータス", "受注納期", "送付者"],
            ["2026-02-15", "", "顧客A", "GL2Z444462", "10",
             "メーカー", "品名1", "未", "未処理", None, "user1"],
            ["2026-02-16", "", "顧客A", "GL2Z444462", "10",
             "メーカー", "品名1", "済", "処理完了", None, "user2"],  # 重複
        ]
        ws = MockWorksheet(data)
        cache = build_confirming_cache(ws)
        assert cache["GL2Z444462|10"][0] == "未"  # 最初の値

    def test_empty_order_detail_skipped(self):
        """注番・明細が空の行はスキップ"""
        data = [
            ["送付日時", "受注日", "顧客名", "受発注伝票", "明細",
             "メーカー名", "品名", "問合せ状況", "ステータス", "受注納期", "送付者"],
            ["2026-02-15", "", "", "", "",
             "", "", "未", "", None, ""],
        ]
        ws = MockWorksheet(data)
        cache = build_confirming_cache(ws)
        assert cache == {}


# ============================================
# build_storage_cache
# ============================================
class TestBuildStorageCache:
    def test_normal(self):
        """保管場所キャッシュの構築"""
        # source_data: 行0-5はダミー、行6からデータ
        source = [[""] * 3 for _ in range(6)]
        source.append(["GL2Z444369", "", "転送中（直送用）"])   # 行6
        source.append(["", "", ""])                             # 行7: 副行
        source.append(["GL2C444510", "", "関東商品センター"])    # 行8

        cols = {"受発注伝票": 0, "保管場所": 2}
        cache = build_storage_cache(source, cols)

        assert cache["GL2Z444369"] == "転送中（直送用）"
        assert cache["GL2C444510"] == "関東商品センター"
        assert len(cache) == 2

    def test_empty_storage_skipped(self):
        """保管場所が空の行はスキップ"""
        source = [[""] * 3 for _ in range(6)]
        source.append(["GL2Z444369", "", ""])  # 保管場所空

        cols = {"受発注伝票": 0, "保管場所": 2}
        cache = build_storage_cache(source, cols)
        assert cache == {}

    def test_duplicate_order_num(self):
        """重複注番は最初の値を使用"""
        source = [[""] * 3 for _ in range(6)]
        source.append(["GL2Z444369", "", "転送中（直送用）"])
        source.append(["GL2Z444369", "", "関東商品センター"])  # 重複

        cols = {"受発注伝票": 0, "保管場所": 2}
        cache = build_storage_cache(source, cols)
        assert cache["GL2Z444369"] == "転送中（直送用）"

    def test_missing_cols(self):
        """必要な列がない場合"""
        source = [[""] * 3 for _ in range(7)]
        cache = build_storage_cache(source, {})
        assert cache == {}

    def test_short_row(self):
        """行の列数が足りない場合"""
        source = [[""] * 3 for _ in range(6)]
        source.append(["GL2Z444369"])  # 1列しかない

        cols = {"受発注伝票": 0, "保管場所": 2}
        cache = build_storage_cache(source, cols)
        assert cache == {}


# ============================================
# build_all_caches
# ============================================
class TestBuildAllCaches:
    def test_all_none(self):
        """全てNoneの場合→空のCacheStore"""
        store = build_all_caches(None, None, None, [], {})
        assert isinstance(store, CacheStore)
        assert store.mfg_name == {}
        assert store.cust_days == {}
        assert store.confirm == {}
        assert store.storage == {}

    def test_with_manufacturer(self):
        """メーカーキャッシュのみ"""
        mfg_data = [
            ["品目Group", "メーカー名", "加算日数"],
            ["D01", "ダイヘン", 2],
        ]
        ws = MockWorksheet(mfg_data)
        mfg_wb = MockWorkbook({"メーカー一覧": ws})

        store = build_all_caches(mfg_wb, None, None, [], {})
        assert store.mfg_name["D01"] == "ダイヘン"
        assert store.mfg_days["D01"] == 2

    def test_with_all_sources(self):
        """全ソースを指定"""
        # メーカー
        mfg_data = [
            ["品目Group", "メーカー名", "加算日数"],
            ["D01", "ダイヘン", 2],
        ]
        mfg_ws = MockWorksheet(mfg_data)
        mfg_wb = MockWorkbook({"メーカー一覧": mfg_ws})

        # 顧客
        cust_data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便"],
            ["テスト顧客", "月水金", 2, ""],
        ]
        cust_ws = MockWorksheet(cust_data)
        cust_wb = MockWorkbook({"顧客マスター": cust_ws})

        # 確認中
        conf_data = [
            ["送付日時", "受注日", "顧客名", "受発注伝票", "明細",
             "メーカー名", "品名", "問合せ状況", "ステータス", "受注納期", "送付者"],
            ["2026-02-15", "", "顧客A", "GL2Z444462", "10",
             "メーカー", "品名1", "未", "未処理", None, "user1"],
        ]
        conf_ws = MockWorksheet(conf_data)

        # 受注一覧
        source = [[""] * 3 for _ in range(6)]
        source.append(["GL2Z444369", "", "転送中（直送用）"])
        cols = {"受発注伝票": 0, "保管場所": 2}

        store = build_all_caches(mfg_wb, cust_wb, conf_ws, source, cols)

        assert store.mfg_name["D01"] == "ダイヘン"
        assert store.cust_days["テスト顧客"] == [2, 4, 6]
        assert "GL2Z444462|10" in store.confirm
        assert store.storage["GL2Z444369"] == "転送中（直送用）"
        assert store.cust_email_start_col == 4  # 旧フォーマット


# ============================================
# build_pattern_cache
# ============================================
class TestBuildPatternCache:
    def test_normal(self):
        """配送パターンの読み込み"""
        data = [
            ["パターン名", "cutoff1", "cutoff1前", "cutoff2", "cutoff2前"],
            ["近隣2便", "11:30", "当日", "16:00", "翌日"],
            ["遠方午前", "16:00", "翌日", None, None],
            ["遠方午後", "11:30", "当日", None, None],
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"配送パターン": ws})

        patterns = build_pattern_cache(wb)

        assert len(patterns) == 3

        kinrin = patterns["近隣2便"]
        assert kinrin.cutoff1 == (11, 30)
        assert kinrin.days_before_cutoff1 == 0
        assert kinrin.cutoff2 == (16, 0)
        assert kinrin.days_between_cutoffs == 1
        assert kinrin.days_after_all == 1

        enpo_am = patterns["遠方午前"]
        assert enpo_am.cutoff1 == (16, 0)
        assert enpo_am.days_before_cutoff1 == 1
        assert enpo_am.cutoff2 is None
        assert enpo_am.days_after_all == 2

        enpo_pm = patterns["遠方午後"]
        assert enpo_pm.cutoff1 == (11, 30)
        assert enpo_pm.days_before_cutoff1 == 0
        assert enpo_pm.cutoff2 is None
        assert enpo_pm.days_after_all == 1

    def test_missing_sheet(self):
        """シートが存在しない場合→空dict"""
        wb = MockWorkbook({})
        patterns = build_pattern_cache(wb)
        assert patterns == {}

    def test_invalid_cutoff(self):
        """cutoff1が不正な場合→スキップ"""
        data = [
            ["パターン名", "cutoff1", "cutoff1前", "cutoff2", "cutoff2前"],
            ["不正パターン", "abc", "当日", None, None],
            ["正常パターン", "11:30", "当日", None, None],
        ]
        ws = MockWorksheet(data)
        wb = MockWorkbook({"配送パターン": ws})

        patterns = build_pattern_cache(wb)
        assert len(patterns) == 1
        assert "正常パターン" in patterns


# ============================================
# _detect_customer_master_format
# ============================================
class TestDetectCustomerMasterFormat:
    def test_old_format_email_header(self):
        """旧フォーマット: E列ヘッダがメール関連"""
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便", "メール"],
            ["顧客A", "月水金", 0, "", "user@example.com"],
        ]
        ws = MockWorksheet(data)
        assert _detect_customer_master_format(ws) is False

    def test_new_format_pattern_header(self):
        """新フォーマット: E列ヘッダが「配送パターン」"""
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便", "配送パターン"],
            ["顧客A", "月水金", 0, "", "近隣2便"],
        ]
        ws = MockWorksheet(data)
        assert _detect_customer_master_format(ws) is True

    def test_empty_e_header(self):
        """E列ヘッダが空→旧フォーマット"""
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便", ""],
            ["顧客A", "月水金", 0, "", "近隣2便"],
        ]
        ws = MockWorksheet(data)
        assert _detect_customer_master_format(ws) is False

    def test_short_header_row(self):
        """ヘッダにE列がない→旧フォーマット"""
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便"],
            ["顧客A", "月水金", 0, ""],
        ]
        ws = MockWorksheet(data)
        assert _detect_customer_master_format(ws) is False

    def test_mailto_in_data_still_new_format(self):
        """データ行にmailto:があってもヘッダが正しければ新フォーマット"""
        data = [
            ["顧客名", "出荷曜日", "保持日数", "路線便", "配送パターン",
             "mailto:user@example.com"],
            ["顧客A", "月水金", 0, "", "近隣2便", "user@example.com"],
        ]
        ws = MockWorksheet(data)
        assert _detect_customer_master_format(ws) is True

    def test_empty_worksheet(self):
        """空のワークシート→旧フォーマット"""
        data = []
        ws = MockWorksheet(data)
        assert _detect_customer_master_format(ws) is False
