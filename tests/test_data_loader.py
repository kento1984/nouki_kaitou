"""data_loader.py のユニットテスト"""

import datetime
import tempfile
from pathlib import Path

import pytest

from nouki_kaitou.data_loader import (
    get_column_positions,
    get_data_rows_range,
    group_order_numbers_by_customer,
    is_data_row,
    load_source_file,
    parse_order_row,
)


# ============================================
# テスト用ヘッダー行（10PM.XLSの5行目相当）
# ============================================
# 実際のSAPヘッダーを簡略化したもの
SAMPLE_HEADER = [
    "",              # 0
    "",              # 1
    "",              # 2
    "受発注伝票",     # 3
    "",              # 4
    "明細",          # 5
    "",              # 6
    "",              # 7
    "伝票タイプ",     # 8
    "受注先",         # 9
    "",              # 10
    "テキスト",       # 11 → 品名
    "受注数量",       # 12
    "受注単価",       # 13
    "正味額",         # 14
    "名称",          # 15 → メーカー
    "保管場所",       # 16
    "出荷先名",       # 17
    "出荷ステータス",  # 18
    "受注納期",       # 19
    "品目 Group",    # 20
    "得意先担当者",    # 21
    "得意先発注番号",  # 22
    "コメント（明細）", # 23
    "コメント（社内）", # 24
    "コメント（社外）", # 25
    "伝票タイプ",     # 26 （重複、無視される）
    "時刻",          # 27
    "登録日",         # 28
    "拒否理由",       # 29
    "指定納期",       # 30
]

# ダミーデータ行
SAMPLE_DATA_ROW = [
    "",              # 0
    "",              # 1
    "",              # 2
    "GL2Z444369",    # 3: 受発注伝票
    "",              # 4
    "10",            # 5: 明細
    "",              # 6
    "",              # 7
    "【受注】直送販売",  # 8: 伝票タイプ
    "共同ガス（株）　本社",  # 9: 受注先
    "",              # 10
    "Ｕ１２５４６－４　チップ",  # 11: 品名
    "100",           # 12: 受注数量
    "500",           # 13: 受注単価
    "50000",         # 14: 正味額
    "ダイヘン",       # 15: メーカー
    "転送中（直送用）",  # 16: 保管場所
    "共同ガス（株）　本社",  # 17: 出荷先名
    "未処理",         # 18: 出荷ステータス
    "2026/1/15",     # 19: 受注納期
    "D01",           # 20: 品目Group
    "田中",          # 21: 得意先担当者
    "PO-12345",      # 22: 得意先発注番号
    "",              # 23: コメント（明細）
    "",              # 24: コメント（社内）
    "",              # 25: コメント（社外）
    "",              # 26
    "8:54:58",       # 27: 時刻
    "2026/1/6",      # 28: 登録日
    "",              # 29: 拒否理由
    "2026/1/10",     # 30: 指定納期
]

# 副行（マツモト担当者行）
SAMPLE_SUB_ROW = [
    "",              # 0
    "",              # 1
    "",              # 2
    "",              # 3: 受発注伝票（空＝副行）
    "",              # 4
    "",              # 5
    "",              # 6
    "",              # 7
    "",              # 8
    "",              # 9
    "",              # 10
    "柏原",          # 11
]


def _build_source_data() -> list[list[str]]:
    """テスト用の最小限のsource_dataを構築する。"""
    # 行0-3: ヘッダー前の行
    rows = [[""] * 5 for _ in range(4)]
    # 行4: ヘッダー行（5行目）
    rows.append(SAMPLE_HEADER)
    # 行5: 空行
    rows.append([""] * 5)
    # 行6: データ行（7行目）
    rows.append(SAMPLE_DATA_ROW)
    # 行7: 副行
    rows.append(SAMPLE_SUB_ROW)
    return rows


# ============================================
# load_source_file
# ============================================
class TestLoadSourceFile:
    def test_utf16_file(self, tmp_path: Path):
        """UTF-16LE BOM付きタブ区切りファイルの読み込み"""
        file = tmp_path / "test.xls"
        content = "col1\tcol2\tcol3\nval1\tval2\tval3\n"
        file.write_text(content, encoding="utf-16")

        result = load_source_file(file)
        assert len(result) == 2
        assert result[0] == ["col1", "col2", "col3"]
        assert result[1] == ["val1", "val2", "val3"]

    def test_cp932_file(self, tmp_path: Path):
        """CP932ファイルの読み込み"""
        file = tmp_path / "test.xls"
        content = "顧客名\t品名\n田中\tチップ\n"
        file.write_text(content, encoding="cp932")

        result = load_source_file(file)
        assert len(result) == 2
        assert result[0][0] == "顧客名"

    def test_empty_cells(self, tmp_path: Path):
        """空セルの処理"""
        file = tmp_path / "test.xls"
        content = "a\t\tb\n\t\t\n"
        file.write_text(content, encoding="utf-16")

        result = load_source_file(file)
        assert result[0] == ["a", "", "b"]
        assert result[1] == ["", "", ""]


# ============================================
# get_column_positions
# ============================================
class TestGetColumnPositions:
    def test_normal(self):
        """正常なヘッダーから列位置を取得"""
        source = _build_source_data()
        cols = get_column_positions(source)

        assert cols is not None
        assert cols["受発注伝票"] == 3
        assert cols["明細"] == 5
        assert cols["受注先"] == 9
        assert cols["品名"] == 11    # ヘッダーは「テキスト」
        assert cols["メーカー"] == 15  # ヘッダーは「名称」
        assert cols["品目Group"] == 20

    def test_required_columns_present(self):
        """必須列が全て存在"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        required = ["受発注伝票", "明細", "受注先", "品名", "受注数量",
                     "出荷先名", "受注納期", "品目Group", "登録日"]
        for name in required:
            assert name in cols, f"必須列 '{name}' が見つからない"

    def test_too_few_rows(self):
        """行数不足でNone"""
        source = [["a", "b"]] * 3  # 3行しかない
        assert get_column_positions(source) is None

    def test_missing_required_column(self):
        """必須列が不足でNone"""
        # 受発注伝票が含まれないヘッダー
        header = ["品名", "受注数量"]
        source = [[""] * 2 for _ in range(4)] + [header]
        assert get_column_positions(source) is None

    def test_first_only_flag(self):
        """最初の一致のみフラグの動作確認"""
        # 「明細」が5列目と46列目にある場合、最初の一致のみ取得
        header = list(SAMPLE_HEADER) + [""] * 20 + ["明細"]
        source = [[""] * 5 for _ in range(4)] + [header]
        # 最低限の必須列を含める
        cols = get_column_positions(source)
        if cols is not None:
            assert cols["明細"] == 5  # 最初のもの


# ============================================
# parse_order_row
# ============================================
class TestParseOrderRow:
    def test_normal(self):
        """正常なデータ行のパース"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        row = parse_order_row(source, 6, cols)
        assert row.order_number == "GL2Z444369"
        assert row.detail_number == "10"
        assert row.document_type == "【受注】直送販売"
        assert row.customer_name == "共同ガス（株）　本社"
        assert row.product_name == "Ｕ１２５４６－４　チップ"
        assert row.quantity == "100"
        assert row.ship_status == "未処理"
        assert row.manufacturer_name == "ダイヘン"
        assert row.storage_place == "転送中（直送用）"
        assert row.item_group_code == "D01"
        assert row.source_row == 6

    def test_date_parsing(self):
        """日付列のパース"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        row = parse_order_row(source, 6, cols)
        assert row.order_delivery_date == datetime.date(2026, 1, 15)
        assert row.registration_date == datetime.date(2026, 1, 6)
        assert row.specified_delivery_date == datetime.date(2026, 1, 10)

    def test_out_of_range(self):
        """範囲外の行番号"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        row = parse_order_row(source, 999, cols)
        assert row.order_number == ""
        assert row.source_row == 999

    def test_time_value(self):
        """時刻フィールド"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        row = parse_order_row(source, 6, cols)
        assert row.time_value == "8:54:58"


# ============================================
# get_data_rows_range / is_data_row
# ============================================
class TestDataRowDetection:
    def test_get_data_rows_range(self):
        """データ行範囲の取得"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        rng = get_data_rows_range(source, cols)
        assert 6 in rng  # データ行（受発注伝票列に値あり）
        # 副行（行7）は受発注伝票列が空なのでlast_rowにならない
        # rangeは6〜6（最後のデータ行まで）

    def test_is_data_row_true(self):
        """データ行の判定"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        assert is_data_row(source, 6, cols) is True   # データ行

    def test_is_data_row_false_for_sub_row(self):
        """副行はデータ行でない"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        assert is_data_row(source, 7, cols) is False   # 副行

    def test_is_data_row_out_of_range(self):
        """範囲外の行"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        assert is_data_row(source, 999, cols) is False

    def test_empty_source(self):
        """空データの場合"""
        cols = {"受発注伝票": 0}
        assert is_data_row([], 0, cols) is False


# ============================================
# group_order_numbers_by_customer
# ============================================
class TestGroupOrderNumbersByCustomer:
    def test_normal(self):
        """顧客ごとのグループ化"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        result = group_order_numbers_by_customer(
            source, cols, ["GL2Z444369"]
        )
        assert "共同ガス（株）　本社" in result
        assert "GL2Z444369" in result["共同ガス（株）　本社"]

    def test_not_found(self):
        """注番が見つからない場合"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        result = group_order_numbers_by_customer(
            source, cols, ["NOTEXIST"]
        )
        assert result == {}

    def test_no_duplicates(self):
        """同じ注番の重複なし"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        result = group_order_numbers_by_customer(
            source, cols, ["GL2Z444369", "GL2Z444369"]
        )
        customer_orders = result.get("共同ガス（株）　本社", [])
        assert customer_orders.count("GL2Z444369") == 1

    def test_empty_order_numbers(self):
        """空の注番リスト"""
        source = _build_source_data()
        cols = get_column_positions(source)
        assert cols is not None

        result = group_order_numbers_by_customer(source, cols, [])
        assert result == {}
