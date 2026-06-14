"""SAP受注一覧の読込・列位置検索モジュール

VBAの以下の関数を移植:
- GetColumnPositions (L2246): 列位置の動的検索
- GroupOrderNumbersByCustomer (L932): 注番を顧客ごとにグループ化
- 10PM.XLS の読み込み（UTF-16LE BOM付きタブ区切りテキスト）
"""

from __future__ import annotations

import re
from pathlib import Path

from nouki_kaitou.models import ColumnMap, OrderRow
from nouki_kaitou.utils import parse_date


# ============================================
# 10PM.XLS の読み込み
# ============================================
def load_source_file(file_path: str | Path) -> list[list[str]]:
    """SAP受注伝票一覧ファイルを読み込む。

    10PM.XLSは拡張子が.XLSだが実態はUTF-16LE BOM付きタブ区切りテキスト。
    VBAではExcelで開いてシートとして読み込んでいたが、
    Pythonではテキストとして直接パースする。

    Args:
        file_path: ファイルパス

    Returns:
        2次元リスト（行×列）。VBAのg_SourceData相当。
        インデックスは0-based。
    """
    path = Path(file_path)

    # まずUTF-16で試す（10PM.XLSの標準形式）
    try:
        with open(path, "r", encoding="utf-16") as f:
            lines = f.readlines()
    except UnicodeError:
        # UTF-16で読めない場合はcp932で試す
        with open(path, "r", encoding="cp932") as f:
            lines = f.readlines()

    rows: list[list[str]] = []
    for line in lines:
        # タブで分割（末尾の改行を除去）
        cells = line.rstrip("\n\r").split("\t")
        rows.append(cells)

    return rows


# ============================================
# VBA: GetColumnPositions (L2246-2318)
# ヘッダー行（5行目）から列位置を動的検索
# ============================================

# VBAの検索条件: InStr(headerValue, キーワード) > 0
# 左辺がVBA内部名、右辺がヘッダーの検索キーワード
_COLUMN_MAPPINGS: list[tuple[str, str, bool]] = [
    # (内部名, 検索キーワード, 最初の一致のみフラグ)
    ("受発注伝票", "受発注伝票", False),
    ("明細", "明細", True),          # 最初の一致のみ（5列目と46列目の両方にある）
    ("受注先", "受注先", True),      # 最初の一致のみ（9列目と36列目の両方にある）
    ("品名", "テキスト", True),      # 内部名は「品名」、ヘッダーは「テキスト」
    ("受注数量", "受注数量", False),
    ("受注単価", "受注単価", False),
    ("正味額", "正味額", False),
    ("発注単価", "発注単価", False),     # 仕入単価（手配台帳の金額列用）
    ("メーカー", "名称", False),     # 内部名は「メーカー」、ヘッダーは「名称」
    ("保管場所", "保管場所", False),
    ("出荷先名", "出荷先名", False),
    ("出荷ステータス", "出荷ステータス", False),
    ("受注納期", "受注納期", False),
    ("品目Group", "品目 Group", False),  # スペースあり版
    ("品目Group", "品目Group", False),   # スペースなし版
    ("得意先担当者", "得意先担当者", False),
    ("得意先発注番号", "得意先発注番号", False),
    ("コメント（明細）", "コメント（明細）", False),
    ("コメント（社内）", "コメント（社内）", False),
    ("コメント（社外）", "コメント（社外）", False),
    ("伝票タイプ", "伝票タイプ", False),
    ("マツモト担当者名", "マツモト担当者名", False),  # 社内手配担当（TWF台帳用）
    ("時刻", "時刻", False),
    ("登録日", "登録日", False),
    ("拒否理由", "拒否理由", False),
    ("指定納期", "指定納期", False),
]

# 必須列
_REQUIRED_COLUMNS = frozenset([
    "受発注伝票", "明細", "受注先", "品名", "受注数量",
    "出荷先名", "受注納期", "品目Group", "登録日",
])


def get_column_positions(
    source_data: list[list[str]],
) -> tuple[ColumnMap, int] | None:
    """ヘッダー行から列位置を動的検索する。

    先頭10行を走査して「受発注伝票」を含む行をヘッダー行とみなす。
    旧フォーマット（5行目ヘッダー）でも新フォーマット（2行目ヘッダー）でも対応可能。

    Args:
        source_data: load_source_fileの返り値

    Returns:
        (ColumnMap, ヘッダー行インデックス) のタプル。
        ヘッダーが見つからない、または必須列が揃わなければNone。
    """
    header_row_idx: int | None = None
    for i in range(min(10, len(source_data))):
        for cell in source_data[i]:
            if "受発注伝票" in cell.strip():
                header_row_idx = i
                break
        if header_row_idx is not None:
            break

    if header_row_idx is None:
        return None

    header_row = source_data[header_row_idx]
    cols: ColumnMap = {}

    for col_idx, header_val in enumerate(header_row):
        header_val = header_val.strip()
        if not header_val:
            continue

        for internal_name, keyword, first_only in _COLUMN_MAPPINGS:
            if keyword in header_val:
                if first_only and internal_name in cols:
                    continue  # 既に見つかっていればスキップ
                if internal_name not in cols:
                    cols[internal_name] = col_idx

    # 必須列チェック
    if not _REQUIRED_COLUMNS.issubset(cols.keys()):
        return None

    return (cols, header_row_idx)


# ============================================
# 受注データ行の構造化読み取り
# ============================================
def parse_order_row(
    source_data: list[list[str]],
    row_idx: int,
    cols: ColumnMap,
) -> OrderRow:
    """受注一覧の1行をOrderRowに変換する。

    Args:
        source_data: 生データ
        row_idx: 行番号（0-indexed）
        cols: 列位置マッピング

    Returns:
        OrderRow
    """
    row = source_data[row_idx] if row_idx < len(source_data) else []

    def get_val(col_name: str) -> str:
        idx = cols.get(col_name)
        if idx is None or idx >= len(row):
            return ""
        val = row[idx]
        if val is None:
            return ""
        # カンマ区切り数値のクォート除去（例: '"44,000"' → '44000'）
        return str(val).strip().strip('"')

    return OrderRow(
        order_number=get_val("受発注伝票"),
        detail_number=get_val("明細"),
        document_type=get_val("伝票タイプ"),
        customer_name=get_val("受注先"),
        product_name=get_val("品名"),
        ship_status=get_val("出荷ステータス"),
        quantity=get_val("受注数量"),
        unit_price=get_val("受注単価"),
        net_amount=get_val("正味額"),
        purchase_unit_price=get_val("発注単価"),
        manufacturer_name=get_val("メーカー"),
        storage_place=get_val("保管場所"),
        customer_order_number=get_val("得意先発注番号"),
        customer_contact=get_val("得意先担当者"),
        comment_detail=get_val("コメント（明細）"),
        comment_external=get_val("コメント（社外）"),
        comment_internal=get_val("コメント（社内）"),
        rejection_reason=get_val("拒否理由"),
        ship_to_name=get_val("出荷先名"),
        registration_date=parse_date(get_val("登録日")),
        time_value=get_val("時刻"),
        order_delivery_date=parse_date(get_val("受注納期")),
        specified_delivery_date=parse_date(get_val("指定納期")),
        item_group_code=get_val("品目Group"),
        rep_name=get_val("マツモト担当者名"),
        source_row=row_idx,
    )


# ============================================
# データ行の判定
# ============================================
def get_data_rows_range(
    source_data: list[list[str]], cols: ColumnMap, header_row_idx: int
) -> range:
    """データ行の範囲を返す。

    ヘッダー行の直後からスキャンし、受発注伝票列に値がある
    最初の行をデータ開始行とする。
    旧フォーマット（ヘッダー後に空行1行）でも新フォーマット（ヘッダー直後がデータ）でも対応。

    Returns:
        有効なデータ行のrangeオブジェクト
    """
    order_col = cols.get("受発注伝票")
    if order_col is None:
        return range(0)

    # ヘッダー直後からデータ開始行を探す
    data_start = header_row_idx + 1
    for i in range(header_row_idx + 1, min(header_row_idx + 5, len(source_data))):
        row = source_data[i]
        if order_col < len(row) and str(row[order_col]).strip():
            data_start = i
            break

    # 最終行を求める
    last_row = data_start
    for i in range(len(source_data) - 1, data_start - 1, -1):
        row = source_data[i]
        if order_col < len(row) and str(row[order_col]).strip():
            last_row = i
            break

    return range(data_start, last_row + 1)


def is_data_row(
    source_data: list[list[str]], row_idx: int, cols: ColumnMap
) -> bool:
    """指定行がデータ行かどうか判定する。

    副行（マツモト担当者行）は受発注伝票列が空なのでFalse。
    """
    order_col = cols.get("受発注伝票")
    if order_col is None:
        return False

    if row_idx >= len(source_data):
        return False

    row = source_data[row_idx]
    if order_col >= len(row):
        return False

    return str(row[order_col]).strip() != ""


# ============================================
# VBA: GroupOrderNumbersByCustomer (L932-982)
# 注番を顧客ごとにグループ化
# ============================================
def group_order_numbers_by_customer(
    source_data: list[list[str]],
    cols: ColumnMap,
    order_numbers: list[str],
    header_row_idx: int = 4,
) -> dict[str, list[str]]:
    """注番リストを顧客名でグループ化する。

    Args:
        source_data: 生データ
        cols: 列位置マッピング
        order_numbers: 注番リスト
        header_row_idx: ヘッダー行インデックス（デフォルト4=旧フォーマット）

    Returns:
        {顧客名: [注番リスト]}（重複なし）
    """
    order_col = cols.get("受発注伝票")
    customer_col = cols.get("受注先")
    if order_col is None or customer_col is None:
        return {}

    customer_groups: dict[str, list[str]] = {}

    for order_num in order_numbers:
        order_num = order_num.strip()
        customer_name = ""

        # 注番から顧客名を検索
        for row in source_data[header_row_idx + 1:]:
            if order_col < len(row) and customer_col < len(row):
                if str(row[order_col]).strip() == order_num:
                    customer_name = str(row[customer_col]).strip()
                    break

        if not customer_name:
            continue

        if customer_name not in customer_groups:
            customer_groups[customer_name] = []

        # 重複チェック
        if order_num not in customer_groups[customer_name]:
            customer_groups[customer_name].append(order_num)

    return customer_groups
