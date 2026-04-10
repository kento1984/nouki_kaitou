"""設定読込モジュール

VBAの以下の関数を移植:
- Loadholidays (L27): 特別日カレンダー読み込み
- LoadBranchSettings (L106): 営業所設定シートからグローバル設定読み込み
- GetBranchSettings (L72): 営業所設定取得（特別締切時間の動的上書き）
"""

from __future__ import annotations

import datetime
import re
from typing import TYPE_CHECKING

from nouki_kaitou.models import BranchSettings, HolidayMap
from nouki_kaitou.utils import parse_date

if TYPE_CHECKING:
    import openpyxl.worksheet.worksheet as ws_type


# ============================================
# VBA: Loadholidays (L27-67)
# メーカー一覧.xlsx「特別日カレンダー」シートから祝日・特別締切時間を読み込み
# ============================================
def load_holidays(manufacturer_master_wb: object) -> HolidayMap:
    """特別日カレンダーから祝日・特別締切時間を読み込む。

    Args:
        manufacturer_master_wb: メーカー一覧.xlsxのWorkbook

    Returns:
        HolidayMap: {日付: None(=祝日) or int(=特別締切時間)}

    VBA版のロジック:
    - A列: 日付
    - C列: 空欄→祝日、数値→特別締切時間（営業日扱い）
    """
    holidays: HolidayMap = {}

    try:
        calendar_ws = manufacturer_master_wb["特別日カレンダー"]
    except KeyError:
        return holidays

    for row in calendar_ws.iter_rows(min_row=2, max_col=3, values_only=True):
        date_val = row[0]
        cutoff_val = row[2] if len(row) > 2 else None

        special_date = parse_date(date_val)
        if special_date is None:
            continue

        if cutoff_val is not None and cutoff_val != "":
            try:
                cutoff_int = int(cutoff_val)
                holidays[special_date] = cutoff_int
            except (ValueError, TypeError):
                holidays[special_date] = None
        else:
            holidays[special_date] = None

    return holidays


# ============================================
# VBA: LoadBranchSettings (L106-167)
# 受注一覧の注番先頭2文字 → 営業所設定シートから設定読み込み
# ============================================
def load_branch_settings(
    manufacturer_master_wb: object,
    source_data: list[list[str]],
    cols: dict[str, int],
    header_row_idx: int = 4,
) -> BranchSettings:
    """営業所設定を読み込む。

    注番（受発注伝票）の先頭2文字が英字2文字のものを探し、
    メーカー一覧.xlsx「営業所設定」シートから対応する設定を取得する。

    Args:
        manufacturer_master_wb: メーカー一覧.xlsxのWorkbook
        source_data: 受注一覧データ（0-indexed行リスト、各行はタブ区切り列リスト）
        cols: 列位置マッピング
        header_row_idx: ヘッダー行インデックス（デフォルト4=旧フォーマット）

    Returns:
        BranchSettings（見つからない場合はデフォルト値）
    """
    settings = BranchSettings()

    # 受注一覧から先頭2文字が英字の注番を探す
    order_col = cols.get("受発注伝票")
    if order_col is None:
        return settings

    branch_code = ""
    # ヘッダー行の次からデータ行をスキャン
    for row in source_data[header_row_idx + 1:]:
        if order_col < len(row):
            order_num = str(row[order_col]).strip()
            if order_num and len(order_num) >= 2:
                first_two = order_num[:2]
                if re.match(r"^[A-Z][A-Z0-9]$", first_two):
                    branch_code = first_two
                    break

    if not branch_code:
        return settings

    # 営業所設定シートから検索
    try:
        branch_ws = manufacturer_master_wb["営業所設定"]
    except KeyError:
        return settings

    for row in branch_ws.iter_rows(min_row=2, max_col=7, values_only=True):
        code = str(row[0]).strip() if row[0] else ""
        if code == branch_code:
            settings.name = str(row[1]).strip() if row[1] else ""
            try:
                settings.default_cutoff = int(row[2]) if row[2] else 15
            except (ValueError, TypeError):
                settings.default_cutoff = 15
            settings.base_center = str(row[3]).strip() if row[3] else ""
            settings.shared_email = str(row[4]).strip() if row[4] else ""
            settings.signature = f"マツモト産業\n{settings.name}"
            if row[5]:
                d = parse_date(row[5])
                if d:
                    settings.start_date = d.strftime("%Y/%m/%d")
            # G列: 注番列の表示（「連絡事項」→L列を社外コメントに切替）
            if len(row) > 6 and row[6]:
                mode = str(row[6]).strip()
                if mode == "連絡事項":
                    settings.remarks_mode = "external"
            break

    return settings


# ============================================
# VBA: GetBranchSettings (L72-102)
# 営業所名・締切時間・署名をタプルで返す（特別締切時間対応）
# ============================================
def get_branch_settings(
    branch: BranchSettings,
    holidays: HolidayMap | None = None,
    target_date: datetime.date | None = None,
) -> tuple[str, int, str]:
    """営業所名・締切時間・署名を取得する。

    特別日カレンダーに締切時間が設定されている場合はそちらを優先。

    Args:
        branch: 営業所設定
        holidays: 祝日辞書
        target_date: 対象日（Noneなら今日）

    Returns:
        (営業所名, 締切時間, 署名) のタプル
    """
    if target_date is None:
        target_date = datetime.date.today()

    cutoff_hour = branch.default_cutoff

    if holidays and target_date in holidays:
        special_value = holidays[target_date]
        if special_value is not None:
            cutoff_hour = special_value

    return branch.name, cutoff_hour, branch.signature
