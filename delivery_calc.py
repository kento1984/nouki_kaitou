"""納期計算コアモジュール

VBAの以下の関数を移植:
- CalculateDeliveryDate (L3034): 納期計算メイン
- ExtractPickupDate (L5129): 引取日抽出
- ExtractArrivalDateFromInternal (L7129): @@着日指定抽出
"""

from __future__ import annotations

import datetime
import re
from typing import Optional

from nouki_kaitou.business_days import (
    add_business_days,
    get_next_business_day,
    get_next_delivery_day,
    get_previous_business_day,
)
from nouki_kaitou.config import get_branch_settings
from nouki_kaitou.confirming import get_confirmed_delivery_date
from nouki_kaitou.customer import (
    get_customer_delivery_days,
    is_route_delivery,
)
from nouki_kaitou.manufacturer import get_delivery_days_to_add
from nouki_kaitou.models import (
    BranchSettings,
    CacheStore,
    HolidayMap,
    OrderRow,
)
from nouki_kaitou.stockout import get_storage_place_from_same_order
from nouki_kaitou.utils import (
    convert_to_half_width,
    format_date_japanese,
    is_december_31,
    normalize_name_for_comparison,
    parse_time,
)


# ============================================
# 内部ヘルパー
# ============================================
def _result(
    d: datetime.date, past: str, future: str, today: datetime.date
) -> str:
    """日付 + 過去/未来サフィックスでフォーマット"""
    return format_date_japanese(d) + (past if d <= today else future)


def _before_cutoff(hour: int, minute: int, cutoff: int) -> bool:
    """締切時間より前かどうか（VBA: timeHour < cutoffHour Or (= And minute=0)）"""
    return hour < cutoff or (hour == cutoff and minute == 0)


def _resolve_storage_place(row: OrderRow, cache: CacheStore) -> str:
    """保管場所を解決（空なら同一注番から取得）"""
    sp = row.storage_place.strip()
    if not sp:
        sp = get_storage_place_from_same_order(row.order_number, cache)
    return sp


def _is_same_customer(customer_name: str, ship_to_name: str) -> bool:
    """受注先と出荷先が同一かどうかを判定する。

    SAPでは同じ顧客でも全角/半角の揺れがある。
    例: 受注先「（有）三橋機工」vs 出荷先「(有)三橋機工」
    """
    return (
        normalize_name_for_comparison(customer_name)
        == normalize_name_for_comparison(ship_to_name)
    )


# ============================================
# VBA: ExtractPickupDate (L5129-5198)
# コメントから引取日を抽出
# ============================================
def extract_pickup_date(
    comment: str,
    today: datetime.date | None = None,
) -> Optional[datetime.date]:
    """コメントから引取日を抽出する。

    「引取」または「引き取り」を含むコメントからM/D形式の日付を抽出。
    180日以上過去の日付は翌年として扱う。

    Args:
        comment: コメント文字列（社外+社内を結合したもの）
        today: 今日の日付（テスト用）

    Returns:
        引取日（抽出できなければNone）
    """
    if not comment:
        return None

    if "引取" not in comment and "引き取り" not in comment:
        return None

    if today is None:
        today = datetime.date.today()

    # 数字とスラッシュ(半角/全角)を収集して日付をパース
    date_str = ""
    for ch in comment:
        if ch.isdigit() or ch in ("/", "／"):
            date_str += ch
        else:
            if date_str:
                result = _parse_slash_date(date_str, today)
                if result is not None:
                    return result
                date_str = ""

    # 末尾に残っている場合
    if date_str:
        result = _parse_slash_date(date_str, today)
        if result is not None:
            return result

    return None


def _parse_slash_date(
    date_str: str, today: datetime.date
) -> Optional[datetime.date]:
    """「M/D」形式の文字列をパースする（全角スラッシュ対応）"""
    date_str = date_str.replace("／", "/")
    slash_pos = date_str.find("/")
    if slash_pos <= 0:
        return None

    try:
        month_num = int(date_str[:slash_pos])
        day_num = int(date_str[slash_pos + 1:])
    except ValueError:
        return None

    if not (1 <= month_num <= 12 and 1 <= day_num <= 31):
        return None

    try:
        result = datetime.date(today.year, month_num, day_num)
    except ValueError:
        return None

    # 180日以上過去なら翌年
    if result < today and (today - result).days > 180:
        try:
            result = datetime.date(today.year + 1, month_num, day_num)
        except ValueError:
            return None

    return result


# ============================================
# VBA: ExtractArrivalDateFromInternal (L7129-7185)
# コメント（社内）から@@着日を抽出
# ============================================
def extract_arrival_date_from_internal(
    comment: str,
    today: datetime.date | None = None,
) -> Optional[datetime.date]:
    """コメント（社内）から@@着日指定を抽出する。

    「@@」（半角/全角）の後ろにあるM/D形式の日付を抽出。
    例: 「@@12/20」→ 12月20日

    Args:
        comment: コメント（社内）
        today: 今日の日付（テスト用）

    Returns:
        着日（抽出できなければNone）
    """
    if not comment:
        return None

    if today is None:
        today = datetime.date.today()

    # 「@@」を探す（半角・全角両対応）
    start_pos = comment.find("@@")
    if start_pos < 0:
        start_pos = comment.find("＠＠")
    if start_pos < 0:
        return None

    # 「@@」の後ろを取得
    after = comment[start_pos + 2:]
    after = convert_to_half_width(after)

    # 数字とスラッシュを抽出
    date_str = ""
    for ch in after:
        if ch.isdigit() or ch == "/":
            date_str += ch
        elif date_str:
            break

    if "/" not in date_str:
        return None

    return _parse_slash_date(date_str, today)


# ============================================
# VBA: CalculateDeliveryDate (L3034-3783)
# 納期計算メイン関数
# ============================================
def calculate_delivery_date(
    row: OrderRow,
    cache: CacheStore,
    holidays: HolidayMap | None = None,
    branch: BranchSettings | None = None,
    execution_time: datetime.datetime | None = None,
    today: datetime.date | None = None,
) -> str:
    """納期を計算して表示文字列を返す。

    VBAのCalculateDeliveryDate関数を移植。
    伝票タイプ・出荷ステータス・指定納期・受注納期等から
    納期を計算し「○月○日出荷予定」等の文字列を返す。

    処理優先順位:
    1. &&作業（コメント社内に「&&」or「＆＆」）
    2. @@着日指定（コメント社内に「@@M/D」）
    3. 引取（コメントに「引取」or「引き取り」＋日付）
    4. 指定納期あり（12/31以外）→ 在庫販売/直送販売別処理
    5. 在庫販売+処理完了 → 時刻ベース計算
    6. 紐付き+処理完了 → マクロ実行時刻ベース計算
    7. 受注納期なし → 「日程調整中」
    8. 受注納期=12/31 → 確認中一覧 or 「確認中」/「日程調整中」
    9. 通常 → +営業日+曜日制限

    Args:
        row: 受注データ行
        cache: マスターキャッシュ
        holidays: 祝日辞書
        branch: 営業所設定
        execution_time: マクロ実行時刻（紐付き処理完了で使用）
        today: 今日の日付（テスト用）

    Returns:
        納期表示文字列
    """
    if today is None:
        today = datetime.date.today()
    if branch is None:
        branch = BranchSettings()
    if execution_time is None:
        execution_time = datetime.datetime.now()

    # ============================================
    # 1. &&作業チェック
    # ============================================
    result = _check_work_order(row, today)
    if result is not None:
        return result

    # ============================================
    # 2. @@着日指定チェック
    # ============================================
    result = _check_arrival_date(row, today)
    if result is not None:
        return result

    # ============================================
    # 3. 引取チェック
    # ============================================
    result = _check_pickup(row, today)
    if result is not None:
        return result

    # ============================================
    # フラグ設定
    # ============================================
    # 保管場所解決
    storage_place = _resolve_storage_place(row, cache)

    # useShipRule: 在庫販売 + 受注先≠出荷先
    use_ship_rule = False
    if row.document_type == "【受注】在庫販売":
        if not _is_same_customer(row.customer_name, row.ship_to_name):
            use_ship_rule = True

    # 路線便フラグ
    is_rosenbin = is_route_delivery(row.customer_name, cache)

    # originalを保存してから路線便でuseShipRuleを上書き
    original_use_ship_rule = use_ship_rule
    if not use_ship_rule and is_rosenbin:
        use_ship_rule = True

    # ============================================
    # 4. 指定納期 early check（処理完了より前に判定）
    # ============================================
    result = _check_specified_date(
        row, cache, holidays, today,
        storage_place, use_ship_rule, is_rosenbin,
    )
    if result is not None:
        return result

    # ============================================
    # 5. 在庫販売 + 処理完了
    # ============================================
    result = _check_stock_completed(
        row, cache, holidays, branch, today,
        storage_place, use_ship_rule,
    )
    if result is not None:
        return result

    # ============================================
    # 6. 紐付き + 処理完了
    # ============================================
    result = _check_himozuki_completed(
        row, cache, holidays, branch, execution_time, today,
        storage_place, is_rosenbin,
    )
    if result is not None:
        return result

    # ============================================
    # 7. 受注納期なし → 日程調整中
    # ============================================
    delivery_date = row.order_delivery_date
    if delivery_date is None:
        return "日程調整中"

    # ============================================
    # 8. 受注納期 = 12/31 → 確認中一覧 or 未確定
    # ============================================
    if is_december_31(delivery_date):
        return _check_dec31(
            row, cache, holidays, today,
            storage_place, original_use_ship_rule, is_rosenbin,
        )

    # ============================================
    # 9. 通常の納期計算（受注納期ベース）
    # ============================================
    return _calc_normal(
        row, cache, holidays, today,
        delivery_date, storage_place,
        use_ship_rule, original_use_ship_rule, is_rosenbin,
    )


# ============================================
# 各パスの内部関数
# ============================================

def _check_work_order(
    row: OrderRow, today: datetime.date
) -> Optional[str]:
    """&&作業チェック"""
    internal = row.comment_internal.strip()
    if "&&" not in internal and "＆＆" not in internal:
        return None

    # 指定納期 → 受注納期 の順で日付を取得（12/31は無視）
    work_date = row.specified_delivery_date
    if work_date is not None and is_december_31(work_date):
        work_date = None

    if work_date is None:
        work_date = row.order_delivery_date
        if work_date is not None and is_december_31(work_date):
            work_date = None

    if work_date is None:
        return "日程調整中"

    if work_date <= today:
        return format_date_japanese(work_date) + "作業済み"
    return format_date_japanese(work_date) + "作業予定"


def _check_arrival_date(
    row: OrderRow, today: datetime.date
) -> Optional[str]:
    """@@着日指定チェック"""
    internal = row.comment_internal.strip()
    arrival_date = extract_arrival_date_from_internal(internal, today)
    if arrival_date is None:
        return None

    # 出荷日: 指定納期 → 受注納期（12/31は無視）
    ship_date = row.specified_delivery_date
    if ship_date is not None and is_december_31(ship_date):
        ship_date = None

    if ship_date is None:
        ship_date = row.order_delivery_date
        if ship_date is not None and is_december_31(ship_date):
            ship_date = None

    if ship_date is None:
        return None

    # M/D形式で表示
    ship_fmt = f"{ship_date.month}/{ship_date.day}"
    arrival_fmt = f"{arrival_date.month}/{arrival_date.day}"

    if ship_date <= today:
        return f"{ship_fmt}出荷済→{arrival_fmt}着"
    return f"{ship_fmt}出荷→{arrival_fmt}着予定"


def _check_pickup(
    row: OrderRow, today: datetime.date
) -> Optional[str]:
    """引取チェック"""
    comment = (
        (row.comment_external.strip() + " " + row.comment_internal.strip())
        .strip()
    )
    pickup_date = extract_pickup_date(comment, today)
    if pickup_date is None:
        return None

    # 引取は < today（他は <= today）
    if pickup_date < today:
        return format_date_japanese(pickup_date) + "引取済み"
    return format_date_japanese(pickup_date) + "引取予定"


def _check_specified_date(
    row: OrderRow,
    cache: CacheStore,
    holidays: HolidayMap | None,
    today: datetime.date,
    storage_place: str,
    use_ship_rule: bool,
    is_rosenbin: bool,
) -> Optional[str]:
    """指定納期 early check（処理完了より前）"""
    spec_date = row.specified_delivery_date
    if spec_date is None or is_december_31(spec_date):
        return None

    # ---- 在庫販売 ----
    if row.document_type == "【受注】在庫販売":
        if storage_place == "転送中（直送用）":
            return _result(spec_date, "出荷済み", "出荷予定", today)
        elif use_ship_rule:
            ship_date = get_previous_business_day(spec_date, holidays)
            return _result(ship_date, "出荷済み", "出荷予定", today)
        else:
            return _result(spec_date, "配達済み", "配達予定", today)

    # ---- 直送販売（紐付き含む）----
    days_to_add = get_delivery_days_to_add(row.item_group_code, cache)

    if storage_place == "転送中（直送用）":
        return _result(spec_date, "出荷済み", "出荷予定", today)

    # 出荷曜日制限チェック
    delivery_days = get_customer_delivery_days(row.customer_name, cache)
    adjusted = add_business_days(spec_date, days_to_add, holidays)

    if delivery_days:
        adjusted = get_next_delivery_day(adjusted, delivery_days, holidays)
        return _result(adjusted, "出荷済み", "出荷予定", today)

    # 曜日制限なし
    if is_rosenbin:
        rosenbin_date = add_business_days(
            spec_date, max(days_to_add - 1, 0), holidays
        )
        return _result(rosenbin_date, "出荷済み", "出荷予定", today)

    # 受注先≠出荷先 → 1営業日前を出荷日として「出荷予定」
    # （直送販売ではuse_ship_ruleが常にFalseなので、直接比較する）
    if not _is_same_customer(row.customer_name, row.ship_to_name):
        ship_date = get_previous_business_day(adjusted, holidays)
        return _result(ship_date, "出荷済み", "出荷予定", today)

    return _result(adjusted, "配達済み", "配達予定", today)


def _check_stock_completed(
    row: OrderRow,
    cache: CacheStore,
    holidays: HolidayMap | None,
    branch: BranchSettings,
    today: datetime.date,
    storage_place: str,
    use_ship_rule: bool,
) -> Optional[str]:
    """在庫販売 + 処理完了"""
    if row.document_type != "【受注】在庫販売":
        return None
    if row.ship_status != "処理完了":
        return None

    # 登録日・時刻が必要
    reg_date = row.registration_date
    if reg_date is None:
        return None

    time_parts = parse_time(row.time_value)
    if time_parts is None:
        return None

    hour, minute = time_parts
    _, cutoff, _ = get_branch_settings(branch, holidays, reg_date)

    # 他支店在庫チェック（受注先≠出荷先より優先）
    base_center = branch.base_center
    if base_center and storage_place and storage_place != base_center:
        if _before_cutoff(hour, minute, cutoff):
            ship_date = reg_date
        else:
            ship_date = add_business_days(reg_date, 1, holidays)
        return _result(ship_date, "他拠点より出荷済み", "他拠点より出荷予定", today)

    # 受注先≠出荷先 → 出荷ルール
    if use_ship_rule:
        if _before_cutoff(hour, minute, cutoff):
            ship_date = reg_date
        else:
            ship_date = add_business_days(reg_date, 1, holidays)
        return _result(ship_date, "出荷済み", "出荷予定", today)

    # 配達日計算のベース
    if _before_cutoff(hour, minute, cutoff):
        adjusted = add_business_days(reg_date, 1, holidays)
    else:
        adjusted = add_business_days(reg_date, 2, holidays)

    # 自拠点: 曜日制限チェック
    delivery_days = get_customer_delivery_days(row.customer_name, cache)
    if delivery_days:
        if _before_cutoff(hour, minute, cutoff):
            base_date = reg_date
        else:
            base_date = add_business_days(reg_date, 1, holidays)
        adjusted = get_next_delivery_day(base_date, delivery_days, holidays)
        return _result(adjusted, "出荷済み", "出荷予定", today)

    # 曜日制限なし → 配達予定
    return _result(adjusted, "配達済み", "配達予定", today)


def _check_himozuki_completed(
    row: OrderRow,
    cache: CacheStore,
    holidays: HolidayMap | None,
    branch: BranchSettings,
    execution_time: datetime.datetime,
    today: datetime.date,
    storage_place: str,
    is_rosenbin: bool,
) -> Optional[str]:
    """紐付き + 処理完了（直送販売 + 処理完了 + 非転送中）"""
    if row.document_type != "【受注】直送販売":
        return None
    if row.ship_status != "処理完了":
        return None
    if storage_place == "転送中（直送用）":
        return None

    # マクロ実行時刻ベースで計算
    _, cutoff, _ = get_branch_settings(branch, holidays, today)
    exec_hour = execution_time.hour

    if exec_hour < cutoff:
        adjusted = add_business_days(today, 1, holidays)
    else:
        adjusted = add_business_days(today, 2, holidays)

    # 受注先≠出荷先 → 配達日から1営業日逆算して出荷日
    if not _is_same_customer(row.customer_name, row.ship_to_name):
        ship_date = get_previous_business_day(adjusted, holidays)
        return _result(ship_date, "出荷済み", "出荷予定", today)

    # 曜日制限チェック
    delivery_days = get_customer_delivery_days(row.customer_name, cache)
    if delivery_days:
        adjusted = get_next_delivery_day(adjusted, delivery_days, holidays)
        return _result(adjusted, "出荷済み", "出荷予定", today)

    # 路線便 → 1営業日逆算して出荷日
    if is_rosenbin:
        ship_date = get_previous_business_day(adjusted, holidays)
        return _result(ship_date, "出荷済み", "出荷予定", today)

    # 自社便配達
    return _result(adjusted, "配達済み", "配達予定", today)


def _check_dec31(
    row: OrderRow,
    cache: CacheStore,
    holidays: HolidayMap | None,
    today: datetime.date,
    storage_place: str,
    original_use_ship_rule: bool,
    is_rosenbin: bool,
) -> str:
    """受注納期=12/31の処理（確認中一覧チェック or 未確定）"""
    # 確認中一覧から確定納期を取得
    confirmed_date = get_confirmed_delivery_date(
        row.order_number, row.detail_number, cache
    )

    if confirmed_date is not None:
        days_to_add = get_delivery_days_to_add(row.item_group_code, cache)

        # 紐付き+受注先≠出荷先チェック
        is_himozuki_diff = (
            row.document_type == "【受注】直送販売"
            and storage_place != "転送中（直送用）"
            and not _is_same_customer(row.customer_name, row.ship_to_name)
        )

        if storage_place == "転送中（直送用）":
            return _result(confirmed_date, "出荷済み", "出荷予定", today)

        if original_use_ship_rule:
            return _result(confirmed_date, "出荷済み", "出荷予定", today)

        if is_rosenbin:
            adjusted = add_business_days(
                confirmed_date, max(days_to_add - 1, 0), holidays
            )
            return _result(adjusted, "出荷済み", "出荷予定", today)

        # 通常: +営業日 + 曜日制限
        adjusted = add_business_days(confirmed_date, days_to_add, holidays)
        delivery_days = get_customer_delivery_days(row.customer_name, cache)

        if delivery_days:
            adjusted = get_next_delivery_day(adjusted, delivery_days, holidays)
            return _result(adjusted, "出荷済み", "出荷予定", today)

        # 曜日制限なし
        if is_himozuki_diff:
            return _result(adjusted, "出荷済み", "出荷予定", today)
        return _result(adjusted, "配達済み", "配達予定", today)

    # 確定納期なし
    if row.document_type == "【受注】在庫販売":
        return "日程調整中"
    return "確認中"


def _calc_normal(
    row: OrderRow,
    cache: CacheStore,
    holidays: HolidayMap | None,
    today: datetime.date,
    delivery_date: datetime.date,
    storage_place: str,
    use_ship_rule: bool,
    original_use_ship_rule: bool,
    is_rosenbin: bool,
) -> str:
    """通常の納期計算（受注納期ベース）"""
    days_to_add = get_delivery_days_to_add(row.item_group_code, cache)

    # 転送中（直送） → そのまま出荷予定
    if storage_place == "転送中（直送用）":
        return _result(delivery_date, "出荷済み", "出荷予定", today)

    # 紐付き/路線便/受注先≠出荷先
    if use_ship_rule or is_rosenbin:
        delivery_days = get_customer_delivery_days(row.customer_name, cache)

        if delivery_days:
            # 曜日制限あり
            if is_rosenbin and not original_use_ship_rule:
                base = add_business_days(delivery_date, days_to_add, holidays)
            else:
                base = delivery_date
            adjusted = get_next_delivery_day(base, delivery_days, holidays)
            return _result(adjusted, "出荷済み", "出荷予定", today)

        # 曜日制限なし
        if is_rosenbin and not original_use_ship_rule:
            rosenbin_date = add_business_days(
                delivery_date, max(days_to_add - 1, 0), holidays
            )
        else:
            rosenbin_date = delivery_date
        return _result(rosenbin_date, "出荷済み", "出荷予定", today)

    # 通常（自社便配達）
    delivery_days = get_customer_delivery_days(row.customer_name, cache)

    if delivery_days:
        # 曜日制限あり → +営業日 → 次の出荷曜日 → 「出荷予定」
        adjusted = add_business_days(delivery_date, days_to_add, holidays)
        adjusted = get_next_delivery_day(adjusted, delivery_days, holidays)
        return _result(adjusted, "出荷済み", "出荷予定", today)

    # 曜日制限なし → +営業日 → 「配達予定」
    adjusted = add_business_days(delivery_date, days_to_add, holidays)
    return _result(adjusted, "配達済み", "配達予定", today)
