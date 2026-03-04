"""仙台営業所 配送パターンベースの納期計算

既存の _check_stock_completed（単一cutoff→翌日or翌々日）では対応できない
仙台営業所の配送パターン（近隣2便/遠方午前/遠方午後/路線便）に対応する。

calculate_delivery_date から呼ばれ、仙台以外は即 None を返して
既存ロジックにフォールスルーする。
"""

from __future__ import annotations

import datetime
from typing import Optional

from nouki_kaitou.business_days import (
    _is_weekend,
    add_business_days,
    get_next_business_day,
    get_next_delivery_day,
    is_holiday,
)
from nouki_kaitou.customer import get_customer_delivery_days, get_customer_pattern
from nouki_kaitou.models import (
    BranchSettings,
    CacheStore,
    DeliveryPattern,
    HolidayMap,
    OrderRow,
)
from nouki_kaitou.business_days import get_previous_business_day as _get_prev_bday
from nouki_kaitou.utils import format_date_japanese, normalize_name_for_comparison, parse_time


def _before_cutoff_hm(
    hour: int, minute: int, cutoff: tuple[int, int]
) -> bool:
    """cutoff時刻より前かどうか（cutoffちょうど＝締切後）。"""
    return (hour, minute) < cutoff


def _calc_pattern_days(
    hour: int, minute: int, pattern: DeliveryPattern
) -> int:
    """パターン定義から配達までの営業日数を算出する。"""
    if (hour, minute) < pattern.cutoff1:
        return pattern.days_before_cutoff1
    if pattern.cutoff2 is not None and (hour, minute) < pattern.cutoff2:
        return pattern.days_between_cutoffs
    return pattern.days_after_all


def _calc_pattern_period(
    hour: int, minute: int, pattern: DeliveryPattern
) -> str:
    """パターン定義から配達時間帯ラベルを算出する（空文字＝表示なし）。

    2段階cutoff（cutoff2あり）のパターンは午前便・午後便の区別があるため
    PM/AM/PMを自動設定する。1段階パターンは常に空文字。
    """
    if pattern.cutoff2 is None:
        return ""
    if (hour, minute) < pattern.cutoff1:
        return "PM"
    if (hour, minute) < pattern.cutoff2:
        return "AM"
    return "PM"


def check_sendai_stock_completed(
    row: OrderRow,
    cache: CacheStore,
    holidays: HolidayMap | None,
    branch: BranchSettings,
    today: datetime.date,
    storage_place: str,
    use_ship_rule: bool,
) -> Optional[str]:
    """仙台営業所の在庫販売+処理完了を配送パターンで計算する。

    仙台以外 / パターン未設定 / 在庫販売+処理完了でない場合は None を返し、
    呼び出し元の既存ロジックにフォールスルーする。

    Args:
        row: 受注データ行
        cache: マスターキャッシュ
        holidays: 祝日辞書
        branch: 営業所設定
        today: 今日の日付
        storage_place: 保管場所
        use_ship_rule: 受注先≠出荷先 or 路線便フラグ

    Returns:
        納期文字列。仙台以外・非該当なら None。
    """
    # 1. 仙台以外は即フォールスルー
    if "仙台" not in branch.name:
        return None

    # 2. 在庫販売+処理完了ガード
    if row.document_type != "【受注】在庫販売":
        return None
    if row.ship_status != "処理完了":
        return None

    # 3. 登録日・時刻取得
    reg_date = row.registration_date
    if reg_date is None:
        return None

    time_parts = parse_time(row.time_value)
    if time_parts is None:
        return None

    hour, minute = time_parts

    # 4. 土日祝→翌営業日起算
    if _is_weekend(reg_date) or is_holiday(reg_date, holidays):
        reg_date = get_next_business_day(reg_date, holidays)

    # 5. 顧客の配送パターン取得
    pattern_name = get_customer_pattern(row.customer_name, cache)
    if not pattern_name:
        return None
    pattern = cache.delivery_patterns.get(pattern_name)
    if pattern is None:
        return None

    # 6. 他支店在庫
    # 出荷可否はセンターの締切（branch.default_cutoff）で判定する。
    # pattern.cutoff1は配達ルートの便の締切であり、出荷能力とは無関係。
    branch_cutoff = (branch.default_cutoff, 0)
    base_center = branch.base_center
    if base_center and storage_place and storage_place != base_center:
        if _before_cutoff_hm(hour, minute, branch_cutoff):
            ship_date = reg_date
        else:
            ship_date = add_business_days(reg_date, 1, holidays)
        return _sendai_result(
            ship_date, "他拠点より出荷済み", "他拠点より出荷予定", today
        )

    # 7. use_ship_rule（受注先≠出荷先 or 路線便）
    if use_ship_rule:
        if _before_cutoff_hm(hour, minute, branch_cutoff):
            ship_date = reg_date
        else:
            ship_date = add_business_days(reg_date, 1, holidays)
        return _sendai_result(ship_date, "出荷済み", "出荷予定", today)

    # 8. パターンから営業日数・時間帯ラベルを算出
    biz_days = _calc_pattern_days(hour, minute, pattern)
    period = _calc_pattern_period(hour, minute, pattern)

    if biz_days == 0:
        adjusted = reg_date
    else:
        adjusted = add_business_days(reg_date, biz_days, holidays)

    # 9. 曜日制限チェック
    delivery_days = get_customer_delivery_days(row.customer_name, cache)
    if delivery_days:
        adjusted = get_next_delivery_day(adjusted, delivery_days, holidays)
        return _sendai_result(adjusted, "出荷済み", "出荷予定", today)

    # 10. 曜日制限なし（配達パス: periodあり）
    # biz_days==0: 配達日が今日以降なら「配達予定」、過去なら「配達済み」
    if biz_days == 0 and adjusted >= today:
        past_suffix = "配達予定"
    else:
        past_suffix = "配達済み"
    return _sendai_result(adjusted, past_suffix, "配達予定", today, period)


def check_sendai_himozuki_completed(
    row: OrderRow,
    cache: CacheStore,
    holidays: HolidayMap | None,
    branch: BranchSettings,
    execution_time: datetime.datetime,
    today: datetime.date,
    storage_place: str,
    is_rosenbin: bool,
) -> Optional[str]:
    """仙台営業所の紐付き+処理完了を配送パターンで計算する。

    既存の _check_himozuki_completed と同じガード条件だが、
    単一cutoffの代わりにパターンのcutoffsで営業日数を決定する。
    execution_time（マクロ実行時刻）ベースで計算。

    仙台以外 / パターン未設定 / 紐付き+処理完了でない場合は None を返し、
    呼び出し元の既存ロジックにフォールスルーする。

    Args:
        row: 受注データ行
        cache: マスターキャッシュ
        holidays: 祝日辞書
        branch: 営業所設定
        execution_time: マクロ実行時刻
        today: 今日の日付
        storage_place: 保管場所
        is_rosenbin: 路線便フラグ

    Returns:
        納期文字列。仙台以外・非該当なら None。
    """
    # 1. 仙台以外は即フォールスルー
    if "仙台" not in branch.name:
        return None

    # 2. 紐付き+処理完了ガード（直送販売+処理完了+非転送中）
    if row.document_type != "【受注】直送販売":
        return None
    if row.ship_status != "処理完了":
        return None
    if storage_place == "転送中（直送用）":
        return None

    # 3. 顧客の配送パターン取得
    pattern_name = get_customer_pattern(row.customer_name, cache)
    if not pattern_name:
        return None
    pattern = cache.delivery_patterns.get(pattern_name)
    if pattern is None:
        return None

    # 4. execution_timeからhour/minuteを取得
    hour, minute = execution_time.hour, execution_time.minute

    # 5. パターンから営業日数・時間帯ラベルを算出（todayベース）
    biz_days = _calc_pattern_days(hour, minute, pattern)
    period = _calc_pattern_period(hour, minute, pattern)

    if biz_days == 0:
        adjusted = today
    else:
        adjusted = add_business_days(today, biz_days, holidays)

    # 6. 受注先≠出荷先 → 配達日から1営業日逆算して出荷日
    if normalize_name_for_comparison(row.customer_name) != normalize_name_for_comparison(row.ship_to_name):
        ship_date = _get_prev_bday(adjusted, holidays)
        return _sendai_result(ship_date, "出荷済み", "出荷予定", today)

    # 7. 曜日制限チェック
    delivery_days = get_customer_delivery_days(row.customer_name, cache)
    if delivery_days:
        adjusted = get_next_delivery_day(adjusted, delivery_days, holidays)
        return _sendai_result(adjusted, "出荷済み", "出荷予定", today)

    # 8. 路線便 → 1営業日逆算して出荷日
    if is_rosenbin:
        ship_date = _get_prev_bday(adjusted, holidays)
        return _sendai_result(ship_date, "出荷済み", "出荷予定", today)

    # 9. 自社便配達（配達パス: periodあり）
    # biz_days==0: 配達日が今日以降なら「配達予定」、過去なら「配達済み」
    if biz_days == 0 and adjusted >= today:
        past_suffix = "配達予定"
    else:
        past_suffix = "配達済み"
    return _sendai_result(adjusted, past_suffix, "配達予定", today, period)


def _sendai_result(
    d: datetime.date, past: str, future: str, today: datetime.date,
    period: str = "",
) -> str:
    """日付 + 時間帯 + 過去/未来サフィックスでフォーマット（delivery_calc._resultと同等）"""
    return format_date_japanese(d) + period + (past if d <= today else future)
