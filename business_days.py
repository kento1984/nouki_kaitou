"""営業日・カレンダー計算モジュール

VBAの以下の関数を移植:
- IsHoliday (L5417): 祝日判定
- AddBusinessDays (L5202): N営業日後の日付計算
- GetNextBusinessDay (L5433): 翌営業日取得
- GetPreviousBusinessDay (L5454): 前営業日取得
- GetNextDeliveryDay (L5280): 次の配送可能日（曜日制限・振替・間隔ルール対応）
- ShouldSkipDueToInterval (L5237): 間隔1日ルール判定
- IsShiftedDeliveryDay (L5349): 振替出荷日判定
- CheckIntervalRule (L5384): 間隔ルールチェック
- CountBusinessDaysBetween (L7526): 営業日数カウント

VBA Weekday番号: 日=1, 月=2, 火=3, 水=4, 木=5, 金=6, 土=7
Python weekday(): 月=0, 火=1, 水=2, 木=3, 金=4, 土=5, 日=6

変換: vba_weekday = (python_weekday + 2) % 7 or 7 if 0
"""

from __future__ import annotations

import datetime
from typing import Optional

from nouki_kaitou.models import HolidayMap


def _to_vba_weekday(d: datetime.date) -> int:
    """Python dateのweekday()をVBA Weekday番号に変換する。

    Python: 月=0, 火=1, 水=2, 木=3, 金=4, 土=5, 日=6
    VBA:    日=1, 月=2, 火=3, 水=4, 木=5, 金=6, 土=7
    """
    return (d.weekday() + 2) % 7 or 7


def _is_weekend(d: datetime.date) -> bool:
    """土曜日(5)または日曜日(6)かどうか。"""
    return d.weekday() >= 5


# ============================================
# VBA: IsHoliday (L5417-5428)
# 祝日かどうか判定
# ============================================
def is_holiday(
    check_date: datetime.date,
    holidays: HolidayMap | None = None,
) -> bool:
    """指定日が祝日かどうかを判定する。

    HolidayMapの値がNone → 祝日（休日）
    HolidayMapの値がint → 特別締切日（営業日扱い、祝日ではない）

    Args:
        check_date: 判定対象日
        holidays: 祝日辞書（Noneなら祝日なし）

    Returns:
        True = 祝日（休日）
    """
    if holidays is None:
        return False

    if check_date in holidays:
        # 値がNoneなら祝日、intなら特別締切日（営業日）
        return holidays[check_date] is None

    return False


# ============================================
# VBA: AddBusinessDays (L5202-5233)
# 開始日からN営業日後の日付を計算
# ============================================
def add_business_days(
    start_date: datetime.date,
    business_days: int,
    holidays: HolidayMap | None = None,
) -> datetime.date:
    """開始日からN営業日後の日付を返す。

    土日・祝日をスキップ。特別締切日（HolidayMap値がint）は営業日としてカウント。

    Args:
        start_date: 開始日
        business_days: 加算する営業日数
        holidays: 祝日辞書

    Returns:
        N営業日後の日付
    """
    current = start_date
    days_added = 0

    while days_added < business_days:
        current += datetime.timedelta(days=1)

        # 土日はスキップ
        if _is_weekend(current):
            continue

        # 祝日チェック
        if holidays is not None and current in holidays:
            if holidays[current] is None:
                # 祝日 → スキップ
                continue
            # 特別締切日 → 営業日としてカウント（fallthrough）

        days_added += 1

    return current


# ============================================
# VBA: GetNextBusinessDay (L5433-5450)
# 翌営業日を取得
# ============================================
def get_next_business_day(
    start_date: datetime.date,
    holidays: HolidayMap | None = None,
) -> datetime.date:
    """指定日の翌営業日を返す。土日・祝日をスキップ。

    Args:
        start_date: 基準日
        holidays: 祝日辞書

    Returns:
        翌営業日
    """
    check = start_date + datetime.timedelta(days=1)

    while True:
        if _is_weekend(check):
            check += datetime.timedelta(days=1)
        elif is_holiday(check, holidays):
            check += datetime.timedelta(days=1)
        else:
            break

    return check


# ============================================
# VBA: GetPreviousBusinessDay (L5454-5469)
# 前営業日を取得
# ============================================
def get_previous_business_day(
    target_date: datetime.date,
    holidays: HolidayMap | None = None,
) -> datetime.date:
    """指定日の前営業日を返す。土日・祝日をスキップして1日戻る。

    Args:
        target_date: 基準日
        holidays: 祝日辞書

    Returns:
        前営業日
    """
    check = target_date - datetime.timedelta(days=1)

    while True:
        if _is_weekend(check):
            check -= datetime.timedelta(days=1)
        elif is_holiday(check, holidays):
            check -= datetime.timedelta(days=1)
        else:
            break

    return check


# ============================================
# VBA: ShouldSkipDueToInterval (L5237-5275)
# 間隔1日ルール：2日前が出荷曜日かつ祝日だった場合スキップ
# ============================================
def should_skip_due_to_interval(
    check_date: datetime.date,
    delivery_days: list[int],
    holidays: HolidayMap | None = None,
) -> bool:
    """間隔1日ルール判定。

    2日前が出荷曜日かつ祝日だった場合、振替出荷が行われたはずなので
    間隔が近すぎるためスキップすべき。

    Args:
        check_date: 判定対象日
        delivery_days: 出荷曜日リスト（VBA Weekday番号: 日=1..土=7）
        holidays: 祝日辞書

    Returns:
        True = スキップすべき
    """
    two_days_ago = check_date - datetime.timedelta(days=2)

    # 2日前が土日ならルール適用外
    if _is_weekend(two_days_ago):
        return False

    # 2日前が出荷曜日かチェック
    two_days_ago_weekday = _to_vba_weekday(two_days_ago)
    if two_days_ago_weekday not in delivery_days:
        return False

    # 2日前が祝日ならスキップすべき
    return is_holiday(two_days_ago, holidays)


# ============================================
# VBA: CheckIntervalRule (L5384-5413)
# 間隔ルールチェック（ShouldSkipDueToIntervalとほぼ同じロジック）
# ============================================
def check_interval_rule(
    check_date: datetime.date,
    delivery_days: list[int],
    holidays: HolidayMap | None = None,
) -> bool:
    """間隔ルール判定。

    2日前が出荷曜日かつ祝日の場合True。
    振替出荷があったため間隔が近すぎることを示す。

    Args:
        check_date: 判定対象日
        delivery_days: 出荷曜日リスト（VBA Weekday番号）
        holidays: 祝日辞書

    Returns:
        True = 間隔ルールに該当
    """
    two_days_ago = check_date - datetime.timedelta(days=2)

    # 土日ならルール適用外
    if _is_weekend(two_days_ago):
        return False

    # 2日前が出荷曜日かチェック
    two_days_ago_weekday = _to_vba_weekday(two_days_ago)
    if two_days_ago_weekday not in delivery_days:
        return False

    # 2日前が祝日ならルール適用
    return is_holiday(two_days_ago, holidays)


# ============================================
# VBA: IsShiftedDeliveryDay (L5349-5379)
# 振替出荷日かどうか判定
# ============================================
def is_shifted_delivery_day(
    check_date: datetime.date,
    delivery_days: list[int],
    holidays: HolidayMap | None = None,
) -> bool:
    """振替出荷日かどうかを判定する。

    前日が出荷曜日で祝日、または前日が出荷曜日で間隔ルールに該当する場合、
    当日は振替出荷日になる。

    Args:
        check_date: 判定対象日
        delivery_days: 出荷曜日リスト（VBA Weekday番号）
        holidays: 祝日辞書

    Returns:
        True = 振替出荷日
    """
    yesterday = check_date - datetime.timedelta(days=1)

    # 前日が出荷曜日かチェック
    yesterday_weekday = _to_vba_weekday(yesterday)
    if yesterday_weekday not in delivery_days:
        return False

    # 前日が間隔ルールに該当 → 今日は振替出荷日
    if check_interval_rule(yesterday, delivery_days, holidays):
        return True

    # 前日が祝日 → 今日は振替出荷日
    if is_holiday(yesterday, holidays):
        return True

    return False


# ============================================
# VBA: GetNextDeliveryDay (L5280-5343)
# 次の配送可能日を計算（曜日制限・振替出荷日対応）
# ============================================
def get_next_delivery_day(
    base_date: datetime.date,
    delivery_days: list[int],
    holidays: HolidayMap | None = None,
) -> datetime.date:
    """基準日から次の出荷日を取得する。

    出荷曜日制限・振替出荷日・間隔ルールを考慮して
    次の出荷可能日を返す。

    ロジック:
    1. 土日はスキップ
    2. 当日が出荷曜日なら:
       - 祝日 → 翌営業日を返す
       - 間隔ルール該当 → 翌営業日を返す
       - それ以外 → 当日を返す
    3. 当日が振替出荷日なら → 当日を返す
    4. 上記以外 → 翌日へ進む

    Args:
        base_date: 基準日
        delivery_days: 出荷曜日リスト（VBA Weekday番号: 日=1..土=7）。
                       空リストの場合は制限なしとして営業日を返す。
        holidays: 祝日辞書

    Returns:
        次の出荷日
    """
    # 出荷曜日が空なら曜日制限なし → 基準日が営業日ならそのまま返す
    if not delivery_days:
        check = base_date
        while True:
            if _is_weekend(check):
                check += datetime.timedelta(days=1)
            elif is_holiday(check, holidays):
                check += datetime.timedelta(days=1)
            else:
                return check

    check = base_date
    max_loop = 30

    for _ in range(max_loop):
        # 土日はスキップ
        if _is_weekend(check):
            check += datetime.timedelta(days=1)
            continue

        # 出荷曜日かチェック
        check_weekday = _to_vba_weekday(check)
        is_delivery = check_weekday in delivery_days

        # 振替出荷日かチェック
        is_shifted = is_shifted_delivery_day(check, delivery_days, holidays)

        if is_delivery:
            # 通常の出荷曜日
            if is_holiday(check, holidays):
                return get_next_business_day(check, holidays)

            if check_interval_rule(check, delivery_days, holidays):
                return get_next_business_day(check, holidays)

            return check

        elif is_shifted:
            # 振替出荷日 → そのまま出荷OK
            return check

        check += datetime.timedelta(days=1)

    # 30日以内に見つからなければ基準日を返す（フォールバック）
    return base_date


# ============================================
# VBA: CountBusinessDaysBetween (L7526-7554)
# 2つの日付間の営業日数をカウント
# ============================================
def count_business_days_between(
    start_date: datetime.date,
    end_date: datetime.date,
    holidays: HolidayMap | None = None,
) -> int:
    """2つの日付間の営業日数をカウントする。

    start_dateの翌日からend_dateまでの営業日数。
    特別締切日（HolidayMap値がint）は営業日としてカウント。

    数学的に平日数を計算し、祝日を差し引く。
    O(日数) → O(祝日数) に最適化。

    Args:
        start_date: 開始日（この日は含まない）
        end_date: 終了日（この日を含む）
        holidays: 祝日辞書

    Returns:
        営業日数
    """
    if start_date >= end_date:
        return 0

    # start_date+1 から end_date までの範囲
    first_day = start_date + datetime.timedelta(days=1)
    total_days = (end_date - first_day).days + 1

    if total_days <= 0:
        return 0

    # 平日数を数学的に計算
    full_weeks = total_days // 7
    weekday_count = full_weeks * 5
    remainder = total_days % 7
    start_wd = first_day.weekday()  # 月=0, 日=6
    for i in range(remainder):
        if (start_wd + i) % 7 < 5:
            weekday_count += 1

    # 祝日（平日のみ）を差し引く
    if holidays:
        for hdate, hvalue in holidays.items():
            if hvalue is not None:
                # 特別締切日 → 営業日（差し引かない）
                continue
            if first_day <= hdate <= end_date and hdate.weekday() < 5:
                weekday_count -= 1

    return weekday_count
