"""business_days.py のユニットテスト

テストで使用するカレンダー想定（2026年1月〜3月）:
- 1/1(木) 元旦 → 祝日
- 1/2(金) → 祝日
- 1/3(土) → 土曜
- 1/4(日) → 日曜
- 1/5(月) → 営業日
- 1/12(月) 成人の日 → 祝日
- 1/13(火) → 営業日
- 2/11(水) 建国記念日 → 祝日
- 2/23(月) 天皇誕生日 → 祝日
- 3/20(金) 春分の日 → 祝日
- 12/30(水) 年末 → 特別締切日（12時）= 営業日扱い

VBA Weekday番号: 日=1, 月=2, 火=3, 水=4, 木=5, 金=6, 土=7
"""

import datetime

import pytest

from nouki_kaitou.business_days import (
    _to_vba_weekday,
    add_business_days,
    check_interval_rule,
    count_business_days_between,
    get_next_business_day,
    get_next_delivery_day,
    get_previous_business_day,
    is_holiday,
    is_shifted_delivery_day,
    should_skip_due_to_interval,
)
from nouki_kaitou.models import HolidayMap


# ============================================
# テスト用祝日データ
# ============================================
@pytest.fixture
def holidays_2026() -> HolidayMap:
    """2026年のテスト用祝日辞書"""
    return {
        datetime.date(2026, 1, 1): None,     # 元旦（木）
        datetime.date(2026, 1, 2): None,     # 休み（金）
        datetime.date(2026, 1, 12): None,    # 成人の日（月）
        datetime.date(2026, 2, 11): None,    # 建国記念日（水）
        datetime.date(2026, 2, 23): None,    # 天皇誕生日（月）
        datetime.date(2026, 3, 20): None,    # 春分の日（金）
        datetime.date(2026, 12, 30): 12,     # 年末特別締切12時（水）= 営業日
    }


# ============================================
# _to_vba_weekday
# ============================================
class TestToVbaWeekday:
    def test_monday(self):
        # 2026/1/5 = 月曜日
        assert _to_vba_weekday(datetime.date(2026, 1, 5)) == 2

    def test_tuesday(self):
        assert _to_vba_weekday(datetime.date(2026, 1, 6)) == 3

    def test_wednesday(self):
        assert _to_vba_weekday(datetime.date(2026, 1, 7)) == 4

    def test_thursday(self):
        assert _to_vba_weekday(datetime.date(2026, 1, 8)) == 5

    def test_friday(self):
        assert _to_vba_weekday(datetime.date(2026, 1, 9)) == 6

    def test_saturday(self):
        assert _to_vba_weekday(datetime.date(2026, 1, 10)) == 7

    def test_sunday(self):
        assert _to_vba_weekday(datetime.date(2026, 1, 11)) == 1


# ============================================
# is_holiday
# ============================================
class TestIsHoliday:
    def test_holiday(self, holidays_2026: HolidayMap):
        """祝日判定"""
        assert is_holiday(datetime.date(2026, 1, 1), holidays_2026) is True

    def test_special_cutoff_is_not_holiday(self, holidays_2026: HolidayMap):
        """特別締切日は祝日ではない（営業日扱い）"""
        assert is_holiday(datetime.date(2026, 12, 30), holidays_2026) is False

    def test_normal_day(self, holidays_2026: HolidayMap):
        """通常営業日"""
        assert is_holiday(datetime.date(2026, 1, 5), holidays_2026) is False

    def test_no_holidays(self):
        """祝日辞書なし"""
        assert is_holiday(datetime.date(2026, 1, 1), None) is False

    def test_empty_holidays(self):
        """空の祝日辞書"""
        assert is_holiday(datetime.date(2026, 1, 1), {}) is False


# ============================================
# add_business_days
# ============================================
class TestAddBusinessDays:
    def test_simple(self):
        """月曜から1営業日後 = 火曜"""
        # 2026/1/5(月) + 1営業日 = 1/6(火)
        result = add_business_days(datetime.date(2026, 1, 5), 1)
        assert result == datetime.date(2026, 1, 6)

    def test_skip_weekend(self):
        """金曜から1営業日後 = 月曜"""
        # 2026/1/9(金) + 1営業日 = 1/12(月)
        result = add_business_days(datetime.date(2026, 1, 9), 1)
        assert result == datetime.date(2026, 1, 12)

    def test_skip_holiday(self, holidays_2026: HolidayMap):
        """祝日をスキップ"""
        # 2026/1/9(金) + 1営業日 = 1/12(月)は成人の日 → 1/13(火)
        result = add_business_days(
            datetime.date(2026, 1, 9), 1, holidays_2026
        )
        assert result == datetime.date(2026, 1, 13)

    def test_multiple_days(self, holidays_2026: HolidayMap):
        """複数営業日"""
        # 2026/1/5(月) + 3営業日 = 1/8(木)
        result = add_business_days(
            datetime.date(2026, 1, 5), 3, holidays_2026
        )
        assert result == datetime.date(2026, 1, 8)

    def test_special_cutoff_is_business_day(self, holidays_2026: HolidayMap):
        """特別締切日は営業日としてカウント"""
        # 2026/12/29(火) + 1営業日 = 12/30(水) 特別締切日
        result = add_business_days(
            datetime.date(2026, 12, 29), 1, holidays_2026
        )
        assert result == datetime.date(2026, 12, 30)

    def test_zero_days(self):
        """0営業日 = 当日"""
        result = add_business_days(datetime.date(2026, 1, 5), 0)
        assert result == datetime.date(2026, 1, 5)

    def test_new_year_holiday(self, holidays_2026: HolidayMap):
        """年末年始の連休スキップ"""
        # 2025/12/31(水) + 1営業日 → 1/1祝日, 1/2祝日, 1/3土, 1/4日 → 1/5(月)
        result = add_business_days(
            datetime.date(2025, 12, 31), 1, holidays_2026
        )
        assert result == datetime.date(2026, 1, 5)

    def test_no_holidays_dict(self):
        """祝日辞書なしでも動作"""
        result = add_business_days(datetime.date(2026, 1, 5), 2, None)
        assert result == datetime.date(2026, 1, 7)


# ============================================
# get_next_business_day
# ============================================
class TestGetNextBusinessDay:
    def test_weekday(self):
        """平日の翌営業日"""
        # 2026/1/5(月) → 1/6(火)
        result = get_next_business_day(datetime.date(2026, 1, 5))
        assert result == datetime.date(2026, 1, 6)

    def test_friday(self):
        """金曜の翌営業日 = 月曜"""
        # 2026/1/9(金) → 1/12(月)
        result = get_next_business_day(datetime.date(2026, 1, 9))
        assert result == datetime.date(2026, 1, 12)

    def test_skip_holiday(self, holidays_2026: HolidayMap):
        """祝日スキップ"""
        # 2026/1/9(金) → 1/12(月)成人の日スキップ → 1/13(火)
        result = get_next_business_day(
            datetime.date(2026, 1, 9), holidays_2026
        )
        assert result == datetime.date(2026, 1, 13)

    def test_from_saturday(self, holidays_2026: HolidayMap):
        """土曜から"""
        # 2026/1/10(土) → 1/12(月)成人の日スキップ → 1/13(火)
        result = get_next_business_day(
            datetime.date(2026, 1, 10), holidays_2026
        )
        assert result == datetime.date(2026, 1, 13)

    def test_special_cutoff_not_skipped(self, holidays_2026: HolidayMap):
        """特別締切日はスキップされない"""
        # 2026/12/29(火) → 12/30(水) 特別締切日 = 営業日
        result = get_next_business_day(
            datetime.date(2026, 12, 29), holidays_2026
        )
        assert result == datetime.date(2026, 12, 30)


# ============================================
# get_previous_business_day
# ============================================
class TestGetPreviousBusinessDay:
    def test_weekday(self):
        """平日の前営業日"""
        # 2026/1/6(火) → 1/5(月)
        result = get_previous_business_day(datetime.date(2026, 1, 6))
        assert result == datetime.date(2026, 1, 5)

    def test_monday(self):
        """月曜の前営業日 = 金曜"""
        # 2026/1/12(月) → 1/9(金)
        result = get_previous_business_day(datetime.date(2026, 1, 12))
        assert result == datetime.date(2026, 1, 9)

    def test_skip_holiday(self, holidays_2026: HolidayMap):
        """祝日スキップ"""
        # 2026/1/13(火) → 1/12(月)成人の日 → 1/9(金)
        result = get_previous_business_day(
            datetime.date(2026, 1, 13), holidays_2026
        )
        assert result == datetime.date(2026, 1, 9)

    def test_new_year(self, holidays_2026: HolidayMap):
        """年始の前営業日"""
        # 2026/1/5(月) → 1/4(日),1/3(土),1/2(金)祝日,1/1(木)祝日 → 12/31(水)
        result = get_previous_business_day(
            datetime.date(2026, 1, 5), holidays_2026
        )
        assert result == datetime.date(2025, 12, 31)

    def test_special_cutoff_not_skipped(self, holidays_2026: HolidayMap):
        """特別締切日はスキップされない"""
        # 2026/12/31(木) → 12/30(水) 特別締切日 = 営業日
        result = get_previous_business_day(
            datetime.date(2026, 12, 31), holidays_2026
        )
        assert result == datetime.date(2026, 12, 30)


# ============================================
# should_skip_due_to_interval
# ============================================
class TestShouldSkipDueToInterval:
    def test_no_skip(self):
        """通常日はスキップしない"""
        # 2026/1/7(水): 2日前=1/5(月), 出荷曜日=[2(月),4(水),6(金)]
        # 1/5(月)は出荷曜日だが祝日でない → False
        delivery_days = [2, 4, 6]  # 月水金
        result = should_skip_due_to_interval(
            datetime.date(2026, 1, 7), delivery_days
        )
        assert result is False

    def test_skip_due_to_holiday(self, holidays_2026: HolidayMap):
        """2日前が出荷曜日かつ祝日 → スキップ"""
        # 2026/1/14(水): 2日前=1/12(月)=成人の日
        # 出荷曜日=[2(月),4(水),6(金)], 1/12は月曜で出荷曜日かつ祝日 → True
        delivery_days = [2, 4, 6]
        result = should_skip_due_to_interval(
            datetime.date(2026, 1, 14), delivery_days, holidays_2026
        )
        assert result is True

    def test_two_days_ago_weekend(self):
        """2日前が土日→ルール適用外"""
        # 2026/1/12(月): 2日前=1/10(土) → False
        delivery_days = [2, 4, 6]
        result = should_skip_due_to_interval(
            datetime.date(2026, 1, 12), delivery_days
        )
        assert result is False

    def test_two_days_ago_not_delivery_day(self, holidays_2026: HolidayMap):
        """2日前が出荷曜日でない→ルール適用外"""
        # 出荷曜日=[3(火),5(木)] のとき
        # 2026/1/14(水): 2日前=1/12(月)=成人の日だが、月は出荷曜日でない → False
        delivery_days = [3, 5]
        result = should_skip_due_to_interval(
            datetime.date(2026, 1, 14), delivery_days, holidays_2026
        )
        assert result is False


# ============================================
# check_interval_rule
# ============================================
class TestCheckIntervalRule:
    def test_interval_rule_hit(self, holidays_2026: HolidayMap):
        """間隔ルールに該当"""
        # 2026/1/14(水): 2日前=1/12(月)成人の日、出荷曜日に月曜含む → True
        delivery_days = [2, 4, 6]  # 月水金
        result = check_interval_rule(
            datetime.date(2026, 1, 14), delivery_days, holidays_2026
        )
        assert result is True

    def test_interval_rule_miss(self, holidays_2026: HolidayMap):
        """間隔ルールに非該当"""
        # 2026/1/8(木): 2日前=1/6(火)、火は出荷曜日[2,4,6]に含まない → False
        delivery_days = [2, 4, 6]
        result = check_interval_rule(
            datetime.date(2026, 1, 8), delivery_days, holidays_2026
        )
        assert result is False

    def test_two_days_ago_weekend(self):
        """2日前が土日→ルール非該当"""
        delivery_days = [2, 7]  # 月土
        result = check_interval_rule(
            datetime.date(2026, 1, 12), delivery_days  # 2日前=1/10(土)
        )
        assert result is False


# ============================================
# is_shifted_delivery_day
# ============================================
class TestIsShiftedDeliveryDay:
    def test_shifted_due_to_holiday(self, holidays_2026: HolidayMap):
        """前日が出荷曜日で祝日 → 振替出荷日"""
        # 2026/1/13(火): 前日=1/12(月)成人の日, 出荷曜日に月曜含む → True
        delivery_days = [2, 4, 6]  # 月水金
        result = is_shifted_delivery_day(
            datetime.date(2026, 1, 13), delivery_days, holidays_2026
        )
        assert result is True

    def test_not_shifted_normal(self, holidays_2026: HolidayMap):
        """前日が通常営業日 → 振替なし"""
        # 2026/1/7(水): 前日=1/6(火), 火は出荷曜日に含まない → False
        delivery_days = [2, 4, 6]
        result = is_shifted_delivery_day(
            datetime.date(2026, 1, 7), delivery_days, holidays_2026
        )
        assert result is False

    def test_not_shifted_yesterday_not_delivery_day(self, holidays_2026: HolidayMap):
        """前日が出荷曜日でない → 振替なし"""
        # 出荷曜日=[3(火),5(木)]
        # 2026/1/13(火): 前日=1/12(月)は祝日だが月曜は出荷曜日でない → False
        delivery_days = [3, 5]
        result = is_shifted_delivery_day(
            datetime.date(2026, 1, 13), delivery_days, holidays_2026
        )
        assert result is False

    def test_shifted_due_to_interval_rule(self, holidays_2026: HolidayMap):
        """前日が間隔ルール該当 → 振替出荷日"""
        # 前日が出荷曜日で、前日のCheckIntervalRuleがTrueになるケース
        # 1/14(水)は出荷曜日, CheckIntervalRule(1/14)=True(2日前1/12成人の日)
        # → 1/15(木)は振替出荷日
        delivery_days = [2, 4, 6]  # 月水金
        result = is_shifted_delivery_day(
            datetime.date(2026, 1, 15), delivery_days, holidays_2026
        )
        # 1/15の前日=1/14(水)は出荷曜日。CheckIntervalRule(1/14)は
        # 2日前=1/12(月)成人の日,月曜は出荷曜日 → True
        # よって1/15は振替出荷日
        assert result is True


# ============================================
# get_next_delivery_day
# ============================================
class TestGetNextDeliveryDay:
    def test_today_is_delivery_day(self):
        """当日が出荷曜日なら当日を返す"""
        # 2026/1/5(月), 出荷曜日=[2(月),4(水),6(金)]
        delivery_days = [2, 4, 6]
        result = get_next_delivery_day(datetime.date(2026, 1, 5), delivery_days)
        assert result == datetime.date(2026, 1, 5)

    def test_skip_non_delivery_day(self):
        """当日が出荷曜日でなければ次の出荷曜日へ"""
        # 2026/1/6(火), 出荷曜日=[2(月),4(水),6(金)] → 1/7(水)
        delivery_days = [2, 4, 6]
        result = get_next_delivery_day(datetime.date(2026, 1, 6), delivery_days)
        assert result == datetime.date(2026, 1, 7)

    def test_skip_weekend(self):
        """土日をスキップ"""
        # 2026/1/10(土), 出荷曜日=[2(月)] → 1/12(月)
        delivery_days = [2]
        result = get_next_delivery_day(datetime.date(2026, 1, 10), delivery_days)
        assert result == datetime.date(2026, 1, 12)

    def test_delivery_day_on_holiday(self, holidays_2026: HolidayMap):
        """出荷曜日が祝日 → 翌営業日を返す"""
        # 2026/1/12(月)=成人の日, 出荷曜日=[2(月)]
        # → 祝日なので翌営業日 = 1/13(火)
        delivery_days = [2]
        result = get_next_delivery_day(
            datetime.date(2026, 1, 12), delivery_days, holidays_2026
        )
        assert result == datetime.date(2026, 1, 13)

    def test_shifted_delivery_day(self, holidays_2026: HolidayMap):
        """振替出荷日の検出"""
        # 2026/1/12(月)=成人の日, 出荷曜日=[2(月),4(水),6(金)]
        # 1/12は祝日 → 翌営業日1/13(火)を返す
        # 1/13は振替出荷日だが、1/12が祝日なのでget_next_business_dayで1/13が返る
        delivery_days = [2, 4, 6]
        result = get_next_delivery_day(
            datetime.date(2026, 1, 12), delivery_days, holidays_2026
        )
        assert result == datetime.date(2026, 1, 13)

    def test_empty_delivery_days(self, holidays_2026: HolidayMap):
        """出荷曜日が空 → 曜日制限なし、営業日を返す"""
        # 2026/1/5(月) → 営業日なのでそのまま
        result = get_next_delivery_day(
            datetime.date(2026, 1, 5), [], holidays_2026
        )
        assert result == datetime.date(2026, 1, 5)

    def test_empty_delivery_days_on_weekend(self, holidays_2026: HolidayMap):
        """出荷曜日が空で土日 → 翌営業日"""
        # 2026/1/10(土) → 1/12(月)成人の日 → 1/13(火)
        result = get_next_delivery_day(
            datetime.date(2026, 1, 10), [], holidays_2026
        )
        assert result == datetime.date(2026, 1, 13)

    def test_interval_rule_skip(self, holidays_2026: HolidayMap):
        """間隔ルールで出荷曜日をスキップ"""
        # 2026/1/14(水): 出荷曜日=[2(月),4(水),6(金)]
        # CheckIntervalRule: 2日前=1/12(月)成人の日, 月は出荷曜日 → True
        # → 翌営業日=1/15(木)へ。1/15は振替出荷日なのでそこで返る。
        # ただし1/15は木曜で出荷曜日[月水金]ではない。IsShiftedDeliveryDay(1/15)を確認:
        # 前日1/14(水)は出荷曜日、CheckIntervalRule(1/14)=True → 振替出荷日
        delivery_days = [2, 4, 6]
        result = get_next_delivery_day(
            datetime.date(2026, 1, 14), delivery_days, holidays_2026
        )
        # 1/14(水)は出荷曜日だが間隔ルール該当→翌営業日1/15(木)
        # 1/15(木)は振替出荷日なのでOK
        assert result == datetime.date(2026, 1, 15)

    def test_consecutive_weeks(self):
        """連続する週で正しい出荷日を返す"""
        delivery_days = [2, 6]  # 月金
        # 2026/1/6(火) → 次の出荷曜日=1/9(金)
        result = get_next_delivery_day(datetime.date(2026, 1, 6), delivery_days)
        assert result == datetime.date(2026, 1, 9)


# ============================================
# count_business_days_between
# ============================================
class TestCountBusinessDaysBetween:
    def test_same_week(self):
        """同一週内"""
        # 1/5(月) → 1/9(金) = 4営業日
        result = count_business_days_between(
            datetime.date(2026, 1, 5), datetime.date(2026, 1, 9)
        )
        assert result == 4

    def test_over_weekend(self):
        """土日をまたぐ"""
        # 1/9(金) → 1/13(火) = 火のみ1営業日...
        # 1/10(土)NG, 1/11(日)NG, 1/12(月)OK, 1/13(火)OK = 2営業日
        result = count_business_days_between(
            datetime.date(2026, 1, 9), datetime.date(2026, 1, 13)
        )
        assert result == 2

    def test_skip_holidays(self, holidays_2026: HolidayMap):
        """祝日をスキップしてカウント"""
        # 1/9(金) → 1/13(火): 1/10(土)NG, 1/11(日)NG, 1/12(月)祝日NG, 1/13(火)OK
        result = count_business_days_between(
            datetime.date(2026, 1, 9), datetime.date(2026, 1, 13), holidays_2026
        )
        assert result == 1

    def test_special_cutoff_counted(self, holidays_2026: HolidayMap):
        """特別締切日は営業日としてカウント"""
        # 12/29(火) → 12/30(水)特別締切日 = 1営業日
        result = count_business_days_between(
            datetime.date(2026, 12, 29),
            datetime.date(2026, 12, 30),
            holidays_2026,
        )
        assert result == 1

    def test_same_day(self):
        """同日 = 0営業日"""
        result = count_business_days_between(
            datetime.date(2026, 1, 5), datetime.date(2026, 1, 5)
        )
        assert result == 0

    def test_no_holidays_dict(self):
        """祝日辞書なし"""
        result = count_business_days_between(
            datetime.date(2026, 1, 5), datetime.date(2026, 1, 7), None
        )
        assert result == 2

    def test_full_week(self):
        """1週間 = 5営業日"""
        # 1/5(月) → 1/12(月) = 5営業日
        result = count_business_days_between(
            datetime.date(2026, 1, 5), datetime.date(2026, 1, 12)
        )
        assert result == 5


# ============================================
# 統合テスト：実際のビジネスシナリオ
# ============================================
class TestBusinessScenarios:
    def test_new_year_first_business_day(self, holidays_2026: HolidayMap):
        """年末年始後の最初の営業日"""
        # 2025/12/31(水) → 翌営業日: 1/1祝日, 1/2祝日, 1/3土, 1/4日 → 1/5(月)
        result = get_next_business_day(
            datetime.date(2025, 12, 31), holidays_2026
        )
        assert result == datetime.date(2026, 1, 5)

    def test_delivery_after_holiday_shifted(self, holidays_2026: HolidayMap):
        """祝日後の振替出荷"""
        # 出荷曜日=月水金、1/12(月)成人の日
        # 1/12から次の出荷日 → 1/12は月で出荷曜日だが祝日 → 翌営業日1/13(火)
        delivery_days = [2, 4, 6]
        result = get_next_delivery_day(
            datetime.date(2026, 1, 12), delivery_days, holidays_2026
        )
        assert result == datetime.date(2026, 1, 13)

    def test_add_business_days_over_long_holiday(self, holidays_2026: HolidayMap):
        """長期休暇をまたぐ営業日加算"""
        # 2025/12/31(水) + 2営業日
        # → 1/1祝日,1/2祝日,1/3土,1/4日 → 1/5(月)=1日目, 1/6(火)=2日目
        result = add_business_days(
            datetime.date(2025, 12, 31), 2, holidays_2026
        )
        assert result == datetime.date(2026, 1, 6)

    def test_previous_business_day_before_holiday(self, holidays_2026: HolidayMap):
        """祝日前の営業日"""
        # 2026/2/11(水)建国記念日 → 前営業日 = 2/10(火)
        result = get_previous_business_day(
            datetime.date(2026, 2, 11), holidays_2026
        )
        assert result == datetime.date(2026, 2, 10)

    def test_delivery_day_with_no_restriction(self, holidays_2026: HolidayMap):
        """曜日制限なしの出荷日"""
        # 曜日制限なし、平日ならOK
        result = get_next_delivery_day(
            datetime.date(2026, 1, 6), [], holidays_2026
        )
        assert result == datetime.date(2026, 1, 6)
