"""delivery_calc.py のユニットテスト

CalculateDeliveryDateの全条件分岐を網羅するパラメタライズドテスト。
"""

import datetime

import pytest

from nouki_kaitou.delivery_calc import (
    _is_same_customer,
    calculate_delivery_date,
    extract_arrival_date_from_internal,
    extract_pickup_date,
)
from nouki_kaitou.models import BranchSettings, CacheStore, HolidayMap, OrderRow


# ============================================
# テスト用ヘルパー
# ============================================
TODAY = datetime.date(2026, 2, 16)  # 月曜日
BRANCH = BranchSettings(
    name="京葉営業所",
    default_cutoff=15,
    base_center="関東商品センター",
)
EXEC_TIME = datetime.datetime(2026, 2, 16, 12, 0, 0)  # 12時実行


def _holidays() -> HolidayMap:
    """テスト用祝日辞書"""
    return {
        datetime.date(2026, 2, 11): None,  # 建国記念日（水）
        datetime.date(2026, 2, 23): None,  # 天皇誕生日（月）
        datetime.date(2026, 3, 21): None,  # 春分の日（土）
    }


def _make_row(**kwargs) -> OrderRow:
    """テスト用OrderRow生成"""
    return OrderRow(**kwargs)


def _make_cache(
    mfg_days: dict | None = None,
    cust_days: dict | None = None,
    cust_route: dict | None = None,
    confirm: dict | None = None,
    storage: dict | None = None,
) -> CacheStore:
    """テスト用CacheStore生成"""
    cache = CacheStore()
    if mfg_days:
        cache.mfg_days = mfg_days
    if cust_days:
        cache.cust_days = cust_days
    if cust_route:
        cache.cust_route = cust_route
    if confirm:
        cache.confirm = confirm
    if storage:
        cache.storage = storage
    return cache


# ============================================
# extract_pickup_date
# ============================================
class TestExtractPickupDate:
    def test_basic(self):
        result = extract_pickup_date("引取 3/15", TODAY)
        assert result == datetime.date(2026, 3, 15)

    def test_hikitori_kanji(self):
        result = extract_pickup_date("引き取り 3/20", TODAY)
        assert result == datetime.date(2026, 3, 20)

    def test_full_width_slash(self):
        result = extract_pickup_date("引取 ３／１５", TODAY)
        assert result == datetime.date(2026, 3, 15)

    def test_no_keyword(self):
        assert extract_pickup_date("通常コメント 3/15", TODAY) is None

    def test_no_date(self):
        assert extract_pickup_date("引取", TODAY) is None

    def test_empty(self):
        assert extract_pickup_date("", TODAY) is None

    def test_year_rollover(self):
        """180日以上過去なら翌年"""
        result = extract_pickup_date("引取 7/1", datetime.date(2026, 2, 16))
        # 7/1は2026年7月1日 → 未来なのでそのまま
        assert result == datetime.date(2026, 7, 1)

    def test_date_at_end(self):
        """日付がコメント末尾"""
        result = extract_pickup_date("引取3/15", TODAY)
        assert result == datetime.date(2026, 3, 15)

    def test_invalid_date(self):
        assert extract_pickup_date("引取 13/32", TODAY) is None


# ============================================
# extract_arrival_date_from_internal
# ============================================
class TestExtractArrivalDateFromInternal:
    def test_basic(self):
        result = extract_arrival_date_from_internal("@@12/20", TODAY)
        assert result == datetime.date(2026, 12, 20)

    def test_full_width_at(self):
        result = extract_arrival_date_from_internal("＠＠12/20", TODAY)
        assert result == datetime.date(2026, 12, 20)

    def test_full_width_digits(self):
        result = extract_arrival_date_from_internal("@@１２／２０", TODAY)
        assert result == datetime.date(2026, 12, 20)

    def test_with_text_after(self):
        result = extract_arrival_date_from_internal("memo @@3/5 着指定", TODAY)
        assert result == datetime.date(2026, 3, 5)

    def test_no_marker(self):
        assert extract_arrival_date_from_internal("12/20着", TODAY) is None

    def test_no_date(self):
        assert extract_arrival_date_from_internal("@@着指定", TODAY) is None

    def test_empty(self):
        assert extract_arrival_date_from_internal("", TODAY) is None

    def test_year_rollover(self):
        """180日以上過去の日付は翌年"""
        today = datetime.date(2026, 12, 1)
        result = extract_arrival_date_from_internal("@@3/10", today)
        # 3/10は2026年3月10日、12月からは180日以上前 → 2027年
        assert result == datetime.date(2027, 3, 10)


# ============================================
# calculate_delivery_date - &&作業
# ============================================
class TestWorkOrder:
    def test_work_future(self):
        row = _make_row(
            comment_internal="&&",
            specified_delivery_date=datetime.date(2026, 3, 10),
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "3月10日作業予定"

    def test_work_past(self):
        row = _make_row(
            comment_internal="&&",
            specified_delivery_date=datetime.date(2026, 2, 10),
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "2月10日作業済み"

    def test_work_fullwidth(self):
        row = _make_row(
            comment_internal="＆＆",
            order_delivery_date=datetime.date(2026, 3, 5),
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "3月5日作業予定"

    def test_work_spec_date_1231_fallback(self):
        """指定納期が12/31なら受注納期にフォールバック"""
        row = _make_row(
            comment_internal="&&",
            specified_delivery_date=datetime.date(2026, 12, 31),
            order_delivery_date=datetime.date(2026, 4, 1),
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "4月1日作業予定"

    def test_work_no_date(self):
        row = _make_row(
            comment_internal="&&",
            specified_delivery_date=datetime.date(2026, 12, 31),
            order_delivery_date=datetime.date(2026, 12, 31),
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "日程調整中"


# ============================================
# calculate_delivery_date - @@着日指定
# ============================================
class TestArrivalDate:
    def test_arrival_future(self):
        row = _make_row(
            comment_internal="@@3/20",
            specified_delivery_date=datetime.date(2026, 3, 15),
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "3/15出荷→3/20着予定"

    def test_arrival_past(self):
        row = _make_row(
            comment_internal="@@2/14",
            specified_delivery_date=datetime.date(2026, 2, 12),
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "2/12出荷済→2/14着"


# ============================================
# calculate_delivery_date - 引取
# ============================================
class TestPickup:
    def test_pickup_future(self):
        row = _make_row(
            comment_external="引取 3/10",
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "3月10日引取予定"

    def test_pickup_past(self):
        row = _make_row(
            comment_internal="引取 2/10",
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "2月10日引取済み"

    def test_pickup_today(self):
        """引取当日は「引取予定」（< today ではないので）"""
        row = _make_row(
            comment_external="引取 2/16",
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "2月16日引取予定"


# ============================================
# calculate_delivery_date - 指定納期（在庫販売）
# ============================================
class TestSpecifiedDateStock:
    def test_stock_direct_transfer(self):
        """在庫販売 + 転送中 → 出荷予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            specified_delivery_date=datetime.date(2026, 3, 10),
            storage_place="転送中（直送用）",
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "3月10日出荷予定"

    def test_stock_use_ship_rule(self):
        """在庫販売 + 受注先≠出荷先 → 1営業日逆算して出荷予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            specified_delivery_date=datetime.date(2026, 3, 10),  # 火
            customer_name="顧客A",
            ship_to_name="出荷先B",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 3/10(火)の前営業日 = 3/9(月)
        assert result == "3月9日出荷予定"

    def test_stock_normal(self):
        """在庫販売 + 自社便 → 配達予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            specified_delivery_date=datetime.date(2026, 3, 10),
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "3月10日配達予定"


# ============================================
# calculate_delivery_date - 指定納期（直送販売）
# ============================================
class TestSpecifiedDateDirect:
    def test_direct_transfer(self):
        """直送 + 転送中 → そのまま出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            specified_delivery_date=datetime.date(2026, 3, 10),
            storage_place="転送中（直送用）",
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "3月10日出荷予定"

    def test_direct_with_days(self):
        """直送 + 紐付き → +2営業日 → 配達予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            specified_delivery_date=datetime.date(2026, 3, 10),  # 火
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(mfg_days={"D01": 2})
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 3/10 + 2営業日 = 3/12(木)
        assert result == "3月12日配達予定"

    def test_direct_with_delivery_days(self):
        """直送 + 曜日制限 → +2営業日→次の出荷曜日 → 出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            specified_delivery_date=datetime.date(2026, 3, 10),  # 火
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            mfg_days={"D01": 2},
            cust_days={"顧客A": [2, 6]},  # 月・金
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 3/10+2営業日=3/12(木) → 次の月or金 = 3/13(金)
        assert result == "3月13日出荷予定"

    def test_direct_rosenbin(self):
        """直送 + 路線便 + 曜日制限なし → daysToAdd-1営業日 → 出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            specified_delivery_date=datetime.date(2026, 3, 10),  # 火
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            mfg_days={"D01": 3},
            cust_route={"顧客A": True},
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 路線便 + 曜日制限なし: max(3-1,0)=2営業日 → 3/10+2営業日=3/12(木)
        assert result == "3月12日出荷予定"


# ============================================
# calculate_delivery_date - 在庫販売+処理完了
# ============================================
class TestStockCompleted:
    def test_before_cutoff(self):
        """在庫+処理完了+自社便+締切前 → 翌営業日配達予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 16),  # 月
            time_value="10:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 2/16(月) 10時 < 15時 → 翌営業日 = 2/17(火)
        assert result == "2月17日配達予定"

    def test_after_cutoff(self):
        """在庫+処理完了+自社便+締切後 → 翌々営業日配達予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 16),
            time_value="16:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 16時 >= 15時 → 2営業日後 = 2/18(水)
        assert result == "2月18日配達予定"

    def test_use_ship_rule(self):
        """在庫+処理完了+受注先≠出荷先 → 出荷日"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 16),
            time_value="16:00:00",
            customer_name="顧客A",
            ship_to_name="出荷先B",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 16時 >= 15時 → 出荷日=翌営業日=2/17(火)
        assert result == "2月17日出荷予定"

    def test_other_branch(self):
        """在庫+処理完了+他拠点在庫 → 他拠点より出荷予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 16),
            time_value="16:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関西商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 16時 >= 15時 → 出荷日=翌営業日=2/17
        assert result == "2月17日他拠点より出荷予定"

    def test_other_branch_with_use_ship_rule(self):
        """在庫+処理完了+他拠点在庫+受注先≠出荷先 → 他拠点が優先"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 13),
            time_value="10:38:05",
            customer_name="顧客A",
            ship_to_name="出荷先B",  # 受注先≠出荷先 → use_ship_rule
            storage_place="関西商品センター",  # ≠ base_center → 他拠点
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 他拠点チェックがuse_ship_ruleより優先される
        # 10時 < 15時 → 出荷日=登録日=2/13, 2/13 < today(2/16) → 済み
        assert result == "2月13日他拠点より出荷済み"

    def test_with_delivery_days(self):
        """在庫+処理完了+曜日制限 → 次の出荷曜日"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 16),  # 月
            time_value="10:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        cache = _make_cache(cust_days={"顧客A": [4, 6]})  # 水・金
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # base_date=2/16(月), 次の水or金 = 2/18(水)
        assert result == "2月18日出荷予定"

    def test_saturday_before_cutoff(self):
        """在庫+処理完了+土曜10時 → 月曜起算→火曜配達予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 21),  # 土
            time_value="10:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays,
            today=datetime.date(2026, 2, 21), branch=BRANCH,
        )
        # 土→翌営業日=2/23(月)…祝日→2/24(火) 起算、10時<15時 → +1営業日=2/25(水)
        assert result == "2月25日配達予定"

    def test_saturday_after_cutoff(self):
        """在庫+処理完了+土曜16時 → 月曜起算→翌々営業日配達予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 21),  # 土
            time_value="16:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays,
            today=datetime.date(2026, 2, 21), branch=BRANCH,
        )
        # 土→翌営業日=2/24(火) 起算、16時>=15時 → +2営業日=2/26(木)
        assert result == "2月26日配達予定"

    def test_sunday_before_cutoff(self):
        """在庫+処理完了+日曜10時 → 月曜起算→火曜配達予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 22),  # 日
            time_value="10:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays,
            today=datetime.date(2026, 2, 22), branch=BRANCH,
        )
        # 日→翌営業日=2/23(月)…祝日→2/24(火) 起算、10時<15時 → +1営業日=2/25(水)
        assert result == "2月25日配達予定"

    def test_holiday_before_cutoff(self):
        """在庫+処理完了+祝日10時 → 翌営業日起算→配達予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 23),  # 月・天皇誕生日
            time_value="10:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays,
            today=datetime.date(2026, 2, 23), branch=BRANCH,
        )
        # 祝日→翌営業日=2/24(火) 起算、10時<15時 → +1営業日=2/25(水)
        assert result == "2月25日配達予定"

    def test_saturday_use_ship_rule(self):
        """在庫+処理完了+土曜+受注先≠出荷先 → 月曜起算→出荷予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 21),  # 土
            time_value="10:00:00",
            customer_name="顧客A",
            ship_to_name="出荷先B",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays,
            today=datetime.date(2026, 2, 21), branch=BRANCH,
        )
        # 土→翌営業日=2/24(火) 起算、10時<15時 → 出荷日=2/24(火)
        assert result == "2月24日出荷予定"

    def test_saturday_other_branch(self):
        """在庫+処理完了+土曜+他拠点 → 月曜起算→他拠点出荷予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 21),  # 土
            time_value="10:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関西商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays,
            today=datetime.date(2026, 2, 21), branch=BRANCH,
        )
        # 土→翌営業日=2/24(火) 起算、10時<15時 → 出荷日=2/24(火)
        assert result == "2月24日他拠点より出荷予定"

    def test_saturday_no_holiday(self):
        """在庫+処理完了+土曜（祝日なし） → 月曜起算→火曜配達予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 2, 14),  # 土
            time_value="10:00:00",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays,
            today=datetime.date(2026, 2, 14), branch=BRANCH,
        )
        # 土→翌営業日=2/16(月) 起算、10時<15時 → +1営業日=2/17(火)
        assert result == "2月17日配達予定"


# ============================================
# calculate_delivery_date - 紐付き+処理完了
# ============================================
class TestHimozukiCompleted:
    def test_before_cutoff_delivery(self):
        """紐付き+処理完了+自社便+12時実行 → 翌営業日配達予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            ship_status="処理完了",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        exec_time = datetime.datetime(2026, 2, 16, 12, 0, 0)
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays, today=TODAY,
            branch=BRANCH, execution_time=exec_time,
        )
        # 12時 < 15時 → 翌営業日=2/17(火) → 配達予定
        assert result == "2月17日配達予定"

    def test_after_cutoff_delivery(self):
        """紐付き+処理完了+自社便+17時実行 → 翌々営業日配達予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            ship_status="処理完了",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        exec_time = datetime.datetime(2026, 2, 16, 17, 0, 0)
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays, today=TODAY,
            branch=BRANCH, execution_time=exec_time,
        )
        # 17時 >= 15時 → 2営業日後=2/18(水) → 配達予定
        assert result == "2月18日配達予定"

    def test_diff_place(self):
        """紐付き+処理完了+受注先≠出荷先 → 1営業日逆算して出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            ship_status="処理完了",
            customer_name="顧客A",
            ship_to_name="出荷先B",
            storage_place="関東商品センター",
        )
        exec_time = datetime.datetime(2026, 2, 16, 17, 0, 0)
        holidays = _holidays()
        result = calculate_delivery_date(
            row, CacheStore(), holidays=holidays, today=TODAY,
            branch=BRANCH, execution_time=exec_time,
        )
        # 17時>=15時 → 2営業日後=2/18(水), 前営業日=2/17(火)
        assert result == "2月17日出荷予定"

    def test_rosenbin(self):
        """紐付き+処理完了+路線便 → 1営業日逆算して出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            ship_status="処理完了",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="関東商品センター",
        )
        cache = _make_cache(cust_route={"顧客A": True})
        exec_time = datetime.datetime(2026, 2, 16, 17, 0, 0)
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY,
            branch=BRANCH, execution_time=exec_time,
        )
        # 17時>=15時 → 2営業日後=2/18(水), 前営業日=2/17(火)
        assert result == "2月17日出荷予定"

    def test_transfer_not_himozuki(self):
        """直送販売+処理完了+転送中 → 紐付き処理完了パスに入らない"""
        row = _make_row(
            document_type="【受注】直送販売",
            ship_status="処理完了",
            customer_name="顧客A",
            ship_to_name="顧客A",
            storage_place="転送中（直送用）",
            order_delivery_date=datetime.date(2026, 3, 10),
            item_group_code="D01",
        )
        cache = _make_cache(mfg_days={"D01": 2})
        # 転送中なので紐付きではなく通常計算へ
        result = calculate_delivery_date(
            row, cache, today=TODAY, branch=BRANCH,
        )
        assert result == "3月10日出荷予定"


# ============================================
# calculate_delivery_date - 受注納期なし
# ============================================
class TestNoDeliveryDate:
    def test_no_dates(self):
        row = _make_row(document_type="【受注】在庫販売")
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "日程調整中"


# ============================================
# calculate_delivery_date - 12/31（未確定）
# ============================================
class TestDec31:
    def test_stock_1231(self):
        """在庫販売+12/31 → 日程調整中"""
        row = _make_row(
            document_type="【受注】在庫販売",
            order_delivery_date=datetime.date(2026, 12, 31),
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "日程調整中"

    def test_direct_1231(self):
        """直送販売+12/31 → 確認中"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 12, 31),
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "確認中"


# ============================================
# calculate_delivery_date - 確認中確定パス
# ============================================
class TestConfirmedDate:
    def test_confirmed_transfer(self):
        """確定+転送中 → そのまま出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 12, 31),
            order_number="GL2Z001",
            detail_number="10",
            storage_place="転送中（直送用）",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            confirm={"GL2Z001|10": ("済", "回答待ち", datetime.date(2026, 3, 10))},
        )
        result = calculate_delivery_date(
            row, cache, today=TODAY, branch=BRANCH
        )
        assert result == "3月10日出荷予定"

    def test_confirmed_original_ship_rule(self):
        """確定+在庫販売+受注先≠出荷先 → そのまま出荷予定"""
        row = _make_row(
            document_type="【受注】在庫販売",
            order_delivery_date=datetime.date(2026, 12, 31),
            order_number="GL2Z002",
            detail_number="10",
            customer_name="顧客A",
            ship_to_name="出荷先B",
            storage_place="関東商品センター",
            item_group_code="D01",
        )
        cache = _make_cache(
            confirm={"GL2Z002|10": ("済", "回答待ち", datetime.date(2026, 3, 10))},
        )
        result = calculate_delivery_date(
            row, cache, today=TODAY, branch=BRANCH
        )
        assert result == "3月10日出荷予定"

    def test_confirmed_rosenbin(self):
        """確定+路線便 → max(daysToAdd-1, 0)営業日 → 出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 12, 31),
            order_number="GL2Z003",
            detail_number="10",
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            mfg_days={"D01": 3},
            cust_route={"顧客A": True},
            confirm={"GL2Z003|10": ("済", "回答待ち", datetime.date(2026, 3, 10))},
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 3/10+max(3-1,0)=2営業日=3/12(木)
        assert result == "3月12日出荷予定"

    def test_confirmed_normal_delivery(self):
        """確定+通常 → +営業日 → 配達予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 12, 31),
            order_number="GL2Z004",
            detail_number="10",
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            mfg_days={"D01": 2},
            confirm={"GL2Z004|10": ("済", "回答待ち", datetime.date(2026, 3, 10))},
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 3/10+2営業日=3/12(木) → 配達予定
        assert result == "3月12日配達予定"

    def test_confirmed_with_delivery_days(self):
        """確定+曜日制限 → 出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 12, 31),
            order_number="GL2Z005",
            detail_number="10",
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            mfg_days={"D01": 2},
            cust_days={"顧客A": [2, 6]},  # 月・金
            confirm={"GL2Z005|10": ("済", "回答待ち", datetime.date(2026, 3, 10))},
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 3/10+2営業日=3/12(木) → 次の月or金=3/13(金)
        assert result == "3月13日出荷予定"

    def test_confirmed_himozuki_diff_place(self):
        """確定+紐付き+受注先≠出荷先 → 出荷予定（配達ではない）"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 12, 31),
            order_number="GL2Z006",
            detail_number="10",
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="出荷先B",
        )
        cache = _make_cache(
            mfg_days={"D01": 2},
            confirm={"GL2Z006|10": ("済", "回答待ち", datetime.date(2026, 3, 10))},
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 3/10+2営業日=3/12(木) → 出荷予定（紐付き+diff place）
        assert result == "3月12日出荷予定"


# ============================================
# calculate_delivery_date - 通常計算
# ============================================
class TestNormalCalc:
    def test_transfer_direct(self):
        """転送中 → そのまま出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 3, 10),
            storage_place="転送中（直送用）",
            item_group_code="D01",
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "3月10日出荷予定"

    def test_normal_delivery(self):
        """通常 + 曜日制限なし → +2営業日 → 配達予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 3, 10),  # 火
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(mfg_days={"D01": 2})
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 3/10+2営業日=3/12(木)
        assert result == "3月12日配達予定"

    def test_with_delivery_days(self):
        """通常 + 曜日制限あり → 出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 3, 10),
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            mfg_days={"D01": 2},
            cust_days={"顧客A": [2, 6]},  # 月・金
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 3/10+2営業日=3/12(木) → 次の月or金=3/13(金)
        assert result == "3月13日出荷予定"

    def test_rosenbin_with_delivery_days(self):
        """路線便 + 曜日制限あり → +営業日→曜日 → 出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 3, 10),
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            mfg_days={"D01": 2},
            cust_route={"顧客A": True},
            cust_days={"顧客A": [4]},  # 水のみ
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 路線便+曜日あり: base=3/10+2営業日=3/12(木) → 次の水=3/18(水)
        assert result == "3月18日出荷予定"

    def test_rosenbin_no_delivery_days(self):
        """路線便 + 曜日制限なし → max(daysToAdd-1,0)営業日 → 出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 3, 10),
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            mfg_days={"D01": 3},
            cust_route={"顧客A": True},
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 路線便+曜日なし: max(3-1,0)=2営業日 → 3/12(木)
        assert result == "3月12日出荷予定"

    def test_use_ship_rule_with_delivery_days(self):
        """受注先≠出荷先(在庫) + 曜日制限あり → base=deliveryDate"""
        row = _make_row(
            document_type="【受注】在庫販売",
            order_delivery_date=datetime.date(2026, 3, 10),  # 火
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="出荷先B",
        )
        cache = _make_cache(
            mfg_days={"D01": 2},
            cust_days={"顧客A": [4]},  # 水のみ
        )
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # originalUseShipRule=True → base=deliveryDate=3/10(火)
        # 次の水=3/11(水)
        assert result == "3月11日出荷予定"

    def test_past_date(self):
        """過去の日付 → 済み"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 2, 10),
            storage_place="転送中（直送用）",
        )
        result = calculate_delivery_date(
            row, CacheStore(), today=TODAY, branch=BRANCH
        )
        assert result == "2月10日出荷済み"

    def test_storage_fallback_to_cache(self):
        """保管場所が空 → キャッシュから取得"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 3, 10),
            storage_place="",
            order_number="GL2Z001",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(
            mfg_days={"D01": 2},
            storage={"GL2Z001": "転送中（直送用）"},
        )
        result = calculate_delivery_date(
            row, cache, today=TODAY, branch=BRANCH
        )
        # 保管場所=転送中（キャッシュから）→ 出荷予定
        assert result == "3月10日出荷予定"

    def test_himozuki_diff_ship_to(self):
        """紐付き + 受注先≠出荷先 + 曜日制限なし → 出荷予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 3, 10),  # 火
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="出荷先B",
        )
        cache = _make_cache(mfg_days={"D01": 2})
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 紐付き+受注先≠出荷先: 3/10+2営業日=3/12(木) → 1営業日前=3/11(水)出荷予定
        assert result == "3月11日出荷予定"

    def test_himozuki_same_ship_to(self):
        """紐付き + 受注先=出荷先 + 曜日制限なし → 配達予定（変更なし）"""
        row = _make_row(
            document_type="【受注】直送販売",
            order_delivery_date=datetime.date(2026, 3, 10),  # 火
            storage_place="関東商品センター",
            item_group_code="D01",
            customer_name="顧客A",
            ship_to_name="顧客A",
        )
        cache = _make_cache(mfg_days={"D01": 2})
        holidays = _holidays()
        result = calculate_delivery_date(
            row, cache, holidays=holidays, today=TODAY, branch=BRANCH
        )
        # 紐付き+受注先=出荷先: 3/10+2営業日=3/12(木) → 配達予定
        assert result == "3月12日配達予定"


# ============================================
# _is_same_customer（全角/半角正規化）
# ============================================
class TestIsSameCustomer:
    def test_identical(self):
        assert _is_same_customer("テスト商事", "テスト商事") is True

    def test_fullwidth_vs_halfwidth_parens(self):
        """（有）vs (有) - SAPで実際に発生するケース"""
        assert _is_same_customer("（有）三橋機工", "(有)三橋機工") is True

    def test_fullwidth_vs_halfwidth_kabu(self):
        """（株）vs (株)"""
        assert _is_same_customer("テスト（株）　本社", "テスト(株) 本社") is True

    def test_different_names(self):
        assert _is_same_customer("テスト商事", "別会社") is False

    def test_strip_whitespace(self):
        assert _is_same_customer("  テスト商事  ", "テスト商事") is True

    def test_fullwidth_space(self):
        """全角スペースと半角スペースの揺れ"""
        assert _is_same_customer("本多酸素（株）　八潮営業所", "本多酸素(株) 八潮営業所") is True


# ============================================
# 受注先=出荷先の全角/半角正規化が納期計算に反映されるテスト
# ============================================
class TestCustomerNameNormalizationInCalc:
    def test_himozuki_completed_same_customer_fullwidth_halfwidth(self):
        """紐付き+処理完了: （有）vs (有) が同一顧客と判定され配達予定になる"""
        row = _make_row(
            document_type="【受注】直送販売",
            ship_status="処理完了",
            storage_place="関東商品センター",
            customer_name="（有）三橋機工",
            ship_to_name="(有)三橋機工",
            specified_delivery_date=datetime.date(2026, 12, 31),
            order_delivery_date=datetime.date(2026, 12, 31),
        )
        result = calculate_delivery_date(
            row, CacheStore(), _holidays(), BRANCH,
            execution_time=datetime.datetime(2026, 2, 16, 12, 0, 0),
            today=TODAY,
        )
        # 受注先=出荷先（正規化後）→ 配達予定（出荷予定ではない）
        assert "配達予定" in result

    def test_himozuki_completed_different_customer(self):
        """紐付き+処理完了: 本当に受注先≠出荷先なら出荷（配達ではない）"""
        row = _make_row(
            document_type="【受注】直送販売",
            ship_status="処理完了",
            storage_place="関東商品センター",
            customer_name="（有）三橋機工",
            ship_to_name="別会社工業",
            specified_delivery_date=datetime.date(2026, 12, 31),
            order_delivery_date=datetime.date(2026, 12, 31),
        )
        result = calculate_delivery_date(
            row, CacheStore(), _holidays(), BRANCH,
            execution_time=datetime.datetime(2026, 2, 16, 12, 0, 0),
            today=TODAY,
        )
        # 受注先≠出荷先 → 出荷（配達ではない）
        assert "出荷" in result
        assert "配達" not in result

    def test_specified_date_same_customer_fullwidth(self):
        """指定納期あり+紐付き: 全角/半角が同一顧客なら配達予定"""
        row = _make_row(
            document_type="【受注】直送販売",
            ship_status="未処理",
            storage_place="関東商品センター",
            customer_name="（有）三橋機工",
            ship_to_name="(有)三橋機工",
            specified_delivery_date=datetime.date(2026, 3, 10),
            order_delivery_date=datetime.date(2026, 12, 31),
            item_group_code="D01",
        )
        cache = _make_cache(mfg_days={"D01": 1})
        result = calculate_delivery_date(
            row, cache, _holidays(), BRANCH, today=TODAY,
        )
        # 受注先=出荷先 → 配達予定
        assert "配達予定" in result

    def test_stock_sale_same_customer_fullwidth(self):
        """在庫販売: 全角/半角が同一顧客ならuse_ship_rule=False→配達"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="未処理",
            storage_place="関東商品センター",
            customer_name="（有）三橋機工",
            ship_to_name="(有)三橋機工",
            specified_delivery_date=datetime.date(2026, 3, 10),
        )
        result = calculate_delivery_date(
            row, CacheStore(), _holidays(), BRANCH, today=TODAY,
        )
        # 受注先=出荷先（正規化後）→ 配達予定
        assert "配達予定" in result
