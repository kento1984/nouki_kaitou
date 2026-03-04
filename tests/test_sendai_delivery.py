"""sendai_delivery.py のユニットテスト"""

import datetime

import pytest

from nouki_kaitou.models import (
    BranchSettings,
    CacheStore,
    DeliveryPattern,
    OrderRow,
)
from nouki_kaitou.sendai_delivery import (
    _before_cutoff_hm,
    _calc_pattern_days,
    _calc_pattern_period,
    check_sendai_himozuki_completed,
    check_sendai_stock_completed,
)
from nouki_kaitou.utils import format_date_japanese


# ============================================
# テスト共通定数
# ============================================
SENDAI_BRANCH = BranchSettings(
    name="仙台営業所",
    default_cutoff=15,
    base_center="東北商品センター",
)

KEIYO_BRANCH = BranchSettings(
    name="京葉営業所",
    default_cutoff=15,
    base_center="関東商品センター",
)

# パターン定義
KINRIN_2BIN = DeliveryPattern(
    name="近隣2便",
    cutoff1=(11, 30),
    days_before_cutoff1=0,
    cutoff2=(16, 0),
    days_between_cutoffs=1,
)

ENPO_GOZEN = DeliveryPattern(
    name="遠方午前",
    cutoff1=(16, 0),
    days_before_cutoff1=1,
    cutoff2=None,
    days_between_cutoffs=1,
)

ENPO_GOGO = DeliveryPattern(
    name="遠方午後",
    cutoff1=(11, 30),
    days_before_cutoff1=0,
    cutoff2=None,
    days_between_cutoffs=1,
)

TODAY = datetime.date(2026, 2, 23)  # 月曜日


def _make_cache(pattern_name: str = "近隣2便") -> CacheStore:
    """テスト用CacheStoreを構築する。"""
    cache = CacheStore()
    cache.delivery_patterns = {
        "近隣2便": KINRIN_2BIN,
        "遠方午前": ENPO_GOZEN,
        "遠方午後": ENPO_GOGO,
    }
    cache.cust_pattern = {"テスト顧客": pattern_name} if pattern_name else {}
    cache.cust_days = {"テスト顧客": []}
    cache.cust_route = {"テスト顧客": False}
    return cache


def _make_row(
    *,
    time_value: str = "10:00:00",
    reg_date: datetime.date | None = None,
    ship_status: str = "処理完了",
    document_type: str = "【受注】在庫販売",
    customer_name: str = "テスト顧客",
    ship_to_name: str = "テスト顧客",
    storage_place: str = "東北商品センター",
) -> OrderRow:
    if reg_date is None:
        reg_date = TODAY
    return OrderRow(
        order_number="SD2Z000001",
        detail_number="10",
        document_type=document_type,
        customer_name=customer_name,
        ship_to_name=ship_to_name,
        ship_status=ship_status,
        storage_place=storage_place,
        registration_date=reg_date,
        time_value=time_value,
    )


# ============================================
# TestBeforeCutoffHM
# ============================================
class TestBeforeCutoffHM:
    def test_before(self):
        assert _before_cutoff_hm(10, 0, (11, 30)) is True

    def test_exactly_at(self):
        """ちょうど＝締切後"""
        assert _before_cutoff_hm(11, 30, (11, 30)) is False

    def test_after(self):
        assert _before_cutoff_hm(12, 0, (11, 30)) is False

    def test_same_hour_before_minute(self):
        assert _before_cutoff_hm(11, 29, (11, 30)) is True

    def test_same_hour_after_minute(self):
        assert _before_cutoff_hm(11, 31, (11, 30)) is False


# ============================================
# TestCalcPatternDays
# ============================================
class TestCalcPatternDays:
    def test_kinrin_before_cutoff1(self):
        """近隣2便: cutoff1(11:30)前 → 当日(0)"""
        assert _calc_pattern_days(10, 0, KINRIN_2BIN) == 0

    def test_kinrin_at_cutoff1(self):
        """近隣2便: cutoff1(11:30)ちょうど → cutoff1～cutoff2間(1)"""
        assert _calc_pattern_days(11, 30, KINRIN_2BIN) == 1

    def test_kinrin_between_cutoffs(self):
        """近隣2便: cutoff1～cutoff2間 → 翌日(1)"""
        assert _calc_pattern_days(14, 0, KINRIN_2BIN) == 1

    def test_kinrin_at_cutoff2(self):
        """近隣2便: cutoff2(16:00)ちょうど → 全超過(1)"""
        assert _calc_pattern_days(16, 0, KINRIN_2BIN) == 1

    def test_kinrin_after_cutoff2(self):
        """近隣2便: cutoff2(16:00)超過 → 全超過(1)"""
        assert _calc_pattern_days(17, 0, KINRIN_2BIN) == 1

    def test_enpo_gozen_before(self):
        """遠方午前: cutoff1(16:00)前 → 翌日(1)"""
        assert _calc_pattern_days(10, 0, ENPO_GOZEN) == 1

    def test_enpo_gozen_after(self):
        """遠方午前: cutoff1(16:00)超過 → 翌々日(2)"""
        assert _calc_pattern_days(16, 0, ENPO_GOZEN) == 2

    def test_enpo_gogo_before(self):
        """遠方午後: cutoff1(11:30)前 → 当日(0)"""
        assert _calc_pattern_days(10, 0, ENPO_GOGO) == 0

    def test_enpo_gogo_after(self):
        """遠方午後: cutoff1(11:30)超過 → 翌日(1)"""
        assert _calc_pattern_days(11, 30, ENPO_GOGO) == 1


# ============================================
# TestSendaiStockCompleted — 結合テスト
# ============================================
class TestSendaiStockCompleted:
    """check_sendai_stock_completedの結合テスト"""

    # --- 非仙台/非該当のフォールスルー ---

    def test_non_sendai_returns_none(self):
        """非仙台ブランチ → None"""
        row = _make_row()
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, KEIYO_BRANCH, TODAY,
            "関東商品センター", False,
        )
        assert result is None

    def test_sendai_no_pattern_returns_none(self):
        """仙台+パターンなし → None"""
        row = _make_row()
        cache = _make_cache("")  # パターンなし
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    def test_sendai_unknown_pattern_returns_none(self):
        """仙台+未定義パターン → None"""
        row = _make_row()
        cache = _make_cache("近隣2便")
        cache.cust_pattern["テスト顧客"] = "存在しないパターン"
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    def test_not_stock_sale_returns_none(self):
        """仙台+直送販売 → None"""
        row = _make_row(document_type="【受注】直送販売")
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    def test_not_completed_returns_none(self):
        """仙台+未処理 → None"""
        row = _make_row(ship_status="未処理")
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    # --- 近隣2便 ---

    def test_kinrin_before_cutoff1(self):
        """近隣2便 10:00 → 当日PM配達予定（当日便はまだ届いていない）"""
        row = _make_row(time_value="10:00:00")
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "PM配達予定"

    def test_kinrin_at_cutoff1(self):
        """近隣2便 11:30(ちょうど=超過) → 翌営業日AM配達予定"""
        row = _make_row(time_value="11:30:00")
        cache = _make_cache("近隣2便")
        # TODAY=2/23(月) → 翌営業日=2/24(火)
        expected_date = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected_date) + "AM配達予定"

    def test_kinrin_between_cutoffs(self):
        """近隣2便 14:00 → 翌営業日AM配達予定"""
        row = _make_row(time_value="14:00:00")
        cache = _make_cache("近隣2便")
        expected_date = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected_date) + "AM配達予定"

    def test_kinrin_at_cutoff2(self):
        """近隣2便 16:00(ちょうど=超過) → 翌営業日PM配達予定"""
        row = _make_row(time_value="16:00:00")
        cache = _make_cache("近隣2便")
        expected_date = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected_date) + "PM配達予定"

    def test_kinrin_after_cutoff2(self):
        """近隣2便 17:00 → 翌営業日PM配達予定"""
        row = _make_row(time_value="17:00:00")
        cache = _make_cache("近隣2便")
        expected_date = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected_date) + "PM配達予定"

    # --- 遠方午前 ---

    def test_enpo_gozen_before_cutoff(self):
        """遠方午前 10:00 → 翌営業日配達予定(+1)"""
        row = _make_row(time_value="10:00:00")
        cache = _make_cache("遠方午前")
        expected_date = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected_date) + "配達予定"

    def test_enpo_gozen_after_cutoff(self):
        """遠方午前 16:00 → 翌々営業日配達予定(+2)"""
        row = _make_row(time_value="16:00:00")
        cache = _make_cache("遠方午前")
        expected_date = datetime.date(2026, 2, 25)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected_date) + "配達予定"

    # --- 遠方午後 ---

    def test_enpo_gogo_before_cutoff(self):
        """遠方午後 10:00 → 当日配達予定(+0)"""
        row = _make_row(time_value="10:00:00")
        cache = _make_cache("遠方午後")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "配達予定"

    def test_enpo_gogo_at_cutoff(self):
        """遠方午後 11:30(ちょうど=超過) → 翌営業日配達予定(+1)"""
        row = _make_row(time_value="11:30:00")
        cache = _make_cache("遠方午後")
        expected_date = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected_date) + "配達予定"

    # --- 他拠点在庫 ---

    def test_other_branch_stock(self):
        """仙台+他拠点在庫 → 他拠点出荷"""
        row = _make_row(time_value="10:00:00", storage_place="関東商品センター")
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "関東商品センター", False,
        )
        assert "他拠点より" in result

    def test_other_branch_stock_before_cutoff(self):
        """他拠点: branch.default_cutoff(15時)前 → 当日出荷"""
        row = _make_row(time_value="10:00:00", storage_place="関東商品センター")
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "関東商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "他拠点より出荷済み"

    def test_other_branch_stock_after_branch_cutoff(self):
        """他拠点: branch.default_cutoff(15時)後 → 翌営業日出荷"""
        row = _make_row(time_value="15:00:00", storage_place="関東商品センター")
        cache = _make_cache("近隣2便")
        expected_date = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "関東商品センター", False,
        )
        assert result == format_date_japanese(expected_date) + "他拠点より出荷予定"

    def test_other_branch_stock_between_pattern_and_branch_cutoff(self):
        """他拠点: pattern.cutoff1(11:30)後だがbranch(15時)前 → 当日出荷
        出荷可否はセンターの締切で判定するため、配達ルートの11:30は無関係"""
        row = _make_row(time_value="14:00:00", storage_place="関東商品センター")
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "関東商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "他拠点より出荷済み"

    # --- 受注先≠出荷先 ---

    def test_ship_rule(self):
        """仙台+受注先≠出荷先: branch.default_cutoff(15時)前 → 当日出荷"""
        row = _make_row(time_value="10:00:00")
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", True,  # use_ship_rule=True
        )
        assert result == format_date_japanese(TODAY) + "出荷済み"

    def test_ship_rule_after_branch_cutoff(self):
        """受注先≠出荷先: branch.default_cutoff(15時)後 → 翌営業日出荷"""
        row = _make_row(time_value="15:00:00")
        cache = _make_cache("近隣2便")
        expected_date = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", True,
        )
        assert result == format_date_japanese(expected_date) + "出荷予定"

    def test_ship_rule_between_pattern_and_branch_cutoff(self):
        """受注先≠出荷先: pattern.cutoff1(11:30)後だがbranch(15時)前 → 当日出荷"""
        row = _make_row(time_value="14:00:00")
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", True,
        )
        assert result == format_date_japanese(TODAY) + "出荷済み"

    # --- 曜日制限 ---

    def test_delivery_days_restriction(self):
        """仙台+曜日制限 → 次の配送日で出荷予定"""
        row = _make_row(time_value="10:00:00")
        cache = _make_cache("近隣2便")
        # 月水金配達 → 当日(月)は配達曜日 → 当日出荷
        cache.cust_days["テスト顧客"] = [2, 4, 6]  # 月水金
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert "出荷" in result

    def test_delivery_days_next_day(self):
        """曜日制限: 当日が配送曜日でない → 次の配送日"""
        row = _make_row(time_value="10:00:00")
        cache = _make_cache("近隣2便")
        # 火木のみ → 月曜は配送日でない → 火曜(2/24)になる
        cache.cust_days["テスト顧客"] = [3, 5]  # 火木
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        expected_date = datetime.date(2026, 2, 24)  # 火曜日
        assert result == format_date_japanese(expected_date) + "出荷予定"

    # --- 土曜登録 ---

    def test_saturday_registration(self):
        """土曜登録 → 翌営業日起算"""
        saturday = datetime.date(2026, 2, 21)  # 土曜日
        row = _make_row(time_value="10:00:00", reg_date=saturday)
        cache = _make_cache("近隣2便")
        # 土曜→翌営業日(月曜2/23)が起算、10:00<11:30 → 当日PM配達予定
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "PM配達予定"

    # --- 祝日登録 ---

    def test_holiday_registration(self):
        """祝日登録 → 翌営業日起算"""
        holiday_date = datetime.date(2026, 2, 23)  # 祝日扱い
        holidays = {holiday_date: None}
        row = _make_row(time_value="10:00:00", reg_date=holiday_date)
        cache = _make_cache("近隣2便")
        # 祝日→翌営業日(2/24火曜)が起算、10:00<11:30 → 当日PM配達
        expected_date = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, holidays, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected_date) + "PM配達予定"

    # --- 登録日/時刻なし ---

    def test_no_registration_date(self):
        """登録日なし → None"""
        row = _make_row(reg_date=None)
        row.registration_date = None
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    def test_no_time_value(self):
        """時刻なし → None"""
        row = _make_row(time_value="")
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    # --- 当日配達(biz_days==0)の過去日 ---

    def test_kinrin_biz0_past_date(self):
        """近隣2便 biz_days=0 + 登録日が過去 → PM配達済み"""
        past_date = datetime.date(2026, 2, 20)  # 金曜日（過去）
        row = _make_row(time_value="10:00:00", reg_date=past_date)
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(past_date) + "PM配達済み"

    def test_enpo_gogo_biz0_past_date(self):
        """遠方午後 biz_days=0 + 登録日が過去 → 配達済み"""
        past_date = datetime.date(2026, 2, 19)  # 木曜日（過去）
        row = _make_row(time_value="10:00:00", reg_date=past_date)
        cache = _make_cache("遠方午後")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(past_date) + "配達済み"

    def test_kinrin_biz0_today_still_yotei(self):
        """近隣2便 biz_days=0 + 登録日が今日 → PM配達予定（変わらず）"""
        row = _make_row(time_value="10:00:00", reg_date=TODAY)
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "PM配達予定"


# ============================================
# TestSendaiHimozukiCompleted — 紐付き+処理完了 結合テスト
# ============================================
def _make_himozuki_row(
    *,
    customer_name: str = "テスト顧客",
    ship_to_name: str = "テスト顧客",
    ship_status: str = "処理完了",
    storage_place: str = "東北商品センター",
) -> OrderRow:
    """紐付き用テスト行を作成する。"""
    return OrderRow(
        order_number="SD2Z000002",
        detail_number="10",
        document_type="【受注】直送販売",
        customer_name=customer_name,
        ship_to_name=ship_to_name,
        ship_status=ship_status,
        storage_place=storage_place,
        registration_date=TODAY,
        time_value="10:00:00",
    )


class TestSendaiHimozukiCompleted:
    """check_sendai_himozuki_completedの結合テスト"""

    # --- 非仙台/非該当のフォールスルー ---

    def test_non_sendai_returns_none(self):
        """非仙台ブランチ → None"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, KEIYO_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    def test_not_chokuso_returns_none(self):
        """在庫販売 → None"""
        row = _make_himozuki_row()
        row.document_type = "【受注】在庫販売"
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    def test_not_completed_returns_none(self):
        """未処理 → None"""
        row = _make_himozuki_row(ship_status="未処理")
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    def test_tensouchu_returns_none(self):
        """転送中（直送用）→ None（直送扱い、紐付きでない）"""
        row = _make_himozuki_row(storage_place="転送中（直送用）")
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "転送中（直送用）", False,
        )
        assert result is None

    def test_no_pattern_returns_none(self):
        """パターンなし → None"""
        row = _make_himozuki_row()
        cache = _make_cache("")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result is None

    # --- 近隣2便: execution_timeベース ---

    def test_kinrin_before_cutoff1(self):
        """近隣2便 実行10:00 → 当日PM配達予定(+0)"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "PM配達予定"

    def test_kinrin_at_cutoff1(self):
        """近隣2便 実行11:30(ちょうど=超過) → 翌営業日AM配達予定(+1)"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 11, 30)
        expected = datetime.date(2026, 2, 24)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "AM配達予定"

    def test_kinrin_between_cutoffs(self):
        """近隣2便 実行14:00 → 翌営業日AM配達予定(+1)"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 14, 0)
        expected = datetime.date(2026, 2, 24)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "AM配達予定"

    def test_kinrin_at_cutoff2(self):
        """近隣2便 実行16:00(ちょうど=超過) → 翌営業日PM配達予定(+1)"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 16, 0)
        expected = datetime.date(2026, 2, 24)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "PM配達予定"

    def test_kinrin_after_cutoff2(self):
        """近隣2便 実行17:00 → 翌営業日PM配達予定(+1)"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 17, 0)
        expected = datetime.date(2026, 2, 24)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "PM配達予定"

    # --- 遠方午前 ---

    def test_enpo_gozen_before(self):
        """遠方午前 実行10:00 → 翌営業日配達予定(+1)"""
        row = _make_himozuki_row()
        cache = _make_cache("遠方午前")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        expected = datetime.date(2026, 2, 24)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "配達予定"

    def test_enpo_gozen_after(self):
        """遠方午前 実行16:00 → 翌々営業日配達予定(+2)"""
        row = _make_himozuki_row()
        cache = _make_cache("遠方午前")
        exec_time = datetime.datetime(2026, 2, 23, 16, 0)
        expected = datetime.date(2026, 2, 25)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "配達予定"

    # --- 遠方午後 ---

    def test_enpo_gogo_before(self):
        """遠方午後 実行10:00 → 当日配達予定(+0)"""
        row = _make_himozuki_row()
        cache = _make_cache("遠方午後")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "配達予定"

    def test_enpo_gogo_at_cutoff(self):
        """遠方午後 実行11:30(ちょうど=超過) → 翌営業日配達予定(+1)"""
        row = _make_himozuki_row()
        cache = _make_cache("遠方午後")
        exec_time = datetime.datetime(2026, 2, 23, 11, 30)
        expected = datetime.date(2026, 2, 24)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "配達予定"

    # --- 受注先≠出荷先 ---

    def test_diff_customer_ship(self):
        """受注先≠出荷先 → 配達日から逆算して出荷予定"""
        row = _make_himozuki_row(
            customer_name="テスト顧客",
            ship_to_name="別の出荷先",
        )
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        # 近隣2便 10:00 → biz_days=0 → adjusted=TODAY
        # 受注先≠出荷先 → previous_business_day(TODAY=2/23月) = 2/20(金)
        expected = datetime.date(2026, 2, 20)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "出荷済み"

    def test_diff_customer_ship_after_cutoff(self):
        """受注先≠出荷先 + cutoff後 → 翌営業日から逆算"""
        row = _make_himozuki_row(
            customer_name="テスト顧客",
            ship_to_name="別の出荷先",
        )
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 14, 0)
        # 近隣2便 14:00 → biz_days=1 → adjusted=2/24(火)
        # previous_business_day(2/24) = 2/23(月)
        expected = datetime.date(2026, 2, 23)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "出荷済み"

    # --- 曜日制限 ---

    def test_delivery_days(self):
        """曜日制限あり → 次の配送日で出荷予定"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        cache.cust_days["テスト顧客"] = [3, 5]  # 火木
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        # 近隣2便 10:00 → biz_days=0 → adjusted=TODAY(月)
        # 月曜は配送日でない → 火曜(2/24)
        expected = datetime.date(2026, 2, 24)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "出荷予定"

    def test_delivery_days_after_cutoff(self):
        """曜日制限 + cutoff後 → 翌日から次の配送日"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        cache.cust_days["テスト顧客"] = [4, 6]  # 水金
        exec_time = datetime.datetime(2026, 2, 23, 14, 0)
        # 近隣2便 14:00 → biz_days=1 → adjusted=2/24(火)
        # 火曜は配送日でない → 水曜(2/25)
        expected = datetime.date(2026, 2, 25)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "出荷予定"

    # --- 路線便 ---

    def test_rosenbin(self):
        """路線便 → 配達日から逆算して出荷予定"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 14, 0)
        # 近隣2便 14:00 → biz_days=1 → adjusted=2/24(火)
        # 路線便 → previous_business_day(2/24) = 2/23(月)
        expected = datetime.date(2026, 2, 23)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", True,  # is_rosenbin=True
        )
        assert result == format_date_japanese(expected) + "出荷済み"

    # --- 当日配達(biz_days==0)の確認 ---

    def test_kinrin_biz0_today_still_yotei(self):
        """近隣2便 biz_days=0 + adjusted==today → PM配達予定（変わらず）"""
        row = _make_himozuki_row()
        cache = _make_cache("近隣2便")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "PM配達予定"

    # --- 全角/半角の顧客名比較 ---

    def test_fullwidth_halfwidth_same_customer(self):
        """全角半角が揺れても受注先=出荷先と判定 → PM配達予定"""
        row = _make_himozuki_row(
            customer_name="（有）三橋機工",
            ship_to_name="(有)三橋機工",
        )
        cache = _make_cache("近隣2便")
        cache.cust_pattern["（有）三橋機工"] = "近隣2便"
        cache.cust_days["（有）三橋機工"] = []
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "PM配達予定"


# ============================================
# TestCalcPatternPeriod
# ============================================
class TestCalcPatternPeriod:
    """_calc_pattern_periodのユニットテスト"""

    def test_kinrin_before_cutoff1(self):
        """近隣2便: cutoff1前 → PM"""
        assert _calc_pattern_period(10, 0, KINRIN_2BIN) == "PM"

    def test_kinrin_at_cutoff1(self):
        """近隣2便: cutoff1ちょうど → AM（cutoff1～cutoff2間）"""
        assert _calc_pattern_period(11, 30, KINRIN_2BIN) == "AM"

    def test_kinrin_between_cutoffs(self):
        """近隣2便: cutoff1～cutoff2間 → AM"""
        assert _calc_pattern_period(14, 0, KINRIN_2BIN) == "AM"

    def test_kinrin_at_cutoff2(self):
        """近隣2便: cutoff2ちょうど → PM（cutoff2以降）"""
        assert _calc_pattern_period(16, 0, KINRIN_2BIN) == "PM"

    def test_kinrin_after_cutoff2(self):
        """近隣2便: cutoff2以降 → PM"""
        assert _calc_pattern_period(17, 0, KINRIN_2BIN) == "PM"

    def test_enpo_gozen_before(self):
        """遠方午前: cutoff1前 → 空文字（表示なし）"""
        assert _calc_pattern_period(10, 0, ENPO_GOZEN) == ""

    def test_enpo_gozen_after(self):
        """遠方午前: cutoff1以降 → 空文字（1段階パターンは常に空）"""
        assert _calc_pattern_period(16, 0, ENPO_GOZEN) == ""

    def test_enpo_gogo_before(self):
        """遠方午後: cutoff1前 → 空文字"""
        assert _calc_pattern_period(10, 0, ENPO_GOGO) == ""

    def test_enpo_gogo_after(self):
        """遠方午後: cutoff1以降 → 空文字"""
        assert _calc_pattern_period(11, 30, ENPO_GOGO) == ""


# ============================================
# 遠方パターン回帰テスト（AM/PM表示なし確認）
# ============================================
class TestEnpoNoPeriod:
    """遠方午前・遠方午後でAM/PM表示がないことを確認する回帰テスト"""

    def test_enpo_gozen_stock_no_period(self):
        """遠方午前 在庫販売 → AM/PMなし"""
        row = _make_row(time_value="10:00:00")
        cache = _make_cache("遠方午前")
        expected = datetime.date(2026, 2, 24)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "配達予定"
        assert "AM" not in result
        assert "PM" not in result

    def test_enpo_gogo_stock_no_period(self):
        """遠方午後 在庫販売 → AM/PMなし"""
        row = _make_row(time_value="10:00:00")
        cache = _make_cache("遠方午後")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "配達予定"
        assert "AM" not in result
        assert "PM" not in result

    def test_enpo_gozen_himozuki_no_period(self):
        """遠方午前 紐付き → AM/PMなし"""
        row = _make_himozuki_row()
        cache = _make_cache("遠方午前")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        expected = datetime.date(2026, 2, 24)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "配達予定"
        assert "AM" not in result
        assert "PM" not in result

    def test_enpo_gogo_himozuki_no_period(self):
        """遠方午後 紐付き → AM/PMなし"""
        row = _make_himozuki_row()
        cache = _make_cache("遠方午後")
        exec_time = datetime.datetime(2026, 2, 23, 10, 0)
        result = check_sendai_himozuki_completed(
            row, cache, {}, SENDAI_BRANCH, exec_time, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(TODAY) + "配達予定"
        assert "AM" not in result
        assert "PM" not in result


# ============================================
# 過去日のAM/PMテスト
# ============================================
class TestPastDateWithPeriod:
    """配達済みでもAM/PMが付くことを確認するテスト"""

    def test_stock_kinrin_past_pm(self):
        """在庫販売 近隣2便 過去日 10:00 → PM配達済み"""
        past_date = datetime.date(2026, 2, 20)  # 金曜日（過去）
        row = _make_row(time_value="10:00:00", reg_date=past_date)
        cache = _make_cache("近隣2便")
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(past_date) + "PM配達済み"

    def test_stock_kinrin_past_am(self):
        """在庫販売 近隣2便 過去日 14:00 → AM配達済み（翌営業日だが過去）"""
        past_date = datetime.date(2026, 2, 19)  # 木曜日（過去）
        row = _make_row(time_value="14:00:00", reg_date=past_date)
        cache = _make_cache("近隣2便")
        # 14:00 → cutoff1～cutoff2間 → biz_days=1 → 2/20(金)、AM
        expected = datetime.date(2026, 2, 20)
        result = check_sendai_stock_completed(
            row, cache, {}, SENDAI_BRANCH, TODAY,
            "東北商品センター", False,
        )
        assert result == format_date_japanese(expected) + "AM配達済み"
