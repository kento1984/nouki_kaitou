"""report_generator モジュールのテスト"""

import datetime
import tempfile
from pathlib import Path

import pytest
from openpyxl import load_workbook

from nouki_kaitou.models import (
    BranchSettings,
    BunnoEntry,
    CacheStore,
    ConfirmingRecord,
    HolidayMap,
    HistoryRecord,
    OrderRow,
    ReportResult,
    ReportRow,
    StockoutEntry,
    TrackingEntry,
)
from nouki_kaitou.report_generator import (
    build_report_row,
    create_delivery_report,
    create_delivery_report_by_order_numbers,
    group_order_numbers_by_customer,
    _resolve_manufacturer_name,
    _resolve_product_name,
    _resolve_delivery_place,
    _resolve_price,
    _pass_basic_filter,
    _is_excluded,
    _determine_flags,
    _classify_order,
)


# ============================================
# テスト用ヘルパー
# ============================================
TODAY = datetime.date(2026, 2, 16)
EXEC_TIME = datetime.datetime(2026, 2, 16, 10, 0, 0)
BRANCH = BranchSettings(
    name="京葉営業所",
    default_cutoff=15,
    base_center="関東商品センター",
)


def _make_cache(**kwargs) -> CacheStore:
    """テスト用キャッシュを作成する。"""
    return CacheStore(
        mfg_name=kwargs.get("mfg_name", {"D01": "ダイヘン", "Z99": ""}),
        mfg_days=kwargs.get("mfg_days", {"D01": 2}),
        cust_days=kwargs.get("cust_days", {}),
        cust_retention=kwargs.get("cust_retention", {}),
        cust_route=kwargs.get("cust_route", {}),
        confirm=kwargs.get("confirm", {}),
        storage=kwargs.get("storage", {}),
    )


def _make_row(**kwargs) -> OrderRow:
    """テスト用OrderRowを作成する。"""
    defaults = {
        "order_number": "1000001",
        "detail_number": "10",
        "document_type": "【受注】在庫販売",
        "customer_name": "テスト商事",
        "product_name": "溶接棒 3.2mm",
        "ship_status": "未処理",
        "quantity": "100",
        "unit_price": "500",
        "net_amount": "50000",
        "manufacturer_name": "ダイヘン",
        "storage_place": "関東商品センター",
        "customer_order_number": "PO-001",
        "customer_contact": "田中",
        "comment_detail": "",
        "comment_external": "",
        "comment_internal": "",
        "rejection_reason": "",
        "ship_to_name": "テスト商事",
        "registration_date": TODAY,
        "time_value": "10:00:00",
        "order_delivery_date": datetime.date(2026, 2, 20),
        "specified_delivery_date": None,
        "item_group_code": "D01",
    }
    defaults.update(kwargs)
    return OrderRow(**defaults)


# ============================================
# group_order_numbers_by_customer
# ============================================
class TestGroupOrderNumbersByCustomer:
    """注番の顧客別グループ化テスト"""

    def test_basic_grouping(self):
        data = [
            _make_row(order_number="100", customer_name="A社"),
            _make_row(order_number="200", customer_name="B社"),
            _make_row(order_number="300", customer_name="A社"),
        ]
        result = group_order_numbers_by_customer(data, ["100", "200", "300"])
        assert result == {"A社": ["100", "300"], "B社": ["200"]}

    def test_dedup(self):
        data = [
            _make_row(order_number="100", customer_name="A社"),
        ]
        result = group_order_numbers_by_customer(data, ["100", "100"])
        assert result == {"A社": ["100"]}

    def test_not_found(self):
        data = [
            _make_row(order_number="100", customer_name="A社"),
        ]
        result = group_order_numbers_by_customer(data, ["999"])
        assert result == {}

    def test_empty(self):
        result = group_order_numbers_by_customer([], ["100"])
        assert result == {}


# ============================================
# _resolve_manufacturer_name
# ============================================
class TestResolveManufacturerName:
    """メーカー名解決テスト"""

    def test_normal(self):
        row = _make_row(item_group_code="D01")
        cache = _make_cache()
        assert _resolve_manufacturer_name(row, cache) == "ダイヘン"

    def test_z99_with_space(self):
        row = _make_row(
            item_group_code="Z99",
            product_name="ABC商事 特殊溶接棒"
        )
        cache = _make_cache()
        assert _resolve_manufacturer_name(row, cache) == "ABC商事"

    def test_z99_full_width_space(self):
        row = _make_row(
            item_group_code="Z99",
            product_name="ABC商事\u3000特殊溶接棒"
        )
        cache = _make_cache()
        assert _resolve_manufacturer_name(row, cache) == "ABC商事"

    def test_z99_no_space(self):
        row = _make_row(
            item_group_code="Z99",
            product_name="特殊溶接棒"
        )
        cache = _make_cache()
        assert _resolve_manufacturer_name(row, cache) == ""

    def test_fallback_to_manufacturer_field(self):
        row = _make_row(
            item_group_code="X01",
            manufacturer_name="フォールバック社"
        )
        cache = _make_cache(mfg_name={})
        assert _resolve_manufacturer_name(row, cache) == "フォールバック社"

    def test_fallback_to_code(self):
        row = _make_row(
            item_group_code="X01",
            manufacturer_name=""
        )
        cache = _make_cache(mfg_name={})
        assert _resolve_manufacturer_name(row, cache) == "X01"


# ============================================
# _resolve_product_name
# ============================================
class TestResolveProductName:
    """品名解決テスト"""

    def test_normal(self):
        row = _make_row(item_group_code="D01", product_name="溶接棒 3.2mm")
        assert _resolve_product_name(row, "ダイヘン") == "溶接棒 3.2mm"

    def test_z99(self):
        row = _make_row(
            item_group_code="Z99",
            product_name="ABC商事 特殊溶接棒"
        )
        assert _resolve_product_name(row, "ABC商事") == "特殊溶接棒"

    def test_z97_full_width_space(self):
        row = _make_row(
            item_group_code="Z97",
            product_name="XYZ\u3000特殊品"
        )
        assert _resolve_product_name(row, "XYZ") == "特殊品"


# ============================================
# _resolve_delivery_place
# ============================================
class TestResolveDeliveryPlace:
    """納入先名解決テスト"""

    def test_same_as_customer(self):
        row = _make_row(customer_name="テスト商事", ship_to_name="テスト商事")
        assert _resolve_delivery_place(row, TODAY) == "貴社"

    def test_different_no_sama(self):
        row = _make_row(ship_to_name="東京工場")
        assert _resolve_delivery_place(row, TODAY) == "東京工場様"

    def test_different_with_sama(self):
        row = _make_row(ship_to_name="東京工場様")
        assert _resolve_delivery_place(row, TODAY) == "東京工場様"

    def test_pickup(self):
        row = _make_row(
            comment_external="引取 2/20",
            ship_to_name="テスト商事"
        )
        assert _resolve_delivery_place(row, TODAY) == "お引き取り"


# ============================================
# _resolve_price
# ============================================
class TestResolvePrice:
    """価格解決テスト"""

    def test_normal(self):
        row = _make_row(unit_price="500", net_amount="50000")
        result = _resolve_price(row, "2月20日配達予定", False)
        assert result == ("500", "50000")

    def test_unit_price_1(self):
        row = _make_row(unit_price=1)
        result = _resolve_price(row, "2月20日配達予定", False)
        assert result == ("確認中", "確認中")

    def test_unit_price_str_1(self):
        row = _make_row(unit_price="1")
        result = _resolve_price(row, "2月20日配達予定", False)
        assert result == ("確認中", "確認中")

    def test_confirming_no_dollar(self):
        row = _make_row(comment_internal="")
        result = _resolve_price(row, "確認中", False)
        assert result == ("確認中", "確認中")

    def test_confirming_with_dollar(self):
        row = _make_row(comment_internal="$$")
        result = _resolve_price(row, "確認中", False)
        assert result == ("500", "50000")

    def test_force_delivered_overrides(self):
        row = _make_row(comment_internal="")
        result = _resolve_price(row, "確認中", True)
        assert result == ("500", "50000")


# ============================================
# _pass_basic_filter
# ============================================
class TestPassBasicFilter:
    """基本フィルタテスト"""

    def test_pass(self):
        row = _make_row()
        assert _pass_basic_filter(row) is True

    def test_reject_deleted(self):
        row = _make_row(rejection_reason="明細削除")
        assert _pass_basic_filter(row) is False

    def test_reject_wrong_doc_type(self):
        row = _make_row(document_type="【受注】返品")
        assert _pass_basic_filter(row) is False

    def test_reject_hash_exclude(self):
        row = _make_row(comment_internal="##除外")
        assert _pass_basic_filter(row) is False

    def test_reject_fullwidth_hash(self):
        row = _make_row(comment_internal="＃＃除外")
        assert _pass_basic_filter(row) is False

    def test_chokusouhan(self):
        row = _make_row(document_type="【受注】直送販売")
        assert _pass_basic_filter(row) is True


# ============================================
# _is_excluded
# ============================================
class TestIsExcluded:
    """除外判定テスト"""

    def test_excluded_from_history(self):
        row = _make_row()
        cache = _make_cache()
        assert _is_excluded(row, "除外", cache) is True

    def test_excluded_from_confirming(self):
        row = _make_row(order_number="100", detail_number="10")
        cache = _make_cache(confirm={"100|10": ("除外", "", None)})
        assert _is_excluded(row, "", cache) is True

    def test_not_excluded(self):
        row = _make_row()
        cache = _make_cache()
        assert _is_excluded(row, "", cache) is False


# ============================================
# _determine_flags
# ============================================
class TestDetermineFlags:
    """フラグ判定テスト"""

    def test_stock_normal(self):
        row = _make_row(ship_status="未処理")
        cache = _make_cache()
        fd, hm, bc = _determine_flags(row, cache, {}, "100|10", "")
        assert fd is False
        assert hm is False
        assert bc is False

    def test_chokusouhan_completed_force_delivered(self):
        """直送+処理完了+転送中 → forceDelivered"""
        row = _make_row(
            document_type="【受注】直送販売",
            storage_place="転送中（直送用）",
            ship_status="処理完了",
        )
        cache = _make_cache()
        fd, hm, bc = _determine_flags(row, cache, {}, "100|10", "")
        assert fd is True
        assert hm is False

    def test_himozuki_no_force(self):
        """紐付き（直送+非転送中） → isHimozuki, forceDelivered=False"""
        row = _make_row(
            document_type="【受注】直送販売",
            storage_place="関東商品センター",
            ship_status="処理完了",
        )
        cache = _make_cache()
        fd, hm, bc = _determine_flags(row, cache, {}, "100|10", "")
        assert fd is False
        assert hm is True

    def test_bunno_completed(self):
        """分納+処理完了 → isBunnoCompleted"""
        row = _make_row(
            ship_status="処理完了",
            order_number="100",
            detail_number="10",
        )
        cache = _make_cache(
            confirm={"100|10": ("未", "分納", None)}
        )
        fd, hm, bc = _determine_flags(row, cache, {}, "100|10", "")
        assert bc is True

    def test_previous_confirming_forces(self):
        """前回確認中+処理完了 → forceDelivered"""
        row = _make_row(
            document_type="【受注】在庫販売",
            ship_status="処理完了",
        )
        cache = _make_cache()
        sent = {"1000001|10": "確認中"}
        fd, hm, bc = _determine_flags(
            row, cache, sent, "1000001|10", "確認中"
        )
        assert fd is True


# ============================================
# build_report_row
# ============================================
class TestBuildReportRow:
    """OrderRow → ReportRow変換テスト"""

    def test_basic(self):
        row = _make_row()
        cache = _make_cache()
        report, status = build_report_row(
            row, cache, {}, BRANCH, EXEC_TIME, False, TODAY
        )
        assert isinstance(report, ReportRow)
        assert report.manufacturer_name == "ダイヘン"
        assert report.order_number == "1000001"
        assert report.delivery_place == "貴社"
        assert "予定" in status or "済み" in status

    def test_stockout_confirming(self):
        """欠品中 + 確認中 → 欠品中"""
        row = _make_row(
            comment_detail="欠品中",
            order_delivery_date=datetime.date(2026, 12, 31),
        )
        cache = _make_cache()
        _, status = build_report_row(
            row, cache, {}, BRANCH, EXEC_TIME, False, TODAY
        )
        assert status == "欠品中"

    def test_stockout_with_date(self):
        """欠品中 + 納期あり → ○月○日配達予定（欠品）"""
        row = _make_row(
            comment_detail="欠品中",
            order_delivery_date=datetime.date(2026, 2, 20),
        )
        cache = _make_cache()
        _, status = build_report_row(
            row, cache, {}, BRANCH, EXEC_TIME, False, TODAY
        )
        assert "（欠品）" in status

    def test_bunno(self):
        """分納コメントあり → 分納"""
        row = _make_row(
            comment_detail="分納:50個 2/20,50個 未定",
            ship_status="未処理",
        )
        cache = _make_cache()
        _, status = build_report_row(
            row, cache, {}, BRANCH, EXEC_TIME, False, TODAY
        )
        assert status == "分納"

    def test_bunno_completed_ignored(self):
        """分納+処理完了 → 分納は無視（通常の納期計算）"""
        row = _make_row(
            comment_detail="分納:50個 2/20,50個 未定",
            ship_status="処理完了",
        )
        cache = _make_cache()
        _, status = build_report_row(
            row, cache, {}, BRANCH, EXEC_TIME, False, TODAY
        )
        assert status != "分納"

    def test_remark_cleaned(self):
        """備考から欠品テキスト・分納テキスト除去"""
        row = _make_row(
            comment_detail="欠品中 1月上旬予定 分納:50個 2/20"
        )
        cache = _make_cache()
        report, _ = build_report_row(
            row, cache, {}, BRANCH, EXEC_TIME, False, TODAY
        )
        assert "欠品中" not in report.remarks
        assert "分納:" not in report.remarks

    def test_price_confirming(self):
        """単価=1 → 確認中"""
        row = _make_row(unit_price=1, net_amount=100)
        cache = _make_cache()
        report, _ = build_report_row(
            row, cache, {}, BRANCH, EXEC_TIME, False, TODAY
        )
        assert report.unit_price == "確認中"
        assert report.net_amount == "確認中"


# ============================================
# _classify_order
# ============================================
class TestClassifyOrder:
    """確定/確認中分類テスト"""

    def test_confirmed(self):
        """通常の確定伝票 → confirmed_orders"""
        confirmed: list[HistoryRecord] = []
        confirming: list[ConfirmingRecord] = []
        row = _make_row()
        cache = _make_cache()
        _classify_order(
            row, "2月20日配達予定", cache, [],
            False, "ダイヘン", "溶接棒",
            confirmed, confirming,
        )
        assert len(confirmed) == 1
        assert len(confirming) == 0
        assert confirmed[0].delivery_answer == "2月20日配達予定"

    def test_confirming(self):
        """確認中 → confirming_orders"""
        confirmed: list[HistoryRecord] = []
        confirming: list[ConfirmingRecord] = []
        row = _make_row()
        cache = _make_cache()
        _classify_order(
            row, "確認中", cache, [],
            False, "ダイヘン", "溶接棒",
            confirmed, confirming,
        )
        assert len(confirmed) == 0
        assert len(confirming) == 1

    def test_stockout(self):
        """欠品中 → confirming_orders with status=欠品中"""
        confirmed: list[HistoryRecord] = []
        confirming: list[ConfirmingRecord] = []
        row = _make_row()
        cache = _make_cache()
        _classify_order(
            row, "欠品中", cache, [],
            False, "ダイヘン", "溶接棒",
            confirmed, confirming,
        )
        assert len(confirming) == 1
        assert confirming[0].status == "欠品中"

    def test_bunno_with_mitei(self):
        """分納（未定あり） → confirming_orders with status=分納"""
        confirmed: list[HistoryRecord] = []
        confirming: list[ConfirmingRecord] = []
        row = _make_row()
        cache = _make_cache()
        bunno = [BunnoEntry("50個", "2/20", ""), BunnoEntry("50個", "未定", "")]
        _classify_order(
            row, "分納", cache, bunno,
            False, "ダイヘン", "溶接棒",
            confirmed, confirming,
        )
        assert len(confirming) == 1
        assert confirming[0].status == "分納"

    def test_bunno_all_confirmed(self):
        """分納（全確定） → confirmed_orders with 分納完了"""
        confirmed: list[HistoryRecord] = []
        confirming: list[ConfirmingRecord] = []
        row = _make_row()
        cache = _make_cache()
        bunno = [BunnoEntry("50個", "2/20", ""), BunnoEntry("50個", "2/25", "")]
        _classify_order(
            row, "分納", cache, bunno,
            False, "ダイヘン", "溶接棒",
            confirmed, confirming,
        )
        assert len(confirmed) == 1
        assert confirmed[0].delivery_answer == "分納完了"

    def test_bunno_completed(self):
        """分納完了 → confirmed_orders with 分納完了"""
        confirmed: list[HistoryRecord] = []
        confirming: list[ConfirmingRecord] = []
        row = _make_row()
        cache = _make_cache()
        _classify_order(
            row, "納品済み", cache, [],
            True, "ダイヘン", "溶接棒",
            confirmed, confirming,
        )
        assert len(confirmed) == 1
        assert confirmed[0].delivery_answer == "分納完了"

    def test_scheduling(self):
        """日程調整中 → confirming_orders"""
        confirmed: list[HistoryRecord] = []
        confirming: list[ConfirmingRecord] = []
        row = _make_row()
        cache = _make_cache()
        _classify_order(
            row, "日程調整中", cache, [],
            False, "ダイヘン", "溶接棒",
            confirmed, confirming,
        )
        assert len(confirming) == 1


# ============================================
# create_delivery_report_by_order_numbers
# ============================================
class TestCreateDeliveryReportByOrderNumbers:
    """注番指定モード回答書作成テスト"""

    def test_basic(self, tmp_path):
        """基本的な回答書作成"""
        data = [_make_row(order_number="100")]
        cache = _make_cache()
        result = create_delivery_report_by_order_numbers(
            data, "テスト商事", ["100"],
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is not None
        assert result.customer_name == "テスト商事"
        assert Path(result.file_path).exists()

    def test_no_matching_data(self, tmp_path):
        """該当なし → None"""
        data = [_make_row(order_number="100")]
        cache = _make_cache()
        result = create_delivery_report_by_order_numbers(
            data, "テスト商事", ["999"],
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is None

    def test_rejects_deleted(self, tmp_path):
        """明細削除は除外"""
        data = [_make_row(order_number="100", rejection_reason="明細削除")]
        cache = _make_cache()
        result = create_delivery_report_by_order_numbers(
            data, "テスト商事", ["100"],
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is None

    def test_multiple_orders(self, tmp_path):
        """複数注番"""
        data = [
            _make_row(order_number="100", detail_number="10"),
            _make_row(order_number="200", detail_number="10",
                      product_name="ガス 10L"),
        ]
        cache = _make_cache()
        result = create_delivery_report_by_order_numbers(
            data, "テスト商事", ["100", "200"],
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is not None
        # Excelを読んで2行あることを確認
        wb = load_workbook(result.file_path)
        ws = wb.active
        assert ws.cell(row=7, column=12).value == "100"
        assert ws.cell(row=8, column=12).value == "200"

    def test_tracking_info_collected(self, tmp_path):
        """送り状情報収集"""
        data = [
            _make_row(
                order_number="100",
                comment_external="佐川 1234567890",
            )
        ]
        cache = _make_cache()
        result = create_delivery_report_by_order_numbers(
            data, "テスト商事", ["100"],
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is not None
        assert len(result.tracking_info_list) == 1

    def test_stockout_info_collected(self, tmp_path):
        """欠品情報収集"""
        data = [
            _make_row(
                order_number="100",
                comment_detail="欠品中 3月上旬予定",
                order_delivery_date=datetime.date(2026, 12, 31),
            )
        ]
        cache = _make_cache()
        result = create_delivery_report_by_order_numbers(
            data, "テスト商事", ["100"],
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is not None
        assert len(result.stockout_info_list) == 1


# ============================================
# create_delivery_report
# ============================================
class TestCreateDeliveryReport:
    """期間指定モード回答書作成テスト"""

    def test_basic(self, tmp_path):
        """基本的な回答書作成"""
        data = [_make_row()]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            date_from=datetime.date(2026, 2, 1),
            date_to=datetime.date(2026, 2, 28),
            today=TODAY,
        )
        assert result is not None
        assert result.customer_name == "テスト商事"
        assert Path(result.file_path).exists()
        assert len(result.confirmed_orders) == 1

    def test_no_matching_customer(self, tmp_path):
        """顧客名不一致 → None"""
        data = [_make_row(customer_name="別会社")]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is None

    def test_date_range_filter(self, tmp_path):
        """期間外はフィルタされる"""
        data = [
            _make_row(registration_date=datetime.date(2026, 1, 15)),
        ]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            date_from=datetime.date(2026, 2, 1),
            date_to=datetime.date(2026, 2, 28),
            today=TODAY,
        )
        assert result is None

    def test_already_sent_skip(self, tmp_path):
        """送付済み伝票はスキップ（登録日が今日より前）"""
        data = [
            _make_row(
                registration_date=datetime.date(2026, 2, 10),
            ),
        ]
        sent = {"1000001|10": "2月15日配達予定"}
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", sent,
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            date_from=datetime.date(2026, 2, 1),
            date_to=datetime.date(2026, 2, 28),
            today=TODAY,
        )
        assert result is None

    def test_today_registration_resend(self, tmp_path):
        """当日登録の送付済み → 再送OK"""
        data = [
            _make_row(registration_date=TODAY),
        ]
        sent = {"1000001|10": "2月16日配達予定"}
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", sent,
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            date_from=datetime.date(2026, 2, 1),
            date_to=datetime.date(2026, 2, 28),
            today=TODAY,
        )
        # 当日登録の場合、registration_date < today が False なので再送される
        assert result is not None

    def test_force_delivered(self, tmp_path):
        """直送+処理完了 → 納品済み"""
        data = [
            _make_row(
                document_type="【受注】直送販売",
                storage_place="転送中（直送用）",
                ship_status="処理完了",
                order_delivery_date=datetime.date(2026, 12, 31),
            ),
        ]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is not None
        # 確定伝票として記録
        assert len(result.confirmed_orders) == 1
        assert result.confirmed_orders[0].delivery_answer == "納品済み"

    def test_confirming_orders(self, tmp_path):
        """確認中 → confirming_orders"""
        data = [
            _make_row(
                order_delivery_date=datetime.date(2026, 12, 31),
                document_type="【受注】直送販売",
                storage_place="転送中（直送用）",
            ),
        ]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is not None
        assert len(result.confirming_orders) == 1
        assert result.has_confirming is True

    def test_excluded_skip(self, tmp_path):
        """除外伝票はスキップ"""
        data = [_make_row()]
        sent = {"1000001|10": "除外"}
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", sent,
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is None

    def test_bunno_completed_notification(self, tmp_path):
        """分納完了の通知収集"""
        data = [
            _make_row(
                ship_status="処理完了",
                order_number="100",
                detail_number="10",
                comment_detail="分納:50個 2/10,50個 2/15",
            ),
        ]
        cache = _make_cache(
            confirm={"100|10": ("未", "分納", None)}
        )
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is not None
        assert len(result.bunno_completed_list) == 1

    def test_hash_exclude(self, tmp_path):
        """##除外"""
        data = [
            _make_row(comment_internal="##この伝票は除外"),
        ]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is None

    def test_bunno_kanryo_skip(self, tmp_path):
        """分納完了は送付済みとしてスキップ"""
        data = [_make_row(registration_date=datetime.date(2026, 2, 10))]
        sent = {"1000001|10": "分納完了"}
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", sent,
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            date_from=datetime.date(2026, 2, 1),
            date_to=datetime.date(2026, 2, 28),
            today=TODAY,
        )
        assert result is None


# ============================================
# 結合テスト
# ============================================
class TestIntegration:
    """結合テスト"""

    def test_full_flow_period_mode(self, tmp_path):
        """期間モードの完全フロー"""
        data = [
            # 確定伝票
            _make_row(
                order_number="100", detail_number="10",
                product_name="溶接棒A", quantity="50",
            ),
            # 確認中伝票
            _make_row(
                order_number="200", detail_number="10",
                product_name="ガスB", quantity="10",
                order_delivery_date=datetime.date(2026, 12, 31),
                document_type="【受注】直送販売",
                storage_place="転送中（直送用）",
            ),
            # 別顧客（フィルタされる）
            _make_row(
                order_number="300", detail_number="10",
                customer_name="別会社",
            ),
        ]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            date_from=datetime.date(2026, 2, 1),
            date_to=datetime.date(2026, 2, 28),
            today=TODAY,
        )
        assert result is not None
        assert len(result.confirmed_orders) == 1
        assert len(result.confirming_orders) == 1

        # Excelファイルを検証
        wb = load_workbook(result.file_path)
        ws = wb.active
        assert ws.cell(row=7, column=12).value == "100"
        assert ws.cell(row=8, column=12).value == "200"

    def test_full_flow_order_number_mode(self, tmp_path):
        """注番モードの完全フロー"""
        data = [
            _make_row(order_number="100", detail_number="10"),
            _make_row(order_number="100", detail_number="20",
                      product_name="溶接ワイヤ"),
            _make_row(order_number="200", detail_number="10",
                      customer_name="別会社"),
        ]
        cache = _make_cache()

        # まず注番をグループ化
        groups = group_order_numbers_by_customer(data, ["100"])
        assert "テスト商事" in groups

        # 回答書作成
        result = create_delivery_report_by_order_numbers(
            data, "テスト商事", groups["テスト商事"],
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is not None
        wb = load_workbook(result.file_path)
        ws = wb.active
        # 2行（同じ注番の2明細）
        assert ws.cell(row=7, column=12).value == "100"
        assert ws.cell(row=8, column=12).value == "100"
