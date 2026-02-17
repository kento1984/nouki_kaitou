"""models.py のユニットテスト"""

import datetime

from nouki_kaitou.models import (
    AppConfig,
    BranchSettings,
    BunnoEntry,
    CacheStore,
    ConfirmingRecord,
    HistoryRecord,
    OrderRow,
    ReportResult,
    ReportRow,
    StockoutEntry,
    TrackingEntry,
)


class TestBranchSettings:
    def test_defaults(self):
        bs = BranchSettings()
        assert bs.name == ""
        assert bs.default_cutoff == 15
        assert bs.base_center == ""

    def test_creation(self):
        bs = BranchSettings(
            name="京葉営業所",
            default_cutoff=15,
            base_center="関東商品センター",
            shared_email="keiyou@mac-exe.co.jp",
            signature="マツモト産業\n京葉営業所",
            start_date="2026/01/06",
        )
        assert bs.name == "京葉営業所"
        assert bs.default_cutoff == 15


class TestAppConfig:
    def test_defaults(self):
        config = AppConfig()
        assert config.current_subfolder == ""
        assert isinstance(config.execution_time, datetime.datetime)
        assert isinstance(config.branch, BranchSettings)


class TestCacheStore:
    def test_defaults(self):
        cache = CacheStore()
        assert cache.mfg_name == {}
        assert cache.mfg_days == {}
        assert cache.cust_days == {}

    def test_add_data(self):
        cache = CacheStore()
        cache.mfg_name["D01"] = "（株）ダイヘン"
        cache.mfg_days["D01"] = 2
        cache.cust_route["テスト顧客"] = True
        assert cache.mfg_name["D01"] == "（株）ダイヘン"
        assert cache.cust_route["テスト顧客"] is True


class TestOrderRow:
    def test_creation(self):
        row = OrderRow(
            order_number="GL2Z444369",
            detail_number="10",
            document_type="【受注】直送販売",
            customer_name="共同ガス（株）　本社",
            product_name="Ｕ１２５４６－４　チップ",
            ship_status="処理完了",
            registration_date=datetime.date(2026, 1, 5),
            item_group_code="D01",
        )
        assert row.order_number == "GL2Z444369"
        assert row.ship_status == "処理完了"


class TestReportRow:
    def test_creation(self):
        row = ReportRow(
            manufacturer_name="ダイヘン",
            product_name="チップ",
            quantity="100",
            delivery_answer="1月9日配達済み",
        )
        assert row.delivery_answer == "1月9日配達済み"


class TestBunnoEntry:
    def test_creation(self):
        entry = BunnoEntry(quantity="30個", date_str="3/10", location="貴社")
        assert entry.quantity == "30個"
        assert entry.date_str == "3/10"


class TestTrackingEntry:
    def test_creation(self):
        entry = TrackingEntry(
            carrier_name="佐川急便", tracking_number="452710942423"
        )
        assert entry.carrier_name == "佐川急便"
        assert len(entry.tracking_number) == 12


class TestStockoutEntry:
    def test_creation(self):
        entry = StockoutEntry(
            manufacturer_name="ＮＡＣＨＩ",
            product_name="オイルホールドリル",
            quantity="2",
            approx_delivery="3月上旬入荷予定",
        )
        assert entry.approx_delivery == "3月上旬入荷予定"


class TestHistoryRecord:
    def test_creation(self):
        rec = HistoryRecord(
            order_number="GL2C444510",
            detail_number="10",
            customer_name="テスト顧客",
            delivery_answer="1月14日配達済み",
            sender="boxeo",
        )
        assert rec.order_number == "GL2C444510"
        assert isinstance(rec.sent_datetime, datetime.datetime)


class TestConfirmingRecord:
    def test_defaults(self):
        rec = ConfirmingRecord()
        assert rec.inquiry_status == "未"
        assert rec.status == ""


class TestReportResult:
    def test_defaults(self):
        result = ReportResult()
        assert result.confirmed_orders == []
        assert result.confirming_orders == []
        assert result.has_confirming is False
