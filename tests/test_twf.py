"""twf モジュール（TWF展示会受注の専用回答書）のテスト

期間限定機能。機能削除時はこのテストファイルごと削除してよい。
"""

import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from nouki_kaitou.models import (
    BranchSettings,
    CacheStore,
    OrderRow,
    ReportResult,
    StockoutEntry,
)
from nouki_kaitou.report_generator import create_delivery_report
from nouki_kaitou.twf import (
    TWF_END_DATE,
    TWF_FILENAME_TAG,
    TWF_NOTICE_EMAIL,
    TWF_NOTICE_EXCEL,
    TWF_REPORT_TITLE,
    TWF_SHEET_PREFIX,
    build_twf_remark_map,
    collect_twf_orders,
    format_twf_remark,
    is_twf_active,
    is_twf_comment,
    merge_email_input,
    normalize_twf_text,
    remove_twf_text,
)


# ============================================
# テスト用ヘルパー（test_report_generator.py と同パターン）
# ============================================
TODAY = datetime.date(2026, 2, 16)
EXEC_TIME = datetime.datetime(2026, 2, 16, 10, 0, 0)
BRANCH = BranchSettings(
    name="京葉営業所",
    default_cutoff=15,
    base_center="関東商品センター",
)


def _make_cache(**kwargs) -> CacheStore:
    return CacheStore(
        mfg_name=kwargs.get("mfg_name", {"D01": "ダイヘン"}),
        mfg_days=kwargs.get("mfg_days", {"D01": 2}),
        cust_days=kwargs.get("cust_days", {}),
        cust_retention=kwargs.get("cust_retention", {}),
        cust_route=kwargs.get("cust_route", {}),
        confirm=kwargs.get("confirm", {}),
        storage=kwargs.get("storage", {}),
    )


def _make_row(**kwargs) -> OrderRow:
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
# normalize_twf_text
# ============================================
class TestNormalizeTwfText:
    def test_fullwidth_to_halfwidth(self):
        assert normalize_twf_text("ＴＷＦＮｏ．００３２４３") == "TWFNO.003243"

    def test_lowercase_to_upper(self):
        assert normalize_twf_text("twfno.003243") == "TWFNO.003243"

    def test_remove_spaces(self):
        assert normalize_twf_text("TWF No. 003243") == "TWFNO.003243"

    def test_remove_fullwidth_spaces(self):
        assert normalize_twf_text("TWF　No．003243　新成（株）") == "TWFNO.003243新成(株)"

    def test_numero_sign(self):
        # NFKC正規化で № → No に展開される
        assert normalize_twf_text("TWF№003243") == "TWFNO003243"

    def test_empty(self):
        assert normalize_twf_text("") == ""


# ============================================
# is_twf_comment
# ============================================
class TestIsTwfComment:
    """正規化後に「TWFNO」を含むか（数字の有無は問わない）"""

    @pytest.mark.parametrize("comment", [
        "TWFNo.003243　新成（株）",
        "ＴＷＦＮｏ．００３２４３　新成（株）",
        "TWF No. 003243",
        "TWF　No　3243",
        "twf no.3243",
        "TWF№003243",
        "ＴＷＦ№３２４３",
        "TWFNo.",          # 番号の書き忘れ・省略も検知
        "TWFNO3243",
        "分納:50個 2/20、TWFNo.003243",  # 他コメントとの併記
    ])
    def test_positive(self, comment):
        assert is_twf_comment(comment) is True

    @pytest.mark.parametrize("comment", [
        "ＴＷＦ特価",          # 展示会前の特価先行受注（対象外）
        "TWF特価",
        "ダイヘン期間限定ＷＦキャンペーン",
        "TWF",                # Noなし単独は対象外
        "欠品中 3月上旬入荷予定",
        "",
    ])
    def test_negative(self, comment):
        assert is_twf_comment(comment) is False


# ============================================
# collect_twf_orders（注番単位の伝播）
# ============================================
class TestCollectTwfOrders:
    def test_propagates_to_all_details_of_order(self):
        """同一注番の1明細にTWF記載 → 注番全体が対象"""
        orders = [
            _make_row(order_number="100", detail_number="10",
                      comment_detail="TWFNo.003243　新成（株）"),
            _make_row(order_number="100", detail_number="20",
                      comment_detail=""),  # 入れ忘れ
            _make_row(order_number="200", detail_number="10",
                      comment_detail=""),
        ]
        twf_orders, detected = collect_twf_orders(orders)
        assert twf_orders == {"100"}
        assert len(detected) == 1
        assert detected[0] == ("100", "10", "TWFNo.003243　新成（株）")

    def test_multiple_orders(self):
        orders = [
            _make_row(order_number="100", comment_detail="TWFNo.001"),
            _make_row(order_number="200", comment_detail="ＴＷＦ№002"),
            _make_row(order_number="300", comment_detail="ＴＷＦ特価"),  # 対象外
        ]
        twf_orders, detected = collect_twf_orders(orders)
        assert twf_orders == {"100", "200"}
        assert len(detected) == 2

    def test_empty(self):
        twf_orders, detected = collect_twf_orders([])
        assert twf_orders == set()
        assert detected == []


# ============================================
# remove_twf_text
# ============================================
class TestRemoveTwfText:
    def test_remove_to_end_of_line(self):
        """TWF記載から行末（得意先名含む）まで除去"""
        assert remove_twf_text("TWFNo.003243　新成（株）") == ""

    def test_keep_preceding_text(self):
        result = remove_twf_text("至急対応 TWFNo.003243　新成（株）")
        assert result.strip() == "至急対応"

    def test_fullwidth(self):
        assert remove_twf_text("ＴＷＦＮｏ．００３２４３　新成（株）") == ""

    def test_numero_sign(self):
        assert remove_twf_text("TWF№003243 新成") == ""

    def test_no_twf_text_unchanged(self):
        assert remove_twf_text("欠品中 3月上旬入荷予定") == "欠品中 3月上旬入荷予定"

    def test_twf_tokka_unchanged(self):
        """ＴＷＦ特価（Noなし）は除去しない（判定対象外のため）"""
        assert remove_twf_text("ＴＷＦ特価") == "ＴＷＦ特価"

    def test_empty(self):
        assert remove_twf_text("") == ""

    def test_multiline_keeps_other_lines(self):
        result = remove_twf_text("分納:50個 2/20\nTWFNo.003243　新成（株）")
        assert "分納:50個 2/20" in result
        assert "TWF" not in result


# ============================================
# format_twf_remark（K列備考用の整形）
# ============================================
class TestFormatTwfRemark:
    def test_basic(self):
        assert format_twf_remark("TWFNo.003243　新成（株）様") == "No.003243 新成（株）様"

    def test_number_only(self):
        assert format_twf_remark("TWFNo.005210") == "No.005210"

    def test_fumei(self):
        assert format_twf_remark("TWFNo.不明　新成㈱第三工場様") == "No.不明 新成㈱第三工場様"

    def test_fullwidth(self):
        assert format_twf_remark("ＴＷＦＮｏ．００３２４３　新成（株）様") == "No.００３２４３ 新成（株）様"

    def test_numero_sign(self):
        assert format_twf_remark("TWF№003243 新成様") == "No.003243 新成様"

    def test_no_period_no_space(self):
        assert format_twf_remark("TWFNo.004662先進機設（株）様") == "No.004662先進機設（株）様"

    def test_collapse_spaces(self):
        assert format_twf_remark("TWF No.  003243　　新成様") == "No.003243 新成様"

    def test_status_memo(self):
        assert format_twf_remark("TWFNo.004982　(有)狩野溶接工業様お持ち帰り") == "No.004982 (有)狩野溶接工業様お持ち帰り"

    def test_no_twf(self):
        assert format_twf_remark("欠品中 3月上旬入荷予定") == ""

    def test_twf_tokka_not_formatted(self):
        assert format_twf_remark("ＴＷＦ特価") == ""

    def test_empty(self):
        assert format_twf_remark("") == ""

    def test_bare_no(self):
        assert format_twf_remark("TWFNo.") == "No."


# ============================================
# build_twf_remark_map（注番→TWF情報の引き継ぎマップ）
# ============================================
class TestBuildTwfRemarkMap:
    def test_first_occurrence_wins(self):
        orders = [
            _make_row(order_number="100", detail_number="10",
                      comment_detail="TWFNo.001　新成様"),
            _make_row(order_number="100", detail_number="20",
                      comment_detail="TWFNo.001　別テキスト"),
            _make_row(order_number="200", detail_number="10",
                      comment_detail=""),
        ]
        m = build_twf_remark_map(orders)
        assert m == {"100": "No.001 新成様"}

    def test_empty(self):
        assert build_twf_remark_map([]) == {}


# ============================================
# is_twf_active（期間ゲート）
# ============================================
class TestIsTwfActive:
    def test_before_end_date(self):
        assert is_twf_active(datetime.date(2026, 6, 12)) is True

    def test_on_end_date(self):
        assert is_twf_active(TWF_END_DATE) is True

    def test_after_end_date(self):
        assert is_twf_active(TWF_END_DATE + datetime.timedelta(days=1)) is False


# ============================================
# merge_email_input
# ============================================
class TestMergeEmailInput:
    def _result(self, file_path="r.xlsx", is_twf=False, **kwargs) -> ReportResult:
        return ReportResult(
            file_path=file_path,
            customer_name=kwargs.get("customer_name", "テスト商事"),
            rep_name=kwargs.get("rep_name", ""),
            stockout_info_list=kwargs.get("stockout_info_list", []),
            tracking_info_list=kwargs.get("tracking_info_list", []),
            bunno_info_list=kwargs.get("bunno_info_list", []),
            bunno_completed_list=kwargs.get("bunno_completed_list", []),
            is_twf=is_twf,
        )

    def test_normal_only(self):
        merged = merge_email_input(self._result("normal.xlsx"), None)
        assert merged["attachments"] == ["normal.xlsx"]
        assert merged["twf_notice"] is None
        assert merged["customer_name"] == "テスト商事"

    def test_normal_and_twf(self):
        merged = merge_email_input(
            self._result("normal.xlsx"),
            self._result("twf.xlsx", is_twf=True),
        )
        assert merged["attachments"] == ["normal.xlsx", "twf.xlsx"]
        assert merged["twf_notice"] == TWF_NOTICE_EMAIL

    def test_twf_only(self):
        """TWF受注のみの顧客（通常回答書なし）"""
        merged = merge_email_input(None, self._result("twf.xlsx", is_twf=True))
        assert merged["attachments"] == ["twf.xlsx"]
        assert merged["twf_notice"] == TWF_NOTICE_EMAIL
        assert merged["customer_name"] == "テスト商事"

    def test_info_lists_merged(self):
        normal = self._result(
            "normal.xlsx",
            stockout_info_list=[StockoutEntry(product_name="商品A")],
        )
        twf = self._result(
            "twf.xlsx", is_twf=True,
            stockout_info_list=[StockoutEntry(product_name="商品B")],
        )
        merged = merge_email_input(normal, twf)
        names = [s.product_name for s in merged["stockout_info_list"]]
        assert names == ["商品A", "商品B"]

    def test_both_none_raises(self):
        with pytest.raises(ValueError):
            merge_email_input(None, None)


# ============================================
# create_delivery_report のTWF関連パラメータ
# ============================================
class TestCreateDeliveryReportTwf:
    def test_exclude_orders(self, tmp_path):
        """exclude_orders指定の注番は通常回答書から除外される"""
        data = [
            _make_row(order_number="100", comment_detail="TWFNo.001"),
            _make_row(order_number="200"),
        ]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            date_from=datetime.date(2026, 2, 1),
            date_to=datetime.date(2026, 2, 28),
            today=TODAY,
            exclude_orders={"100"},
        )
        assert result is not None
        wb = load_workbook(result.file_path)
        ws = wb.active
        order_nums = [ws.cell(row=r, column=12).value for r in range(7, 9)]
        assert "200" in order_nums
        assert "100" not in order_nums

    def test_exclude_all_returns_none(self, tmp_path):
        """全注番が除外されたら None（回答書なし）"""
        data = [_make_row(order_number="100")]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            exclude_orders={"100"},
        )
        assert result is None

    def test_include_only_orders(self, tmp_path):
        """include_only_orders指定の注番のみが出力される"""
        data = [
            _make_row(order_number="100", comment_detail="TWFNo.001"),
            _make_row(order_number="200"),
        ]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            include_only_orders={"100"},
        )
        assert result is not None
        wb = load_workbook(result.file_path)
        ws = wb.active
        assert ws.cell(row=7, column=12).value == "100"
        assert ws.cell(row=8, column=12).value is None

    def test_filter_already_sent_false_shows_sent(self, tmp_path):
        """filter_already_sent=False なら送付済みでも表示される"""
        data = [
            _make_row(registration_date=datetime.date(2026, 2, 10)),
        ]
        sent = {"1000001|10": "2月15日配達予定"}
        cache = _make_cache()

        # 従来動作（True）: スキップされ None
        result_default = create_delivery_report(
            data, "テスト商事", sent,
            cache, tmp_path / "a", {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result_default is None

        # False: 表示される
        (tmp_path / "b").mkdir()
        result = create_delivery_report(
            data, "テスト商事", sent,
            cache, tmp_path / "b", {}, BRANCH, EXEC_TIME,
            today=TODAY,
            filter_already_sent=False,
        )
        assert result is not None

    def test_filter_already_sent_false_excluded_still_skipped(self, tmp_path):
        """filter_already_sent=False でも「除外」ステータスは常にスキップ"""
        data = [_make_row()]
        sent = {"1000001|10": "除外"}
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", sent,
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            filter_already_sent=False,
        )
        assert result is None

    def test_filter_already_sent_false_hash_marker_still_skipped(self, tmp_path):
        """filter_already_sent=False でも##除外マーカーは常にスキップ"""
        data = [_make_row(comment_internal="##テスト除外")]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            filter_already_sent=False,
        )
        assert result is None

    def test_twf_mode_title_and_sheet_name(self, tmp_path):
        """twf_mode: タイトル・シート名・ファイル名がTWF用になる"""
        data = [_make_row(order_number="100", comment_detail="TWFNo.001")]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            include_only_orders={"100"},
            filter_already_sent=False,
            twf_mode=True,
        )
        assert result is not None
        assert result.is_twf is True
        assert TWF_FILENAME_TAG in Path(result.file_path).name

        wb = load_workbook(result.file_path)
        ws = wb.active
        assert ws.title.startswith(TWF_SHEET_PREFIX)
        assert ws["A1"].value == TWF_REPORT_TITLE

    def test_twf_mode_notice_in_info_section(self, tmp_path):
        """twf_mode: ご連絡事項にTWF注記が出る"""
        data = [_make_row(order_number="100", comment_detail="TWFNo.001")]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            include_only_orders={"100"},
            filter_already_sent=False,
            twf_mode=True,
        )
        wb = load_workbook(result.file_path)
        ws = wb.active
        values = [
            ws.cell(row=r, column=1).value
            for r in range(7, ws.max_row + 1)
            if ws.cell(row=r, column=1).value
        ]
        assert TWF_NOTICE_EXCEL in values

    def _twf_report(self, data, tmp_path, sent=None):
        """TWF専用回答書を生成して(result, ws)を返すヘルパー。"""
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", sent or {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            include_only_orders={r.order_number.strip() for r in data},
            filter_already_sent=False,
            twf_mode=True,
        )
        wb = load_workbook(result.file_path)
        return result, wb.active

    def test_twf_mode_remark_formatted(self, tmp_path):
        """TWF記載はK列備考に「No.～」形式で整形表示される"""
        data = [_make_row(order_number="100",
                          comment_detail="TWFNo.003243　新成（株）様")]
        result, ws = self._twf_report(data, tmp_path)
        assert ws.cell(row=7, column=11).value == "No.003243 新成（株）様"

    def test_twf_mode_remark_inherited(self, tmp_path):
        """TWF記載のない明細（入れ忘れ）は同注番の他明細から引き継ぐ"""
        data = [
            _make_row(order_number="100", detail_number="10",
                      comment_detail="TWFNo.003243　新成（株）様"),
            _make_row(order_number="100", detail_number="20",
                      comment_detail=""),  # 入れ忘れ
        ]
        result, ws = self._twf_report(data, tmp_path)
        assert ws.cell(row=7, column=11).value == "No.003243 新成（株）様"
        assert ws.cell(row=8, column=11).value == "No.003243 新成（株）様"

    def test_twf_mode_remark_coexists_with_other_text(self, tmp_path):
        """既存の備考内容がある場合は「TWF情報 ／ 既存備考」で共存"""
        data = [_make_row(
            order_number="100",
            comment_detail="至急対応\nTWFNo.003243　新成（株）様",
        )]
        result, ws = self._twf_report(data, tmp_path)
        assert ws.cell(row=7, column=11).value == "No.003243 新成（株）様 ／ 至急対応"

    def test_normal_mode_remark_still_removed(self, tmp_path):
        """通常モードでは従来どおりTWF記載は備考から除去される"""
        data = [_make_row(order_number="100",
                          comment_detail="至急対応\nTWFNo.003243　新成（株）様")]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        wb = load_workbook(result.file_path)
        ws = wb.active
        assert ws.cell(row=7, column=11).value == "至急対応"

    def test_twf_mode_auto_filter_set(self, tmp_path):
        """TWF専用回答書にはオートフィルタが設定される"""
        data = [
            _make_row(order_number="100", detail_number="10",
                      comment_detail="TWFNo.001"),
            _make_row(order_number="100", detail_number="20",
                      comment_detail="TWFNo.001"),
        ]
        result, ws = self._twf_report(data, tmp_path)
        assert ws.auto_filter.ref == "A6:L8"

    def test_normal_mode_no_auto_filter(self, tmp_path):
        """通常回答書にはオートフィルタを設定しない（従来どおり）"""
        data = [_make_row()]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        wb = load_workbook(result.file_path)
        ws = wb.active
        assert ws.auto_filter.ref is None

    def test_twf_mode_delivered_override(self, tmp_path):
        """twf_mode: 処理完了+回答済み（履歴に確定記録）→「納品済み」表示"""
        data = [
            _make_row(
                order_number="100",
                document_type="【受注】直送販売",
                storage_place="転送中（直送用）",
                ship_status="処理完了",
                registration_date=datetime.date(2026, 2, 10),
                order_delivery_date=datetime.date(2026, 12, 31),
            ),
        ]
        # 既に「2月13日出荷予定」で回答済み
        sent = {"100|10": "2月13日出荷予定"}
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", sent,
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            include_only_orders={"100"},
            filter_already_sent=False,
            twf_mode=True,
        )
        assert result is not None
        wb = load_workbook(result.file_path)
        ws = wb.active
        assert ws.cell(row=7, column=9).value == "納品済み"

    def test_twf_mode_no_override_when_confirming(self, tmp_path):
        """twf_mode: 前回「確認中」+処理完了は既存ルール（force_delivered）で処理"""
        data = [
            _make_row(
                order_number="100",
                document_type="【受注】直送販売",
                storage_place="転送中（直送用）",
                ship_status="処理完了",
                registration_date=datetime.date(2026, 2, 10),
                order_delivery_date=datetime.date(2026, 12, 31),
            ),
        ]
        sent = {"100|10": "確認中"}
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", sent,
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            include_only_orders={"100"},
            filter_already_sent=False,
            twf_mode=True,
        )
        assert result is not None
        wb = load_workbook(result.file_path)
        ws = wb.active
        # 既存のforce_deliveredルールで納品済みになる
        assert ws.cell(row=7, column=9).value == "納品済み"

    def test_twf_mode_classification_still_recorded(self, tmp_path):
        """案Y: TWF回答書でも確定/確認中の分類（履歴記録用）は行われる"""
        data = [
            _make_row(order_number="100", comment_detail="TWFNo.001"),
            _make_row(
                order_number="100", detail_number="20",
                order_delivery_date=datetime.date(2026, 12, 31),  # 未確定
                specified_delivery_date=datetime.date(2026, 12, 31),
            ),
        ]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
            include_only_orders={"100"},
            filter_already_sent=False,
            twf_mode=True,
        )
        assert result is not None
        assert len(result.confirmed_orders) == 1   # 明細10: 確定
        assert len(result.confirming_orders) == 1  # 明細20: 日程調整中

    def test_default_behavior_unchanged(self, tmp_path):
        """新パラメータ未指定なら従来と同一の出力"""
        data = [_make_row()]
        cache = _make_cache()
        result = create_delivery_report(
            data, "テスト商事", {},
            cache, tmp_path, {}, BRANCH, EXEC_TIME,
            today=TODAY,
        )
        assert result is not None
        assert result.is_twf is False
        assert TWF_FILENAME_TAG not in Path(result.file_path).name
        wb = load_workbook(result.file_path)
        ws = wb.active
        assert ws["A1"].value == "納　期　回　答　書"
        assert not ws.title.startswith(TWF_SHEET_PREFIX)


# ============================================
# email_builder のTWF対応
# ============================================
class TestEmailBuilderTwf:
    def _make_customer_ws(self, rows):
        wb = Workbook()
        ws = wb.active
        ws.cell(row=1, column=1).value = "顧客名"
        ws.cell(row=1, column=2).value = "住所"
        ws.cell(row=1, column=3).value = "電話"
        ws.cell(row=1, column=4).value = "担当"
        ws.cell(row=1, column=5).value = "メール1"
        for row_idx, data in enumerate(rows, 2):
            for col_idx, val in enumerate(data, 1):
                ws.cell(row=row_idx, column=col_idx).value = val
        return ws

    def test_body_contains_twf_notice(self):
        from nouki_kaitou.email_builder import build_email_body_html

        body = build_email_body_html(
            "テスト商事", BRANCH, today=TODAY,
            twf_notice=TWF_NOTICE_EMAIL,
        )
        assert "東京ウェルディングフェスタ" in body

    def test_body_without_twf_notice(self):
        from nouki_kaitou.email_builder import build_email_body_html

        body = build_email_body_html("テスト商事", BRANCH, today=TODAY)
        assert "東京ウェルディングフェスタ" not in body

    def test_create_emails_multiple_attachments(self):
        from nouki_kaitou.email_builder import create_emails

        cust_ws = self._make_customer_ws([
            ("テスト商事", "", "", "", "test@example.com"),
        ])
        files = [{
            "customer_name": "テスト商事",
            "file_path": "normal.xlsx",
            "attachments": ["normal.xlsx", "twf.xlsx"],
            "twf_notice": TWF_NOTICE_EMAIL,
        }]
        results, skipped = create_emails(files, BRANCH, cust_ws, today=TODAY)
        assert len(results) == 1
        assert results[0]["attachments"] == ["normal.xlsx", "twf.xlsx"]
        assert "東京ウェルディングフェスタ" in results[0]["html_body"]

    def test_create_emails_fallback_single_attachment(self):
        """attachments未指定なら従来どおりfile_path 1件添付"""
        from nouki_kaitou.email_builder import create_emails

        cust_ws = self._make_customer_ws([
            ("テスト商事", "", "", "", "test@example.com"),
        ])
        files = [{
            "customer_name": "テスト商事",
            "file_path": "normal.xlsx",
        }]
        results, skipped = create_emails(files, BRANCH, cust_ws, today=TODAY)
        assert len(results) == 1
        assert results[0]["attachments"] == ["normal.xlsx"]
        assert "東京ウェルディングフェスタ" not in results[0]["html_body"]


# ============================================
# utils のTWF対応（ファイル名タグ・シート名プレフィックス）
# ============================================
class TestUtilsTwf:
    def test_filename_with_tag(self):
        from nouki_kaitou.utils import build_report_filename

        name = build_report_filename(
            "テスト商事", EXEC_TIME, filename_tag=TWF_FILENAME_TAG,
        )
        assert name == "納期回答書【展示会】_テスト商事様_20260216.xlsx"

    def test_filename_with_tag_and_rep(self):
        from nouki_kaitou.utils import build_report_filename

        name = build_report_filename(
            "テスト商事", EXEC_TIME, rep_name="田中",
            filename_tag=TWF_FILENAME_TAG,
        )
        assert name == "納期回答書【展示会】_テスト商事様_田中様_20260216.xlsx"

    def test_filename_without_tag_unchanged(self):
        from nouki_kaitou.utils import build_report_filename

        name = build_report_filename("テスト商事", EXEC_TIME)
        assert name == "納期回答書_テスト商事様_20260216.xlsx"

    def test_sheet_name_with_prefix(self):
        from nouki_kaitou.utils import build_sheet_name

        assert build_sheet_name("テスト商事", prefix=TWF_SHEET_PREFIX) == "展示会_テスト商事様"

    def test_sheet_name_prefix_31_chars_limit(self):
        from nouki_kaitou.utils import build_sheet_name

        name = build_sheet_name("あ" * 40, prefix=TWF_SHEET_PREFIX)
        assert len(name) <= 31
        assert name.startswith(TWF_SHEET_PREFIX)

    def test_sheet_name_without_prefix_unchanged(self):
        from nouki_kaitou.utils import build_sheet_name

        assert build_sheet_name("テスト商事") == "テスト商事様"
