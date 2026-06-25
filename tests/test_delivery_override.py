"""納期上書き（納期上書きシート＋逐語文字列）のテスト

SAP施錠等でツール計算の納期を直せない明細を、運用者が手書きした逐語文字列で
最優先に差し替える機能。期間限定運用。機能削除時はこのファイルごと削除してよい。

検証の柱:
- 上書きシート空 → 全行が従来どおり（完全no-op）
- 上書き1行 → その明細だけ逐語文字列で出る（指定納期もTWF納品済み上書きも越える）
- 同じ注番の他明細・他注番・他客は一切影響を受けない
- 納期上書きシートのラウンドトリップ保全（保存で消えない）
"""

import datetime

from openpyxl import load_workbook

from nouki_kaitou.delivery_calc import (
    calculate_delivery_date,
    get_delivery_override,
)
from nouki_kaitou.history import (
    OVERRIDE_SHEET_NAME,
    extract_override_rows,
    initialize_delivery_history,
    load_delivery_overrides,
    save_history_batch,
)
from nouki_kaitou.models import BranchSettings, CacheStore, OrderRow
from nouki_kaitou.report_generator import create_delivery_report

# ============================================
# テスト用ヘルパー（test_twf.py と同パターン）
# ============================================
TODAY = datetime.date(2026, 6, 26)
EXEC = datetime.datetime(2026, 6, 26, 10, 0, 0)
WRONG = datetime.date(2026, 6, 20)  # 計上済み行に焼き付いた間違った指定納期（過去日）
BRANCH = BranchSettings(
    name="京葉営業所", default_cutoff=15, base_center="関東商品センター"
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
        delivery_overrides=kwargs.get("delivery_overrides", {}),
    )


def _make_row(**kwargs) -> OrderRow:
    defaults = {
        "order_number": "9999001",
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
        "order_delivery_date": datetime.date(2026, 6, 20),
        "specified_delivery_date": None,
        "item_group_code": "D01",
    }
    defaults.update(kwargs)
    return OrderRow(**defaults)


def _twf_answers(data, cache, sent, tmp_path) -> list[str]:
    """TWF専用回答書を生成し、データ行のI列(納期回答)を順に返す。"""
    result = create_delivery_report(
        data, "テスト商事", sent,
        cache, tmp_path, {}, BRANCH, EXEC,
        today=TODAY,
        include_only_orders={r.order_number.strip() for r in data},
        filter_already_sent=False,
        twf_mode=True,
    )
    ws = load_workbook(result.file_path).active
    answers = []
    for r in range(7, ws.max_row + 1):
        if ws.cell(r, 12).value:  # L列=弊社注番 がある＝データ行
            answers.append(ws.cell(r, 9).value)
    return answers


# ============================================
# get_delivery_override（単体）
# ============================================
class TestGetDeliveryOverride:
    def test_empty_cache_returns_none(self):
        """上書き辞書が空なら常にNone（no-op）"""
        cache = _make_cache()
        assert get_delivery_override("9999001", "10", cache) is None

    def test_hit_returns_verbatim(self):
        """ヒットすれば逐語文字列を返す"""
        cache = _make_cache(delivery_overrides={"9999001|10": "6月30日納品予定"})
        assert get_delivery_override("9999001", "10", cache) == "6月30日納品予定"

    def test_key_is_stripped(self):
        """注番・明細の前後空白を除去してキー照合する"""
        cache = _make_cache(delivery_overrides={"9999001|10": "X"})
        assert get_delivery_override(" 9999001 ", " 10 ", cache) == "X"

    def test_other_key_unaffected(self):
        """別の明細・別の注番はヒットしない"""
        cache = _make_cache(delivery_overrides={"9999001|10": "X"})
        assert get_delivery_override("9999001", "20", cache) is None
        assert get_delivery_override("9999002", "10", cache) is None

    def test_blank_value_treated_as_none(self):
        """空白だけの上書き値はNone扱い"""
        cache = _make_cache(delivery_overrides={"9999001|10": "   "})
        assert get_delivery_override("9999001", "10", cache) is None


# ============================================
# calculate_delivery_date（優先0で最優先）
# ============================================
class TestCalculatePriority:
    def test_override_beats_specified_date_on_completed_row(self):
        """計上済み（処理完了）＋指定納期入りでも、上書きが最優先で勝つ"""
        row = _make_row(
            ship_status="処理完了", specified_delivery_date=WRONG
        )
        # 上書きなし → 指定納期パスで間違った過去日
        base = _make_cache()
        assert calculate_delivery_date(row, base, today=TODAY, execution_time=EXEC) \
            == "6月20日配達済み"
        # 上書きあり → 逐語文字列が最優先
        over = _make_cache(delivery_overrides={"9999001|10": "6月30日納品予定"})
        assert calculate_delivery_date(row, over, today=TODAY, execution_time=EXEC) \
            == "6月30日納品予定"

    def test_no_override_is_noop_for_normal_row(self):
        """上書き辞書が空なら通常行の計算結果は一切変わらない"""
        row = _make_row(specified_delivery_date=WRONG)
        empty = _make_cache()
        # delivery_overrides={} は CacheStore のデフォルト。明示有無で差が出ないこと
        assert calculate_delivery_date(row, empty, today=TODAY, execution_time=EXEC) \
            == calculate_delivery_date(row, _make_cache(delivery_overrides={}),
                                       today=TODAY, execution_time=EXEC)


# ============================================
# TWF回答書（report_generator 統合）
# ============================================
class TestTwfIntegration:
    def _problem_set(self):
        """計上済み・指定納期入りの問題行(明細10) ＋ 正常な同注番3行"""
        problem = _make_row(
            detail_number="10", ship_status="処理完了",
            specified_delivery_date=WRONG, comment_detail="TWFNo.001 テスト商事様",
        )
        others = [
            _make_row(
                detail_number=str(n), ship_status="処理完了",
                specified_delivery_date=WRONG, product_name=f"部材{n}",
                comment_detail="TWFNo.001 テスト商事様",
            )
            for n in (20, 30, 40)
        ]
        return problem, others

    def test_override_empty_is_noop(self, tmp_path):
        """上書き空 → 全4行が従来どおり（処理完了+履歴あり→納品済み）"""
        problem, others = self._problem_set()
        data = [problem] + others
        sent = {f"9999001|{n}": "6月20日配達済み" for n in (10, 20, 30, 40)}
        cache = _make_cache()  # delivery_overrides 空
        answers = _twf_answers(data, cache, sent, tmp_path)
        assert answers == ["納品済み"] * 4

    def test_override_wins_over_nouhinzumi(self, tmp_path):
        """上書き1行 → その行だけ逐語、他3行は従来どおり納品済み"""
        problem, others = self._problem_set()
        data = [problem] + others
        sent = {f"9999001|{n}": "6月20日配達済み" for n in (10, 20, 30, 40)}
        cache = _make_cache(
            delivery_overrides={"9999001|10": "6月30日納品予定"}
        )
        answers = _twf_answers(data, cache, sent, tmp_path)
        # 明細10だけ上書き、残りは納品済み
        assert answers.count("6月30日納品予定") == 1
        assert answers.count("納品済み") == 3

    def test_other_customer_unaffected(self, tmp_path):
        """別客の同形の行は、上書きキーに無いので影響なし"""
        # 別客の注番8888001（上書き辞書には9999001しか無い）
        row = _make_row(
            order_number="8888001", detail_number="10", ship_status="処理完了",
            specified_delivery_date=WRONG, customer_name="別商事",
            ship_to_name="別商事", comment_detail="TWFNo.099 別商事様",
        )
        sent = {"8888001|10": "6月20日配達済み"}
        cache = _make_cache(delivery_overrides={"9999001|10": "6月30日納品予定"})
        result = create_delivery_report(
            [row], "別商事", sent,
            cache, tmp_path, {}, BRANCH, EXEC,
            today=TODAY, include_only_orders={"8888001"},
            filter_already_sent=False, twf_mode=True,
        )
        ws = load_workbook(result.file_path).active
        answers = [ws.cell(r, 9).value for r in range(7, ws.max_row + 1)
                   if ws.cell(r, 12).value]
        assert answers == ["納品済み"]  # 上書きされない


# ============================================
# 納期上書きシートのI/O・ラウンドトリップ保全
# ============================================
class TestSheetIO:
    def test_initialize_creates_override_sheet(self, tmp_path):
        """新規送付履歴.xlsxに納期上書きシートが作られる"""
        path = str(tmp_path / "送付履歴.xlsx")
        wb = initialize_delivery_history(path)
        assert OVERRIDE_SHEET_NAME in wb.sheetnames
        ws = wb[OVERRIDE_SHEET_NAME]
        assert ws.cell(1, 1).value == "受発注伝票"
        assert ws.cell(1, 3).value == "納期回答(上書き)"

    def test_load_empty_sheet_is_empty_dict(self, tmp_path):
        """空シート → 空dict（完全no-op）"""
        path = str(tmp_path / "送付履歴.xlsx")
        wb = initialize_delivery_history(path)
        assert load_delivery_overrides(wb[OVERRIDE_SHEET_NAME]) == {}

    def test_load_none_sheet_is_empty_dict(self):
        """シートNone（既存ファイルに無い）→ 空dict"""
        assert load_delivery_overrides(None) == {}

    def test_load_skips_blank_answer(self, tmp_path):
        """納期回答(上書き)が空の行は無視（注番だけ書いても上書きしない）"""
        path = str(tmp_path / "送付履歴.xlsx")
        wb = initialize_delivery_history(path)
        ws = wb[OVERRIDE_SHEET_NAME]
        ws.cell(2, 1).value = "9999001"
        ws.cell(2, 2).value = 10
        ws.cell(2, 3).value = "6月30日納品予定"
        ws.cell(3, 1).value = "9999001"     # 注番だけ・回答空
        ws.cell(3, 2).value = 20
        overrides = load_delivery_overrides(ws)
        assert overrides == {"9999001|10": "6月30日納品予定"}

    def test_int_detail_normalized(self, tmp_path):
        """明細が数値セルでも文字列キーに正規化される"""
        path = str(tmp_path / "送付履歴.xlsx")
        wb = initialize_delivery_history(path)
        ws = wb[OVERRIDE_SHEET_NAME]
        ws.cell(2, 1).value = "9999001"
        ws.cell(2, 2).value = 10           # int
        ws.cell(2, 3).value = "X"
        assert load_delivery_overrides(ws) == {"9999001|10": "X"}

    def test_roundtrip_survives_save(self, tmp_path):
        """手入力した上書き行が save_history_batch（WB新規作成）後も残る"""
        path = str(tmp_path / "送付履歴.xlsx")
        wb = initialize_delivery_history(path)
        ws = wb[OVERRIDE_SHEET_NAME]
        ws.cell(2, 1).value = "9999001"
        ws.cell(2, 2).value = 10
        ws.cell(2, 3).value = "6月30日納品予定"
        ws.cell(2, 4).value = "計上済みSAP施錠"
        wb.save(path)

        # 読み直し → 保全用に行を抽出 → 保存（履歴更新なし）
        wb_ro = load_workbook(path, read_only=True)
        override_rows = extract_override_rows(wb_ro[OVERRIDE_SHEET_NAME])
        wb_ro.close()
        assert len(override_rows) == 1

        save_history_batch(
            path, [], [], [], [], EXEC, "tester",
            today=TODAY, override_rows=override_rows,
        )

        # 保存後も上書きが生きている
        wb2 = load_workbook(path)
        assert OVERRIDE_SHEET_NAME in wb2.sheetnames
        assert load_delivery_overrides(wb2[OVERRIDE_SHEET_NAME]) == {
            "9999001|10": "6月30日納品予定"
        }

    def test_save_without_override_rows_makes_empty_sheet(self, tmp_path):
        """override_rows未指定でも空シートが作られクラッシュしない"""
        path = str(tmp_path / "送付履歴.xlsx")
        initialize_delivery_history(path)
        save_history_batch(path, [], [], [], [], EXEC, "tester", today=TODAY)
        wb = load_workbook(path)
        assert OVERRIDE_SHEET_NAME in wb.sheetnames
        assert load_delivery_overrides(wb[OVERRIDE_SHEET_NAME]) == {}
