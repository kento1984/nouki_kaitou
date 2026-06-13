"""twf_ledger（TWF2026 社内手配管理表）のテスト

期間限定機能。TWF運用終了時はこのテストファイルごと削除してよい。
"""

from pathlib import Path

from openpyxl import load_workbook

from nouki_kaitou.models import OrderRow
from nouki_kaitou.twf_ledger import (
    MANUFACTURER_UNKNOWN,
    STATUS_CHOICES,
    STATUS_DEFAULT,
    LedgerRow,
    apply_carryover,
    build_ledger_rows,
    classify_tehai,
    format_twf_no,
    read_existing_status,
    resolve_manufacturer,
    write_ledger,
)


def _order(
    order_number="GL001",
    detail_number="10",
    document_type="【受注】直送販売",
    storage_place="転送中（直送用）",
    comment_detail="TWFNo.003243　新成（株）様",
    rep_name="柏原　賢人",
    manufacturer_name="ベルテクノ（株）",
    product_name="溶接機",
    quantity="2",
    item_group_code="0001",
):
    return OrderRow(
        order_number=order_number,
        detail_number=detail_number,
        document_type=document_type,
        storage_place=storage_place,
        comment_detail=comment_detail,
        rep_name=rep_name,
        manufacturer_name=manufacturer_name,
        product_name=product_name,
        quantity=quantity,
        item_group_code=item_group_code,
    )


# ============================================
# 手配区分の判定
# ============================================
def test_classify_chokusou():
    o = _order(document_type="【受注】直送販売", storage_place="転送中（直送用）")
    assert classify_tehai(o) == "直送"


def test_classify_himozuki():
    o = _order(document_type="【受注】直送販売", storage_place="GL0")
    assert classify_tehai(o) == "紐付き"


def test_classify_zaiko():
    o = _order(document_type="【受注】在庫販売", storage_place="")
    assert classify_tehai(o) == "在庫販売"


# ============================================
# メーカー名の補完（在庫販売で「名称」が空のとき品目Groupで補完）
# ============================================
def test_resolve_keeps_existing_name():
    """名称があればマスターに関係なくそのまま（直送・紐付き）。"""
    o = _order(manufacturer_name="日鉄溶接工業（株）", item_group_code="0316")
    assert resolve_manufacturer(o, {"0316": "別メーカー"}) == "日鉄溶接工業（株）"


def test_resolve_backfills_blank_from_master():
    """名称が空（在庫販売）→ 品目Groupマスターで補完。"""
    o = _order(document_type="【受注】在庫販売", storage_place="関東商品センター",
               manufacturer_name="", item_group_code="0075")
    assert resolve_manufacturer(o, {"0075": "ダイヘン"}) == "ダイヘン"


def test_resolve_unknown_when_not_in_master():
    """名称が空でマスターにも無ければ「（要確認）」。"""
    o = _order(document_type="【受注】在庫販売", storage_place="関東商品センター",
               manufacturer_name="", item_group_code="1774")
    assert resolve_manufacturer(o, {"0075": "ダイヘン"}) == MANUFACTURER_UNKNOWN


def test_resolve_no_master_leaves_blank():
    """マスター未提供（None）なら補完せず空欄のまま。"""
    o = _order(document_type="【受注】在庫販売", manufacturer_name="",
               item_group_code="0075")
    assert resolve_manufacturer(o, None) == ""


def test_build_applies_backfill():
    orders = [
        _order(order_number="GLZ1", document_type="【受注】在庫販売",
               storage_place="関東商品センター", manufacturer_name="",
               item_group_code="0075",
               comment_detail="TWFNo.001　甲社様"),
        _order(order_number="GLZ2", document_type="【受注】在庫販売",
               storage_place="関東商品センター", manufacturer_name="",
               item_group_code="1774",
               comment_detail="TWFNo.002　乙社様"),
    ]
    rows = build_ledger_rows(orders, {"0075": "ダイヘン"})
    by_order = {r.order_number: r for r in rows}
    assert by_order["GLZ1"].manufacturer == "ダイヘン"
    assert by_order["GLZ2"].manufacturer == MANUFACTURER_UNKNOWN


# ============================================
# 台帳行の構築
# ============================================
def test_build_filters_non_twf():
    orders = [
        _order(order_number="GL001", comment_detail="TWFNo.003243　新成（株）様"),
        _order(order_number="GL999", comment_detail="通常受注（TWFなし）"),
    ]
    rows = build_ledger_rows(orders)
    assert len(rows) == 1
    assert rows[0].order_number == "GL001"
    assert rows[0].twf_no == "003243"
    assert rows[0].customer == "新成（株）様"
    assert rows[0].status == STATUS_DEFAULT


def test_build_propagates_within_order():
    """同一注番の入れ忘れ明細にも TWF No.・お客様名が引き継がれる。"""
    orders = [
        _order(order_number="GL010", detail_number="10",
               comment_detail="TWFNo.004041　金安工業様"),
        _order(order_number="GL010", detail_number="20",
               comment_detail="（TWF記載なし・入れ忘れ）"),
    ]
    rows = build_ledger_rows(orders)
    assert len(rows) == 2
    by_detail = {r.detail_number: r for r in rows}
    assert by_detail["20"].twf_no == "004041"
    assert by_detail["20"].customer == "金安工業様"


def test_build_question_mark_number_goes_to_number_column():
    """？はTWF№側へ。お客様名には？が付かない（4行の改善）。"""
    orders = [
        _order(order_number="GLQ1", detail_number="10",
               comment_detail="TWFNo.0014？？　㈱ハイプラン様"),
    ]
    rows = build_ledger_rows(orders)
    assert rows[0].twf_no == "0014??"
    assert rows[0].customer == "㈱ハイプラン様"


# ============================================
# 桁揃え（台帳のみ・6桁ゼロ埋め）
# ============================================
def test_format_twf_no_pads_short_pure_digits():
    assert format_twf_no("0014") == "000014"
    assert format_twf_no("2") == "000002"


def test_format_twf_no_keeps_six_or_more():
    assert format_twf_no("000022") == "000022"
    assert format_twf_no("0005017") == "0005017"  # 7桁はそのまま


def test_format_twf_no_skips_question_unknown_blank():
    assert format_twf_no("0014??") == "0014??"   # ？含みは埋めない
    assert format_twf_no("002???") == "002???"
    assert format_twf_no("不明") == "不明"
    assert format_twf_no("") == ""


def test_build_applies_zero_pad_display():
    orders = [
        _order(order_number="GLP1", comment_detail="TWFNo.0014　甲社様"),
    ]
    rows = build_ledger_rows(orders)
    assert rows[0].twf_no == "000014"


def test_build_sort_order():
    """TWF No.昇順 → 番号なし・不明は末尾。"""
    orders = [
        _order(order_number="A", comment_detail="TWFNo.005　甲社様"),
        _order(order_number="B", comment_detail="TWFNo.001　乙社様"),
        _order(order_number="C", comment_detail="TWFNo.不明　丙社様"),
    ]
    rows = build_ledger_rows(orders)
    # 台帳表示は6桁ゼロ埋め。不明は末尾。
    assert [r.twf_no for r in rows] == ["000001", "000005", "不明"]


# ============================================
# 引き継ぎ
# ============================================
def test_carryover_round_trip(tmp_path):
    orders = [
        _order(order_number="GL100", detail_number="10",
               comment_detail="TWFNo.003243　新成（株）様"),
    ]
    rows = build_ledger_rows(orders)
    # 手入力を模してステータス・備考を設定 → 書き出し
    rows[0].status = "メーカー発注済"
    rows[0].note = "6/20入荷予定"
    first = write_ledger(rows, tmp_path / "first.xlsx")

    # 再生成（手入力はデフォルトに戻る）→ 引き継ぎ適用
    rows2 = build_ledger_rows(orders)
    assert rows2[0].status == STATUS_DEFAULT
    existing = read_existing_status(first)
    applied = apply_carryover(rows2, existing)
    assert applied == 1
    assert rows2[0].status == "メーカー発注済"
    assert rows2[0].note == "6/20入荷予定"


def test_read_existing_missing_file_returns_empty(tmp_path):
    assert read_existing_status(tmp_path / "nope.xlsx") == {}


# ============================================
# Excel 出力の構造
# ============================================
def test_write_creates_table_and_dropdown(tmp_path):
    orders = [
        _order(order_number="GL200", comment_detail="TWFNo.001　甲社様"),
        _order(order_number="GL201", comment_detail="TWFNo.002　乙社様"),
    ]
    rows = build_ledger_rows(orders)
    out = write_ledger(rows, tmp_path / "ledger.xlsx")
    assert out.exists()

    wb = load_workbook(out)
    ws = wb.active
    # テーブルが1つ存在し、B8 始まり
    assert len(ws.tables) == 1
    tbl = list(ws.tables.values())[0]
    assert tbl.ref.startswith("B8")
    # ステータスのドロップダウンが存在
    dvs = list(ws.data_validations.dataValidation)
    assert any(dv.type == "list" for dv in dvs)
    # ヘッダーはB8（左A列・上の余白を確保している）
    assert ws["B8"].value == "TWF No."
    assert ws["A8"].value is None  # 左余白
    # ヘッダー固定
    assert ws.freeze_panes == "A9"


def test_status_choices_count():
    # 5項目に統一（メーカー手配済み/在庫計上済み/保留/その他 ＋ 初期値の未着手）
    assert STATUS_CHOICES == [
        "未着手", "メーカー手配済み", "在庫計上済み", "保留", "その他",
    ]
    assert STATUS_DEFAULT == "未着手"
    # 廃止した選択肢が残っていないこと
    for removed in ("一部手配", "対象外・キャンセル", "メーカー発注済", "完了"):
        assert removed not in STATUS_CHOICES
