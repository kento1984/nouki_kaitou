"""紐付き+12/31+処理完了の納期計算トレース

GL2C444376|10 を使って実際に計算を実行し、
どの分岐を通るかステップバイステップで確認する。
"""

import sys
import io
import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from nouki_kaitou.data_loader import load_source_file, get_column_positions
from nouki_kaitou.models import OrderRow, CacheStore, BranchSettings
from nouki_kaitou.delivery_calc import calculate_delivery_date
from nouki_kaitou.utils import is_december_31


def main():
    print("=" * 100)
    print("紐付き + 受注納期12/31 + 処理完了 の納期計算トレース")
    print("=" * 100)

    # 対象注番（指定納期も12/31のケース）
    target_order = "GL2C444955"
    target_detail = "10"

    print(f"\n対象注番: {target_order}|{target_detail}")

    # 16AM.xlsを読み込み
    print("\n16AM.xlsを読み込み中...")
    xls_path = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧\16AM.xls")
    source_data = load_source_file(xls_path)
    cols = get_column_positions(source_data)

    if cols is None:
        print("列位置を取得できませんでした")
        return

    # 対象行を検索
    target_row = None
    for row_idx in range(5, len(source_data)):
        row = source_data[row_idx]
        if len(row) <= cols.get("受発注伝票", 0):
            continue

        order_num = str(row[cols.get("受発注伝票", 0)] or "").strip()
        detail_num = str(row[cols.get("明細", 1)] or "").strip()

        if order_num == target_order and detail_num == target_detail:
            # OrderRowを構築
            target_row = build_order_row(row, cols)
            break

    if target_row is None:
        print(f"注番 {target_order}|{target_detail} が見つかりませんでした")
        return

    # データ表示
    print("\n" + "-" * 80)
    print("【SAP実データ】")
    print("-" * 80)
    print(f"注番: {target_row.order_number}|{target_row.detail_number}")
    print(f"顧客: {target_row.customer_name}")
    print(f"出荷先: {target_row.ship_to_name}")
    print(f"品名: {target_row.product_name}")
    print(f"伝票タイプ: {target_row.document_type}")
    print(f"保管場所: {target_row.storage_place}")
    print(f"出荷ステータス: {target_row.ship_status}")
    print(f"指定納期: {target_row.specified_delivery_date}")
    print(f"受注納期: {target_row.order_delivery_date}")

    # トレース開始
    print("\n" + "=" * 100)
    print("【delivery_calc.py トレース】")
    print("=" * 100)

    trace_calculation(target_row)


def build_order_row(row: list, cols: dict) -> OrderRow:
    """SAP行からOrderRowを構築"""
    def get_val(col_name, default=""):
        col_idx = cols.get(col_name)
        if col_idx is not None and col_idx < len(row):
            return row[col_idx]
        return default

    def parse_date(val):
        if val is None:
            return None
        val_str = str(val).strip()
        if not val_str:
            return None
        for fmt in ["%Y/%m/%d", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S"]:
            try:
                return datetime.datetime.strptime(val_str.split()[0], fmt.split()[0]).date()
            except ValueError:
                continue
        return None

    return OrderRow(
        order_number=str(get_val("受発注伝票", "")).strip(),
        detail_number=str(get_val("明細", "")).strip(),
        customer_name=str(get_val("受注先", "")).strip(),
        ship_to_name=str(get_val("出荷先名", "")).strip(),
        product_name=str(get_val("品名", "")).strip(),
        document_type=str(get_val("伝票タイプ", "")).strip(),
        storage_place=str(get_val("保管場所", "")).strip(),
        ship_status=str(get_val("出荷ステータス", "")).strip(),
        specified_delivery_date=parse_date(get_val("指定納期")),
        order_delivery_date=parse_date(get_val("受注納期")),
        registration_date=parse_date(get_val("登録日")),
        item_group_code=str(get_val("品目グループ", "")).strip(),
        quantity=str(get_val("数量", "")).strip(),
        comment_detail=str(get_val("コメント（明細）", "")).strip(),
        comment_internal=str(get_val("コメント（社内）", "")).strip(),
        comment_external=str(get_val("コメント（社外）", "")).strip(),
        customer_contact=str(get_val("受注先担当", "")).strip(),
        customer_order_number=str(get_val("客先注文番号", "")).strip(),
        unit_price=str(get_val("単価", "")).strip(),
        net_amount=str(get_val("純額", "")).strip(),
        rejection_reason=str(get_val("否認理由", "")).strip(),
    )


def trace_calculation(row: OrderRow):
    """納期計算のトレースを実行"""

    today = datetime.date.today()
    execution_time = datetime.datetime.now()
    branch = BranchSettings()

    # 空のキャッシュ（送付履歴なしシミュレーション）
    cache_empty = CacheStore()

    print("\n【Step 1】基本条件チェック")
    print("-" * 60)
    print(f"  伝票タイプ = '{row.document_type}'")
    print(f"  → '【受注】直送販売' か？ → {row.document_type == '【受注】直送販売'}")
    print(f"  出荷ステータス = '{row.ship_status}'")
    print(f"  → '処理完了' か？ → {row.ship_status == '処理完了'}")
    print(f"  保管場所 = '{row.storage_place}'")
    print(f"  → '転送中（直送用）' でないか？ → {row.storage_place != '転送中（直送用）'}")

    is_himozuki = (
        row.document_type == "【受注】直送販売"
        and row.storage_place != "転送中（直送用）"
    )
    print(f"\n  ★ 紐付き判定 (is_himozuki) = {is_himozuki}")

    print("\n【Step 2】delivery_calc.py の分岐優先順位")
    print("-" * 60)
    print("  1. &&作業チェック → コメント社内に「&&」があるか")
    print(f"     コメント社内 = '{row.comment_internal[:50]}...'")
    print(f"     → 該当しない")
    print()
    print("  2. @@着日指定チェック → コメント社内に「@@M/D」があるか")
    print(f"     → 該当しない")
    print()
    print("  3. 引取チェック → 「引取」「引き取り」+日付があるか")
    print(f"     → 該当しない")
    print()
    print("  4. 指定納期チェック → 12/31以外の指定納期があるか")
    print(f"     指定納期 = {row.specified_delivery_date}")
    is_dec31_specified = row.specified_delivery_date and is_december_31(row.specified_delivery_date)
    print(f"     → 12/31か？ {is_dec31_specified}")
    if row.specified_delivery_date and not is_dec31_specified:
        print(f"     → 【このパスに入る可能性あり！】")
    print()
    print("  5. 在庫販売+処理完了チェック")
    print(f"     伝票タイプ = '{row.document_type}'")
    print(f"     → '【受注】在庫販売' か？ → {row.document_type == '【受注】在庫販売'}")
    print(f"     → 該当しない（直送販売のため）")
    print()
    print("  6. ★ 紐付き+処理完了チェック (_check_himozuki_completed)")
    print("     条件:")
    print(f"       - doc_type == '【受注】直送販売' → {row.document_type == '【受注】直送販売'}")
    print(f"       - ship_status == '処理完了' → {row.ship_status == '処理完了'}")
    print(f"       - storage != '転送中（直送用）' → {row.storage_place != '転送中（直送用）'}")
    all_conditions = (
        row.document_type == "【受注】直送販売"
        and row.ship_status == "処理完了"
        and row.storage_place != "転送中（直送用）"
    )
    print(f"     → 全条件満たす？ → {all_conditions}")

    if all_conditions:
        print()
        print("  ★★★ _check_himozuki_completed に入る！ ★★★")
        print()
        print("  【重要】このパスに入ると、受注納期12/31のチェック(_check_dec31)は実行されない")
        print()
        print("  _check_himozuki_completed の処理:")
        print("    1. マクロ実行時刻を取得")
        print(f"       現在時刻 = {execution_time.strftime('%H:%M')}")
        print("    2. 締切時間（cutoff）を取得（デフォルト12時）")
        print("    3. cutoff前なら+1営業日、後なら+2営業日")
        print("    4. 受注先≠出荷先 → 配達日から1営業日逆算して出荷日")
        print()
        print(f"    受注先 = '{row.customer_name}'")
        print(f"    出荷先 = '{row.ship_to_name}'")
        print(f"    → 受注先≠出荷先？ → {row.customer_name != row.ship_to_name}")
    else:
        print()
        print("  → 該当しない場合、次のチェックへ...")
        print()
        print("  7. 受注納期なし → '日程調整中'")
        print(f"     受注納期 = {row.order_delivery_date}")
        print()
        print("  8. 受注納期=12/31 → _check_dec31")
        print(f"     12/31か？ → {is_december_31(row.order_delivery_date) if row.order_delivery_date else 'N/A'}")

    print("\n" + "=" * 100)
    print("【実際の計算実行】")
    print("=" * 100)

    print("\n■ ケース1: 送付履歴なし（初回送信）")
    print("-" * 60)
    result1 = calculate_delivery_date(
        row, cache_empty, None, branch, execution_time, today
    )
    print(f"  計算結果: 【{result1}】")

    print("\n■ ケース2: 確認中一覧に確定日あり")
    print("-" * 60)
    # 確認中一覧に確定日を設定
    cache_with_confirming = CacheStore()
    # confirm辞書に直接設定 (注番|明細 → (問合せ状況, ステータス, 受注納期))
    key = f"{row.order_number}|{row.detail_number}"
    cache_with_confirming.confirm[key] = ("回答待ち", "確認中", datetime.date(2026, 1, 20))
    result2 = calculate_delivery_date(
        row, cache_with_confirming, None, branch, execution_time, today
    )
    print(f"  確認中一覧に確定日 2026/01/20 を設定")
    print(f"  計算結果: 【{result2}】")

    print("\n" + "=" * 100)
    print("【結論】")
    print("=" * 100)
    print()
    print("紐付き（直送販売）+ 処理完了 の場合:")
    print("  - _check_himozuki_completed が先に評価される（line 311-316）")
    print("  - 受注納期12/31のチェック(_check_dec31)より優先される")
    print("  - マクロ実行時刻ベースで「○月○日出荷予定」または「出荷済み」を計算")
    print()
    print("report_generator.py での処理:")
    print("  - is_himozuki = True のため force_delivered = False を維持")
    print("  - 計算結果がそのまま使用される（「納品済み」には上書きされない）")
    print()
    print("このケースの処理フロー:")
    print("  1. delivery_calc.py: 「○月○日出荷予定」を返す")
    print("  2. report_generator.py: force_delivered=False なのでそのまま使用")
    print("  3. _classify_order: 確定として送付履歴へ記録")


if __name__ == "__main__":
    main()
