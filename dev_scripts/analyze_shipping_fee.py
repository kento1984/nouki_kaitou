"""送料行の出荷ステータス分析

送料52件について:
1. 明細削除や削除になっているものがあるか
2. 同注番の他明細（本体）の出荷ステータス
3. 送料だけ残って本体は処理完了のパターンを集計
"""

import sys
from pathlib import Path
# プロジェクトルートをパスに追加（nouki_kaitouの親ディレクトリ）
project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from collections import defaultdict
from nouki_kaitou.data_loader import load_source_file, get_column_positions

# 送料行の注番リスト（ghost_analysis.txtから抽出）
SHIPPING_FEE_ORDERS = [
    ("GL2C444374", "20"),
    ("GL2C444729", "50"),
    ("GL2C444779", "20"),
    ("GL2V444923", "20"),
    ("GL2C445035", "90"),
    ("GL2C445037", "20"),
    ("GL2V445086", "20"),
    ("GL2C445181", "30"),
    ("GL2C445484", "20"),
    ("GL2Z445548", "30"),
    ("GL2C445636", "20"),
    ("GL2F445707", "40"),
    ("GL2F445857", "20"),
    ("GL2H446077", "20"),
    ("GL2C446105", "20"),
    ("GL2A446117", "20"),
    ("GL2D446204", "20"),
    ("GL2C446247", "20"),
    ("GL2H446251", "30"),
    ("GL2R446317", "20"),
    ("GL2Z446321", "30"),
    ("GL2C446328", "20"),
    ("GL2H446398", "20"),
    ("GL2C446427", "30"),
    ("GL2A446604", "20"),
    ("GL2H446684", "30"),
    ("GL2V446775", "20"),
    ("GL2H446779", "30"),
    ("GL2C446783", "20"),
    ("GL2H446788", "20"),
    ("GL2F446793", "20"),
    ("GL2H446805", "60"),
    ("GL2M446871", "20"),
    ("GL2V446908", "30"),
    ("GL2Z446911", "20"),
    ("GL2H446937", "20"),
    ("GL2F446945", "20"),
    ("GL2F446950", "20"),
    ("GL2F446969", "20"),
    ("GL2R447028", "50"),
    ("GL2R447038", "20"),
    ("GL2H447044", "20"),
    ("GL2H447050", "20"),
    ("GL2F447052", "20"),
    ("GL2S447058", "20"),
    ("GL2S447086", "20"),
    ("GL2F447109", "20"),
    ("GL2H447112", "20"),
    ("GL2Z447113", "20"),
    ("GL2V447123", "20"),
]

def load_order_data():
    """16AM.xlsから受注データを読み込む"""
    xls_path = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧\16AM.xls")

    source_data = load_source_file(xls_path)
    cols = get_column_positions(source_data)

    if cols is None:
        raise ValueError("列位置を取得できませんでした")

    # データを読み込み（注番をキーにした辞書）
    orders = defaultdict(list)

    # colsは辞書型（ColumnMap）
    header_row = 4  # 固定: 5行目（0-indexed: 4）
    col_order = cols.get("受発注伝票", 0)
    col_detail = cols.get("明細", 1)
    col_product = cols.get("品名", 2)
    col_ship_status = cols.get("出荷ステータス")
    col_rejection = cols.get("拒否理由")
    col_comment = cols.get("コメント（社内）")
    col_amount = cols.get("正味額")

    for row_idx in range(header_row + 1, len(source_data)):
        row = source_data[row_idx]
        if len(row) <= col_order:
            continue

        order_num = str(row[col_order] or "").strip()
        if not order_num or order_num == "受発注伝票":
            continue

        detail_num = str(row[col_detail] if col_detail < len(row) else "").strip()
        product_name = str(row[col_product] if col_product < len(row) else "").strip()
        ship_status = str(row[col_ship_status] if col_ship_status and col_ship_status < len(row) else "").strip()
        rejection_reason = str(row[col_rejection] if col_rejection and col_rejection < len(row) else "").strip()
        comment_internal = str(row[col_comment] if col_comment and col_comment < len(row) else "").strip()
        amount = row[col_amount] if col_amount and col_amount < len(row) else 0
        try:
            amount_str = str(amount).replace(",", "").replace('"', '')
            amount = float(amount_str) if amount_str else 0
        except:
            amount = 0

        orders[order_num].append({
            "detail_num": detail_num,
            "product_name": product_name[:40],  # 長すぎる場合は切り詰め
            "ship_status": ship_status,
            "rejection_reason": rejection_reason,
            "comment_internal": comment_internal[:30] if comment_internal else "",
            "amount": amount,
        })

    return orders

def analyze_shipping_fees():
    """送料行の分析"""
    print("16AM.xlsを読み込み中...")
    orders = load_order_data()
    print(f"読み込み完了: {len(orders)}件の注番\n")

    # 集計用
    deleted_count = 0  # 明細削除の送料
    main_completed_count = 0  # 本体完了で送料残り
    both_pending_count = 0  # 両方未処理
    shipping_only_count = 0  # 送料のみの注番

    results = []

    print("=" * 100)
    print("送料行と同注番明細の出荷ステータス一覧")
    print("=" * 100)

    for order_num, detail_num in SHIPPING_FEE_ORDERS:
        if order_num not in orders:
            print(f"\n【{order_num}】データなし")
            continue

        details = orders[order_num]
        shipping_row = None
        other_rows = []

        for d in details:
            if d["detail_num"] == detail_num:
                shipping_row = d
            else:
                other_rows.append(d)

        if not shipping_row:
            print(f"\n【{order_num}|{detail_num}】送料行が見つかりません")
            continue

        # 送料行の状態確認
        is_deleted = "削除" in shipping_row["rejection_reason"] or "削除" in shipping_row["ship_status"]
        shipping_status = shipping_row["ship_status"] or "（空白）"
        rejection = shipping_row["rejection_reason"]

        if is_deleted:
            deleted_count += 1

        # 本体（送料以外）の状態確認
        if not other_rows:
            shipping_only_count += 1
            pattern = "送料のみ"
        else:
            completed_count = sum(1 for r in other_rows if r["ship_status"] == "処理完了")
            pending_count = sum(1 for r in other_rows if r["ship_status"] in ("未処理", "", "（空白）") or not r["ship_status"])

            if completed_count == len(other_rows):
                main_completed_count += 1
                pattern = "★本体完了・送料残り"
            elif pending_count == len(other_rows):
                both_pending_count += 1
                pattern = "両方未処理"
            else:
                pattern = f"本体混在（完了{completed_count}/未処理{pending_count}）"

        # 結果表示
        print(f"\n【{order_num}】{pattern}")
        print(f"  送料行 |{detail_num}| ステータス: {shipping_status}", end="")
        if rejection:
            print(f" / 拒否理由: {rejection}", end="")
        if shipping_row["comment_internal"]:
            print(f" / コメント: {shipping_row['comment_internal']}", end="")
        print()

        if other_rows:
            print("  ---- 他明細 ----")
            for r in other_rows:
                status = r["ship_status"] or "（空白）"
                print(f"    |{r['detail_num']}| {r['product_name'][:30]} - {status}")

        results.append({
            "order_num": order_num,
            "detail_num": detail_num,
            "pattern": pattern,
            "is_deleted": is_deleted,
            "shipping_status": shipping_status,
            "other_count": len(other_rows),
        })

    # 集計結果
    print("\n" + "=" * 100)
    print("集計結果")
    print("=" * 100)
    print(f"送料行の総数: {len(SHIPPING_FEE_ORDERS)}件")
    print()
    print(f"【パターン別】")
    print(f"  ★本体処理完了・送料だけ未処理: {main_completed_count}件")
    print(f"  両方未処理: {both_pending_count}件")
    print(f"  送料のみ（本体なし）: {shipping_only_count}件")
    print(f"  本体混在: {len(results) - main_completed_count - both_pending_count - shipping_only_count}件")
    print()
    print(f"【明細削除/削除になっている送料行】: {deleted_count}件")

    # 本体完了・送料残りの詳細リスト
    if main_completed_count > 0:
        print("\n" + "=" * 100)
        print("★本体処理完了・送料だけ残りパターンの詳細")
        print("=" * 100)
        for r in results:
            if "本体完了" in r["pattern"]:
                print(f"  {r['order_num']}|{r['detail_num']}")

if __name__ == "__main__":
    analyze_shipping_fees()
