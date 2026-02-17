"""エッジケースの実データ検索

欠品、分納、処理完了タイミングの具体例を検索
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from nouki_kaitou.data_loader import load_source_file, get_column_positions

def main():
    print("16AM.xlsを読み込み中...")
    xls_path = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧\16AM.xls")
    source_data = load_source_file(xls_path)
    cols = get_column_positions(source_data)

    if cols is None:
        print("列位置を取得できませんでした")
        return

    col_order = cols.get("受発注伝票", 0)
    col_detail = cols.get("明細", 1)
    col_customer = cols.get("受注先")
    col_product = cols.get("品名")
    col_ship_status = cols.get("出荷ステータス")
    col_comment_detail = cols.get("コメント（明細）")
    col_comment_internal = cols.get("コメント（社内）")
    col_specified = cols.get("指定納期")
    col_order_date = cols.get("受注納期")
    col_doc_type = cols.get("伝票タイプ")

    # カテゴリ別に収集
    kepppin_orders = []  # 欠品中
    bunno_orders = []    # 分納
    dec31_completed = [] # 12/31のまま処理完了
    mitei_completed = [] # 未定のまま処理完了
    bunno_mixed = []     # 分納で処理完了と未処理が混在

    # 注番別に明細を収集
    order_details = {}

    for row_idx in range(5, len(source_data)):
        row = source_data[row_idx]
        if len(row) <= col_order:
            continue

        order_num = str(row[col_order] or "").strip()
        if not order_num:
            continue

        detail_num = str(row[col_detail] if col_detail < len(row) else "").strip()
        customer = str(row[col_customer] if col_customer and col_customer < len(row) else "").strip()[:25]
        product = str(row[col_product] if col_product and col_product < len(row) else "").strip()[:35]
        ship_status = str(row[col_ship_status] if col_ship_status and col_ship_status < len(row) else "").strip()
        comment_detail = str(row[col_comment_detail] if col_comment_detail and col_comment_detail < len(row) else "").strip()
        comment_internal = str(row[col_comment_internal] if col_comment_internal and col_comment_internal < len(row) else "").strip()
        specified = str(row[col_specified] if col_specified and col_specified < len(row) else "").strip()
        order_date = str(row[col_order_date] if col_order_date and col_order_date < len(row) else "").strip()
        doc_type = str(row[col_doc_type] if col_doc_type and col_doc_type < len(row) else "").strip()

        item = {
            "order": order_num,
            "detail": detail_num,
            "customer": customer,
            "product": product,
            "ship_status": ship_status,
            "comment_detail": comment_detail[:50],
            "comment_internal": comment_internal[:30],
            "specified": specified,
            "order_date": order_date,
            "doc_type": doc_type,
        }

        # 注番別に収集
        if order_num not in order_details:
            order_details[order_num] = []
        order_details[order_num].append(item)

        # 欠品中（コメント明細に「欠品」）
        if "欠品" in comment_detail:
            kepppin_orders.append(item)

        # 分納（コメント明細に「分納:」）
        if "分納:" in comment_detail or "分納：" in comment_detail:
            bunno_orders.append(item)

        # 12/31のまま処理完了
        if ship_status == "処理完了" and "12/31" in specified:
            dec31_completed.append(item)

        # 未定のまま処理完了
        if ship_status == "処理完了" and "未定" in comment_internal:
            mitei_completed.append(item)

    # 分納で処理完了と未処理が混在する注番を探す
    for order_num, details in order_details.items():
        has_bunno = any("分納:" in d["comment_detail"] or "分納：" in d["comment_detail"] for d in details)
        if not has_bunno:
            continue
        completed = sum(1 for d in details if d["ship_status"] == "処理完了")
        pending = sum(1 for d in details if d["ship_status"] == "未処理")
        if completed > 0 and pending > 0:
            bunno_mixed.append({
                "order": order_num,
                "details": details,
                "completed": completed,
                "pending": pending,
            })

    # 結果表示
    print("\n" + "=" * 100)
    print("【1】欠品中の注番（コメント明細に「欠品」）")
    print("=" * 100)
    print(f"該当件数: {len(kepppin_orders)}件\n")

    for i, item in enumerate(kepppin_orders[:10]):
        print(f"{i+1}. {item['order']}|{item['detail']}")
        print(f"   顧客: {item['customer']}")
        print(f"   品名: {item['product']}")
        print(f"   出荷ステータス: {item['ship_status']}")
        print(f"   伝票タイプ: {item['doc_type']}")
        print(f"   コメント(明細): {item['comment_detail']}")
        print(f"   コメント(社内): {item['comment_internal']}")
        print(f"   指定納期: {item['specified']} / 受注納期: {item['order_date']}")
        print()

    if len(kepppin_orders) > 10:
        print(f"... 他 {len(kepppin_orders) - 10}件\n")

    print("\n" + "=" * 100)
    print("【2】分納の注番（コメント明細に「分納:」）")
    print("=" * 100)
    print(f"該当件数: {len(bunno_orders)}件\n")

    for i, item in enumerate(bunno_orders[:10]):
        print(f"{i+1}. {item['order']}|{item['detail']}")
        print(f"   顧客: {item['customer']}")
        print(f"   品名: {item['product']}")
        print(f"   出荷ステータス: {item['ship_status']}")
        print(f"   伝票タイプ: {item['doc_type']}")
        print(f"   コメント(明細): {item['comment_detail']}")
        print(f"   コメント(社内): {item['comment_internal']}")
        print()

    if len(bunno_orders) > 10:
        print(f"... 他 {len(bunno_orders) - 10}件\n")

    print("\n" + "=" * 100)
    print("【3】分納で処理完了と未処理が混在する注番")
    print("=" * 100)
    print(f"該当件数: {len(bunno_mixed)}件\n")

    for i, mixed in enumerate(bunno_mixed[:5]):
        print(f"{i+1}. 注番: {mixed['order']} (完了{mixed['completed']}件 / 未処理{mixed['pending']}件)")
        for d in mixed['details'][:5]:
            status_mark = "○" if d['ship_status'] == "処理完了" else "×"
            print(f"   {status_mark} |{d['detail']}| {d['product'][:25]} - {d['ship_status']}")
        if len(mixed['details']) > 5:
            print(f"   ... 他 {len(mixed['details']) - 5}明細")
        print()

    print("\n" + "=" * 100)
    print("【4】指定納期12/31のまま処理完了")
    print("=" * 100)
    print(f"該当件数: {len(dec31_completed)}件\n")

    for i, item in enumerate(dec31_completed[:10]):
        print(f"{i+1}. {item['order']}|{item['detail']}")
        print(f"   品名: {item['product']}")
        print(f"   指定納期: {item['specified']} / 受注納期: {item['order_date']}")
        print(f"   コメント(社内): {item['comment_internal']}")
        print()

    if len(dec31_completed) > 10:
        print(f"... 他 {len(dec31_completed) - 10}件\n")

    print("\n" + "=" * 100)
    print("【5】「未定」がコメントに残ったまま処理完了")
    print("=" * 100)
    print(f"該当件数: {len(mitei_completed)}件\n")

    for i, item in enumerate(mitei_completed[:10]):
        print(f"{i+1}. {item['order']}|{item['detail']}")
        print(f"   品名: {item['product']}")
        print(f"   コメント(社内): {item['comment_internal']}")
        print()

if __name__ == "__main__":
    main()
