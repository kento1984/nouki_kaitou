"""社外コメントのパターン分類"""
import sys
import re
from pathlib import Path
from collections import Counter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from nouki_kaitou.data_loader import load_source_file, get_column_positions, parse_order_row

source_dir = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧")
xls_files = sorted(source_dir.glob("*.xls"), key=lambda p: p.stat().st_mtime, reverse=True)
source_file = xls_files[0]
print(f"ファイル: {source_file.name}\n")

data = load_source_file(str(source_file))
cols = get_column_positions(data)

# パターン分類
tracking_pattern = re.compile(r"(西濃|佐川|ヤマト|福通|トナミ|ゆうぱっく|JP|ＪＰ|日通)[\s：:]+\S*")
pickup_pattern = re.compile(r"引取")

categories = Counter()
non_tracking = []

for i in range(5, len(data)):
    row = parse_order_row(data, i, cols)
    ext = row.comment_external.strip()
    if not ext:
        continue

    if tracking_pattern.search(ext):
        # 送り状番号だけか、+αがあるか
        cleaned = tracking_pattern.sub("", ext).strip()
        if cleaned:
            categories["送り状+追加情報"] += 1
            non_tracking.append((row.order_number, ext, cleaned))
        else:
            categories["送り状番号のみ"] += 1
    elif pickup_pattern.search(ext):
        categories["引取"] += 1
    else:
        categories["その他（送り状・引取以外）"] += 1
        non_tracking.append((row.order_number, ext, ext))

print("=== パターン分類 ===")
for cat, count in categories.most_common():
    print(f"  {cat}: {count} 件")
print(f"  合計: {sum(categories.values())} 件")
print()

# 送り状以外の内容を表示
seen = set()
print(f"=== 送り状・引取以外の内容（備考欄に載せる価値のある情報）===")
count = 0
for order_num, full, extra in non_tracking:
    if full not in seen:
        seen.add(full)
        count += 1
        print(f"  [{count}] 注番={order_num}")
        print(f"       全文: {full}")
        if extra != full:
            print(f"       送り状除去後: {extra}")
        print()
        if count >= 20:
            break
