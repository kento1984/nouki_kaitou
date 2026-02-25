"""社外コメントのサンプルデータを取得するスクリプト"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from nouki_kaitou.data_loader import load_source_file, get_column_positions, parse_order_row

# 最新の受注一覧を使用
source_dir = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧")
xls_files = sorted(source_dir.glob("*.xls"), key=lambda p: p.stat().st_mtime, reverse=True)
if not xls_files:
    print("受注一覧ファイルが見つかりません")
    sys.exit(1)

source_file = xls_files[0]
print(f"ファイル: {source_file.name}")
print()

data = load_source_file(str(source_file))
cols = get_column_positions(data)
if cols is None:
    print("列位置が取得できません")
    sys.exit(1)

# 全行を読み込んで社外コメントを収集
externals = []
details = []
both = []

for i in range(5, len(data)):
    row = parse_order_row(data, i, cols)
    ext = row.comment_external.strip()
    det = row.comment_detail.strip()
    if ext:
        externals.append((row.order_number, ext, det))
    if ext and det:
        both.append((row.order_number, ext, det))

print(f"=== 統計 ===")
print(f"全行数: {len(data) - 5}")
print(f"社外コメントあり: {len(externals)} 件")
print(f"社外+明細両方あり: {len(both)} 件")
print()

if externals:
    lengths = [len(e[1]) for e in externals]
    print(f"社外コメント文字数: 最小={min(lengths)}, 最大={max(lengths)}, 平均={sum(lengths)/len(lengths):.1f}")
    print()

# ユニークな社外コメントをサンプル表示（最大30件）
seen = set()
unique_samples = []
for order_num, ext, det in externals:
    if ext not in seen:
        seen.add(ext)
        unique_samples.append((order_num, ext, det))

print(f"=== ユニークな社外コメント: {len(unique_samples)} 種類 ===")
for i, (order_num, ext, det) in enumerate(unique_samples[:30]):
    print(f"  [{i+1}] 注番={order_num} | 社外({len(ext)}文字): {ext}")
    if det:
        print(f"       明細({len(det)}文字): {det}")
    print()

# 両方ある場合のサンプル
if both:
    print(f"=== 社外+明細 両方ある場合のサンプル（最大10件）===")
    seen2 = set()
    count = 0
    for order_num, ext, det in both:
        key = (ext, det)
        if key not in seen2:
            seen2.add(key)
            count += 1
            print(f"  [{count}] 注番={order_num}")
            print(f"       社外({len(ext)}文字): {ext}")
            print(f"       明細({len(det)}文字): {det}")
            print()
            if count >= 10:
                break
