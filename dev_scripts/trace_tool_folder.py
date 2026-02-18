"""_resolve_tool_folder のトレース: 17PM.xlsを指定した場合のパス解決を確認"""

import sys
import os
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def _normalize(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def _find_file_in_dir(directory: Path, target_name: str) -> Path | None:
    normalized_target = _normalize(target_name)
    direct = directory / target_name
    if direct.exists():
        return direct
    try:
        for p in directory.iterdir():
            if p.is_file() and _normalize(p.name) == normalized_target:
                return p
    except OSError:
        pass
    return None


def trace():
    source_path = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧\17PM.xls")
    marker = "メーカー一覧.xlsx"

    # main.pyと同じ候補リスト（修正後）
    candidates = [
        source_path.parent.resolve(),
        source_path.parent.parent.resolve(),
        Path(__file__).resolve().parent.parent,
        Path(os.getcwd()).resolve(),
    ]

    print("=" * 70)
    print("_resolve_tool_folder トレース")
    print(f"  source_path = {source_path}")
    print(f"  marker = {marker}")
    print("=" * 70)

    for i, folder in enumerate(candidates, 1):
        print(f"\n候補{i}: {folder}")
        print(f"  存在するか: {folder.exists()}")
        if folder.exists():
            found = _find_file_in_dir(folder, marker)
            print(f"  {marker} あるか: {found is not None}")
            if found:
                print(f"  → パス: {found}")
                # ここが tool_folder になる
                print(f"\n  ★ tool_folder = {folder}")

                # 送付履歴の解決
                hist = _find_file_in_dir(folder, "送付履歴.xlsx")
                print(f"\n  送付履歴.xlsx あるか: {hist is not None}")
                if hist:
                    print(f"  → history_path = {hist}")
                    print(f"  → VBAの送付履歴に直接書き込み")
                else:
                    new_path = folder / "送付履歴.xlsx"
                    print(f"  → history_path = {new_path} (新規作成)")

                # 顧客マスターの確認
                cust = _find_file_in_dir(folder, "顧客マスター_v2.xlsm")
                print(f"\n  顧客マスター_v2.xlsm あるか: {cust is not None}")
                if cust:
                    print(f"  → パス: {cust}")

                # 回答書出力先
                output_base = folder / "納期回答書"
                print(f"\n  回答書出力先ベース: {output_base}")
                print(f"  存在するか: {output_base.exists()}")

                return folder
        else:
            print(f"  → スキップ（フォルダが存在しない）")

    # どこにも見つからない場合
    cwd = Path(os.getcwd()).resolve()
    print(f"\n★ どこにも見つからず → tool_folder = CWD = {cwd}")
    return cwd


if __name__ == "__main__":
    trace()
