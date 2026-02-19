"""exe化用エントリポイント

PyInstallerでビルドする際のエントリポイント。
1. コマンドライン引数があればそのまま main.main() に渡す
2. 引数なしならtkinterファイルダイアログで10PM.XLSを選択 → GUIモードへ
"""

import sys
import os
import tkinter as tk
from tkinter import filedialog

# 起動時にnouki_kaitouパッケージのインポートを検証
from nouki_kaitou.main import main as app_main


def select_source_file() -> str | None:
    """tkinterのファイルダイアログで10PM.XLSを選択する。"""
    root = tk.Tk()
    root.withdraw()

    # デスクトップを初期ディレクトリに
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")

    file_path = filedialog.askopenfilename(
        title="SAPの受注一覧ファイルを選択してください",
        filetypes=[
            ("Excelファイル", "*.xls;*.xlsx"),
            ("すべてのファイル", "*.*"),
        ],
        initialdir=desktop,
    )

    root.destroy()
    return file_path if file_path else None


def main():
    """エントリポイント。"""
    # --source が既にある場合（バッチ等からの呼び出し）はそのままmainへ
    if "--source" in sys.argv:
        app_main()
        return

    # コンソールに案内表示
    print("納期回答書作成ツール（Python版）")
    print("SAPの受注一覧を選択してください...")
    print()

    # ファイル選択ダイアログ
    source = select_source_file()
    if not source:
        print("キャンセルされました。")
        sys.exit(0)

    # main.main() を --source 付きで呼び出す
    sys.argv = ["nouki_kaitou", "--source", source]
    app_main()


if __name__ == "__main__":
    main()
