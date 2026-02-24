"""受注データ選択ダイアログ（VBA frmSelection相当）

tkinterでVBAフォームの機能を再現する。
- 期間指定モード: 日付範囲で受注データを絞り込み → 顧客リスト表示 → 複数選択
- 伝票番号指定モード: 受発注伝票番号を直接入力（改行区切り）
"""

from __future__ import annotations

import datetime
import tkinter as tk
from tkinter import messagebox
from collections import Counter
from typing import Optional

from nouki_kaitou.models import BranchSettings, OrderRow


class SelectionDialog:
    """受注データ選択ダイアログ（VBA frmSelection相当）"""

    def __init__(
        self,
        orders: list[OrderRow],
        branch: BranchSettings,
        master_customers: set[str],
        email_customers: set[str],
    ) -> None:
        """初期化。

        Args:
            orders: 全受注データ
            branch: 営業所設定（start_date用）
            master_customers: 顧客マスター登録済み顧客名セット
            email_customers: メールアドレス登録済み顧客名セット
        """
        self._orders = orders
        self._branch = branch
        self._master_customers = master_customers
        self._email_customers = email_customers
        self._result: Optional[dict] = None

        # 顧客リストに表示中のデータ（顧客名のみ、件数除去済み）
        self._current_customer_names: list[str] = []

    def show(self) -> Optional[dict]:
        """ダイアログを表示し、結果を返す。

        Returns:
            キャンセル時: None
            期間モード: {
                "mode": "period",
                "date_from": datetime.date,
                "date_to": datetime.date,
                "customers": list[str],
                "email_mode": "send" | "draft" | "none",
            }
            伝票番号モード: {
                "mode": "ordernumber",
                "order_numbers": list[str],
                "email_mode": "send" | "draft" | "none",
            }
        """
        self._root = tk.Tk()
        self._root.title("受注データ抽出")
        # 画面サイズに応じてウィンドウ高さを調整（小画面対応）
        screen_h = self._root.winfo_screenheight()
        win_h = min(770, screen_h - 100)
        self._ops_y = win_h - 80
        self._cust_frame_h = win_h - 290
        self._order_frame_h = win_h - 154
        self._cust_list_h = self._cust_frame_h - 66
        self._order_text_h = self._order_frame_h - 60

        self._root.geometry(f"600x{win_h}")
        self._root.resizable(False, False)
        self._root.configure(bg="#ECF5FF")

        # ウィンドウを閉じた場合はキャンセル扱い
        self._root.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self._build_ui()

        # デフォルトは全期間で顧客リスト表示
        self._on_all_period()

        self._root.mainloop()
        return self._result

    # ========================================
    # UI構築
    # ========================================
    def _build_ui(self) -> None:
        """UIウィジェットを構築する。"""
        root = self._root

        # --- モード選択タイトルバー ---
        lbl_title = tk.Label(
            root,
            text="▼ 作成モードを選択してください",
            bg="#FF9800",
            fg="white",
            font=("Yu Gothic UI", 13, "bold"),
            relief="solid",
            borderwidth=1,
            anchor="center",
        )
        lbl_title.place(x=12, y=6, width=564, height=24)

        # --- モード切替ラジオボタン ---
        self._mode_var = tk.StringVar(value="period")

        self._opt_period = tk.Radiobutton(
            root,
            text="期間指定",
            variable=self._mode_var,
            value="period",
            command=lambda: self._switch_to_mode("period"),
            bg="#1E88E5",
            fg="white",
            selectcolor="#1E88E5",
            activebackground="#1E88E5",
            activeforeground="white",
            font=("Yu Gothic UI", 11, "bold"),
            indicatoron=False,
            relief="raised",
            borderwidth=2,
        )
        self._opt_period.place(x=12, y=34, width=277, height=28)

        self._opt_ordernumber = tk.Radiobutton(
            root,
            text="伝票番号指定",
            variable=self._mode_var,
            value="ordernumber",
            command=lambda: self._switch_to_mode("ordernumber"),
            bg="#B0BEC5",
            fg="#455A64",
            selectcolor="#3F51B5",
            activebackground="#3F51B5",
            activeforeground="white",
            font=("Yu Gothic UI", 11, "bold"),
            indicatoron=False,
            relief="raised",
            borderwidth=2,
        )
        self._opt_ordernumber.place(x=299, y=34, width=277, height=28)

        # --- 期間フレーム ---
        self._frame_period = tk.LabelFrame(
            root,
            text="期間指定",
            bg="#E3F2FD",
            font=("Yu Gothic UI", 10, "bold"),
            padx=8,
            pady=4,
        )
        self._frame_period.place(x=12, y=68, width=564, height=130)

        self._build_period_frame()

        # --- 顧客選択フレーム ---
        self._frame_customers = tk.LabelFrame(
            root,
            text="顧客選択",
            bg="#E8F4FD",
            font=("Yu Gothic UI", 10, "bold"),
            padx=8,
            pady=4,
        )
        self._frame_customers.place(x=12, y=204, width=564, height=self._cust_frame_h)

        self._build_customer_frame()

        # --- 伝票番号フレーム ---
        self._frame_ordernumbers = tk.LabelFrame(
            root,
            text="受発注伝票番号入力（改行区切り）",
            bg="#E1EFFC",
            font=("Yu Gothic UI", 10, "bold"),
            padx=8,
            pady=4,
        )
        self._frame_ordernumbers.place(x=12, y=68, width=564, height=self._order_frame_h)

        self._build_ordernumber_frame()

        # --- 操作ボタンフレーム ---
        self._frame_operations = tk.Frame(root, bg="#F0F8FF")
        self._frame_operations.place(x=12, y=self._ops_y, width=564, height=70)

        self._build_operation_frame()

        # デフォルトは期間モード
        self._switch_to_mode("period")

    def _build_period_frame(self) -> None:
        """期間入力フレームのウィジェットを構築する。"""
        frame = self._frame_period

        # --- 上段: ショートカットボタン5個（左） + 全期間ボタン（右・2行またぎ正方形） ---
        shortcuts = [
            ("本日", self._on_today, "#0D47A1"),
            ("3日前～", self._on_3days, "#1565C0"),
            ("1週間", self._on_1week, "#1E88E5"),
            ("月初～", self._on_month_start, "#42A5F5"),
            ("1ヶ月", self._on_1month, "#64B5F6"),
        ]

        x_pos = 4
        btn_width = 88
        for text, cmd, color in shortcuts:
            btn = tk.Button(
                frame, text=text, command=cmd,
                bg=color, fg="white", font=("Yu Gothic UI", 9, "bold"),
                relief="raised", cursor="hand2",
            )
            btn.place(x=x_pos, y=2, width=btn_width, height=28)
            x_pos += btn_width + 4

        # 全期間ボタン（右端・2行にまたがる正方形）
        btn_all_period = tk.Button(
            frame, text="全期間", command=self._on_all_period,
            bg="#9C27B0", fg="white", font=("Yu Gothic UI", 10, "bold"),
            relief="raised", cursor="hand2",
        )
        btn_all_period.place(x=472, y=2, width=64, height=64)

        # --- 下段: 日付入力 + 検索ボタン ---
        lbl_from = tk.Label(frame, text="開始:", bg="#E3F2FD",
                            font=("Yu Gothic UI", 10))
        lbl_from.place(x=4, y=40, width=40, height=24)

        self._txt_date_from = tk.Entry(frame, font=("Yu Gothic UI", 10),
                                       bg="white", width=14)
        self._txt_date_from.place(x=46, y=40, width=120, height=24)

        lbl_to = tk.Label(frame, text="～ 終了:", bg="#E3F2FD",
                          font=("Yu Gothic UI", 10))
        lbl_to.place(x=170, y=40, width=55, height=24)

        self._txt_date_to = tk.Entry(frame, font=("Yu Gothic UI", 10),
                                     bg="white", width=14)
        self._txt_date_to.place(x=227, y=40, width=141, height=24)

        # 検索ボタン（1ヶ月ボタンと同位置・同幅）
        btn_search = tk.Button(
            frame, text="検索", command=self._on_search,
            bg="#0277BD", fg="white", font=("Yu Gothic UI", 10, "bold"),
            relief="raised", cursor="hand2",
        )
        btn_search.place(x=372, y=38, width=88, height=28)

        # ヒントラベル
        lbl_hint = tk.Label(
            frame, text="※ ボタンで期間を設定し、検索で顧客リストを更新します",
            bg="#E3F2FD", fg="#546E7A", font=("Yu Gothic UI", 8),
        )
        lbl_hint.place(x=4, y=78, width=532, height=18)

    def _build_customer_frame(self) -> None:
        """顧客選択フレームのウィジェットを構築する。"""
        frame = self._frame_customers

        # ラベル
        lbl_info = tk.Label(
            frame, text="クリックで選択/解除（複数選択可）",
            bg="#E8F4FD", fg="#37474F", font=("Yu Gothic UI", 9, "bold"),
        )
        lbl_info.place(x=4, y=4, width=532, height=18)

        # リストボックス + スクロールバー
        list_frame = tk.Frame(frame, bg="#E8F4FD")
        list_frame.place(x=4, y=28, width=532, height=self._cust_list_h)

        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._lst_customers = tk.Listbox(
            list_frame,
            selectmode=tk.MULTIPLE,
            bg="#FAFCFF",
            font=("Yu Gothic UI", 10),
            yscrollcommand=scrollbar.set,
            activestyle="none",
        )
        self._lst_customers.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self._lst_customers.yview)

    def _build_ordernumber_frame(self) -> None:
        """伝票番号入力フレームのウィジェットを構築する。"""
        frame = self._frame_ordernumbers

        lbl_info = tk.Label(
            frame, text="受発注伝票番号を1行ずつ入力してください:",
            bg="#E1EFFC", fg="#37474F", font=("Yu Gothic UI", 10),
        )
        lbl_info.place(x=0, y=0, width=540, height=24)

        # テキストエリア + スクロールバー
        text_frame = tk.Frame(frame, bg="#E1EFFC")
        text_frame.place(x=0, y=28, width=540, height=self._order_text_h)

        scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._txt_order_numbers = tk.Text(
            text_frame,
            font=("Yu Gothic UI", 11),
            bg="white",
            yscrollcommand=scrollbar.set,
            wrap=tk.NONE,
        )
        self._txt_order_numbers.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self._txt_order_numbers.yview)

    def _build_operation_frame(self) -> None:
        """操作ボタンフレームのウィジェットを構築する。"""
        frame = self._frame_operations

        # 横一列: 全選択 / 全解除 / アドレス登録済み / OK / キャンセル
        btn_h = 36
        y = 4

        self._btn_select_all = tk.Button(
            frame, text="全選択", command=self._on_select_all,
            bg="#0097A7", fg="white", font=("Yu Gothic UI", 9, "bold"),
            relief="raised", cursor="hand2",
        )
        self._btn_select_all.place(x=0, y=y, width=90, height=btn_h)

        self._btn_deselect_all = tk.Button(
            frame, text="全解除", command=self._on_deselect_all,
            bg="#90A4AE", fg="white", font=("Yu Gothic UI", 9, "bold"),
            relief="raised", cursor="hand2",
        )
        self._btn_deselect_all.place(x=96, y=y, width=90, height=btn_h)

        self._btn_select_with_email = tk.Button(
            frame, text="アドレス登録済み", command=self._on_select_with_email,
            bg="#FF5722", fg="white", font=("Yu Gothic UI", 9, "bold"),
            relief="raised", cursor="hand2",
        )
        self._btn_select_with_email.place(x=192, y=y, width=138, height=btn_h)

        self._btn_ok = tk.Button(
            frame, text="OK", command=self._on_ok,
            bg="#3F51B5", fg="white", font=("Yu Gothic UI", 11, "bold"),
            relief="raised", cursor="hand2",
        )
        self._btn_ok.place(x=340, y=y, width=100, height=btn_h)

        self._btn_cancel = tk.Button(
            frame, text="キャンセル", command=self._on_cancel,
            bg="#546E7A", fg="white", font=("Yu Gothic UI", 9, "bold"),
            relief="raised", cursor="hand2",
        )
        self._btn_cancel.place(x=446, y=y, width=118, height=btn_h)

        # 件数ラベル
        self._lbl_count = tk.Label(
            frame, text="", bg="#F0F8FF", fg="#37474F",
            font=("Yu Gothic UI", 9),
        )
        self._lbl_count.place(x=0, y=44, width=564, height=18)

    # ========================================
    # モード切替
    # ========================================
    def _switch_to_mode(self, mode: str) -> None:
        """期間指定/伝票番号指定のモード切替を行う。"""
        btn_h = 36
        y = 4

        if mode == "period":
            self._frame_period.place(x=12, y=68, width=564, height=130)
            self._frame_customers.place(x=12, y=204, width=564, height=self._cust_frame_h)
            self._frame_ordernumbers.place_forget()

            # 選択系ボタン表示
            self._btn_select_all.place(x=0, y=y, width=90, height=btn_h)
            self._btn_deselect_all.place(x=96, y=y, width=90, height=btn_h)
            self._btn_select_with_email.place(x=192, y=y, width=138, height=btn_h)

            # ラジオボタンの色更新
            self._opt_period.configure(bg="#1E88E5", fg="white")
            self._opt_ordernumber.configure(bg="#B0BEC5", fg="#455A64")
        else:
            self._frame_period.place_forget()
            self._frame_customers.place_forget()
            self._frame_ordernumbers.place(x=12, y=68, width=564, height=self._order_frame_h)

            # 選択系ボタン非表示
            self._btn_select_all.place_forget()
            self._btn_deselect_all.place_forget()
            self._btn_select_with_email.place_forget()

            # ラジオボタンの色更新
            self._opt_ordernumber.configure(bg="#3F51B5", fg="white")
            self._opt_period.configure(bg="#B0BEC5", fg="#455A64")

            # テキストをクリアしてフォーカス
            self._txt_order_numbers.delete("1.0", tk.END)
            self._txt_order_numbers.focus_set()

    # ========================================
    # 顧客リスト更新
    # ========================================
    def _update_customer_list(self) -> None:
        """日付でフィルタし顧客リストを更新する。"""
        # 日付パース
        date_from = self._parse_date_field(self._txt_date_from.get())
        date_to = self._parse_date_field(self._txt_date_to.get())

        if date_from is None:
            messagebox.showwarning("入力エラー", "開始日が正しくありません。")
            return
        if date_to is None:
            messagebox.showwarning("入力エラー", "終了日が正しくありません。")
            return
        if date_from > date_to:
            messagebox.showwarning("入力エラー", "開始日が終了日より後になっています。")
            return

        # registration_dateでフィルタしてグルーピング
        counter: Counter[str] = Counter()
        for order in self._orders:
            if order.registration_date is None:
                continue
            if order.registration_date < date_from or order.registration_date > date_to:
                continue
            name = order.customer_name.strip()
            if name:
                counter[name] += 1

        # 顧客マスターに存在する顧客のみ表示（VBA版と同じ）
        filtered_counter = {
            k: v for k, v in counter.items() if k in self._master_customers
        }

        if not filtered_counter:
            messagebox.showinfo("検索結果", "選択された期間内に該当する受注データがありません。")
            self._lst_customers.delete(0, tk.END)
            self._current_customer_names = []
            self._update_count_label()
            return

        # ソートしてリストボックスに表示
        sorted_names = sorted(filtered_counter.keys())
        self._lst_customers.delete(0, tk.END)
        self._current_customer_names = sorted_names

        for name in sorted_names:
            self._lst_customers.insert(tk.END, f"{name} ({filtered_counter[name]}件)")

        self._update_count_label()

    def _update_count_label(self) -> None:
        """件数ラベルを更新する。"""
        total = self._lst_customers.size()
        selected = len(self._lst_customers.curselection())
        self._lbl_count.configure(text=f"全{total}件中 {selected}件選択中")

    @staticmethod
    def _parse_date_field(text: str) -> Optional[datetime.date]:
        """日付テキストをパースする。"""
        text = text.strip()
        if not text:
            return None
        for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None

    # ========================================
    # 期間ショートカットボタン
    # ========================================
    def _set_period_and_search(
        self, date_from: datetime.date, date_to: datetime.date
    ) -> None:
        """日付フィールドを設定して顧客リストを更新する。"""
        self._txt_date_from.delete(0, tk.END)
        self._txt_date_from.insert(0, date_from.strftime("%Y/%m/%d"))
        self._txt_date_to.delete(0, tk.END)
        self._txt_date_to.insert(0, date_to.strftime("%Y/%m/%d"))
        self._update_customer_list()

    def _on_today(self) -> None:
        """本日ボタン。"""
        today = datetime.date.today()
        self._set_period_and_search(today, today)

    def _on_3days(self) -> None:
        """3日前～ボタン。"""
        today = datetime.date.today()
        self._set_period_and_search(today - datetime.timedelta(days=3), today)

    def _on_1week(self) -> None:
        """1週間ボタン。"""
        today = datetime.date.today()
        self._set_period_and_search(today - datetime.timedelta(days=7), today)

    def _on_month_start(self) -> None:
        """月初～ボタン。"""
        today = datetime.date.today()
        first_day = today.replace(day=1)
        self._set_period_and_search(first_day, today)

    def _on_1month(self) -> None:
        """1ヶ月ボタン。"""
        today = datetime.date.today()
        self._set_period_and_search(today - datetime.timedelta(days=30), today)

    def _on_all_period(self) -> None:
        """全期間ボタン（start_date～今日）。"""
        today = datetime.date.today()
        start = self._parse_date_field(self._branch.start_date)
        if start is None:
            # start_dateが未設定の場合は1年前をデフォルトに
            start = today - datetime.timedelta(days=365)
        self._set_period_and_search(start, today)

    def _on_search(self) -> None:
        """検索ボタン。"""
        self._update_customer_list()

    # ========================================
    # 選択操作ボタン
    # ========================================
    def _on_select_all(self) -> None:
        """全選択ボタン。"""
        self._lst_customers.select_set(0, tk.END)
        self._update_count_label()

    def _on_deselect_all(self) -> None:
        """全解除ボタン。"""
        self._lst_customers.select_clear(0, tk.END)
        self._update_count_label()

    def _on_select_with_email(self) -> None:
        """メールアドレス登録済み顧客のみ選択する。"""
        self._lst_customers.select_clear(0, tk.END)
        selected_count = 0

        for i, name in enumerate(self._current_customer_names):
            if name in self._email_customers:
                self._lst_customers.select_set(i)
                selected_count += 1

        self._update_count_label()
        messagebox.showinfo(
            "メール可能顧客",
            f"{selected_count}件のメール可能な顧客を選択しました。",
        )

    # ========================================
    # メール送信確認ダイアログ（VBA版と同じ3択）
    # ========================================
    def _ask_email_mode(self) -> str | None:
        """メール送信確認ダイアログを表示する。

        VBA版 MsgBox("メールを送信しますか？", vbYesNoCancel) と同じ3択。

        Returns:
            "send": そのまま送信（VBA: はい → sendDirectly=True）
            "draft": 確認してから送信（VBA: いいえ → sendDirectly=False）
            "none": 送信しない（VBA: キャンセル）
            None: ダイアログ自体をキャンセル（×ボタン）→ OK操作を中止
        """
        result: dict[str, str | None] = {"value": None}

        dlg = tk.Toplevel(self._root)
        dlg.title("メール送信確認")
        dlg.geometry("420x210")
        dlg.resizable(False, False)
        dlg.configure(bg="#ECF5FF")
        dlg.transient(self._root)
        dlg.grab_set()

        # ×ボタンで閉じた場合 → None（OK操作中止）
        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

        def _choose(value: str) -> None:
            result["value"] = value
            dlg.destroy()

        # アイコン + テキスト
        lbl_icon = tk.Label(
            dlg, text="?", bg="#ECF5FF", fg="#1565C0",
            font=("Yu Gothic UI", 28, "bold"),
        )
        lbl_icon.place(x=16, y=16, width=48, height=48)

        lbl_msg = tk.Label(
            dlg,
            text="メールを送信しますか？",
            bg="#ECF5FF", fg="#212121",
            font=("Yu Gothic UI", 13, "bold"),
            anchor="w",
        )
        lbl_msg.place(x=72, y=20, width=330, height=28)

        # 説明テキスト
        lbl_desc = tk.Label(
            dlg,
            text=(
                "【そのまま送信】 Outlookから直接送信します\n"
                "【確認してから送信】 下書きを作成します\n"
                "【送信しない】 メール処理をスキップします"
            ),
            bg="#ECF5FF", fg="#455A64",
            font=("Yu Gothic UI", 10),
            anchor="w", justify="left",
        )
        lbl_desc.place(x=72, y=52, width=340, height=60)

        # 3択ボタン
        btn_h = 38
        btn_y = 130

        tk.Button(
            dlg, text="そのまま送信", command=lambda: _choose("send"),
            bg="#2E7D32", fg="white", font=("Yu Gothic UI", 10, "bold"),
            relief="raised", cursor="hand2",
        ).place(x=16, y=btn_y, width=120, height=btn_h)

        tk.Button(
            dlg, text="確認してから送信", command=lambda: _choose("draft"),
            bg="#3F51B5", fg="white", font=("Yu Gothic UI", 10, "bold"),
            relief="raised", cursor="hand2",
        ).place(x=150, y=btn_y, width=140, height=btn_h)

        tk.Button(
            dlg, text="送信しない", command=lambda: _choose("none"),
            bg="#546E7A", fg="white", font=("Yu Gothic UI", 10, "bold"),
            relief="raised", cursor="hand2",
        ).place(x=304, y=btn_y, width=100, height=btn_h)

        # ダイアログを画面中央に配置
        dlg.update_idletasks()
        x = self._root.winfo_x() + (self._root.winfo_width() - 420) // 2
        y = self._root.winfo_y() + (self._root.winfo_height() - 210) // 2
        dlg.geometry(f"420x210+{x}+{y}")

        dlg.wait_window()
        return result["value"]

    # ========================================
    # OK / キャンセル
    # ========================================
    def _on_ok(self) -> None:
        """OKボタン（モード分岐）。"""
        mode = self._mode_var.get()

        if mode == "period":
            self._process_period_selection()
        else:
            self._process_ordernumber_input()

    def _process_period_selection(self) -> None:
        """期間指定モードのOK処理。"""
        selected_indices = self._lst_customers.curselection()

        if not selected_indices:
            messagebox.showwarning(
                "選択エラー",
                "顧客が選択されていません。\n少なくとも1つ選択してください。",
            )
            return

        # 日付パース
        date_from = self._parse_date_field(self._txt_date_from.get())
        date_to = self._parse_date_field(self._txt_date_to.get())
        if date_from is None or date_to is None:
            messagebox.showwarning("入力エラー", "期間の日付が正しくありません。")
            return

        # メール送信確認（VBA版と同じ3択）
        email_mode = self._ask_email_mode()
        if email_mode is None:
            # ×ボタンで閉じた → メイン画面に戻る
            return

        # 選択された顧客名を取得
        customers = [self._current_customer_names[i] for i in selected_indices]

        self._result = {
            "mode": "period",
            "date_from": date_from,
            "date_to": date_to,
            "customers": customers,
            "email_mode": email_mode,
        }
        self._root.destroy()

    def _process_ordernumber_input(self) -> None:
        """伝票番号指定モードのOK処理。"""
        input_text = self._txt_order_numbers.get("1.0", tk.END).strip()

        if not input_text:
            messagebox.showwarning("入力エラー", "伝票番号を入力してください。")
            return

        # 改行で分割、空行除去、トリム
        order_numbers = [
            line.strip() for line in input_text.splitlines() if line.strip()
        ]

        if not order_numbers:
            messagebox.showwarning("入力エラー", "有効な伝票番号が入力されていません。")
            return

        # メール送信確認（VBA版と同じ3択）
        email_mode = self._ask_email_mode()
        if email_mode is None:
            # ×ボタンで閉じた → メイン画面に戻る
            return

        self._result = {
            "mode": "ordernumber",
            "order_numbers": order_numbers,
            "email_mode": email_mode,
        }
        self._root.destroy()

    def _on_cancel(self) -> None:
        """キャンセルボタン。"""
        self._result = None
        self._root.destroy()
