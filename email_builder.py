"""メール生成モジュール

VBAの以下の関数を移植:
- CreateEmails (L5520): メール作成・送信
- BuildEmailBodyHTML (L5790): HTMLメール本文生成
- GetRepEmailAddresses (L7673): 担当者マスターからメアド取得
- HtmlEscape (L7752): HTML特殊文字エスケープ

件名生成はCreateEmails内(L5596-5599)のインライン処理を関数化。
"""

from __future__ import annotations

import datetime
import html
import os
from typing import Optional

from nouki_kaitou.models import (
    BranchSettings,
    BunnoEntry,
    CacheStore,
    HolidayMap,
    StockoutEntry,
    TrackingEntry,
)
from nouki_kaitou.representative import get_rep_email_addresses


# ============================================
# 件名生成（VBA CreateEmails L5596-5599 からインライン抽出）
# ============================================
def build_email_subject(
    customer_name: str,
    rep_name: str = "",
    today: Optional[datetime.date] = None,
) -> str:
    """メール件名を生成する。

    Args:
        customer_name: 顧客名
        rep_name: 担当者名（空なら担当者分なし）
        today: 基準日（テスト用）

    Returns:
        件名文字列
    """
    if today is None:
        today = datetime.date.today()

    date_str = f"{today.month:02d}/{today.day:02d}"

    if rep_name and rep_name != "__OTHER__":
        return (
            f"【マツモト産業】納期回答書_{date_str}受注分_"
            f"{customer_name}様（{rep_name}様担当分）"
        )
    return f"【マツモト産業】納期回答書_{date_str}受注分_{customer_name}様"


# ============================================
# VBA: HtmlEscape (L7752-7758)
# HTML特殊文字エスケープ
# ============================================
def html_escape(text: str) -> str:
    """HTML特殊文字をエスケープする。

    Args:
        text: エスケープ対象文字列

    Returns:
        エスケープ済み文字列
    """
    return html.escape(text, quote=True)



# ============================================
# VBA: BuildEmailBodyHTML (L5790-6146)
# HTMLメール本文生成
# ============================================
def build_email_body_html(
    customer_name: str,
    branch: BranchSettings,
    stockout_info_list: Optional[list[StockoutEntry]] = None,
    tracking_info_list: Optional[list[tuple[str, str, str, TrackingEntry]]] = None,
    bunno_info_list: Optional[list[dict]] = None,
    bunno_completed_list: Optional[list[tuple[str, str, str]]] = None,
    holidays: HolidayMap | None = None,
    cache: Optional[CacheStore] = None,
    today: Optional[datetime.date] = None,
) -> str:
    """HTMLメール本文を生成する。

    Args:
        customer_name: 顧客名
        branch: 営業所設定
        stockout_info_list: 欠品情報リスト
        tracking_info_list: 送り状情報 [(メーカー, 品名, 数量, TrackingEntry), ...]
        bunno_info_list: 分納情報（辞書リスト）
        bunno_completed_list: 分納完了リスト [(メーカー, 品名, 数量), ...]
        holidays: 祝日辞書
        cache: キャッシュストア
        today: 基準日（テスト用）

    Returns:
        HTML形式のメール本文
    """
    if today is None:
        today = datetime.date.today()

    has_stockout = bool(stockout_info_list)
    has_tracking = bool(tracking_info_list)
    has_bunno = bool(bunno_info_list)
    has_bunno_completed = bool(bunno_completed_list)

    parts: list[str] = []

    # HTML開始 + CSSスタイル
    parts.append(_build_html_head())

    # 挨拶文
    parts.append(_build_greeting(customer_name, branch))

    # 送り状情報セクション
    if has_tracking:
        parts.append(_build_tracking_section(tracking_info_list))

    # 欠品情報セクション
    if has_stockout:
        parts.append(_build_stockout_section(stockout_info_list))

    # 分納情報セクション
    if has_bunno:
        parts.append(
            _build_bunno_section(
                bunno_info_list, holidays, cache, today
            )
        )

    # 分納完了通知セクション
    if has_bunno_completed:
        parts.append(_build_bunno_completed_section(bunno_completed_list))

    # 確認中の注記
    parts.append(_build_confirming_note())

    # 署名
    parts.append(_build_signature(branch))

    # HTML終了
    parts.append("</body></html>")

    return "".join(parts)


# ============================================
# VBA: CreateEmails (L5520-5666)
# メール作成処理
# ============================================
def create_emails(
    created_files: list[dict],
    branch: BranchSettings,
    customer_master_ws: object,
    rep_master_ws: object = None,
    holidays: HolidayMap | None = None,
    cache: Optional[CacheStore] = None,
    send_directly: bool = False,
    today: Optional[datetime.date] = None,
) -> list[dict]:
    """メールデータを作成する。

    VBAではOutlook COM連携で直接メール送信していたが、
    Python版ではメールデータ（件名・本文・宛先・添付ファイル）を
    辞書のリストとして返す。実際の送信はUI層で行う。

    Args:
        created_files: 作成済みファイル情報のリスト。各要素は辞書:
            - customer_name: 顧客名
            - file_path: PDFファイルパス
            - stockout_info_list: 欠品情報
            - tracking_info_list: 送り状情報
            - bunno_info_list: 分納情報
            - rep_name: 担当者名
            - bunno_completed_list: 分納完了リスト
        branch: 営業所設定
        customer_master_ws: 顧客マスターシート
        rep_master_ws: 担当者マスターシート
        holidays: 祝日辞書
        cache: キャッシュストア
        send_directly: True=送信 / False=下書き（メタデータに記録）
        today: 基準日（テスト用）

    Returns:
        メールデータのリスト。各要素は辞書:
            - to: 宛先メールアドレス
            - subject: 件名
            - html_body: HTML本文
            - attachments: 添付ファイルパスのリスト
            - shared_email: 共有メールアドレス（差出人）
            - send_directly: 送信/下書きフラグ
        スキップされた顧客は含まれない。
    """
    from nouki_kaitou.customer import get_email_addresses

    if today is None:
        today = datetime.date.today()

    results: list[dict] = []
    skipped_customers: list[str] = []

    for file_info in created_files:
        customer_name = file_info.get("customer_name", "")
        file_path = file_info.get("file_path", "")
        rep_name = file_info.get("rep_name", "")

        # メールアドレス取得
        if rep_name and rep_master_ws is not None:
            mail_addresses = get_rep_email_addresses(
                customer_name, rep_name, rep_master_ws
            )
        else:
            email_col = cache.cust_email_start_col if cache else 4
            mail_addresses = get_email_addresses(
                customer_name, customer_master_ws, email_col
            )

        if not mail_addresses:
            skipped_customers.append(customer_name)
            continue

        # 件名
        subject = build_email_subject(customer_name, rep_name, today)

        # HTML本文
        body = build_email_body_html(
            customer_name=customer_name,
            branch=branch,
            stockout_info_list=file_info.get("stockout_info_list"),
            tracking_info_list=file_info.get("tracking_info_list"),
            bunno_info_list=file_info.get("bunno_info_list"),
            bunno_completed_list=file_info.get("bunno_completed_list"),
            holidays=holidays,
            cache=cache,
            today=today,
        )

        results.append({
            "to": mail_addresses,
            "subject": subject,
            "html_body": body,
            "attachments": [file_path] if file_path else [],
            "shared_email": branch.shared_email,
            "send_directly": send_directly,
            "customer_name": customer_name,
        })

    return results


# ============================================
# HTMLパーツ生成（private）
# ============================================
def _build_html_head() -> str:
    """HTML開始部分（CSS含む）"""
    return (
        "<html><head><style>"
        "body { font-family: 'メイリオ', 'Meiryo', sans-serif; "
        "font-size: 14px; line-height: 1.6; }"
        ".section { margin: 15px 0; padding: 10px; "
        "background-color: #f8f8f8; border-left: 4px solid #c0a040; }"
        ".section-title { font-weight: bold; color: #333; margin-bottom: 8px; }"
        ".stockout { border-left-color: #cc0000; }"
        ".stockout .section-title { color: #cc0000; }"
        ".tracking { border-left-color: #0066cc; }"
        ".tracking-link { color: #0066cc; font-weight: bold; }"
        ".item { margin-left: 20px; font-size: 13px; }"
        "</style></head><body>"
    )


def _build_greeting(customer_name: str, branch: BranchSettings) -> str:
    """挨拶文"""
    h = html_escape
    return (
        f"{h(customer_name)} 御中<br><br>"
        f"いつもお世話になっております。<br>"
        f"マツモト産業㈱{h(branch.name)}です。<br><br>"
        f"ご注文ありがとうございます。<br>"
        f"納期回答書をお送りいたします。<br><br>"
    )


def _build_tracking_section(
    tracking_info_list: list[tuple[str, str, str, TrackingEntry]],
) -> str:
    """送り状情報セクション"""
    from nouki_kaitou.tracking import can_direct_track, get_tracking_url
    from nouki_kaitou.utils import format_quantity

    h = html_escape
    parts: list[str] = []

    parts.append("<div class='section tracking'>")
    parts.append("<div class='section-title'>■ 送り状番号のご連絡</div>")

    # 商品ごとの送り状セットを作成
    product_to_tracking: dict[str, dict[str, TrackingEntry]] = {}

    for mfg, product, qty, entry in tracking_info_list:
        product_key = f"{mfg}|{product}|{qty}"
        if product_key not in product_to_tracking:
            product_to_tracking[product_key] = {}

        tracking_key = f"{entry.carrier_name}|{entry.tracking_number}"
        if tracking_key not in product_to_tracking[product_key]:
            product_to_tracking[product_key][tracking_key] = entry

    # 送り状セットでグループ化
    tracking_set_to_products: dict[str, list[tuple[str, dict[str, TrackingEntry]]]] = {}

    for product_key, tracking_dict in product_to_tracking.items():
        # 送り状キーをソートして一意の文字列を作成
        sorted_keys = sorted(tracking_dict.keys())
        set_key = "||".join(sorted_keys)

        if set_key not in tracking_set_to_products:
            tracking_set_to_products[set_key] = []
        tracking_set_to_products[set_key].append((product_key, tracking_dict))

    # 送り状セットごとに表示
    for set_key, product_list in tracking_set_to_products.items():
        # 最初の商品から送り状情報を取得
        tracking_dict = product_list[0][1]
        is_multi_tracking = len(tracking_dict) >= 2

        # 送り状を表示
        for entry in tracking_dict.values():
            url = get_tracking_url(entry.carrier_name, entry.tracking_number)
            direct = can_direct_track(entry.carrier_name)

            if direct and url:
                parts.append(
                    f"<a href='{url}' class='tracking-link'>"
                    f"{h(entry.carrier_name)}  {entry.tracking_number}</a><br>"
                )
            else:
                parts.append(
                    f"<span class='tracking-link'>"
                    f"{h(entry.carrier_name)}  {entry.tracking_number}</span><br>"
                )
                if url:
                    parts.append(
                        f"<div class='item'>→ <a href='{url}' "
                        f"style='color: #0066cc;'>追跡ページ</a>"
                        f"（番号を入力してください）</div>"
                    )

        # 商品を表示
        for product_key, _ in product_list:
            p = product_key.split("|")
            parts.append(
                f"<div class='item'>・{h(p[0])}  {h(p[1])}  x{format_quantity(p[2])}</div>"
            )

        # 複数送り状の場合は注釈
        if is_multi_tracking:
            parts.append(
                "<div style='margin-left: 20px; color: #666; "
                "font-size: 12px; font-style: italic;'>"
                "※別々の場所からの出荷になります</div>"
            )

        parts.append("<br>")

    parts.append("</div>")
    return "".join(parts)


_UNCONFIRMED_DELIVERY = ("欠品中", "確認中", "日程調整中")


def _is_stockout_confirmed(item: StockoutEntry) -> bool:
    """欠品の入荷日が確定しているか判定する。"""
    return bool(
        item.delivery
        and item.delivery not in _UNCONFIRMED_DELIVERY
        and "（欠品）" not in item.delivery
    )


def _build_stockout_section(
    stockout_info_list: list[StockoutEntry],
) -> str:
    """欠品情報セクション（入荷確定 + 欠品継続の2グループ）"""
    confirmed = [i for i in stockout_info_list if _is_stockout_confirmed(i)]
    pending = [i for i in stockout_info_list if not _is_stockout_confirmed(i)]

    parts: list[str] = []
    if confirmed:
        parts.append(_build_stockout_confirmed_section(confirmed))
    if pending:
        parts.append(_build_stockout_pending_section(pending))
    return "".join(parts)


def _build_stockout_confirmed_section(
    items: list[StockoutEntry],
) -> str:
    """入荷確定グループのセクション"""
    h = html_escape
    from nouki_kaitou.utils import format_quantity

    parts: list[str] = []
    parts.append(
        "<div style='margin: 15px 0; padding: 12px; "
        "background-color: #f0f7f0; border-left: 4px solid #338833;'>"
    )
    parts.append(
        "<div style='font-weight: bold; color: #338833; font-size: 15px; "
        "margin-bottom: 8px;'>"
        "欠品しておりました商品の入荷日が確定いたしました。</div>"
    )

    for item in items:
        base_text = (
            f"・{h(item.manufacturer_name)}  {h(item.product_name)}"
            f"  x{format_quantity(item.quantity)}"
        )
        parts.append(
            f"<div style='margin-left: 20px; font-weight: bold;'>"
            f"{base_text} → {h(item.delivery)}</div>"
        )

    parts.append("</div>")
    return "".join(parts)


def _build_stockout_pending_section(
    items: list[StockoutEntry],
) -> str:
    """欠品継続グループのセクション（従来の欠品セクション）"""
    h = html_escape
    from nouki_kaitou.utils import format_quantity

    parts: list[str] = []
    parts.append(
        "<div style='margin: 15px 0; padding: 12px; "
        "background-color: #fff0f0; border-left: 4px solid #cc0000;'>"
    )
    parts.append(
        "<div style='font-weight: bold; color: #cc0000; font-size: 15px; "
        "margin-bottom: 8px;'>【注意】欠品中の商品について</div>"
    )
    parts.append(
        "<div style='margin-bottom: 10px; color: #cc0000;'>"
        "下記商品は現在欠品中です。ご迷惑をおかけし申し訳ございません。</div>"
    )

    for item in items:
        base_text = (
            f"・{h(item.manufacturer_name)}  {h(item.product_name)}"
            f"  x{format_quantity(item.quantity)}"
        )

        if item.approx_delivery:
            delivery_text = h(item.approx_delivery)
        elif item.delivery and item.delivery not in _UNCONFIRMED_DELIVERY:
            # （欠品）付き日付 → マーカー除去して表示
            delivery_text = h(item.delivery.replace("（欠品）", ""))
        else:
            delivery_text = "入荷次第ご連絡"

        parts.append(
            f"<div style='margin-left: 20px; color: #cc0000; "
            f"font-weight: bold;'>{base_text} → {delivery_text}</div>"
        )

    parts.append("</div>")
    return "".join(parts)


def _build_bunno_section(
    bunno_info_list: list[dict],
    holidays: HolidayMap | None,
    cache: Optional[CacheStore],
    today: datetime.date,
) -> str:
    """分納情報セクション"""
    from nouki_kaitou.bunno import calculate_bunno_date, has_bunno_kakuninchu
    from nouki_kaitou.excel_writer import check_same_date_in_bunno
    from nouki_kaitou.utils import format_quantity, to_circled_number

    h = html_escape
    parts: list[str] = []

    # 未定/確認中があるかチェック
    has_mitei = False
    for item in bunno_info_list:
        calc_details = item.get("calc_details", [])
        if has_bunno_kakuninchu(calc_details):
            has_mitei = True
            break

    parts.append(
        "<div style='margin: 15px 0; padding: 12px; "
        "background-color: #e8f0ff; border-left: 4px solid #0066cc;'>"
    )
    parts.append(
        "<div style='font-weight: bold; color: #0066cc; font-size: 15px; "
        "margin-bottom: 8px;'>■ 分納のご連絡</div>"
    )
    parts.append(
        "<div style='margin-bottom: 10px;'>"
        "下記商品は分納にてお届けいたします。</div>"
    )

    if has_mitei:
        parts.append(
            "<div style='margin-bottom: 10px; color: #cc0000;'>"
            "※一部納期未定のためご迷惑をおかけいたします。"
            "確定次第ご連絡いたします。</div>"
        )

    for item in bunno_info_list:
        mfg = item.get("manufacturer", "")
        product = item.get("product", "")
        qty = item.get("quantity", "")
        entries: list[BunnoEntry] = item.get("entries", [])
        calc_details = item.get("calc_details", [])
        is_ship_rule = item.get("is_ship_rule", False)
        days_to_add = item.get("days_to_add", 0)
        order_num = item.get("order_number", "")
        detail_num = item.get("detail_number", "")
        is_rosenbin = item.get("is_rosenbin", False)

        # 商品ヘッダー
        parts.append(
            f"<div style='margin-left: 10px; font-weight: bold; "
            f"color: #003366;'>・{h(mfg)} {h(product)} x{format_quantity(qty)}</div>"
        )

        # 同じ日付チェック
        has_same_date = check_same_date_in_bunno(entries)

        for idx, entry in enumerate(entries):
            counter = idx + 1

            # 計算済み納期を使用（あれば）
            calc_date = ""
            if idx < len(calc_details) and len(calc_details[idx]) >= 4:
                calc_date = str(calc_details[idx][3])

            if not calc_date:
                calc_date = calculate_bunno_date(
                    entry.date_str, is_ship_rule, days_to_add,
                    holidays, cache, order_num, detail_num,
                    is_rosenbin, today,
                )

            location_text = f"（{h(entry.location)}）" if entry.location else ""
            circled = to_circled_number(counter)

            # 色分け判定
            is_uncertain = (
                calc_date == "確認中"
                or ("予定" in calc_date and "出荷" not in calc_date
                    and "配達" not in calc_date)
            )
            is_from_mitei = (
                entry.date_str == "未定"
                or "欠品" in entry.date_str
                or "確認中" in entry.date_str
            )

            if is_uncertain:
                parts.append(
                    f"<div style='margin-left: 30px; color: #cc0000; "
                    f"font-weight: bold;'>{circled}{h(entry.quantity)} → "
                    f"{h(calc_date)}{location_text}（確定次第ご連絡）</div>"
                )
            elif is_from_mitei:
                # 元は未定だったが今は確定 → オレンジ色
                parts.append(
                    f"<div style='margin-left: 30px; color: #c86400; "
                    f"font-weight: bold;'>{circled}{h(entry.quantity)} → "
                    f"{h(calc_date)}{location_text}</div>"
                )
            else:
                parts.append(
                    f"<div style='margin-left: 30px;'>"
                    f"{circled}{h(entry.quantity)} → "
                    f"{h(calc_date)}{location_text}</div>"
                )

        # 同じ日付の場合は注釈
        if has_same_date:
            parts.append(
                "<div style='margin-left: 20px; color: #666; "
                "font-size: 12px; font-style: italic;'>"
                "※別々の場所からの出荷になります</div>"
            )

    parts.append("</div>")
    return "".join(parts)


def _build_bunno_completed_section(
    bunno_completed_list: list[tuple[str, str, str]],
) -> str:
    """分納完了通知セクション"""
    from nouki_kaitou.utils import format_quantity
    h = html_escape
    parts: list[str] = []

    parts.append(
        "<div style='margin: 15px 0; padding: 12px; "
        "background-color: #e8f5e9; border-left: 4px solid #28a745;'>"
    )
    parts.append(
        "<div style='font-weight: bold; color: #28a745; font-size: 15px; "
        "margin-bottom: 8px;'>■ 分納完了のご連絡</div>"
    )
    parts.append(
        "<div style='margin-bottom: 10px;'>"
        "分納でご注文いただいた商品は全て出荷が完了しました。</div>"
    )

    for mfg, product, qty in bunno_completed_list:
        parts.append(
            f"<div style='margin-left: 10px; font-weight: bold; "
            f"color: #1b5e20;'>・{h(mfg)}  {h(product)}  x{format_quantity(qty)}</div>"
        )

    parts.append("</div>")
    return "".join(parts)


def _build_confirming_note() -> str:
    """確認中の注記"""
    return (
        "<br>※納期が「確認中」の商品については、<br>"
        "　メーカー確認後あらためてご連絡いたします。<br><br>"
    )


def _build_signature(branch: BranchSettings) -> str:
    """署名"""
    h = html_escape
    return (
        "ご確認よろしくお願いいたします。<br>"
        "<div style='margin-top: 25px; color: #666666; font-size: 13px;'>"
        f"マツモト産業株式会社<br>"
        f"{h(branch.name)}"
        "</div>"
    )


# ============================================
# Outlook COM連携（VBA CreateEmails L5544-5644 相当）
# ============================================
def create_outlook_drafts(
    email_data_list: list[dict],
) -> list[str]:
    """メールデータからOutlook下書きを作成する。

    VBAの CreateEmails における .Display / .Save 相当。
    .Send() は絶対に呼ばない。

    Args:
        email_data_list: create_emails() の戻り値。各要素は辞書:
            - to: 宛先メールアドレス
            - subject: 件名
            - html_body: HTML本文
            - attachments: 添付ファイルパスのリスト
            - shared_email: 共有メールアドレス（差出人）
            - customer_name: 顧客名（ログ用）

    Returns:
        作成成功した顧客名のリスト
    """
    import win32com.client

    # Outlookアプリケーション取得（VBA L5544-5551相当）
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            print(f"エラー: Outlookに接続できません: {e}")
            return []

    created: list[str] = []

    for data in email_data_list:
        customer_name = data.get("customer_name", "")
        try:
            # olMailItem = 0（VBA L5593相当）
            mail = outlook.CreateItem(0)

            mail.To = data.get("to", "")
            mail.Subject = data.get("subject", "")
            mail.HTMLBody = data.get("html_body", "")

            # 添付ファイル（VBA L5628相当）
            for attachment_path in data.get("attachments", []):
                if attachment_path:
                    mail.Attachments.Add(os.path.abspath(str(attachment_path)))

            # 共有メールアドレス（VBA L5631-5637相当）
            shared_email = data.get("shared_email", "")
            if shared_email:
                try:
                    mail.SentOnBehalfOfName = shared_email
                except Exception:
                    pass

            # 下書き保存のみ（.Send()は絶対に呼ばない）
            mail.Save()
            created.append(customer_name)

        except Exception as e:
            print(f"  メール作成エラー ({customer_name}): {e}")

    return created


def create_outlook_sends(
    email_data_list: list[dict],
) -> list[str]:
    """メールデータからOutlookで直接送信する。

    VBAの CreateEmails における sendDirectly=True (.Send) 相当。
    宛先（To）が空のメールは送信せずスキップする。

    Args:
        email_data_list: create_emails() の戻り値。各要素は辞書:
            - to: 宛先メールアドレス
            - subject: 件名
            - html_body: HTML本文
            - attachments: 添付ファイルパスのリスト
            - shared_email: 共有メールアドレス（差出人）
            - customer_name: 顧客名（ログ用）

    Returns:
        送信成功した顧客名のリスト
    """
    import win32com.client

    # Outlookアプリケーション取得
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            print(f"エラー: Outlookに接続できません: {e}")
            return []

    sent: list[str] = []

    for data in email_data_list:
        customer_name = data.get("customer_name", "")
        to_addr = data.get("to", "").strip()

        # 宛先が空の場合はスキップ
        if not to_addr:
            print(f"  送信スキップ ({customer_name}): 宛先が空です")
            continue

        try:
            # olMailItem = 0
            mail = outlook.CreateItem(0)

            mail.To = to_addr
            mail.Subject = data.get("subject", "")
            mail.HTMLBody = data.get("html_body", "")

            # 添付ファイル
            for attachment_path in data.get("attachments", []):
                if attachment_path:
                    mail.Attachments.Add(os.path.abspath(str(attachment_path)))

            # 共有メールアドレス（差出人）
            shared_email = data.get("shared_email", "")
            if shared_email:
                try:
                    mail.SentOnBehalfOfName = shared_email
                except Exception:
                    pass

            # 直接送信（VBA L5639: .Send 相当）
            mail.Send()
            sent.append(customer_name)

        except Exception as e:
            print(f"  メール送信エラー ({customer_name}): {e}")

    return sent
