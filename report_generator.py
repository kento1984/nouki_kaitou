"""回答書生成モジュール（統合オーケストレーション）

VBAの以下の関数を移植:
- GroupOrderNumbersByCustomer (L932): 注番を顧客別にグループ化
- CreateDeliveryReportByOrderNumbers (L986): 注番指定モードで回答書作成
- CreateDeliveryReport (L2323): 期間指定モードで回答書作成
- CopyDataRow (L3784) のデータ変換ロジック: OrderRow → ReportRow変換
"""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Optional

from openpyxl import Workbook

from nouki_kaitou.bunno import (
    calculate_bunno_date,
    extract_bunno_info,
    has_bunno_mitei,
    remove_bunno_text,
)
from nouki_kaitou.confirming import get_confirmed_delivery_date, get_confirming_status
from nouki_kaitou.customer import is_route_delivery
from nouki_kaitou.delivery_calc import (
    calculate_delivery_date,
    extract_pickup_date,
)
from nouki_kaitou.excel_writer import (
    copy_data_row,
    copy_twf_data_row,
    create_header,
    format_report,
)
from nouki_kaitou.manufacturer import get_delivery_days_to_add, get_manufacturer_name
from nouki_kaitou.models import (
    BranchSettings,
    BunnoEntry,
    CacheStore,
    ConfirmingRecord,
    HolidayMap,
    HistoryRecord,
    OrderRow,
    ReportResult,
    ReportRow,
    StockoutEntry,
    TrackingEntry,
)
from nouki_kaitou.representative import (
    get_rep_list,
    should_include_for_rep,
)
from nouki_kaitou.stockout import extract_approx_delivery, remove_stockout_text
from nouki_kaitou.tracking import clean_external_comment, extract_tracking_info
from nouki_kaitou.twf import (
    TWF_FILENAME_TAG,
    TWF_NOTICE_EXCEL,
    TWF_REPORT_TITLE,
    TWF_SHEET_PREFIX,
    TwfDetailInfo,
    build_twf_info_map,
    parse_twf_comment,
    remove_twf_text,
    twf_sort_key,
)
from nouki_kaitou.utils import (
    build_report_filename,
    build_sheet_name,
    normalize_item_group_code,
    normalize_name_for_comparison,
)

# 品名先頭からメーカー名を抽出する特殊コード
# 旧: Z99(その他) / Z97(その他(修理))
# 新: 0581(その他) / 0579(その他(修理))
_SPECIAL_MFG_FROM_PRODUCT_NAME = frozenset({"Z99", "Z97", "0581", "0579"})


# ============================================
# VBA: GroupOrderNumbersByCustomer (L932-982)
# 注番を顧客別にグループ化
# ============================================
def group_order_numbers_by_customer(
    source_data: list[OrderRow],
    order_numbers: list[str],
) -> dict[str, list[str]]:
    """注番を顧客名でグループ化する。

    Args:
        source_data: 全受注データ
        order_numbers: グループ化対象の注番リスト

    Returns:
        dict[顧客名, 注番リスト]（重複排除済み）
    """
    result: dict[str, list[str]] = {}

    for order_num in order_numbers:
        order_num_stripped = order_num.strip()
        customer_name = ""

        # 注番から顧客名を検索
        for row in source_data:
            if row.order_number.strip() == order_num_stripped:
                customer_name = row.customer_name.strip()
                break

        if not customer_name:
            continue

        if customer_name not in result:
            result[customer_name] = []

        # 重複チェック
        if order_num_stripped not in result[customer_name]:
            result[customer_name].append(order_num_stripped)

    return result


# ============================================
# VBA: CopyDataRow (L3784-4001) のデータ変換部分
# OrderRow → ReportRow変換
# ============================================
def build_report_row(
    row: OrderRow,
    cache: CacheStore,
    holidays: HolidayMap | None = None,
    branch: BranchSettings | None = None,
    execution_time: datetime.datetime | None = None,
    force_delivered: bool = False,
    today: datetime.date | None = None,
) -> tuple[ReportRow, str]:
    """OrderRowからReportRow（Excel出力用）を構築する。

    VBAのCopyDataRowのデータ変換ロジックを移植。
    - メーカー名解決（Z99/Z97特殊処理含む）
    - 納期計算
    - 分納判定 → 「分納」表示
    - 欠品判定 → 「欠品中」/「（欠品）」表示
    - 納入先名解決（引取/貴社/様）
    - 価格確認中判定
    - 備考テキストクリーニング

    Args:
        row: 受注データ行
        cache: マスターキャッシュ
        holidays: 祝日辞書
        branch: 営業所設定
        execution_time: 実行時刻
        force_delivered: 直送処理完了時の納品済み強制フラグ
        today: 基準日（テスト用）

    Returns:
        (ReportRow, 納期回答文字列) のタプル
    """
    if today is None:
        today = datetime.date.today()

    # --- メーカー名解決 ---
    manufacturer_name = _resolve_manufacturer_name(row, cache)

    # --- 品名解決（Z99/Z97） ---
    product_name = _resolve_product_name(row, manufacturer_name)

    # --- 納期計算 ---
    delivery_answer = calculate_delivery_date(
        row, cache, holidays, branch, execution_time, today
    )

    # --- 分納判定 ---
    bunno_info = extract_bunno_info(row.comment_detail)
    if bunno_info and row.ship_status != "処理完了":
        delivery_answer = "分納"

    # --- 欠品判定 ---
    # 確認中一覧に確定日がある場合は欠品overrideをスキップ
    has_confirmed_date = False
    confirmed_date = get_confirmed_delivery_date(
        row.order_number, row.detail_number, cache
    )
    if confirmed_date is not None:
        has_confirmed_date = True

    if ("欠品中" in row.comment_detail
            and row.ship_status != "処理完了"
            and not has_confirmed_date):
        if delivery_answer in ("確認中", "日程調整中"):
            delivery_answer = "欠品中"
        else:
            delivery_answer = delivery_answer + "（欠品）"

    # --- 納入先名解決 ---
    delivery_place = _resolve_delivery_place(row, today)

    # --- 価格確認中判定 ---
    unit_price, net_amount = _resolve_price(
        row, delivery_answer, force_delivered
    )

    # --- 備考テキストクリーニング ---
    remark = remove_stockout_text(row.comment_detail)
    remark = remove_bunno_text(remark)
    remark = remove_twf_text(remark)
    remark = remark.strip()

    report_row = ReportRow(
        registration_date=row.registration_date,
        customer_contact=row.customer_contact,
        customer_order_number=row.customer_order_number,
        manufacturer_name=manufacturer_name,
        product_name=product_name,
        quantity=row.quantity,
        unit_price=unit_price,
        net_amount=net_amount,
        delivery_answer=delivery_answer,
        delivery_place=delivery_place,
        remarks=remark,
        order_number=row.order_number,
    )

    return report_row, delivery_answer


# ============================================
# VBA: CreateDeliveryReportByOrderNumbers (L986-1345)
# 注番指定モードで回答書作成
# ============================================
def create_delivery_report_by_order_numbers(
    source_data: list[OrderRow],
    customer_name: str,
    order_numbers: list[str],
    cache: CacheStore,
    output_dir: Path | str,
    holidays: HolidayMap | None = None,
    branch: BranchSettings | None = None,
    execution_time: datetime.datetime | None = None,
    rep_name: str = "",
    rep_master_ws: object = None,
    today: datetime.date | None = None,
) -> ReportResult | None:
    """注番指定モードで納期回答書を作成する。

    送付履歴チェックなし。指定された注番の伝票を全て出力。

    Args:
        source_data: 全受注データ
        customer_name: 顧客名
        order_numbers: 対象注番リスト
        cache: マスターキャッシュ
        output_dir: 出力先ディレクトリ
        holidays: 祝日辞書
        branch: 営業所設定
        execution_time: 実行時刻
        rep_name: 担当者名（分割送信時）
        rep_master_ws: 担当者マスターシート
        today: 基準日（テスト用）

    Returns:
        ReportResult（データなしならNone）
    """
    if today is None:
        today = datetime.date.today()
    if branch is None:
        branch = BranchSettings()
    if execution_time is None:
        execution_time = datetime.datetime.now()

    # ワークブック作成
    wb = Workbook()
    ws = wb.active
    ws.title = build_sheet_name(customer_name, rep_name)
    create_header(ws, customer_name, rep_name, today, branch)

    current_row = 7
    is_external_mode = branch is not None and branch.remarks_mode == "external"

    # 担当者フィルタ用キャッシュ
    registered_rep_list: list[str] = []
    if rep_name == "__OTHER__" and rep_master_ws is not None:
        registered_rep_list = get_rep_list(customer_name, rep_master_ws)

    # 情報収集用
    tracking_info_list: list[tuple[str, str, str, TrackingEntry]] = []
    stockout_info_list: list[StockoutEntry] = []
    bunno_info_list: list[dict] = []

    order_set = {n.strip() for n in order_numbers}

    for row in source_data:
        if row.order_number.strip() not in order_set:
            continue

        # --- フィルタ ---
        if not _pass_basic_filter(row):
            continue
        if rep_name and not should_include_for_rep(
            row.customer_contact, rep_name, registered_rep_list
        ):
            continue

        # --- データ変換・書き込み ---
        report_row, delivery_status = build_report_row(
            row, cache, holidays, branch, execution_time, False, today
        )
        ext_comment = None
        if is_external_mode:
            raw = row.comment_external.strip()
            ext_comment = clean_external_comment(raw) if raw else ""
        copy_data_row(ws, current_row, report_row, ext_comment)

        # --- 情報収集 ---
        _collect_tracking_info(row, cache, tracking_info_list)
        _collect_stockout_info(
            row, cache, delivery_status, stockout_info_list
        )
        _collect_bunno_info(
            row, cache, customer_name, holidays,
            bunno_info_list, today
        )

        current_row += 1

    # データなし
    if current_row == 7:
        return None

    # 書式設定
    format_report(
        ws, current_row - 1, branch,
        tracking_info_list, stockout_info_list,
        bunno_info_list, None,
        holidays, cache, today,
    )

    # 保存
    filename = build_report_filename(
        customer_name, execution_time, rep_name, list(order_set)
    )
    file_path = str(Path(output_dir) / filename)
    wb.save(file_path)

    return ReportResult(
        file_path=file_path,
        customer_name=customer_name,
        rep_name=rep_name,
        tracking_info_list=tracking_info_list,
        stockout_info_list=stockout_info_list,
        bunno_info_list=bunno_info_list,
    )


# ============================================
# VBA: CreateDeliveryReport (L2323-2928)
# 期間指定モードで回答書作成
# ============================================
def create_delivery_report(
    source_data: list[OrderRow],
    customer_name: str,
    sent_orders: dict[str, str],
    cache: CacheStore,
    output_dir: Path | str,
    holidays: HolidayMap | None = None,
    branch: BranchSettings | None = None,
    execution_time: datetime.datetime | None = None,
    date_from: datetime.date | None = None,
    date_to: datetime.date | None = None,
    rep_name: str = "",
    rep_master_ws: object = None,
    today: datetime.date | None = None,
    exclude_orders: set[str] | None = None,
    include_only_orders: set[str] | None = None,
    filter_already_sent: bool = True,
    twf_mode: bool = False,
) -> ReportResult | None:
    """期間指定モードで納期回答書を作成する。

    送付履歴チェックあり。確定伝票と確認中伝票を分類して返す。

    Args:
        source_data: 全受注データ
        customer_name: 顧客名
        sent_orders: 送付済み伝票辞書 {注番|明細: ステータス}
        cache: マスターキャッシュ
        output_dir: 出力先ディレクトリ
        holidays: 祝日辞書
        branch: 営業所設定
        execution_time: 実行時刻
        date_from: 期間開始日
        date_to: 期間終了日
        rep_name: 担当者名
        rep_master_ws: 担当者マスターシート
        today: 基準日（テスト用）
        exclude_orders: 除外する注番セット（通常回答書からTWF注番を除く用）
        include_only_orders: この注番セットのみ対象とする（TWF専用回答書用）。
            Noneなら全注番が対象
        filter_already_sent: Falseなら送付済みチェックをスキップして
            毎回表示する（除外マーカー・除外ステータスは常に有効）。
            履歴への記録判定（分類）は通常どおり行う
        twf_mode: TWF展示会専用回答書モード。タイトル・ファイル名・
            シート名・注記をTWF用に切り替え、処理完了+回答済みの伝票を
            「納品済み」表示に上書きする

    Returns:
        ReportResult（データなしならNone）
    """
    if today is None:
        today = datetime.date.today()
    if branch is None:
        branch = BranchSettings()
    if execution_time is None:
        execution_time = datetime.datetime.now()

    # ワークブック作成
    wb = Workbook()
    ws = wb.active
    sheet_prefix = TWF_SHEET_PREFIX if twf_mode else ""
    report_title = TWF_REPORT_TITLE if twf_mode else None
    ws.title = build_sheet_name(customer_name, rep_name, prefix=sheet_prefix)
    create_header(
        ws, customer_name, rep_name, today, branch,
        title=report_title, twf_layout=twf_mode,
    )

    current_row = 7
    is_external_mode = branch is not None and branch.remarks_mode == "external"

    # 担当者フィルタ用キャッシュ
    registered_rep_list: list[str] = []
    if rep_name == "__OTHER__" and rep_master_ws is not None:
        registered_rep_list = get_rep_list(customer_name, rep_master_ws)

    # TWFモード: 注番→TWF情報（入れ忘れ明細への引き継ぎ用）と
    # 表示行の一時保持（TWF No.昇順ソートのため2パスで書き込む）
    twf_info_map: dict[str, TwfDetailInfo] = (
        build_twf_info_map(source_data) if twf_mode else {}
    )
    twf_pending: list[tuple[tuple, TwfDetailInfo, ReportRow]] = []

    # 情報収集用
    confirmed_orders: list[HistoryRecord] = []
    confirming_orders: list[ConfirmingRecord] = []
    tracking_info_list: list[tuple[str, str, str, TrackingEntry]] = []
    stockout_info_list: list[StockoutEntry] = []
    bunno_info_list: list[dict] = []
    bunno_completed_list: list[tuple[str, str, str]] = []

    for row in source_data:
        # --- 顧客名フィルタ ---
        if row.customer_name.strip() != customer_name:
            continue

        # --- TWF注番フィルタ（通常回答書: 除外 / TWF回答書: 限定） ---
        order_num = row.order_number.strip()
        if exclude_orders and order_num in exclude_orders:
            continue
        if include_only_orders is not None and order_num not in include_only_orders:
            continue

        # --- 担当者フィルタ ---
        if rep_name and not should_include_for_rep(
            row.customer_contact, rep_name, registered_rep_list
        ):
            continue

        # --- 期間フィルタ ---
        if row.registration_date is None:
            continue
        if date_from and date_to:
            if row.registration_date < date_from or row.registration_date > date_to:
                continue

        # --- 基本フィルタ ---
        if not _pass_basic_filter(row):
            continue

        # --- 送付済みチェック ---
        history_key = f"{row.order_number}|{row.detail_number}"
        is_already_sent = False
        previous_status = sent_orders.get(history_key, "")

        # 除外判定（送付履歴 or 確認中一覧）。
        # filter_already_sent=False（TWF回答書）でも除外は常に有効
        is_excluded = _is_excluded(
            row, previous_status, cache
        )
        if is_excluded:
            continue

        if previous_status == "分納完了":
            is_already_sent = True
        elif row.ship_status == "処理完了" and previous_status == "確認中":
            is_already_sent = False  # 処理完了で確認中→今回出す
        elif row.ship_status == "処理完了" and previous_status:
            is_already_sent = True
        elif previous_status and row.registration_date < today:
            is_already_sent = True

        if is_already_sent and filter_already_sent:
            continue

        # --- forceDelivered / isHimozuki / isBunnoCompleted ---
        force_delivered, is_himozuki, is_bunno_completed = _determine_flags(
            row, cache, sent_orders, history_key, previous_status
        )

        # --- データ変換・書き込み ---
        report_row, delivery_status = build_report_row(
            row, cache, holidays, branch, execution_time,
            force_delivered, today
        )

        # forceDelivered時の納品済み上書き
        if force_delivered and delivery_status in (
            "確認中", "欠品中", "日程調整中"
        ) or (force_delivered and "（欠品）" in delivery_status):
            delivery_status = "納品済み"
            report_row.delivery_answer = "納品済み"

        # 分納完了 → 納品済み上書き
        if is_bunno_completed:
            delivery_status = "納品済み"
            report_row.delivery_answer = "納品済み"

        # TWFモード: 処理完了かつ回答済み（履歴に確定記録あり）の伝票は
        # 「納品済み」表示に上書きする。
        # 履歴チェックを表示でスキップするため、force_delivered の既存判定
        # （履歴が空 or 確認中のときのみ発火）を通らないケースの補完
        if (twf_mode
                and row.ship_status == "処理完了"
                and previous_status not in ("", "確認中")):
            delivery_status = "納品済み"
            report_row.delivery_answer = "納品済み"

        if twf_mode:
            # --- TWF専用レイアウト用のデータ整形 ---
            # TWF記載を分解（A列=番号, C列=お客様名, K列=メモ+既存備考）。
            # 記載のない明細（入れ忘れ救済分）は同注番の他明細から
            # 番号とお客様名のみ引き継ぐ（memoは明細固有なので引き継がない）
            twf_info = parse_twf_comment(row.comment_detail)
            if twf_info is None:
                base = twf_info_map.get(order_num)
                twf_info = (
                    TwfDetailInfo(number=base.number, customer=base.customer)
                    if base else TwfDetailInfo()
                )

            # K列: TWFメモ + 既存備考（build_report_rowでTWF記載除去済み）
            if twf_info.memo:
                report_row.remarks = (
                    f"{twf_info.memo} ／ {report_row.remarks}"
                    if report_row.remarks else twf_info.memo
                )

            # 納入先名: 「ワンタイム出荷先」（SAP内部表現）→「ご指定先」に置換
            if report_row.delivery_place.startswith("ワンタイム出荷先"):
                report_row.delivery_place = "ご指定先"

            # TWF No.昇順ソートのため書き込みは後段でまとめて行う
            twf_pending.append((
                twf_sort_key(twf_info.number, row.order_number, row.detail_number),
                twf_info,
                report_row,
            ))
        else:
            ext_comment = None
            if is_external_mode:
                raw = row.comment_external.strip()
                ext_comment = clean_external_comment(raw) if raw else ""
            copy_data_row(ws, current_row, report_row, ext_comment)

        # --- メーカー名・品名を解決（情報収集用） ---
        manufacturer_name = report_row.manufacturer_name
        product_name = report_row.product_name

        # --- 送り状情報収集 ---
        _collect_tracking_info(row, cache, tracking_info_list)

        # --- 欠品情報収集 ---
        _collect_stockout_info(
            row, cache, delivery_status, stockout_info_list
        )

        # --- 分納情報収集 ---
        bunno_info = extract_bunno_info(row.comment_detail)

        # 分納完了 → 品名・数量を通知用に収集
        if bunno_info and row.product_name != "送料" and is_bunno_completed:
            bunno_completed_list.append(
                (manufacturer_name, product_name, row.quantity)
            )

        # 処理完了でない場合のみ分納情報を収集
        if bunno_info and row.product_name != "送料" and row.ship_status != "処理完了":
            _collect_bunno_info(
                row, cache, customer_name, holidays,
                bunno_info_list, today
            )

        # --- 確定/確認中の分類 ---
        _classify_order(
            row, delivery_status, cache, bunno_info,
            is_bunno_completed, manufacturer_name, product_name,
            confirmed_orders, confirming_orders,
        )

        current_row += 1

    # データなし
    if current_row == 7:
        return None

    # TWFモード: TWF No.昇順（同一No.内は注番→明細順、番号なしは末尾）で書き込み
    if twf_mode:
        twf_pending.sort(key=lambda t: t[0])
        for i, (_, twf_info, pending_row) in enumerate(twf_pending):
            copy_twf_data_row(
                ws, 7 + i, pending_row, twf_info.number, twf_info.customer
            )

    # 書式設定
    format_report(
        ws, current_row - 1, branch,
        tracking_info_list, stockout_info_list,
        bunno_info_list, bunno_completed_list,
        holidays, cache, today,
        twf_notice=TWF_NOTICE_EXCEL if twf_mode else None,
        with_auto_filter=twf_mode,
    )

    # 保存
    filename = build_report_filename(
        customer_name, execution_time, rep_name,
        filename_tag=TWF_FILENAME_TAG if twf_mode else "",
    )
    file_path = str(Path(output_dir) / filename)
    wb.save(file_path)

    return ReportResult(
        file_path=file_path,
        customer_name=customer_name,
        rep_name=rep_name,
        confirmed_orders=confirmed_orders,
        confirming_orders=confirming_orders,
        tracking_info_list=tracking_info_list,
        stockout_info_list=stockout_info_list,
        bunno_info_list=bunno_info_list,
        bunno_completed_list=bunno_completed_list,
        has_confirming=len(confirming_orders) > 0,
        is_twf=twf_mode,
    )


# ============================================
# ヘルパー関数（private）
# ============================================

def _resolve_manufacturer_name(row: OrderRow, cache: CacheStore) -> str:
    """メーカー名を解決する（Z99/Z97特殊処理含む）。"""
    code = normalize_item_group_code(row.item_group_code)

    if code in _SPECIAL_MFG_FROM_PRODUCT_NAME:
        # 品名の先頭部分からメーカー名を抽出
        full_text = row.product_name.strip()
        # 半角スペース→全角スペースの順で探す
        space_pos = full_text.find(" ")
        if space_pos < 0:
            space_pos = full_text.find("\u3000")
        if space_pos > 0:
            return full_text[:space_pos]
        return ""

    mfg_name = get_manufacturer_name(code, cache)

    if not mfg_name:
        # フォールバック: メーカー列の値 → 品目GroupCode
        if row.manufacturer_name.strip():
            return row.manufacturer_name.strip()
        return code

    return mfg_name


def _resolve_product_name(row: OrderRow, manufacturer_name: str) -> str:
    """品名を解決する（Z99/Z97はメーカー名部分を除去）。"""
    product_name = row.product_name.strip()
    code = normalize_item_group_code(row.item_group_code)

    if code in _SPECIAL_MFG_FROM_PRODUCT_NAME and manufacturer_name:
        # メーカー名の後のスペース以降を品名とする
        space_pos = product_name.find(" ")
        if space_pos < 0:
            space_pos = product_name.find("\u3000")
        if space_pos > 0:
            return product_name[space_pos + 1:].strip()

    return product_name


def _resolve_delivery_place(row: OrderRow, today: datetime.date) -> str:
    """納入先名を解決する。"""
    delivery_place = row.ship_to_name.strip()

    # 引取判定
    comment = (row.comment_external.strip() + " " + row.comment_internal.strip()).strip()
    pickup_date = extract_pickup_date(comment, today)
    if pickup_date is not None:
        return "お引き取り"

    # 受注先と同じ → 貴社（全角/半角正規化して比較）
    if (normalize_name_for_comparison(delivery_place)
            == normalize_name_for_comparison(row.customer_name)):
        return "貴社"

    # 「様」がついていなければ付与
    if delivery_place and not delivery_place.endswith("様"):
        return delivery_place + "様"

    return delivery_place


def _is_provisional_price(unit_price: object) -> bool:
    """単価が仮単価（1円）かどうか判定する。

    SAPデータは "1", "1.00", "1.000" 等の形式があるため、
    float変換で統一的に判定する。
    """
    try:
        return float(str(unit_price).replace(",", "")) == 1.0
    except (ValueError, TypeError):
        return False


def _resolve_price(
    row: OrderRow,
    delivery_answer: str,
    force_delivered: bool,
) -> tuple[object, object]:
    """単価・金額を解決する。"""
    unit_price: object = row.unit_price
    net_amount: object = row.net_amount

    # $$フラグ: コメント（社内）に「$$」があれば価格は確定
    price_confirmed = (
        "$$" in row.comment_internal or "＄＄" in row.comment_internal
    )

    # $$があれば仮単価でもそのまま表示（価格確定済みの明示）
    if price_confirmed:
        return unit_price, net_amount

    # 単価=1（仮単価）→ 確認中表示（小数点付きにも対応）
    if _is_provisional_price(unit_price):
        return "確認中", "確認中"

    # 納期が確認中かつforceDeliveredなし → 確認中表示
    if delivery_answer == "確認中" and not force_delivered:
        return "確認中", "確認中"

    return unit_price, net_amount


def _pass_basic_filter(row: OrderRow) -> bool:
    """基本フィルタ（拒否理由・伝票タイプ・##除外）。"""
    # 明細削除
    if row.rejection_reason.strip() == "明細削除":
        return False

    # 伝票タイプ
    doc_type = row.document_type.strip()
    if doc_type not in ("【受注】直送販売", "【受注】在庫販売"):
        return False

    # ##除外
    internal = row.comment_internal.strip()
    if "##" in internal or "＃＃" in internal:
        return False

    return True


def _is_excluded(
    row: OrderRow,
    previous_status: str,
    cache: CacheStore,
) -> bool:
    """除外判定（送付履歴の「除外」 or 確認中一覧の「除外」）。"""
    if previous_status == "除外":
        return True

    # 確認中一覧から除外チェック
    key = f"{row.order_number}|{row.detail_number}"
    entry = cache.confirm.get(key)
    if entry is not None:
        inquiry_status = entry[0]  # 問合せ状況
        if inquiry_status == "除外":
            return True

    return False


def _determine_flags(
    row: OrderRow,
    cache: CacheStore,
    sent_orders: dict[str, str],
    history_key: str,
    previous_status: str,
) -> tuple[bool, bool, bool]:
    """forceDelivered, isHimozuki, isBunnoCompletedを判定する。

    Returns:
        (force_delivered, is_himozuki, is_bunno_completed)
    """
    force_delivered = False
    is_himozuki = False
    is_bunno_completed = False

    # 紐付き判定
    if row.document_type.strip() == "【受注】直送販売":
        storage = row.storage_place.strip()
        if not storage:
            storage = cache.storage.get(row.order_number, "")
        if storage != "転送中（直送用）":
            is_himozuki = True

    # 確認中一覧のステータス取得
    confirming_status = get_confirming_status(
        row.order_number, row.detail_number, cache
    )

    # 分納+処理完了 → 分納完了
    is_bunno_in_confirming = (confirming_status == "分納")
    if is_bunno_in_confirming and row.ship_status == "処理完了":
        is_bunno_completed = True
        is_bunno_in_confirming = False

    # forceDelivered判定
    if row.ship_status == "処理完了":
        if not sent_orders.get(history_key):
            # 送付履歴にない → 納品済みで出す（紐付き・分納を除く）
            if not is_himozuki and not is_bunno_in_confirming:
                force_delivered = True
        elif previous_status == "確認中":
            # 前回「確認中」→ 今回「納品済み」
            if not is_himozuki and not is_bunno_in_confirming:
                force_delivered = True

    return force_delivered, is_himozuki, is_bunno_completed


def _collect_tracking_info(
    row: OrderRow,
    cache: CacheStore,
    tracking_info_list: list[tuple[str, str, str, TrackingEntry]],
) -> None:
    """送り状情報を収集する。"""
    if row.product_name.strip() == "送料":
        return

    entries = extract_tracking_info(row.comment_external)
    if not entries:
        return

    mfg = _resolve_manufacturer_name(row, cache)
    product = row.product_name.strip()
    short_product = product[:25] + "..." if len(product) > 25 else product

    for entry in entries:
        tracking_info_list.append((mfg, short_product, row.quantity, entry))


def _collect_stockout_info(
    row: OrderRow,
    cache: CacheStore,
    delivery_status: str,
    stockout_info_list: list[StockoutEntry],
) -> None:
    """欠品情報を収集する。"""
    if "欠品中" not in row.comment_detail:
        return
    if row.product_name.strip() == "送料":
        return
    # 分納がある場合は分納セクションで表示
    if "分納:" in row.comment_detail or "分納：" in row.comment_detail:
        return
    # 処理完了なら欠品解消済み
    if row.ship_status == "処理完了":
        return

    mfg = _resolve_manufacturer_name(row, cache)
    product = row.product_name.strip()
    short_product = product[:25] + "..." if len(product) > 25 else product
    approx = extract_approx_delivery(row.comment_detail)

    stockout_info_list.append(StockoutEntry(
        manufacturer_name=mfg,
        product_name=short_product,
        quantity=row.quantity,
        delivery=delivery_status,
        approx_delivery=approx,
        order_number=row.order_number,
    ))


def _collect_bunno_info(
    row: OrderRow,
    cache: CacheStore,
    customer_name: str,
    holidays: HolidayMap | None,
    bunno_info_list: list[dict],
    today: datetime.date,
) -> None:
    """分納情報を収集する。"""
    bunno_entries = extract_bunno_info(row.comment_detail)
    if not bunno_entries:
        return
    if row.product_name.strip() == "送料":
        return
    # 処理完了なら分納コメントは残骸
    if row.ship_status == "処理完了":
        return

    mfg = _resolve_manufacturer_name(row, cache)
    product = row.product_name.strip()

    # isShipRule判定
    is_ship_rule = False
    storage = row.storage_place.strip()
    doc_type = row.document_type.strip()

    if storage == "転送中（直送用）":
        is_ship_rule = True
    elif doc_type == "【受注】在庫販売":
        if normalize_name_for_comparison(customer_name) != normalize_name_for_comparison(row.ship_to_name):
            is_ship_rule = True

    is_rosenbin = is_route_delivery(customer_name, cache)

    # 在庫販売 + 路線便 → 出荷予定扱い
    if not is_ship_rule and is_rosenbin and doc_type == "【受注】在庫販売":
        is_ship_rule = True

    days_to_add = get_delivery_days_to_add(row.item_group_code, cache)

    # 計算済み分納情報を生成
    calc_details: list[list[str]] = []
    for entry in bunno_entries:
        calc_date = calculate_bunno_date(
            entry.date_str, is_ship_rule, days_to_add,
            holidays, cache, row.order_number, row.detail_number,
            is_rosenbin, today,
        )
        location = entry.location or ""
        calc_details.append([entry.quantity, entry.date_str, location, calc_date])

    bunno_info_list.append({
        "manufacturer": mfg,
        "product": product,
        "quantity": row.quantity,
        "entries": bunno_entries,
        "calc_details": calc_details,
        "is_ship_rule": is_ship_rule,
        "days_to_add": days_to_add,
        "order_number": row.order_number,
        "detail_number": row.detail_number,
        "is_rosenbin": is_rosenbin,
    })


def _classify_order(
    row: OrderRow,
    delivery_status: str,
    cache: CacheStore,
    bunno_info: list[BunnoEntry],
    is_bunno_completed: bool,
    manufacturer_name: str,
    product_name: str,
    confirmed_orders: list[HistoryRecord],
    confirming_orders: list[ConfirmingRecord],
) -> None:
    """伝票を確定/確認中に分類する。"""
    if not delivery_status:
        return

    # 確認中一覧のステータスを取得
    prev_confirming_status = get_confirming_status(
        row.order_number, row.detail_number, cache
    )

    # keepInConfirming: 分納で確認中一覧に残すべきか
    keep_in_confirming = (
        prev_confirming_status == "分納" and row.ship_status != "処理完了"
    )

    # 価格確認中: 納期は確定しているが仮単価（1円）のまま
    # 納期未確定の場合は既存分岐（「未処理」等）に任せる
    # $$フラグがあれば価格確定済みなので対象外
    price_confirmed = (
        "$$" in row.comment_internal or "＄＄" in row.comment_internal
    )
    delivery_is_undecided = (
        delivery_status in ("確認中", "欠品中", "日程調整中")
        or "（欠品）" in delivery_status
        or "分納" in delivery_status
    )
    is_price_pending = (
        _is_provisional_price(row.unit_price)
        and not price_confirmed
        and not delivery_is_undecided
    )

    # 確認中一覧に追加すべきか判定
    is_confirming_type = (
        delivery_status in ("確認中", "欠品中", "日程調整中")
        or "（欠品）" in delivery_status
        or "分納" in delivery_status
        or keep_in_confirming
        or is_price_pending
    )

    if is_confirming_type:
        # 分納の場合
        if "分納" in delivery_status or prev_confirming_status == "分納":
            if has_bunno_mitei(bunno_info, cache, row.order_number, row.detail_number):
                ship_status_for_confirm = "分納"
            else:
                # 未定なし（全分確定 or 新規で未定なし）→ 送付履歴へ
                confirmed_orders.append(HistoryRecord(
                    order_date=row.registration_date,
                    customer_name=row.customer_name,
                    order_number=row.order_number,
                    detail_number=row.detail_number,
                    manufacturer_name=manufacturer_name,
                    product_name=product_name,
                    delivery_answer="分納完了",
                ))
                return

        # 欠品の場合
        elif delivery_status == "欠品中" or "（欠品）" in delivery_status:
            ship_status_for_confirm = "欠品中"
        # 価格確認中の場合
        elif is_price_pending:
            ship_status_for_confirm = "価格確認中"
        else:
            ship_status_for_confirm = row.ship_status

        confirming_orders.append(ConfirmingRecord(
            order_date=row.registration_date,
            customer_name=row.customer_name,
            order_number=row.order_number,
            detail_number=row.detail_number,
            manufacturer_name=manufacturer_name,
            product_name=product_name,
            status=ship_status_for_confirm,
            order_delivery_date=row.order_delivery_date,
        ))
    else:
        # 確定 → 送付履歴へ
        status_for_history = "分納完了" if is_bunno_completed else delivery_status
        confirmed_orders.append(HistoryRecord(
            order_date=row.registration_date,
            customer_name=row.customer_name,
            order_number=row.order_number,
            detail_number=row.detail_number,
            manufacturer_name=manufacturer_name,
            product_name=product_name,
            delivery_answer=status_for_history,
        ))
