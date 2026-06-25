"""データクラス定義

VBAのグローバル変数・構造体・Collection/Arrayをdataclass化。
各関数間のデータ受け渡しに使用する。
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from typing import Optional


# ============================================
# 配送パターン（仙台営業所等のマルチ便対応）
# ============================================
@dataclass
class DeliveryPattern:
    """配送パターン定義（メーカー一覧「配送パターン」シートから読込）

    1段階（cutoff2=None）: cutoff1で2択（before/after）
    2段階（cutoff2あり）: cutoff1/cutoff2で3択
    """

    name: str = ""
    cutoff1: tuple[int, int] = (15, 0)
    days_before_cutoff1: int = 0
    cutoff2: tuple[int, int] | None = None
    days_between_cutoffs: int = 1

    @property
    def days_after_all(self) -> int:
        if self.cutoff2 is not None:
            return self.days_between_cutoffs
        return self.days_before_cutoff1 + 1


# ============================================
# 営業所設定（VBA: g_BranchName等のグローバル変数群）
# ============================================
@dataclass
class BranchSettings:
    """営業所設定（メーカー一覧.xlsx「営業所設定」シートから読み込み）"""

    name: str = ""                  # g_BranchName: 営業所名（例: "京葉営業所"）
    default_cutoff: int = 15        # g_DefaultCutoff: デフォルト締切時間
    base_center: str = ""           # g_BaseCenter: 商品センター名（例: "関東商品センター"）
    shared_email: str = ""          # g_SharedEmail: 共有メールアドレス
    signature: str = ""             # g_Signature: メール署名
    start_date: str = ""            # g_StartDate: 運用開始日（yyyy/mm/dd形式）
    remarks_mode: str = "detail"    # L列表示モード（detail=注番, external=連絡事項）


# ============================================
# アプリケーション実行状態（VBA: g_ExecutionTime等）
# ============================================
@dataclass
class AppConfig:
    """アプリケーション実行時の状態管理"""

    execution_time: datetime.datetime = field(
        default_factory=datetime.datetime.now
    )
    current_subfolder: str = ""     # 出力先サブフォルダパス
    tool_folder: str = ""           # ツールフォルダ（.basと同階層）
    branch: BranchSettings = field(default_factory=BranchSettings)


# ============================================
# キャッシュストア（VBA: g_MfgNameCache等の全Dictionaryキャッシュ）
# ============================================
@dataclass
class CacheStore:
    """マスターデータキャッシュ

    VBAの6つのDictionaryキャッシュ + g_SourceDataを集約。
    BuildXxxCache()で構築し、各関数に注入する。
    """

    # メーカー一覧から構築
    mfg_name: dict[str, str] = field(default_factory=dict)
    """品目GroupCode → メーカー名"""

    mfg_days: dict[str, int] = field(default_factory=dict)
    """品目GroupCode → 配送加算日数"""

    # 顧客マスターから構築
    cust_days: dict[str, list[int]] = field(default_factory=dict)
    """顧客名 → 配送曜日リスト（VBA Weekday番号: 日=1..土=7）"""

    cust_retention: dict[str, int] = field(default_factory=dict)
    """顧客名 → 保持日数"""

    cust_route: dict[str, bool] = field(default_factory=dict)
    """顧客名 → 路線便フラグ"""

    # 確認中一覧から構築
    confirm: dict[str, tuple[str, str, Optional[datetime.date]]] = field(
        default_factory=dict
    )
    """注番|明細 → (問合せ状況, ステータス, 受注納期)"""

    # 納期上書きシートから構築（任意・期間限定運用）
    delivery_overrides: dict[str, str] = field(default_factory=dict)
    """注番|明細 → 納期回答の逐語上書き文字列。

    SAP施錠等でツール計算の納期を直せない明細を、運用者が手書きした
    文字列で最優先に差し替えるための仕組み。空（=シート無し/未記入）なら
    完全に無視され、全明細が従来どおりの計算結果になる（完全no-op）。
    """

    # 顧客マスターから構築（配送パターン）
    cust_pattern: dict[str, str] = field(default_factory=dict)
    """顧客名 → 配送パターン名"""

    # メーカー一覧「配送パターン」シートから構築
    delivery_patterns: dict[str, DeliveryPattern] = field(default_factory=dict)
    """パターン名 → DeliveryPattern定義"""

    # 顧客マスターのフォーマット検出結果
    cust_email_start_col: int = 4
    """メールアドレス開始列（旧:4, 新:5）"""

    # 受注一覧から構築
    storage: dict[str, str] = field(default_factory=dict)
    """注番 → 保管場所"""


# ============================================
# 受注データ行（VBA: g_SourceData配列の1行分）
# ============================================
@dataclass
class OrderRow:
    """SAP受注伝票一覧の1データ行

    VBAではg_SourceData(rowNum, cols("列名"))でアクセスしていたものを構造体化。
    """

    order_number: str = ""              # 受発注伝票（注番）
    detail_number: str = ""             # 明細
    document_type: str = ""             # 伝票タイプ（「【受注】直送販売」「【受注】在庫販売」等）
    customer_name: str = ""             # 受注先（顧客名）
    product_name: str = ""              # テキスト（品名）
    ship_status: str = ""               # 出荷ステータス（未処理/一部処理済み/処理完了）
    quantity: str = ""                  # 受注数量
    unit_price: str = ""                # 受注単価（売単価）
    net_amount: str = ""                # 正味額（売上総額＝売価合計）
    purchase_unit_price: str = ""       # 発注単価（仕入単価）
    manufacturer_name: str = ""         # 名称（メーカー）
    storage_place: str = ""             # 保管場所
    customer_order_number: str = ""     # 得意先発注番号
    customer_contact: str = ""          # 得意先担当者
    comment_detail: str = ""            # コメント（明細）— 分納・欠品情報
    comment_external: str = ""          # コメント（社外）— 送り状番号
    comment_internal: str = ""          # コメント（社内）— ##除外、&&作業、@@着日
    rejection_reason: str = ""          # 拒否理由
    ship_to_name: str = ""              # 出荷先名
    registration_date: Optional[datetime.date] = None   # 登録日
    time_value: str = ""                # 時刻（"8:54:58"形式）
    order_delivery_date: Optional[datetime.date] = None  # 受注納期
    specified_delivery_date: Optional[datetime.date] = None  # 指定納期
    item_group_code: str = ""           # 品目 Group
    rep_name: str = ""                  # マツモト担当者名（社内手配担当。例: "柏原　賢人"）
    source_row: int = 0                 # 元データの行番号（デバッグ用）


# ============================================
# 納期回答書の1データ行（出力用）
# ============================================
@dataclass
class ReportRow:
    """納期回答書Excelに書き込む1行分のデータ

    VBAのCopyDataRowが新しいワークシートに書き込む列1〜12に対応。
    """

    registration_date: Optional[datetime.date] = None   # A列: 受注日
    customer_contact: str = ""          # B列: 担当者様
    customer_order_number: str = ""     # C列: 貴社注番
    manufacturer_name: str = ""         # D列: メーカー名
    product_name: str = ""              # E列: 品名
    quantity: str = ""                  # F列: 数量
    unit_price: str = ""                # G列: 単価（"確認中"の場合あり）
    net_amount: str = ""                # H列: 金額（"確認中"の場合あり）
    delivery_answer: str = ""           # I列: 納期回答
    delivery_place: str = ""            # J列: 納入先名
    remarks: str = ""                   # K列: 備考
    order_number: str = ""              # L列: 弊社注番


# ============================================
# 分納情報（VBA: ExtractBunnoInfoの返り値 Array(qty, dateStr, location)）
# ============================================
@dataclass
class BunnoEntry:
    """分納の1回分の情報

    VBAではCollection of Array(数量文字列, 日付文字列, 場所文字列)で管理。
    """

    quantity: str = ""       # 数量文字列（例: "30個"、"50本"）
    date_str: str = ""       # 日付文字列（例: "3/10"、"未定"、"3月中旬"）
    location: str = ""       # 場所文字列（例: "貴社"、"○○工場"）


# ============================================
# 送り状情報（VBA: ExtractTrackingInfoの返り値 Array(carrierName, trackingNum)）
# ============================================
@dataclass
class TrackingEntry:
    """送り状番号情報

    VBAではCollection of Array(運送会社正式名, 追跡番号)で管理。
    """

    carrier_name: str = ""    # 運送会社正式名（例: "佐川急便"、"ヤマト運輸"）
    tracking_number: str = "" # 追跡番号（半角数字、10桁以上）


# ============================================
# 欠品情報（メール本文・ご連絡事項用）
# ============================================
@dataclass
class StockoutEntry:
    """欠品商品の情報（メール本文用）"""

    manufacturer_name: str = ""   # メーカー名
    product_name: str = ""        # 品名
    quantity: str = ""            # 数量
    delivery: str = ""            # 計算済み納期（CopyDataRow戻り値相当）
    approx_delivery: str = ""     # 概算入荷日（例: "3月上旬入荷予定"）
    order_number: str = ""        # 注番（グループ化用）


# 欠品の入荷確定判定（email_builder / excel_writer 共用）
_UNCONFIRMED_DELIVERY = ("欠品中", "確認中", "日程調整中")


def is_stockout_confirmed(item: StockoutEntry) -> bool:
    """欠品の入荷日が確定しているか判定する。"""
    return bool(
        item.delivery
        and item.delivery not in _UNCONFIRMED_DELIVERY
        and "（欠品）" not in item.delivery
    )


# ============================================
# 確定/確認中の伝票（送付履歴・確認中一覧書き込み用）
# ============================================
@dataclass
class HistoryRecord:
    """送付履歴テーブルに書き込む1行分"""

    sent_datetime: datetime.datetime = field(
        default_factory=datetime.datetime.now
    )
    order_date: Optional[datetime.date] = None   # 受注日（登録日）
    customer_name: str = ""
    order_number: str = ""          # 受発注伝票
    detail_number: str = ""         # 明細
    manufacturer_name: str = ""
    product_name: str = ""
    delivery_answer: str = ""       # 納期回答
    sender: str = ""                # 送付者


@dataclass
class ConfirmingRecord:
    """確認中テーブルに書き込む1行分"""

    sent_datetime: datetime.datetime = field(
        default_factory=datetime.datetime.now
    )
    order_date: Optional[datetime.date] = None
    customer_name: str = ""
    order_number: str = ""
    detail_number: str = ""
    manufacturer_name: str = ""
    product_name: str = ""
    inquiry_status: str = "未"      # 問合せ状況（未/済/回答待ち/除外）
    status: str = ""                # ステータス（プログラム自動設定）
    order_delivery_date: Optional[datetime.date] = None  # 受注納期
    sender: str = ""


# ============================================
# 回答書生成結果（VBA: CreateDeliveryReportの返り値 Array(filePath, ...)）
# ============================================
@dataclass
class ReportResult:
    """CreateDeliveryReport / CreateDeliveryReportByOrderNumbersの結果"""

    file_path: str = ""                 # 保存先パス
    customer_name: str = ""             # 顧客名
    rep_name: str = ""                  # 担当者名（分割送信時）
    confirmed_orders: list[HistoryRecord] = field(default_factory=list)
    confirming_orders: list[ConfirmingRecord] = field(default_factory=list)
    tracking_info_list: list[tuple[str, TrackingEntry]] = field(
        default_factory=list
    )
    """(品名情報, TrackingEntry) のリスト"""

    stockout_info_list: list[StockoutEntry] = field(default_factory=list)
    bunno_info_list: list[dict] = field(default_factory=list)
    """分納情報のdictリスト（email_builder/excel_writerが参照するフル情報）"""

    bunno_completed_list: list[str] = field(default_factory=list)
    """分納完了した品名のリスト"""

    has_confirming: bool = False        # 確認中の伝票があるか

    is_twf: bool = False
    """TWF展示会専用回答書ならTrue（期間限定機能。twf.py参照）"""


# ============================================
# 列位置マッピング（VBA: GetColumnPositionsの返り値 Dictionary）
# ============================================
# Pythonではdict[str, int]をそのまま使う。型エイリアスのみ定義。
ColumnMap = dict[str, int]
"""SAP受注一覧の列名→列番号マッピング（VBA: cols Dictionary）"""


# ============================================
# 祝日辞書（VBA: Loadholidaysの返り値 Dictionary）
# ============================================
# キー: datetime.date, 値: Optional[int]（None=祝日、int=特別締切時間）
HolidayMap = dict[datetime.date, Optional[int]]
"""祝日・特別日カレンダー。値がNone=休日、int=特別締切時間"""
