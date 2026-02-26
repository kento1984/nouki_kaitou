"""顧客マスター参照モジュール

VBAの以下の関数を移植:
- GetCustomerDeliveryDays (L5703): 顧客の配送曜日取得
- ConvertDayNameToNumber (L5771): 曜日名→VBA Weekday番号変換
- GetRetentionDays (L7460): 顧客の保持日数取得
- IsRouteDelivery (L7495): 路線便フラグ判定
- GetEmailAddresses (L5668): 顧客メールアドレス取得
- CheckCustomerMaster (L1350): メールアドレス未登録チェック
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from nouki_kaitou.models import CacheStore

if TYPE_CHECKING:
    import openpyxl.worksheet.worksheet as ws_type


# ============================================
# VBA: ConvertDayNameToNumber (L5771-5786)
# 曜日名をVBA Weekday番号に変換
# ============================================
# VBA Weekday番号: 日=1, 月=2, 火=3, 水=4, 木=5, 金=6, 土=7
_DAY_NAME_MAP: dict[str, int] = {
    "月": 2, "月曜": 2, "月曜日": 2,
    "火": 3, "火曜": 3, "火曜日": 3,
    "水": 4, "水曜": 4, "水曜日": 4,
    "木": 5, "木曜": 5, "木曜日": 5,
    "金": 6, "金曜": 6, "金曜日": 6,
    "土": 7, "土曜": 7, "土曜日": 7,
    "日": 1, "日曜": 1, "日曜日": 1,
}


def convert_day_name_to_number(day_name: str) -> int:
    """曜日名をVBA Weekday番号に変換する。

    Args:
        day_name: 曜日名（"月", "火曜", "水曜日" 等）

    Returns:
        VBA Weekday番号（日=1..土=7）。無効なら0。
    """
    return _DAY_NAME_MAP.get(day_name.strip(), 0)


# ============================================
# VBA: GetCustomerDeliveryDays (L5703-5765)
# 顧客の配送曜日を取得
# ============================================
def get_customer_delivery_days(
    customer_name: str,
    cache: CacheStore,
) -> list[int]:
    """顧客の配送曜日リストを取得する。

    キャッシュに格納済みのパース結果を返す。
    空リスト = 曜日制限なし。

    Args:
        customer_name: 顧客名
        cache: キャッシュストア

    Returns:
        VBA Weekday番号のリスト（例: [2, 4, 6] = 月水金）
    """
    return cache.cust_days.get(customer_name, [])


# ============================================
# VBA: GetRetentionDays (L7460-7491)
# 顧客の保持日数（路線便の配達日数）を取得
# ============================================
def get_retention_days(
    customer_name: str,
    cache: CacheStore,
) -> int:
    """顧客の保持日数を取得する。

    路線便で出荷から到着までの日数。
    0の場合は路線便でない（自社便ルート配達）。

    Args:
        customer_name: 顧客名
        cache: キャッシュストア

    Returns:
        保持日数（デフォルト0）
    """
    return cache.cust_retention.get(customer_name, 0)


# ============================================
# VBA: IsRouteDelivery (L7495-7522)
# 路線便フラグ判定
# ============================================
def get_customer_pattern(
    customer_name: str,
    cache: CacheStore,
) -> str:
    """顧客の配送パターン名を取得する。

    Args:
        customer_name: 顧客名
        cache: キャッシュストア

    Returns:
        パターン名（未設定なら空文字）
    """
    return cache.cust_pattern.get(customer_name, "")


def is_route_delivery(
    customer_name: str,
    cache: CacheStore,
) -> bool:
    """顧客が路線便配達かどうかを判定する。

    顧客マスターのD列に値がある = 路線便（True）。
    False = 自社便ルート配達。

    Args:
        customer_name: 顧客名
        cache: キャッシュストア

    Returns:
        True = 路線便
    """
    return cache.cust_route.get(customer_name, False)


# ============================================
# VBA: GetEmailAddresses (L5668-5698)
# 顧客メールアドレス取得
# ============================================
def get_email_addresses(
    customer_name: str,
    customer_master_ws: object,
    email_start_col: int = 4,
) -> str:
    """顧客マスターからメールアドレスを取得する。

    顧客マスターのメール開始列以降にあるメールアドレスを「; 」区切りで返す。

    Args:
        customer_name: 顧客名
        customer_master_ws: 顧客マスターシート
        email_start_col: メール開始列（0-indexed。旧フォーマット:4, 新フォーマット:5）

    Returns:
        メールアドレス（「; 」区切り）。見つからなければ空文字。
    """
    if customer_master_ws is None:
        return ""

    for row in customer_master_ws.iter_rows(min_row=2, values_only=True):
        name = str(row[0]).strip() if row[0] else ""
        if name != customer_name:
            continue

        emails: list[str] = []
        for j in range(email_start_col, len(row)):
            addr = str(row[j]).strip() if row[j] else ""
            if addr:
                emails.append(addr)

        return "; ".join(emails)

    return ""


# ============================================
# VBA: CheckCustomerMaster (L1350-1391)
# メールアドレス未登録チェック
# ============================================
def check_customer_master(
    customer_names: list[str],
    customer_master_ws: object,
    email_start_col: int = 4,
) -> str:
    """顧客のメールアドレス未登録をチェックする。

    メールアドレスが1つも登録されていない顧客のリストを返す。

    Args:
        customer_names: チェック対象の顧客名リスト
        customer_master_ws: 顧客マスターシート
        email_start_col: メール開始列（0-indexed。旧フォーマット:4, 新フォーマット:5）

    Returns:
        未登録顧客のリスト文字列（「・顧客名」形式、改行区切り）。
        全て登録済みなら空文字。
    """
    if customer_master_ws is None:
        # シートがなければ全て未登録とみなす
        return "\n".join(f"・{name}" for name in customer_names)

    # シートデータをキャッシュ化（複数顧客を効率よくチェック）
    master_data: dict[str, bool] = {}  # 顧客名 → メールあり
    for row in customer_master_ws.iter_rows(min_row=2, values_only=True):
        name = str(row[0]).strip() if row[0] else ""
        if not name:
            continue

        has_email = False
        for j in range(email_start_col, len(row)):
            if row[j] and str(row[j]).strip():
                has_email = True
                break

        if name not in master_data:
            master_data[name] = has_email

    missing_list: list[str] = []
    for cust_name in customer_names:
        if not master_data.get(cust_name, False):
            missing_list.append(f"・{cust_name}")

    return "\n".join(missing_list)
