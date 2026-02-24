"""送り状番号処理モジュール

VBAの以下の関数を移植:
- ExtractTrackingInfo (L6373): 送り状番号抽出
- GetCarrierFullName (L6477): 運送会社略称→正式名変換
- GetTrackingUrl (L6577): 追跡URL生成
- CanDirectTrack (L6611): 直接追跡可能か判定
"""

from __future__ import annotations

from nouki_kaitou.models import TrackingEntry
from nouki_kaitou.utils import is_numeric_char, to_half_width_num


# 対応する運送会社キーワード（VBA carriersに対応）
_CARRIER_KEYWORDS = [
    "ヤマト", "クロネコ", "佐川", "西濃", "福山", "福通",
    "郵便", "ゆうパック", "日通", "トナミ",
    "セイノー", "エクスプレス", "SSX", "JPロジ", "ＪＰロジ",
    "第一貨物", "第一",
    "近物", "新潟", "名鉄",
]


# ============================================
# VBA: GetCarrierFullName (L6477-6502)
# 運送会社略称を正式名称に変換
# ============================================
_CARRIER_FULL_NAMES: dict[str, str] = {
    "ヤマト": "ヤマト運輸",
    "クロネコ": "ヤマト運輸",
    "佐川": "佐川急便",
    "西濃": "西濃運輸",
    "福山": "福山通運",
    "福通": "福山通運",
    "郵便": "日本郵便",
    "ゆうパック": "日本郵便",
    "日通": "日本通運",
    "トナミ": "トナミ運輸",
    "セイノー": "セイノースーパーエクスプレス",
    "エクスプレス": "セイノースーパーエクスプレス",
    "SSX": "セイノースーパーエクスプレス",
    "JPロジ": "JPロジスティクス",
    "ＪＰロジ": "JPロジスティクス",
    "第一貨物": "第一貨物",
    "第一": "第一貨物",
    "近物": "近物レックス",
    "新潟": "新潟運輸",
    "名鉄": "名鉄運輸",
}


def get_carrier_full_name(short_name: str) -> str:
    """運送会社略称を正式名称に変換する。

    Args:
        short_name: 略称（例: "ヤマト", "佐川"）

    Returns:
        正式名称（例: "ヤマト運輸", "佐川急便"）。
        マッチしなければ略称をそのまま返す。
    """
    return _CARRIER_FULL_NAMES.get(short_name, short_name)


# ============================================
# VBA: ExtractTrackingInfo (L6373-6450)
# 送り状番号をコメントから抽出
# ============================================
def extract_tracking_info(comment: str) -> list[TrackingEntry]:
    """コメントから送り状番号を抽出する。

    対応運送会社のキーワードを検索し、その後ろの10桁以上の数字を
    送り状番号として抽出。ハイフンはスキップ。

    Args:
        comment: コメント文字列（コメント（社外）等）

    Returns:
        TrackingEntryのリスト
    """
    results: list[TrackingEntry] = []

    if not comment:
        return results

    # 同じ位置で見つかったキャリアの重複を防ぐ
    found_positions: set[int] = set()

    for carrier_keyword in _CARRIER_KEYWORDS:
        start_pos = 0

        while start_pos < len(comment):
            pos = comment.find(carrier_keyword, start_pos)
            if pos < 0:
                break

            # 同じ位置で既に見つけた運送会社があればスキップ
            if pos in found_positions:
                start_pos = pos + 1
                continue

            # 運送会社名の後ろを取得
            after_carrier = comment[pos + len(carrier_keyword):]

            # 先頭のコロン（半角・全角）やスペースをスキップ
            while after_carrier:
                c = after_carrier[0]
                if c in (":", "：", " ", "\u3000"):
                    after_carrier = after_carrier[1:]
                else:
                    break

            # 数字部分を抽出（半角・全角両対応、ハイフンスキップ）
            tracking_num = ""
            for c in after_carrier:
                if is_numeric_char(c):
                    tracking_num += to_half_width_num(c)
                elif c in ("-", "－"):
                    # ハイフンはスキップ
                    pass
                elif tracking_num:
                    break

            # 10桁以上なら送り状番号として追加
            if len(tracking_num) >= 10:
                results.append(TrackingEntry(
                    carrier_name=get_carrier_full_name(carrier_keyword),
                    tracking_number=tracking_num,
                ))
                found_positions.add(pos)

            start_pos = pos + len(carrier_keyword)

    return results


# ============================================
# VBA: GetTrackingUrl (L6577-6607)
# 運送会社の追跡URLを生成
# ============================================
def get_tracking_url(carrier_name: str, tracking_num: str) -> str:
    """運送会社の追跡URLを生成する。

    Args:
        carrier_name: 運送会社正式名称
        tracking_num: 追跡番号

    Returns:
        追跡URL。対応なしの場合は空文字。
    """
    # ハイフン・スペースを除去（全運送会社共通）
    num = tracking_num.replace("-", "").replace(" ", "").replace("\u3000", "")

    if "ヤマト" in carrier_name:
        return f"https://member.kms.kuronekoyamato.co.jp/parcel/detail?pno={num}"

    if "佐川" in carrier_name:
        return f"https://k2k.sagawa-exp.co.jp/p/web/okurijosearch.do?okurijoNo={num}"

    if "西濃" in carrier_name and "スーパー" not in carrier_name:
        # 西濃運輸（セイノースーパーエクスプレスではない）
        return f"https://track.seino.co.jp/cgi-bin/gnpquery.pgm?GNPNO1={num}"

    if "福山" in carrier_name or "福通" in carrier_name:
        return f"https://corp.fukutsu.co.jp/situation/tracking_no_hunt/{num}"

    if "郵便" in carrier_name or "ゆうパック" in carrier_name:
        return f"https://trackings.post.japanpost.jp/services/srv/search/?requestNo1={num}"

    if "日通" in carrier_name or "日本通運" in carrier_name:
        return (
            f"https://lp-trace.nittsu.co.jp/web/webarpaa702.srv"
            f"?LANG=JP&officeselect2=&denpyoNo1={num}"
        )

    if "トナミ" in carrier_name:
        return "https://trc1.tonami.co.jp/trc/search3/excSearch3"

    if "セイノー" in carrier_name or "SSX" in carrier_name or "スーパー" in carrier_name:
        return "http://inquire.trc.ssx.seino.co.jp/"

    if "JPロジ" in carrier_name or "ＪＰロジ" in carrier_name:
        return "https://www.jp-logistics.jp/fwexphp/inquiry/chase/init"

    if "第一貨物" in carrier_name:
        return "https://www.daiichi-kamotsu.co.jp/chase/contact_num/"

    return ""


# ============================================
# VBA: CanDirectTrack (L6611-6627)
# 直接リンク可能な運送会社かどうか判定
# ============================================
def can_direct_track(carrier_name: str) -> bool:
    """直接リンク（追跡番号入りURL）で追跡可能かどうか判定する。

    トナミ・セイノースーパーエクスプレス・JPロジ・第一貨物は
    番号入りURLを生成できないのでFalse。

    Args:
        carrier_name: 運送会社正式名称

    Returns:
        True = 追跡番号入りURLで直接追跡可能
    """
    if "ヤマト" in carrier_name:
        return True
    if "佐川" in carrier_name:
        return True
    if "西濃" in carrier_name and "スーパー" not in carrier_name:
        return True
    if "福山" in carrier_name or "福通" in carrier_name:
        return True
    if "郵便" in carrier_name or "ゆうパック" in carrier_name:
        return True
    if "日通" in carrier_name or "日本通運" in carrier_name:
        return True
    return False
