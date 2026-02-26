"""tracking.py のユニットテスト

送り状番号処理の全4関数を網羅的にテスト。
"""

import pytest

from nouki_kaitou.tracking import (
    can_direct_track,
    clean_external_comment,
    extract_tracking_info,
    get_carrier_full_name,
    get_tracking_url,
)


# ============================================
# GetCarrierFullName
# ============================================
class TestGetCarrierFullName:
    @pytest.mark.parametrize("short,expected", [
        ("ヤマト", "ヤマト運輸"),
        ("クロネコ", "ヤマト運輸"),
        ("佐川", "佐川急便"),
        ("西濃", "西濃運輸"),
        ("福山", "福山通運"),
        ("郵便", "日本郵便"),
        ("ゆうパック", "日本郵便"),
        ("日通", "日本通運"),
        ("トナミ", "トナミ運輸"),
        ("セイノー", "セイノースーパーエクスプレス"),
        ("SSX", "セイノースーパーエクスプレス"),
        ("JPロジ", "JPロジスティクス"),
        ("ＪＰロジ", "JPロジスティクス"),
        ("第一貨物", "第一貨物"),
        ("第一", "第一貨物"),
    ])
    def test_known_carriers(self, short, expected):
        assert get_carrier_full_name(short) == expected

    def test_unknown_carrier(self):
        assert get_carrier_full_name("不明な運送会社") == "不明な運送会社"


# ============================================
# ExtractTrackingInfo
# ============================================
class TestExtractTrackingInfo:
    def test_yamato(self):
        result = extract_tracking_info("ヤマト:1234567890123")
        assert len(result) == 1
        assert result[0].carrier_name == "ヤマト運輸"
        assert result[0].tracking_number == "1234567890123"

    def test_sagawa(self):
        result = extract_tracking_info("佐川 9876543210")
        assert len(result) == 1
        assert result[0].carrier_name == "佐川急便"
        assert result[0].tracking_number == "9876543210"

    def test_kuroneko(self):
        result = extract_tracking_info("クロネコ：1111222233334")
        assert len(result) == 1
        assert result[0].carrier_name == "ヤマト運輸"
        assert result[0].tracking_number == "1111222233334"

    def test_multiple_carriers(self):
        """複数の運送会社が含まれる場合"""
        result = extract_tracking_info(
            "ヤマト:1234567890123 佐川:9876543210"
        )
        assert len(result) == 2

    def test_hyphen_skipped(self):
        """ハイフンは無視"""
        result = extract_tracking_info("ヤマト:1234-5678-90123")
        assert len(result) == 1
        assert result[0].tracking_number == "1234567890123"

    def test_fullwidth_hyphen_skipped(self):
        """全角ハイフンも無視"""
        result = extract_tracking_info("佐川：1234－5678－90")
        assert len(result) == 1
        assert result[0].tracking_number == "1234567890"

    def test_fullwidth_numbers(self):
        """全角数字は半角に変換"""
        result = extract_tracking_info("ヤマト：１２３４５６７８９０１２３")
        assert len(result) == 1
        assert result[0].tracking_number == "1234567890123"

    def test_too_short_number(self):
        """10桁未満は無視"""
        result = extract_tracking_info("ヤマト:123456789")
        assert len(result) == 0

    def test_no_carrier(self):
        result = extract_tracking_info("送り状番号なし")
        assert len(result) == 0

    def test_empty(self):
        result = extract_tracking_info("")
        assert len(result) == 0

    def test_colon_variations(self):
        """コロンの半角・全角バリエーション"""
        for sep in [":", "：", " ", "　"]:
            result = extract_tracking_info(f"佐川{sep}1234567890")
            assert len(result) == 1, f"separator='{sep}' failed"

    def test_fullwidth_space_separator(self):
        """全角スペース区切り"""
        result = extract_tracking_info("ヤマト　1234567890123")
        assert len(result) == 1
        assert result[0].tracking_number == "1234567890123"

    def test_seino(self):
        result = extract_tracking_info("西濃:1234567890")
        assert len(result) == 1
        assert result[0].carrier_name == "西濃運輸"

    def test_fukuyama(self):
        result = extract_tracking_info("福山:1234567890")
        assert len(result) == 1
        assert result[0].carrier_name == "福山通運"

    def test_yubin(self):
        result = extract_tracking_info("郵便:1234567890123")
        assert len(result) == 1
        assert result[0].carrier_name == "日本郵便"

    def test_yuu_pack(self):
        result = extract_tracking_info("ゆうパック:1234567890123")
        assert len(result) == 1
        assert result[0].carrier_name == "日本郵便"

    def test_nittsu(self):
        result = extract_tracking_info("日通:1234567890")
        assert len(result) == 1
        assert result[0].carrier_name == "日本通運"

    def test_tonami(self):
        result = extract_tracking_info("トナミ:1234567890")
        assert len(result) == 1
        assert result[0].carrier_name == "トナミ運輸"

    def test_seino_super(self):
        result = extract_tracking_info("セイノー:1234567890")
        assert len(result) == 1
        assert result[0].carrier_name == "セイノースーパーエクスプレス"

    def test_ssx(self):
        result = extract_tracking_info("SSX:1234567890")
        assert len(result) == 1
        assert result[0].carrier_name == "セイノースーパーエクスプレス"

    def test_jp_logi(self):
        result = extract_tracking_info("JPロジ:1234567890")
        assert len(result) == 1
        assert result[0].carrier_name == "JPロジスティクス"

    def test_daiichi(self):
        result = extract_tracking_info("第一貨物:1234567890")
        assert len(result) == 1
        assert result[0].carrier_name == "第一貨物"


# ============================================
# GetTrackingUrl
# ============================================
class TestGetTrackingUrl:
    def test_yamato(self):
        url = get_tracking_url("ヤマト運輸", "1234567890123")
        assert url == "https://member.kms.kuronekoyamato.co.jp/parcel/detail?pno=1234567890123"

    def test_sagawa(self):
        url = get_tracking_url("佐川急便", "9876543210")
        assert url == "https://k2k.sagawa-exp.co.jp/p/web/okurijosearch.do?okurijoNo=9876543210"

    def test_seino(self):
        url = get_tracking_url("西濃運輸", "1234567890")
        assert url == "https://track.seino.co.jp/cgi-bin/gnpquery.pgm?GNPNO1=1234567890"

    def test_fukuyama(self):
        url = get_tracking_url("福山通運", "1234567890")
        assert url == "https://corp.fukutsu.co.jp/situation/tracking_no_hunt/1234567890"

    def test_yubin(self):
        url = get_tracking_url("日本郵便", "1234567890123")
        assert "trackings.post.japanpost.jp" in url
        assert "1234567890123" in url

    def test_nittsu(self):
        url = get_tracking_url("日本通運", "1234567890")
        assert "lp-trace.nittsu.co.jp" in url
        assert "1234567890" in url

    def test_tonami(self):
        url = get_tracking_url("トナミ運輸", "1234567890")
        assert "tonami.co.jp" in url

    def test_seino_super(self):
        url = get_tracking_url("セイノースーパーエクスプレス", "1234567890")
        assert "inquire.trc.ssx.seino.co.jp" in url

    def test_jp_logi(self):
        url = get_tracking_url("JPロジスティクス", "1234567890")
        assert "jp-logistics.jp" in url

    def test_daiichi(self):
        url = get_tracking_url("第一貨物", "1234567890")
        assert "daiichi-kamotsu.co.jp" in url

    def test_unknown(self):
        url = get_tracking_url("不明な運送会社", "1234567890")
        assert url == ""

    def test_strip_hyphens(self):
        """ハイフンが除去されること"""
        url = get_tracking_url("ヤマト運輸", "1234-5678-90123")
        assert "1234567890123" in url


# ============================================
# CanDirectTrack
# ============================================
class TestCanDirectTrack:
    @pytest.mark.parametrize("carrier,expected", [
        ("ヤマト運輸", True),
        ("佐川急便", True),
        ("西濃運輸", True),
        ("福山通運", True),
        ("日本郵便", True),
        ("日本通運", True),
        ("トナミ運輸", False),
        ("セイノースーパーエクスプレス", False),
        ("JPロジスティクス", False),
        ("第一貨物", False),
        ("不明な運送会社", False),
    ])
    def test_direct_track(self, carrier, expected):
        assert can_direct_track(carrier) is expected

    def test_seino_not_super(self):
        """西濃運輸はTrue、セイノースーパーはFalse"""
        assert can_direct_track("西濃運輸") is True
        assert can_direct_track("セイノースーパーエクスプレス") is False


# ============================================
# CleanExternalComment
# ============================================
class TestCleanExternalComment:
    """社外コメントクリーニングテスト"""

    def test_tracking_only(self):
        """送り状番号のみ → 空文字"""
        assert clean_external_comment("佐川:1234567890") == ""

    def test_tracking_with_extra_text(self):
        """送り状番号 + 追加情報 → 追加情報のみ残る"""
        result = clean_external_comment("佐川:1234567890 納品書同封")
        assert result == "納品書同封"

    def test_pickup_only(self):
        """引取テキストのみ → 空文字"""
        assert clean_external_comment("引取") == ""

    def test_pickup_with_date(self):
        """引取+日付 → 空文字"""
        assert clean_external_comment("2/20 引取") == ""
        assert clean_external_comment("引取 2/20") == ""

    def test_pickup_hikitori(self):
        """引き取り → 空文字"""
        assert clean_external_comment("引き取り") == ""

    def test_no_tracking_no_pickup(self):
        """送り状なし・引取なし → そのまま"""
        assert clean_external_comment("納品書同封") == "納品書同封"

    def test_empty(self):
        """空文字 → 空文字"""
        assert clean_external_comment("") == ""

    def test_multiple_tracking(self):
        """複数の送り状番号 → すべて除去"""
        result = clean_external_comment("ヤマト:1234567890123 佐川:9876543210")
        assert result == ""

    def test_tracking_and_extra(self):
        """送り状+引取+追加情報 → 追加情報のみ"""
        result = clean_external_comment("ヤマト:1234567890123 引取 特記事項あり")
        assert result == "特記事項あり"

    def test_fullwidth_colon(self):
        """全角コロン区切りの送り状番号も除去"""
        assert clean_external_comment("佐川：1234567890") == ""

    def test_yamato_with_message(self):
        """ヤマト送り状+メッセージ"""
        result = clean_external_comment("ヤマト 1234567890123 2/25着指定")
        assert result == "2/25着指定"
