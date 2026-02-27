#!/usr/bin/env python3
"""橋梁点検DXシステム提案書 - PPTX生成スクリプト (16:9, Google Slides対応)"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# ========== 定数 ==========
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# カラーパレット (HTML準拠)
PRIMARY = RGBColor(0x1A, 0x36, 0x5D)
PRIMARY_LIGHT = RGBColor(0x2B, 0x6C, 0xB0)
ACCENT = RGBColor(0xE5, 0x3E, 0x3E)
ACCENT_SOFT = RGBColor(0xFC, 0x81, 0x81)
SUCCESS = RGBColor(0x38, 0xA1, 0x69)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK = RGBColor(0x1A, 0x20, 0x2C)
TEXT_MUTED = RGBColor(0x4A, 0x55, 0x68)
TEXT_LIGHT = RGBColor(0x71, 0x80, 0x96)
BG_LIGHT = RGBColor(0xF7, 0xFA, 0xFC)
BORDER = RGBColor(0xE2, 0xE8, 0xF0)

# フォント (Google Slides互換)
FONT_JP = "Noto Sans JP"
FONT_EN = "Arial"


def set_slide_bg(slide, color):
    """スライド背景色を設定"""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_gradient_bg(slide, color1, color2):
    """グラデーション背景用の矩形を追加"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Emu(0), Emu(0), SLIDE_WIDTH, SLIDE_HEIGHT
    )
    shape.fill.gradient()
    shape.fill.gradient_stops[0].color.rgb = color1
    shape.fill.gradient_stops[0].position = 0.0
    shape.fill.gradient_stops[1].color.rgb = color2
    shape.fill.gradient_stops[1].position = 1.0
    shape.line.fill.background()
    shape.rotation = 0


def add_text_box(slide, left, top, width, height, text, font_size=14,
                 font_color=TEXT_DARK, bold=False, alignment=PP_ALIGN.LEFT,
                 font_name=FONT_JP, line_spacing=1.5):
    """テキストボックスを追加"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = font_color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    p.space_after = Pt(0)
    if line_spacing != 1.0:
        p.line_spacing = Pt(font_size * line_spacing)
    return txBox


def add_rounded_rect(slide, left, top, width, height, fill_color,
                     border_color=None, border_width=Pt(0)):
    """角丸四角形を追加"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = border_width
    else:
        shape.line.fill.background()
    return shape


def add_multiline_textbox(slide, left, top, width, height, lines,
                          default_size=14, default_color=TEXT_DARK,
                          alignment=PP_ALIGN.LEFT, line_spacing=1.5):
    """複数行テキストボックス (各行の書式を個別制御)"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line_info in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line_info.get("text", "")
        p.font.size = Pt(line_info.get("size", default_size))
        p.font.color.rgb = line_info.get("color", default_color)
        p.font.bold = line_info.get("bold", False)
        p.font.name = line_info.get("font", FONT_JP)
        p.alignment = line_info.get("align", alignment)
        sp = line_info.get("space_after", 0)
        p.space_after = Pt(sp)
        if line_spacing != 1.0:
            p.line_spacing = Pt(line_info.get("size", default_size) * line_spacing)
    return txBox


# =====================================================================
# スライド作成関数
# =====================================================================

def slide_cover(prs):
    """スライド1: 表紙"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_gradient_bg(slide, PRIMARY, PRIMARY_LIGHT)

    # 装飾: 薄い半透明風の円
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(8), Inches(-1), Inches(7), Inches(7)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = RGBColor(0x22, 0x44, 0x70)
    circle.line.fill.background()

    # PROPOSAL 2026 バッジ
    badge = add_rounded_rect(slide, Inches(4.8), Inches(1.2), Inches(3.7), Inches(0.5),
                             RGBColor(0x2B, 0x5E, 0x9E), WHITE, Pt(1))
    badge.text_frame.paragraphs[0].text = "PROPOSAL 2026"
    badge.text_frame.paragraphs[0].font.size = Pt(14)
    badge.text_frame.paragraphs[0].font.color.rgb = WHITE
    badge.text_frame.paragraphs[0].font.bold = True
    badge.text_frame.paragraphs[0].font.name = FONT_EN
    badge.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # メインタイトル
    add_text_box(slide, Inches(1.5), Inches(2.2), Inches(10.3), Inches(1.5),
                 "橋梁点検DXシステム", font_size=48, font_color=WHITE,
                 bold=True, alignment=PP_ALIGN.CENTER, line_spacing=1.0)

    # サブタイトル
    add_text_box(slide, Inches(2), Inches(3.6), Inches(9.3), Inches(0.8),
                 "AI画像解析 × IoTセンシングによる次世代インフラ点検",
                 font_size=20, font_color=RGBColor(0xBE, 0xD4, 0xED),
                 alignment=PP_ALIGN.CENTER, line_spacing=1.0)

    # 区切り線
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(5.7), Inches(4.6), Inches(2), Pt(3)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = ACCENT
    line.line.fill.background()

    # メタ情報
    add_multiline_textbox(slide, Inches(2), Inches(5.0), Inches(9.3), Inches(1.8),
                          [
                              {"text": "フレックスデザイン社 御中", "size": 22, "color": WHITE,
                               "bold": True, "align": PP_ALIGN.CENTER, "space_after": 12},
                              {"text": "2026年2月26日", "size": 14,
                               "color": RGBColor(0xBE, 0xD4, 0xED),
                               "align": PP_ALIGN.CENTER, "space_after": 4},
                              {"text": "ご提案書（機密）", "size": 12,
                               "color": RGBColor(0xBE, 0xD4, 0xED),
                               "align": PP_ALIGN.CENTER},
                          ], line_spacing=1.2)


def slide_executive_summary(prs):
    """スライド2: エグゼクティブサマリー"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_LIGHT)

    # セクション番号 + ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(2), Inches(0.5),
                 "01", font_size=48, font_color=RGBColor(0xED, 0xF2, 0xF7),
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "Executive Summary", font_size=11, font_color=PRIMARY_LIGHT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.7), Inches(6), Inches(0.6),
                 "エグゼクティブサマリー", font_size=28, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)

    # アンダーライン
    underline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.25), Inches(1.5), Pt(3)
    )
    underline.fill.solid()
    underline.fill.fore_color.rgb = ACCENT
    underline.line.fill.background()

    # 本文
    body_text = ("本提案は、フレックスデザイン社が管理・受託する橋梁点検業務において、"
                 "AI画像解析とIoTセンシング技術を活用した次世代点検DXシステムの導入をご提案するものです。\n"
                 "従来の目視点検に代わり、ドローン撮影画像のAI自動解析とリアルタイムセンサーモニタリングにより、"
                 "点検品質の向上・コスト削減・安全性の確保を同時に実現します。")
    add_text_box(slide, Inches(0.8), Inches(1.5), Inches(11.7), Inches(1.2),
                 body_text, font_size=13, font_color=TEXT_MUTED, line_spacing=1.6)

    # 数値KPI - 4つのカード
    stats = [
        ("40%", "点検コスト削減"),
        ("3倍", "点検速度向上"),
        ("95%", "損傷検出精度"),
        ("24/7", "常時モニタリング"),
    ]
    card_width = Inches(2.7)
    card_height = Inches(1.6)
    start_x = Inches(0.8)
    y = Inches(3.0)
    gap = Inches(0.25)

    for i, (value, label) in enumerate(stats):
        x = start_x + (card_width + gap) * i
        card = add_rounded_rect(slide, x, y, card_width, card_height, WHITE, BORDER, Pt(1))

        add_text_box(slide, x + Inches(0.3), y + Inches(0.25), card_width - Inches(0.6),
                     Inches(0.8), value, font_size=36, font_color=ACCENT,
                     bold=True, alignment=PP_ALIGN.CENTER, font_name=FONT_EN, line_spacing=1.0)
        add_text_box(slide, x + Inches(0.3), y + Inches(1.05), card_width - Inches(0.6),
                     Inches(0.4), label, font_size=12, font_color=TEXT_LIGHT,
                     alignment=PP_ALIGN.CENTER, line_spacing=1.0)

    # ハイライトボックス
    hl_bg = add_rounded_rect(slide, Inches(0.8), Inches(5.0), Inches(11.7), Inches(1.8),
                             RGBColor(0xEB, 0xF8, 0xFF))
    # 左の縦線
    vline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(5.0), Pt(4), Inches(1.8)
    )
    vline.fill.solid()
    vline.fill.fore_color.rgb = PRIMARY_LIGHT
    vline.line.fill.background()

    add_text_box(slide, Inches(1.2), Inches(5.15), Inches(10.8), Inches(0.4),
                 "伴走型パートナーシップ", font_size=15, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)
    add_text_box(slide, Inches(1.2), Inches(5.55), Inches(10.8), Inches(1.1),
                 "私たちは単なるシステム導入にとどまらず、フレックスデザイン社の業務フローに深く入り込み、"
                 "課題発見から運用定着まで「伴走型」で支援いたします。"
                 "貴社の成長戦略に沿った段階的な展開により、確実な成果創出をお約束します。",
                 font_size=12, font_color=TEXT_MUTED, line_spacing=1.6)


def slide_challenges(prs):
    """スライド3: 現状の課題認識"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, PRIMARY, PRIMARY_LIGHT)

    # ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "02  Current Challenges", font_size=11, font_color=ACCENT_SOFT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.65), Inches(8), Inches(0.6),
                 "現状の課題認識", font_size=28, font_color=WHITE,
                 bold=True, line_spacing=1.0)

    add_text_box(slide, Inches(0.8), Inches(1.3), Inches(10), Inches(0.7),
                 "国土交通省の定期点検義務化（2014年〜）により、5年に1度の近接目視点検が\n"
                 "全橋梁に求められています。しかし、従来手法には以下の構造的課題が存在します。",
                 font_size=12, font_color=RGBColor(0xBE, 0xD4, 0xED), line_spacing=1.5)

    # 6つのカード (3x2)
    challenges = [
        ("⏱", "時間・人材不足", "熟練点検員の高齢化と後継者不足。全国約73万橋の点検需要に対し、技術者が圧倒的に不足"),
        ("⚠", "安全リスク", "高所作業・足場設置に伴う労働災害リスク。河川・交通上のリスクも深刻"),
        ("📊", "品質のばらつき", "点検員の経験・スキルに依存した判定。評価者によって判定区分が異なるケースが散見"),
        ("💰", "コスト増大", "足場設置・交通規制等の付帯コストが点検費用の60%以上を占め、予算を圧迫"),
        ("📁", "データ管理の非効率", "紙ベースの調書作成・管理が主流。過去データとの比較分析が困難"),
        ("🔄", "予防保全への転換遅れ", "事後保全型から予防保全型への転換が求められる中、データ基盤が未整備"),
    ]

    card_w = Inches(3.8)
    card_h = Inches(1.6)
    start_x = Inches(0.8)
    start_y = Inches(2.3)
    gap_x = Inches(0.25)
    gap_y = Inches(0.2)

    for i, (icon, title, desc) in enumerate(challenges):
        col = i % 3
        row = i // 3
        x = start_x + (card_w + gap_x) * col
        y = start_y + (card_h + gap_y) * row

        card = add_rounded_rect(slide, x, y, card_w, card_h,
                                RGBColor(0x1F, 0x3D, 0x6A),
                                RGBColor(0x30, 0x55, 0x85), Pt(1))

        # アイコン
        add_text_box(slide, x + Inches(0.2), y + Inches(0.15), Inches(0.5), Inches(0.5),
                     icon, font_size=22, alignment=PP_ALIGN.CENTER, line_spacing=1.0)

        # タイトル
        add_text_box(slide, x + Inches(0.8), y + Inches(0.15), card_w - Inches(1), Inches(0.4),
                     title, font_size=15, font_color=WHITE, bold=True, line_spacing=1.0)

        # 説明
        add_text_box(slide, x + Inches(0.8), y + Inches(0.55), card_w - Inches(1), Inches(0.9),
                     desc, font_size=11, font_color=RGBColor(0xBE, 0xD4, 0xED), line_spacing=1.4)


def slide_solution(prs):
    """スライド4: ソリューション概要"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_LIGHT)

    # ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "03  Our Solution", font_size=11, font_color=PRIMARY_LIGHT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.65), Inches(8), Inches(0.6),
                 "ソリューション概要", font_size=28, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)
    underline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(1.5), Pt(3)
    )
    underline.fill.solid()
    underline.fill.fore_color.rgb = ACCENT
    underline.line.fill.background()

    add_text_box(slide, Inches(0.8), Inches(1.4), Inches(11), Inches(0.5),
                 "3つのコア技術を統合した橋梁点検DXプラットフォームにより、点検業務の全工程をデジタル化・自動化します。",
                 font_size=13, font_color=TEXT_MUTED, line_spacing=1.5)

    # 3つのカード
    solutions = [
        (PRIMARY, PRIMARY_LIGHT, "🤖", "AI画像解析エンジン",
         "深層学習（CNN/Vision Transformer）を活用した損傷自動検出。0.1mm精度で検出し、損傷種別・グレードを自動判定。",
         ["ひび割れ幅の自動計測（0.1mm単位）", "損傷種別の自動分類（13カテゴリ）",
          "経年変化の自動比較・進行度分析", "調書の自動生成（国交省様式準拠）"]),
        (ACCENT, ACCENT_SOFT, "📡", "IoTセンシング基盤",
         "振動・傾斜・温湿度・ひずみの各種センサーで構造物の健全性をリアルタイム常時監視。",
         ["MEMS加速度センサーによる振動モニタリング", "傾斜計による変位の長期監視",
          "LoRaWAN/LTE-M対応の低消費電力通信", "ソーラーパネル駆動による自立運用"]),
        (RGBColor(0x27, 0x67, 0x49), SUCCESS, "☁", "統合管理クラウド",
         "全橋梁の点検・センサー・画像データを一元管理。GISマップ上でのビジュアル管理。",
         ["GISベースの橋梁台帳・マップ管理", "リアルタイムダッシュボード",
          "劣化予測シミュレーション", "修繕優先度の自動ランキング"]),
    ]

    card_w = Inches(3.8)
    card_h = Inches(4.8)
    start_x = Inches(0.8)
    y = Inches(2.1)
    gap = Inches(0.25)

    for i, (color1, color2, icon, title, desc, features) in enumerate(solutions):
        x = start_x + (card_w + gap) * i

        # カード背景
        card = add_rounded_rect(slide, x, y, card_w, card_h, WHITE, BORDER, Pt(1))

        # 上部のカラーバー
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, x, y, card_w, Pt(4)
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = color1
        bar.line.fill.background()

        # アイコン背景
        icon_bg = add_rounded_rect(slide, x + Inches(0.3), y + Inches(0.3),
                                   Inches(0.6), Inches(0.6), color1)
        add_text_box(slide, x + Inches(0.3), y + Inches(0.32), Inches(0.6), Inches(0.6),
                     icon, font_size=22, alignment=PP_ALIGN.CENTER, font_color=WHITE,
                     line_spacing=1.0)

        # タイトル
        add_text_box(slide, x + Inches(0.3), y + Inches(1.05), card_w - Inches(0.6), Inches(0.4),
                     title, font_size=16, font_color=PRIMARY, bold=True, line_spacing=1.0)

        # 説明
        add_text_box(slide, x + Inches(0.3), y + Inches(1.5), card_w - Inches(0.6), Inches(1.0),
                     desc, font_size=11, font_color=TEXT_MUTED, line_spacing=1.5)

        # 機能リスト
        features_text = "\n".join([f"✓  {f}" for f in features])
        add_text_box(slide, x + Inches(0.3), y + Inches(2.6), card_w - Inches(0.6), Inches(2.0),
                     features_text, font_size=11, font_color=TEXT_MUTED, line_spacing=1.8)


def slide_architecture(prs):
    """スライド5: システム構成"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)

    # ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "04  System Architecture", font_size=11, font_color=PRIMARY_LIGHT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.65), Inches(8), Inches(0.6),
                 "システム構成", font_size=28, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)
    underline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(1.5), Pt(3)
    )
    underline.fill.solid()
    underline.fill.fore_color.rgb = ACCENT
    underline.line.fill.background()

    # 4レイヤーのアーキテクチャ図
    layers = [
        ("DATA COLLECTION", ["ドローン撮影", "IoTセンサー", "タブレット入力", "既存台帳データ"],
         PRIMARY, PRIMARY_LIGHT),
        ("EDGE PROCESSING", ["エッジAI推論", "画像前処理", "データ圧縮・送信"],
         ACCENT, ACCENT_SOFT),
        ("CLOUD PLATFORM", ["AI解析エンジン", "データベース", "GISエンジン", "API Gateway"],
         RGBColor(0x27, 0x67, 0x49), SUCCESS),
        ("APPLICATION", ["管理ダッシュボード", "現場タブレットアプリ", "調書自動生成", "外部連携API"],
         PRIMARY, PRIMARY_LIGHT),
    ]

    layer_h = Inches(0.9)
    layer_w = Inches(10.5)
    start_x = Inches(1.4)
    start_y = Inches(1.6)
    arrow_h = Inches(0.35)

    for i, (label, items, color1, color2) in enumerate(layers):
        y = start_y + (layer_h + arrow_h) * i

        # レイヤーの枠
        layer_rect = add_rounded_rect(slide, start_x, y, layer_w, layer_h,
                                      RGBColor(0xF7, 0xFA, 0xFC), BORDER, Pt(1.5))

        # レイヤーラベル
        add_text_box(slide, start_x + Inches(0.2), y - Inches(0.02), Inches(2.5), Inches(0.35),
                     label, font_size=10, font_color=PRIMARY_LIGHT, bold=True,
                     font_name=FONT_EN, line_spacing=1.0)

        # 各アイテム
        item_w = Inches(2.2)
        item_h = Inches(0.4)
        items_start_x = start_x + Inches(0.3)
        items_y = y + Inches(0.38)
        item_gap = Inches(0.2)

        for j, item_text in enumerate(items):
            ix = items_start_x + (item_w + item_gap) * j
            item_rect = add_rounded_rect(slide, ix, items_y, item_w, item_h, color1)
            tf = item_rect.text_frame
            tf.paragraphs[0].text = item_text
            tf.paragraphs[0].font.size = Pt(11)
            tf.paragraphs[0].font.color.rgb = WHITE
            tf.paragraphs[0].font.bold = True
            tf.paragraphs[0].font.name = FONT_JP
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        # 矢印 (最後のレイヤー以外)
        if i < len(layers) - 1:
            arrow_y = y + layer_h + Inches(0.05)
            add_text_box(slide, Inches(6.2), arrow_y, Inches(1), Inches(0.3),
                         "▼", font_size=18, font_color=TEXT_LIGHT,
                         alignment=PP_ALIGN.CENTER, line_spacing=1.0)

    # セキュリティボックス
    sec_y = start_y + (layer_h + arrow_h) * 4 + Inches(0.1)
    hl_bg = add_rounded_rect(slide, Inches(0.8), sec_y, Inches(11.7), Inches(1.0),
                             RGBColor(0xEB, 0xF8, 0xFF))
    vline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), sec_y, Pt(4), Inches(1.0)
    )
    vline.fill.solid()
    vline.fill.fore_color.rgb = PRIMARY_LIGHT
    vline.line.fill.background()

    add_text_box(slide, Inches(1.2), sec_y + Inches(0.1), Inches(10.8), Inches(0.3),
                 "セキュリティ・準拠規格", font_size=13, font_color=PRIMARY, bold=True, line_spacing=1.0)
    add_text_box(slide, Inches(1.2), sec_y + Inches(0.45), Inches(10.8), Inches(0.5),
                 "ISO 27001準拠 | ISMAP認定クラウド | AES-256暗号化 | 多要素認証 | 監査ログ完備",
                 font_size=11, font_color=TEXT_MUTED, line_spacing=1.3)


def slide_features(prs):
    """スライド6: 主要機能詳細"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_LIGHT)

    # ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "05  Key Features", font_size=11, font_color=PRIMARY_LIGHT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.65), Inches(8), Inches(0.6),
                 "主要機能詳細", font_size=28, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)
    underline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(1.5), Pt(3)
    )
    underline.fill.solid()
    underline.fill.fore_color.rgb = ACCENT
    underline.line.fill.background()

    steps = [
        ("01", "ドローン自動飛行・撮影",
         "3Dモデルに基づく飛行経路の自動生成。橋梁形状に最適化された撮影角度・距離をAIが自動計算。"
         "1橋あたり約30分で撮影完了。RTK-GPS対応によりセンチメートル精度の位置情報を付与。"),
        ("02", "AI自動解析・損傷検出",
         "撮影画像をクラウドへアップロード後、AI解析エンジンが自動で損傷箇所を検出。"
         "13種類の損傷を自動分類。信頼度スコアを付与し、閾値以下は人間のレビュー対象としてフラグ付け。"),
        ("03", "デジタル調書自動生成",
         "国土交通省の定期点検要領に準拠した調書を自動生成。損傷図・写真台帳・所見文・判定区分をAIが作成。"
         "調書作成工数を従来比80%削減。Excel/PDF出力対応。"),
        ("04", "劣化予測・修繕計画支援",
         "蓄積された点検データとセンサーデータをもとに、機械学習モデルが劣化進行を予測。"
         "LCC最適化の観点から修繕優先度を自動ランキングし、中長期修繕計画の策定を支援。"),
    ]

    step_w = Inches(11.7)
    step_h = Inches(1.2)
    start_x = Inches(0.8)
    start_y = Inches(1.6)
    gap = Inches(0.2)

    for i, (num, title, desc) in enumerate(steps):
        y = start_y + (step_h + gap) * i

        # ステップ番号円
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, start_x, y + Inches(0.1), Inches(0.8), Inches(0.8)
        )
        circle.fill.gradient()
        circle.fill.gradient_stops[0].color.rgb = PRIMARY
        circle.fill.gradient_stops[0].position = 0.0
        circle.fill.gradient_stops[1].color.rgb = PRIMARY_LIGHT
        circle.fill.gradient_stops[1].position = 1.0
        circle.line.fill.background()

        # STEP番号
        add_multiline_textbox(slide, start_x, y + Inches(0.15), Inches(0.8), Inches(0.7),
                              [
                                  {"text": "STEP", "size": 8, "color": RGBColor(0xBE, 0xD4, 0xED),
                                   "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_EN},
                                  {"text": num, "size": 18, "color": WHITE,
                                   "bold": True, "align": PP_ALIGN.CENTER, "font": FONT_EN},
                              ], line_spacing=1.0)

        # 縦の接続線 (最後以外)
        if i < len(steps) - 1:
            conn = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                start_x + Inches(0.38), y + Inches(0.9),
                Pt(2), Inches(0.5)
            )
            conn.fill.solid()
            conn.fill.fore_color.rgb = BORDER
            conn.line.fill.background()

        # タイトル
        add_text_box(slide, Inches(1.8), y + Inches(0.1), Inches(10), Inches(0.4),
                     title, font_size=17, font_color=PRIMARY, bold=True, line_spacing=1.0)

        # 説明
        add_text_box(slide, Inches(1.8), y + Inches(0.5), Inches(10.5), Inches(0.7),
                     desc, font_size=11, font_color=TEXT_MUTED, line_spacing=1.5)


def slide_comparison(prs):
    """スライド7: 従来手法との比較"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, PRIMARY, PRIMARY_LIGHT)

    # ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "06  Comparison", font_size=11, font_color=ACCENT_SOFT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.65), Inches(8), Inches(0.6),
                 "従来手法との比較", font_size=28, font_color=WHITE,
                 bold=True, line_spacing=1.0)

    # テーブル
    rows_data = [
        ("点検所要時間（1橋あたり）", "2〜3日", "0.5〜1日"),
        ("足場・規制の要否", "必要（大規模）", "原則不要"),
        ("損傷検出精度", "点検員依存（70〜90%）", "95%以上（一定品質）"),
        ("調書作成時間", "3〜5日/橋", "数時間（自動生成）"),
        ("データ蓄積・活用", "紙ベース（検索困難）", "クラウドDB（即時検索）"),
        ("劣化予測", "経験則に依存", "AI予測モデル"),
        ("常時監視", "不可", "IoTセンサーで24/7"),
        ("コスト（年間/橋）", "約150〜250万円", "約90〜150万円"),
    ]

    table_left = Inches(1.0)
    table_top = Inches(1.5)
    table_width = Inches(11.3)
    table_height = Inches(5.5)

    table = slide.shapes.add_table(
        len(rows_data) + 1, 3, table_left, table_top, table_width, table_height
    ).table

    # 列幅
    table.columns[0].width = Inches(4.0)
    table.columns[1].width = Inches(3.65)
    table.columns[2].width = Inches(3.65)

    # ヘッダ行
    headers = ["比較項目", "従来手法（目視点検）", "本システム（AI+IoT）"]
    header_colors = [PRIMARY, PRIMARY, ACCENT]
    for j, (header_text, hc) in enumerate(zip(headers, header_colors)):
        cell = table.cell(0, j)
        cell.text = header_text
        cell.fill.solid()
        cell.fill.fore_color.rgb = hc
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(13)
            p.font.color.rgb = WHITE
            p.font.bold = True
            p.font.name = FONT_JP
            p.alignment = PP_ALIGN.CENTER

    # データ行
    for i, (item, old_val, new_val) in enumerate(rows_data):
        row_idx = i + 1
        bg = WHITE if i % 2 == 0 else BG_LIGHT

        # 項目名
        cell0 = table.cell(row_idx, 0)
        cell0.text = item
        cell0.fill.solid()
        cell0.fill.fore_color.rgb = bg
        for p in cell0.text_frame.paragraphs:
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_DARK
            p.font.bold = True
            p.font.name = FONT_JP
            p.alignment = PP_ALIGN.LEFT

        # 従来手法
        cell1 = table.cell(row_idx, 1)
        cell1.text = old_val
        cell1.fill.solid()
        cell1.fill.fore_color.rgb = bg
        for p in cell1.text_frame.paragraphs:
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_MUTED
            p.font.name = FONT_JP
            p.alignment = PP_ALIGN.CENTER

        # 本システム
        cell2 = table.cell(row_idx, 2)
        cell2.text = new_val
        cell2.fill.solid()
        cell2.fill.fore_color.rgb = bg
        for p in cell2.text_frame.paragraphs:
            p.font.size = Pt(12)
            p.font.color.rgb = ACCENT
            p.font.bold = True
            p.font.name = FONT_JP
            p.alignment = PP_ALIGN.CENTER


def slide_schedule(prs):
    """スライド8: 導入スケジュール"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_LIGHT)

    # ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "07  Implementation Schedule", font_size=11, font_color=PRIMARY_LIGHT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.65), Inches(8), Inches(0.6),
                 "導入スケジュール", font_size=28, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)
    underline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(1.5), Pt(3)
    )
    underline.fill.solid()
    underline.fill.fore_color.rgb = ACCENT
    underline.line.fill.background()

    add_text_box(slide, Inches(0.8), Inches(1.4), Inches(11), Inches(0.5),
                 "PoC（概念実証）からフル導入まで、約12ヶ月の段階的導入計画をご提案します。",
                 font_size=13, font_color=TEXT_MUTED, line_spacing=1.5)

    # ガントチャート風
    chart_left = Inches(0.8)
    chart_top = Inches(2.1)
    label_w = Inches(2.5)
    month_w = Inches(1.65)
    row_h = Inches(0.65)

    months = ["M1-2", "M3-4", "M5-6", "M7-8", "M9-10", "M11-12"]

    # ヘッダ行
    # フェーズラベルヘッダ
    hdr = add_rounded_rect(slide, chart_left, chart_top, label_w, row_h, PRIMARY)
    tf = hdr.text_frame
    tf.paragraphs[0].text = "フェーズ"
    tf.paragraphs[0].font.size = Pt(12)
    tf.paragraphs[0].font.color.rgb = WHITE
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.name = FONT_JP
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    for j, m in enumerate(months):
        mx = chart_left + label_w + Inches(0.05) + (month_w + Inches(0.05)) * j
        mh = add_rounded_rect(slide, mx, chart_top, month_w, row_h, PRIMARY)
        tf = mh.text_frame
        tf.paragraphs[0].text = m
        tf.paragraphs[0].font.size = Pt(12)
        tf.paragraphs[0].font.color.rgb = WHITE
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.name = FONT_EN
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 各フェーズ行
    phases = [
        ("Phase 1: PoC", [(0, "PoC", PRIMARY, PRIMARY_LIGHT), (1, "評価", PRIMARY, PRIMARY_LIGHT)]),
        ("Phase 2: 開発", [(2, "設計", ACCENT, ACCENT_SOFT), (3, "開発", ACCENT, ACCENT_SOFT)]),
        ("Phase 3: パイロット", [(4, "試験", RGBColor(0x27, 0x67, 0x49), SUCCESS)]),
        ("Phase 4: 本番展開", [(5, "展開", RGBColor(0x27, 0x67, 0x49), SUCCESS)]),
    ]

    for i, (phase_name, bars) in enumerate(phases):
        y = chart_top + (row_h + Inches(0.05)) * (i + 1)

        # フェーズラベル
        lbl = add_rounded_rect(slide, chart_left, y, label_w, row_h, WHITE, BORDER, Pt(1))
        tf = lbl.text_frame
        tf.paragraphs[0].text = phase_name
        tf.paragraphs[0].font.size = Pt(11)
        tf.paragraphs[0].font.color.rgb = TEXT_DARK
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.name = FONT_JP
        tf.paragraphs[0].alignment = PP_ALIGN.LEFT
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        # 各月セル
        for j in range(6):
            mx = chart_left + label_w + Inches(0.05) + (month_w + Inches(0.05)) * j
            cell = add_rounded_rect(slide, mx, y, month_w, row_h, WHITE, BORDER, Pt(0.5))

        # バー
        for col, text, c1, c2 in bars:
            bx = chart_left + label_w + Inches(0.05) + (month_w + Inches(0.05)) * col + Inches(0.1)
            by = y + Inches(0.15)
            bar = add_rounded_rect(slide, bx, by, month_w - Inches(0.2), row_h - Inches(0.3), c1)
            tf = bar.text_frame
            tf.paragraphs[0].text = text
            tf.paragraphs[0].font.size = Pt(10)
            tf.paragraphs[0].font.color.rgb = WHITE
            tf.paragraphs[0].font.bold = True
            tf.paragraphs[0].font.name = FONT_JP
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # フェーズ説明カード (4つ)
    phase_descs = [
        ("Phase 1: PoC（2ヶ月）", "対象橋梁3〜5橋でドローン撮影・AI解析の精度検証。既存点検結果との比較評価。"),
        ("Phase 2: システム開発（4ヶ月）", "PoC結果を踏まえたカスタマイズ開発。管理画面・調書生成・IoT連携の構築。"),
        ("Phase 3: パイロット運用（2ヶ月）", "20〜30橋を対象とした実運用テスト。現場オペレーション確立・操作研修実施。"),
        ("Phase 4: 本番展開（2ヶ月）", "全対象橋梁への展開。運用マニュアル整備・保守サポート体制の確立。"),
    ]

    desc_w = Inches(2.75)
    desc_h = Inches(1.5)
    desc_y = Inches(5.6)
    desc_gap = Inches(0.2)

    for i, (title, desc) in enumerate(phase_descs):
        dx = Inches(0.8) + (desc_w + desc_gap) * i
        card = add_rounded_rect(slide, dx, desc_y, desc_w, desc_h, WHITE, BORDER, Pt(1))
        add_text_box(slide, dx + Inches(0.2), desc_y + Inches(0.15), desc_w - Inches(0.4), Inches(0.4),
                     title, font_size=12, font_color=PRIMARY, bold=True, line_spacing=1.0)
        add_text_box(slide, dx + Inches(0.2), desc_y + Inches(0.55), desc_w - Inches(0.4), Inches(0.85),
                     desc, font_size=10, font_color=TEXT_MUTED, line_spacing=1.5)


def slide_pricing(prs):
    """スライド9: 概算費用"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)

    # ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "08  Investment", font_size=11, font_color=PRIMARY_LIGHT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.65), Inches(8), Inches(0.6),
                 "概算費用", font_size=28, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)
    underline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(1.5), Pt(3)
    )
    underline.fill.solid()
    underline.fill.fore_color.rgb = ACCENT
    underline.line.fill.background()

    # 左カラム: 初期投資
    left_x = Inches(0.8)
    right_x = Inches(7.0)

    # 初期投資カード
    card = add_rounded_rect(slide, left_x, Inches(1.5), Inches(5.8), Inches(5.5),
                            WHITE, BORDER, Pt(1))

    # カードヘッダ
    card_hdr = add_rounded_rect(slide, left_x, Inches(1.5), Inches(5.8), Inches(1.2),
                                PRIMARY)
    add_text_box(slide, left_x + Inches(0.4), Inches(1.6), Inches(5), Inches(0.4),
                 "システム導入 概算総額", font_size=14, font_color=RGBColor(0xBE, 0xD4, 0xED),
                 line_spacing=1.0)
    add_text_box(slide, left_x + Inches(0.4), Inches(1.95), Inches(5), Inches(0.6),
                 "¥48,500,000（税別）", font_size=32, font_color=WHITE,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)

    # 内訳
    items = [
        ("Phase 1: PoC（概念実証）", "¥3,500,000"),
        ("Phase 2: システム設計・開発", "¥22,000,000"),
        ("Phase 3: パイロット運用・調整", "¥8,000,000"),
        ("Phase 4: 本番展開・研修", "¥9,000,000"),
        ("IoTセンサー機器一式（50橋分）", "¥6,000,000"),
    ]

    for i, (name, value) in enumerate(items):
        iy = Inches(2.95) + Inches(0.5) * i
        add_text_box(slide, left_x + Inches(0.4), iy, Inches(3.5), Inches(0.4),
                     name, font_size=12, font_color=TEXT_DARK, line_spacing=1.0)
        add_text_box(slide, left_x + Inches(3.8), iy, Inches(1.6), Inches(0.4),
                     value, font_size=12, font_color=PRIMARY, bold=True,
                     alignment=PP_ALIGN.RIGHT, font_name=FONT_EN, line_spacing=1.0)

        # 区切り線
        if i < len(items) - 1:
            sep = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left_x + Inches(0.4), iy + Inches(0.42),
                Inches(5), Pt(0.75)
            )
            sep.fill.solid()
            sep.fill.fore_color.rgb = BORDER
            sep.line.fill.background()

    # 右カラム: ランニングコスト
    rc_card = add_rounded_rect(slide, right_x, Inches(1.5), Inches(5.8), Inches(3.8),
                               WHITE, BORDER, Pt(1))

    add_text_box(slide, right_x + Inches(0.4), Inches(1.65), Inches(5), Inches(0.4),
                 "ランニングコスト（年額）", font_size=16, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)

    running_costs = [
        ("クラウド利用料", "¥3,600,000"),
        ("保守・サポート", "¥2,400,000"),
        ("AIモデル更新", "¥1,800,000"),
        ("IoT通信費", "¥600,000"),
    ]

    for i, (name, value) in enumerate(running_costs):
        ry = Inches(2.2) + Inches(0.45) * i
        add_text_box(slide, right_x + Inches(0.4), ry, Inches(3.2), Inches(0.35),
                     name, font_size=11, font_color=TEXT_DARK, line_spacing=1.0)
        add_text_box(slide, right_x + Inches(3.5), ry, Inches(1.8), Inches(0.35),
                     value, font_size=11, font_color=PRIMARY, bold=True,
                     alignment=PP_ALIGN.RIGHT, font_name=FONT_EN, line_spacing=1.0)

        sep = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            right_x + Inches(0.4), ry + Inches(0.38),
            Inches(5), Pt(0.75)
        )
        sep.fill.solid()
        sep.fill.fore_color.rgb = BORDER
        sep.line.fill.background()

    # 合計
    total_y = Inches(2.2) + Inches(0.45) * 4 + Inches(0.1)
    add_text_box(slide, right_x + Inches(0.4), total_y, Inches(3.2), Inches(0.4),
                 "合計", font_size=13, font_color=TEXT_DARK, bold=True, line_spacing=1.0)
    add_text_box(slide, right_x + Inches(3.5), total_y, Inches(1.8), Inches(0.4),
                 "¥8,400,000", font_size=16, font_color=ACCENT,
                 bold=True, alignment=PP_ALIGN.RIGHT, font_name=FONT_EN, line_spacing=1.0)

    # ROIボックス
    roi_y = Inches(5.6)
    hl_bg = add_rounded_rect(slide, right_x, roi_y, Inches(5.8), Inches(1.4),
                             RGBColor(0xEB, 0xF8, 0xFF))
    vline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, right_x, roi_y, Pt(4), Inches(1.4)
    )
    vline.fill.solid()
    vline.fill.fore_color.rgb = PRIMARY_LIGHT
    vline.line.fill.background()

    add_text_box(slide, right_x + Inches(0.3), roi_y + Inches(0.1), Inches(5.2), Inches(0.3),
                 "投資回収シミュレーション", font_size=13, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)
    add_text_box(slide, right_x + Inches(0.3), roi_y + Inches(0.45), Inches(5.2), Inches(0.85),
                 "50橋の年間点検費用を従来比40%削減 → 年間約3,000万円の削減効果。"
                 "初期投資は約1.5〜2年で回収可能。10年間で累計約2億円のコスト削減見込み。",
                 font_size=10, font_color=TEXT_MUTED, line_spacing=1.6)


def slide_team(prs):
    """スライド10: プロジェクト体制"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_LIGHT)

    # ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "09  Project Team", font_size=11, font_color=PRIMARY_LIGHT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.65), Inches(8), Inches(0.6),
                 "プロジェクト体制", font_size=28, font_color=PRIMARY,
                 bold=True, line_spacing=1.0)
    underline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(1.5), Pt(3)
    )
    underline.fill.solid()
    underline.fill.fore_color.rgb = ACCENT
    underline.line.fill.background()

    add_text_box(slide, Inches(0.8), Inches(1.4), Inches(11), Inches(0.5),
                 "橋梁工学・AI・IoT・クラウドの各領域の専門家によるクロスファンクショナルチームで推進します。",
                 font_size=13, font_color=TEXT_MUTED, line_spacing=1.5)

    # チームカード (4x2)
    team = [
        ("PM", "プロジェクトマネージャー", "全体統括・進捗管理"),
        ("AI", "AIエンジニアリード", "画像解析・モデル開発"),
        ("IoT", "IoTアーキテクト", "センサー設計・通信基盤"),
        ("BE", "バックエンドエンジニア", "API・データベース開発"),
        ("FE", "フロントエンドエンジニア", "UI/UX・ダッシュボード"),
        ("CE", "土木コンサルタント", "橋梁工学・点検要領"),
        ("DR", "ドローンオペレーター", "飛行計画・撮影実施"),
        ("QA", "品質管理・テスト", "検証・受入試験"),
    ]

    card_w = Inches(2.75)
    card_h = Inches(1.5)
    start_x = Inches(0.8)
    start_y = Inches(2.1)
    gap_x = Inches(0.2)
    gap_y = Inches(0.2)

    for i, (abbr, name, role) in enumerate(team):
        col = i % 4
        row = i // 4
        x = start_x + (card_w + gap_x) * col
        y = start_y + (card_h + gap_y) * row

        card = add_rounded_rect(slide, x, y, card_w, card_h, WHITE, BORDER, Pt(1))

        # アバター円
        avatar = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            x + Inches(0.15), y + Inches(0.3),
            Inches(0.65), Inches(0.65)
        )
        avatar.fill.gradient()
        avatar.fill.gradient_stops[0].color.rgb = PRIMARY
        avatar.fill.gradient_stops[0].position = 0.0
        avatar.fill.gradient_stops[1].color.rgb = PRIMARY_LIGHT
        avatar.fill.gradient_stops[1].position = 1.0
        avatar.line.fill.background()

        add_text_box(slide, x + Inches(0.15), y + Inches(0.42), Inches(0.65), Inches(0.45),
                     abbr, font_size=16, font_color=WHITE, bold=True,
                     alignment=PP_ALIGN.CENTER, font_name=FONT_EN, line_spacing=1.0)

        # 名前・役割
        add_text_box(slide, x + Inches(0.95), y + Inches(0.3), card_w - Inches(1.1), Inches(0.35),
                     name, font_size=13, font_color=PRIMARY, bold=True, line_spacing=1.0)
        add_text_box(slide, x + Inches(0.95), y + Inches(0.7), card_w - Inches(1.1), Inches(0.35),
                     role, font_size=10, font_color=TEXT_LIGHT, line_spacing=1.0)

    # 伴走型支援ボックス
    hl_y = Inches(5.6)
    hl_bg = add_rounded_rect(slide, Inches(0.8), hl_y, Inches(11.7), Inches(1.4),
                             RGBColor(0xEB, 0xF8, 0xFF))
    vline = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), hl_y, Pt(4), Inches(1.4)
    )
    vline.fill.solid()
    vline.fill.fore_color.rgb = PRIMARY_LIGHT
    vline.line.fill.background()

    add_text_box(slide, Inches(1.2), hl_y + Inches(0.1), Inches(10.8), Inches(0.3),
                 "伴走型支援体制", font_size=14, font_color=PRIMARY, bold=True, line_spacing=1.0)
    add_text_box(slide, Inches(1.2), hl_y + Inches(0.5), Inches(10.8), Inches(0.8),
                 "導入後も専任のカスタマーサクセスチームが週1回常駐し、運用定着・活用推進を支援します。"
                 "月次での効果測定レポート提出と改善提案を実施し、システムの継続的な価値向上に努めます。",
                 font_size=11, font_color=TEXT_MUTED, line_spacing=1.6)


def slide_next_steps(prs):
    """スライド11: 今後のステップ"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, PRIMARY, PRIMARY_LIGHT)

    # ヘッダ
    add_text_box(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 "10  Next Steps", font_size=11, font_color=ACCENT_SOFT,
                 bold=True, font_name=FONT_EN, line_spacing=1.0)
    add_text_box(slide, Inches(0.8), Inches(0.65), Inches(8), Inches(0.6),
                 "今後のステップ", font_size=28, font_color=WHITE,
                 bold=True, line_spacing=1.0)

    steps = [
        ("1", "ご質疑・ご要望ヒアリング",
         "本提案に対するご質問・ご要望をお聞かせください。\n"
         "貴社の具体的な課題やご要件に合わせて提案を調整いたします。"),
        ("2", "現地調査・PoC計画策定",
         "対象橋梁の現地調査を実施し、PoCの具体的な\n"
         "実施計画（対象橋梁・スケジュール・評価基準）を策定します。"),
        ("3", "PoC実施・効果検証",
         "3〜5橋を対象にPoCを実施。AI解析精度・工数削減効果・\n"
         "ユーザビリティを定量的に評価し、本導入判断の材料をご提供します。"),
    ]

    card_w = Inches(3.8)
    card_h = Inches(3.5)
    start_x = Inches(0.8)
    y = Inches(1.5)
    gap = Inches(0.25)

    for i, (num, title, desc) in enumerate(steps):
        x = start_x + (card_w + gap) * i

        card = add_rounded_rect(slide, x, y, card_w, card_h,
                                RGBColor(0x1F, 0x3D, 0x6A),
                                RGBColor(0x30, 0x55, 0x85), Pt(1))

        # 番号アイコン
        num_circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            x + Inches(0.3), y + Inches(0.3),
            Inches(0.7), Inches(0.7)
        )
        num_circle.fill.solid()
        num_circle.fill.fore_color.rgb = RGBColor(0x30, 0x55, 0x85)
        num_circle.line.fill.background()

        add_text_box(slide, x + Inches(0.3), y + Inches(0.38), Inches(0.7), Inches(0.55),
                     num, font_size=24, font_color=WHITE, bold=True,
                     alignment=PP_ALIGN.CENTER, font_name=FONT_EN, line_spacing=1.0)

        # タイトル
        add_text_box(slide, x + Inches(0.3), y + Inches(1.2), card_w - Inches(0.6), Inches(0.4),
                     title, font_size=17, font_color=WHITE, bold=True, line_spacing=1.0)

        # 説明
        add_text_box(slide, x + Inches(0.3), y + Inches(1.7), card_w - Inches(0.6), Inches(1.5),
                     desc, font_size=12, font_color=RGBColor(0xBE, 0xD4, 0xED), line_spacing=1.6)


def slide_closing(prs):
    """スライド12: クロージング"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_gradient_bg(slide, PRIMARY, PRIMARY_LIGHT)

    # 装飾
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(-2), Inches(3), Inches(6), Inches(6)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = RGBColor(0x22, 0x44, 0x70)
    circle.line.fill.background()

    circle2 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(10), Inches(-2), Inches(5), Inches(5)
    )
    circle2.fill.solid()
    circle2.fill.fore_color.rgb = RGBColor(0x22, 0x44, 0x70)
    circle2.line.fill.background()

    # メインメッセージ
    add_text_box(slide, Inches(2), Inches(2.2), Inches(9.3), Inches(1.0),
                 "橋梁点検の未来を、共に創りませんか",
                 font_size=32, font_color=WHITE, bold=True,
                 alignment=PP_ALIGN.CENTER, line_spacing=1.0)

    add_text_box(slide, Inches(3), Inches(3.3), Inches(7.3), Inches(1.0),
                 "フレックスデザイン社の技術力と私たちのAI・IoTソリューションを掛け合わせ、\n"
                 "インフラ維持管理の新たなスタンダードを確立しましょう。",
                 font_size=15, font_color=RGBColor(0xBE, 0xD4, 0xED),
                 alignment=PP_ALIGN.CENTER, line_spacing=1.6)

    # CTAボタン
    cta = add_rounded_rect(slide, Inches(4.7), Inches(4.7), Inches(4), Inches(0.7), ACCENT)
    tf = cta.text_frame
    tf.paragraphs[0].text = "お問い合わせ・ご相談"
    tf.paragraphs[0].font.size = Pt(15)
    tf.paragraphs[0].font.color.rgb = WHITE
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.name = FONT_JP
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # フッター
    add_text_box(slide, Inches(2), Inches(6.2), Inches(9.3), Inches(0.5),
                 "本資料はフレックスデザイン社向けに作成された機密文書です。無断複製・転載を禁じます。",
                 font_size=10, font_color=RGBColor(0x8A, 0xA5, 0xC2),
                 alignment=PP_ALIGN.CENTER, line_spacing=1.0)
    add_text_box(slide, Inches(2), Inches(6.6), Inches(9.3), Inches(0.4),
                 "© 2026 All Rights Reserved. | Confidential",
                 font_size=9, font_color=RGBColor(0x8A, 0xA5, 0xC2),
                 alignment=PP_ALIGN.CENTER, font_name=FONT_EN, line_spacing=1.0)


# =====================================================================
# メイン実行
# =====================================================================

def main():
    prs = Presentation()

    # 16:9 スライドサイズ
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # 各スライドを生成
    slide_cover(prs)           # 1. 表紙
    slide_executive_summary(prs)  # 2. エグゼクティブサマリー
    slide_challenges(prs)      # 3. 課題認識
    slide_solution(prs)        # 4. ソリューション概要
    slide_architecture(prs)    # 5. システム構成
    slide_features(prs)        # 6. 主要機能
    slide_comparison(prs)      # 7. 比較
    slide_schedule(prs)        # 8. スケジュール
    slide_pricing(prs)         # 9. 概算費用
    slide_team(prs)            # 10. プロジェクト体制
    slide_next_steps(prs)      # 11. 今後のステップ
    slide_closing(prs)         # 12. クロージング

    output_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "proposal_bridge_inspection_dx.pptx"
    )
    prs.save(output_path)
    print(f"PPTX saved: {output_path}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
