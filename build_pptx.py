# -*- coding: utf-8 -*-
"""受動部品の種類と製作 — 実験レポート用プレゼン生成スクリプト
結果・考察中心。実測値は【要記入】プレースホルダで枠だけ確保している。
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

FONT = "Meiryo"  # 日本語フォント（PowerPoint側で適用）

# カラーパレット
NAVY   = RGBColor(0x1F, 0x35, 0x5E)
BLUE   = RGBColor(0x2E, 0x5B, 0xA8)
RED    = RGBColor(0xC0, 0x39, 0x2B)   # 抵抗
TEAL   = RGBColor(0x1F, 0x77, 0x99)   # コンデンサ
GREEN  = RGBColor(0x2E, 0x7D, 0x32)   # インダクタ
GRAY   = RGBColor(0x55, 0x55, 0x55)
LIGHT  = RGBColor(0xF2, 0xF4, 0xF7)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
BLACK  = RGBColor(0x22, 0x22, 0x22)
SKY    = RGBColor(0xAE, 0xC4, 0xE0)
PALE   = RGBColor(0xCB, 0xD7, 0xE8)

prs = Presentation()
prs.slide_width  = Inches(13.333)   # 16:9
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, x, y, w, h, color, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def textbox(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            space_after=6, line_spacing=1.05):
    """runs: list of paragraphs. Each paragraph = list of (text, size, bold, color, bullet_level)."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        # para is (text, size, bold, color, level)
        text, size, bold, color, level = para
        p.level = level
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = FONT
    return tb


def content_header(slide, title, accent, kicker=None):
    """左に縦アクセントバー＋タイトル。上部ヘッダ。"""
    rect(slide, 0, 0, SW, Inches(1.15), accent)
    rect(slide, 0, Inches(1.15), SW, Pt(3), NAVY)
    if kicker:
        textbox(slide, Inches(0.55), Inches(0.12), Inches(11), Inches(0.35),
                [(kicker, 13, True, RGBColor(0xFF, 0xE6, 0xCC), 0)])
    textbox(slide, Inches(0.55), Inches(0.40), Inches(12.3), Inches(0.7),
            [(title, 28, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)


def body_box(slide, paras, x=Inches(0.7), y=Inches(1.5), w=Inches(12.0), h=Inches(5.6),
             space_after=8, line_spacing=1.12):
    return textbox(slide, x, y, w, h, paras, space_after=space_after, line_spacing=line_spacing)


def bullet(text, size=18, bold=False, color=BLACK, level=0):
    prefix = "" if level == 0 else ""
    mark = ["■ ", "・ ", "– "][min(level, 2)]
    return (mark + text, size, bold, color, level)


def plain(text, size=18, bold=False, color=BLACK, level=0):
    return (text, size, bold, color, level)


def placeholder_note(slide, msg="【要記入】実測値をここに差し込み"):
    rect(slide, Inches(0.7), Inches(6.55), Inches(7.5), Inches(0.5), RGBColor(0xFF, 0xF3, 0xCD))
    textbox(slide, Inches(0.85), Inches(6.58), Inches(7.2), Inches(0.45),
            [(msg, 13, True, RGBColor(0x8A, 0x6D, 0x00), 0)], anchor=MSO_ANCHOR.MIDDLE)


def pagenum(slide, n):
    textbox(slide, Inches(12.4), Inches(6.95), Inches(0.8), Inches(0.4),
            [(str(n), 11, False, GRAY, 0)], align=PP_ALIGN.RIGHT)


# ============================================================== 1. タイトル
s = add_slide()
rect(s, 0, 0, SW, SH, NAVY)
rect(s, 0, Inches(2.7), SW, Pt(4), RED)
rect(s, Inches(3.4), Inches(2.72), Inches(2.2), Pt(4), TEAL)
rect(s, Inches(6.0), Inches(2.72), Inches(2.2), Pt(4), GREEN)
textbox(s, Inches(1.0), Inches(1.7), Inches(11), Inches(1.0),
        [("電子回路 実験レポート", 20, True, SKY, 0)])
textbox(s, Inches(1.0), Inches(2.9), Inches(11.3), Inches(1.6),
        [("受動部品の種類と製作", 46, True, WHITE, 0)])
textbox(s, Inches(1.0), Inches(4.3), Inches(11), Inches(0.8),
        [("― 抵抗器・コンデンサ・インダクタの自作と特性評価 ―", 20, False, PALE, 0)])
textbox(s, Inches(1.0), Inches(5.7), Inches(11), Inches(1.2),
        [("氏名：【要記入】　　学籍番号：【要記入】", 16, False, PALE, 0),
         ("実験日：2026年6月10日", 16, False, PALE, 0)],
        space_after=8)

# ============================================================== 2. 目的
s = add_slide()
content_header(s, "実験の目的", NAVY)
body_box(s, [
    bullet("電気回路の基本素子である「抵抗器 R」「コンデンサ C」「インダクタ L」について理解を深める", 20, True, NAVY),
    plain("", 8, False, BLACK),
    bullet("3 種類の受動部品を実際に製作し、どのような特性のものが実現できるかを調べる", 20, True, NAVY),
    plain("", 8, False, BLACK),
    bullet("製作物の値を テスタ／LCRメータ で測定し、理論値・設計意図と比較・考察する", 20, True, NAVY),
    plain("", 14, False, BLACK),
    bullet("本レポートで扱う 3 つの製作実験", 18, True, GRAY),
    bullet("実験1：抵抗器（鉛筆抵抗・マンガニン線抵抗）", 17, False, RED, 1),
    bullet("実験2：コンデンサ（アルミ箔＋グラフ用紙）", 17, False, TEAL, 1),
    bullet("実験3：インダクタ（ボビンコイル・トロイダルコア）", 17, False, GREEN, 1),
])
pagenum(s, 2)

# ============================================================== 3. 全体像
s = add_slide()
content_header(s, "受動部品の全体像と測定機器", NAVY)
# 3カラム
cols = [("抵抗器 R", RED, "Ω（オーム）", ["電流を制限・分圧", "V = R I", "カラーコードで読む"]),
        ("コンデンサ C", TEAL, "F（ファラッド）", ["電荷を蓄える", "C = εS/d", "μF・pF を多用"]),
        ("インダクタ L", GREEN, "H（ヘンリー）", ["磁束を蓄える", "L = K·D²N²×10⁻⁷", "巻数で調整"])]
cx = Inches(0.7)
cw = Inches(3.9)
gap = Inches(0.18)
for name, col, unit, items in cols:
    rect(s, cx, Inches(1.5), cw, Inches(0.7), col)
    textbox(s, cx, Inches(1.5), cw, Inches(0.7),
            [(name, 20, True, WHITE, 0)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, cx, Inches(2.2), cw, Inches(3.0), LIGHT)
    paras = [(unit, 16, True, col, 0)] + [("・" + it, 15, False, BLACK, 0) for it in items]
    textbox(s, cx + Inches(0.2), Inches(2.4), cw - Inches(0.4), Inches(2.7),
            paras, space_after=10)
    cx = cx + cw + gap
textbox(s, Inches(0.7), Inches(5.5), Inches(12), Inches(1.4),
        [("測定機器：", 17, True, NAVY, 0),
         ("・テスタ … 抵抗値の測定　　・LCRメータ（1kHz）… 容量・インダクタンスの測定", 16, False, BLACK, 0),
         ("・LCRメータはプローブの浮遊インダクタンスを測り、実測値から差し引いて真の値とする", 15, False, GRAY, 0)],
        space_after=6)
pagenum(s, 3)

# ============================================================== 4. 実験1 方法
s = add_slide()
content_header(s, "実験1：抵抗器の製作 ― 方法", RED, kicker="EXPERIMENT 1 / RESISTOR")
body_box(s, [
    bullet("(a) 鉛筆抵抗", 19, True, RED),
    bullet("厚紙上に鉛筆（4B以上）で 20mm×5mm の長方形を描き、塗りつぶす", 17, False, BLACK, 1),
    bullet("両端の抵抗値をテスタで測定", 17, False, BLACK, 1),
    plain("", 8, False, BLACK),
    bullet("(b) マンガニン線抵抗", 19, True, RED),
    bullet("マンガニン線をベークライト棒に巻き付け、両端をハンダで固定", 17, False, BLACK, 1),
    bullet("線同士をクロスさせないよう注意し、抵抗値をテスタで測定", 17, False, BLACK, 1),
    plain("", 10, False, BLACK),
    plain("使用器具：テスタ・工具一式・厚紙・ベークライト棒・クリップ×2・鉛筆・マンガニン線1m", 14, False, GRAY),
])
pagenum(s, 4)

# ============================================================== 5. 実験1 結果
s = add_slide()
content_header(s, "実験1：抵抗器の製作 ― 結果", RED, kicker="EXPERIMENT 1 / RESISTOR")
# 結果テーブル枠
rows = [["製作物", "条件", "測定値 [Ω]"],
        ["鉛筆抵抗", "20mm×5mm（標準）", "【要記入】"],
        ["鉛筆抵抗", "濃く厚塗り", "【要記入】"],
        ["鉛筆抵抗", "薄く / 細く", "【要記入】"],
        ["マンガニン線", "巻数 ○ 回", "【要記入】"]]
from pptx.util import Cm
tbl_shape = s.shapes.add_table(len(rows), 3, Inches(0.9), Inches(1.7), Inches(8.0), Inches(3.4)).table
tbl_shape.columns[0].width = Inches(2.6)
tbl_shape.columns[1].width = Inches(3.0)
tbl_shape.columns[2].width = Inches(2.4)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = tbl_shape.cell(r, c)
        cell.text = val
        para = cell.text_frame.paragraphs[0]
        para.font.size = Pt(15)
        para.font.name = FONT
        para.font.bold = (r == 0)
        para.font.color.rgb = WHITE if r == 0 else BLACK
        cell.fill.solid()
        cell.fill.fore_color.rgb = RED if r == 0 else (WHITE if r % 2 else LIGHT)
textbox(s, Inches(9.2), Inches(1.7), Inches(3.6), Inches(3.4),
        [("観察ポイント", 16, True, RED, 0),
         ("・塗り方で抵抗が大きく変化", 14, False, BLACK, 0),
         ("・線抵抗は値が安定", 14, False, BLACK, 0),
         ("・接触の取り方で値がぶれる", 14, False, BLACK, 0)], space_after=8)
placeholder_note(s)
pagenum(s, 5)

# ============================================================== 6. 実験1 考察
s = add_slide()
content_header(s, "実験1：考察 ― 高抵抗・低抵抗をどう作るか", RED, kicker="EXPERIMENT 1 / DISCUSSION")
body_box(s, [
    bullet("鉛筆抵抗：高抵抗にするには", 19, True, RED),
    bullet("塗る領域を「細く・長く」する（断面積 S を小さく、長さ L を大きく：R = ρL/S）", 17, False, BLACK, 1),
    bullet("薄く塗る（黒鉛層を薄くして実効断面積を減らす）", 17, False, BLACK, 1),
    plain("", 6, False, BLACK),
    bullet("鉛筆抵抗：低抵抗にするには", 19, True, RED),
    bullet("濃く・厚く・幅広に塗り、何度も重ね塗りして黒鉛をつなげる（S を大きく、L を短く）", 17, False, BLACK, 1),
    plain("", 8, False, BLACK),
    bullet("マンガニン線抵抗の特徴", 19, True, RED),
    bullet("抵抗温度係数が小さく値が安定。巻数・線長で抵抗値を設計できるが、鉛筆より低抵抗になりやすい", 17, False, BLACK, 1),
    plain("", 8, False, BLACK),
    plain("→ 抵抗値は材料の抵抗率 ρ と形状（長さ L・断面積 S）で決まる、という基本を実感できた【データで補強】", 16, True, NAVY),
])
pagenum(s, 6)

# ============================================================== 7. 実験2 方法
s = add_slide()
content_header(s, "実験2：コンデンサの製作 ― 方法", TEAL, kicker="EXPERIMENT 2 / CAPACITOR")
body_box(s, [
    bullet("アルミ箔5枚を重ね、端から1cmをホチキスで縦どめ。端に導線もホチキス留め（被膜は紙ヤスリで除去）", 17, False, BLACK),
    bullet("同じものを2組作る", 17, False, BLACK),
    bullet("アルミ箔とグラフ用紙をサンドイッチ状に重ねる（誘電体＝紙）", 17, False, BLACK),
    bullet("2本の導線間の抵抗をテスタで測り、2組のアルミ箔が短絡していないことを確認", 17, False, BLACK),
    bullet("のり巻き状に巻き込み、セロテープで固定", 17, False, BLACK),
    bullet("LCRメータで周波数1kHzの容量 C を測定", 17, False, BLACK),
    plain("", 10, False, BLACK),
    plain("C = ε·S/d　（ε=ε₀εᵣ, ε₀=8.854×10⁻¹² F/m, S:電極面積, d:電極間隔）", 17, True, TEAL),
    plain("使用器具：テスタ・LCRメータ・ホチキス・銅線・アルミ箔(100×100mm)10枚・グラフ用紙10枚 ほか", 13, False, GRAY),
])
pagenum(s, 7)

# ============================================================== 8. 実験2 結果
s = add_slide()
content_header(s, "実験2：コンデンサの製作 ― 結果", TEAL, kicker="EXPERIMENT 2 / CAPACITOR")
rows = [["項目", "値"],
        ["電極面積 S [m²]", "【要記入】"],
        ["電極間隔 d（紙厚）[m]", "【要記入】"],
        ["比誘電率 εᵣ（紙, 約2〜4）", "【要記入】"],
        ["理論値 C = εS/d [F]", "【要記入】"],
        ["測定値 C（1kHz）[F]", "【要記入】"],
        ["導線間抵抗（短絡確認）", "【要記入】"]]
tbl = s.shapes.add_table(len(rows), 2, Inches(0.9), Inches(1.6), Inches(7.2), Inches(4.4)).table
tbl.columns[0].width = Inches(4.4)
tbl.columns[1].width = Inches(2.8)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = tbl.cell(r, c)
        cell.text = val
        para = cell.text_frame.paragraphs[0]
        para.font.size = Pt(15)
        para.font.name = FONT
        para.font.bold = (r == 0)
        para.font.color.rgb = WHITE if r == 0 else BLACK
        cell.fill.solid()
        cell.fill.fore_color.rgb = TEAL if r == 0 else (WHITE if r % 2 else LIGHT)
textbox(s, Inches(8.4), Inches(1.6), Inches(4.4), Inches(4.4),
        [("チェック", 16, True, TEAL, 0),
         ("・理論値と測定値の比", 14, False, BLACK, 0),
         ("・巻きの密着度で d が変動", 14, False, BLACK, 0),
         ("・端の浮きが容量低下要因", 14, False, BLACK, 0)], space_after=8)
placeholder_note(s)
pagenum(s, 8)

# ============================================================== 9. 実験2 考察
s = add_slide()
content_header(s, "実験2：考察 ― 容量の増減と製作の難所", TEAL, kicker="EXPERIMENT 2 / DISCUSSION")
body_box(s, [
    bullet("容量を大きくするには（C = εS/d より）", 19, True, TEAL),
    bullet("電極面積 S を大きく（箔を大きく・枚数を増やす・巻き数を増やす）", 17, False, BLACK, 1),
    bullet("電極間隔 d を小さく（誘電体の紙を薄く、強く密着させて巻く）", 17, False, BLACK, 1),
    bullet("誘電率 ε の高い誘電体を使う（εᵣ の大きい材料に替える）", 17, False, BLACK, 1),
    plain("", 6, False, BLACK),
    bullet("容量を小さくするには", 19, True, TEAL),
    bullet("S を小さく、d を大きく、εᵣ の小さい誘電体にする（上記の逆）", 17, False, BLACK, 1),
    plain("", 8, False, BLACK),
    bullet("製作で大変だった点（記入候補）", 19, True, TEAL),
    bullet("アルミ箔同士を短絡させずに均一な間隔で巻くこと／端のずれ・浮きの防止", 17, False, BLACK, 1),
    plain("→ 測定値が理論値より小さい場合、巻きの緩み＝d 増大や箔のしわが原因と考えられる【データで検証】", 15, True, NAVY),
])
pagenum(s, 9)

# ============================================================== 10. 実験3-1 方法
s = add_slide()
content_header(s, "実験3-1：インダクタの製作 ― 方法（ボビンコイル）", GREEN, kicker="EXPERIMENT 3 / INDUCTOR")
body_box(s, [
    bullet("エナメル線の端を5cm出してボビン端をセロテープで固定", 17, False, BLACK),
    bullet("ボビンにエナメル線を単層密着で20回巻く", 17, False, BLACK),
    bullet("反対側も5cm出して固定し、両端を紙ヤスリで磨きエナメルを除去", 17, False, BLACK),
    bullet("LCRメータ（1kHz）でインダクタンス L を測定", 17, False, BLACK),
    plain("", 10, False, BLACK),
    plain("理論式（ソレノイド）", 18, True, GREEN),
    plain("L = K · D²N² × 10⁻⁷　　K = 1 / { 1 + 0.45(D/ℓ) + 0.005(D/ℓ)² }", 17, True, BLACK),
    plain("D：ボビン外径,　N：巻数,　ℓ：ソレノイド長", 15, False, GRAY),
    plain("使用器具：ボビン(8φ)・トロイダルフェライトコア・エナメル線(0.5φ,80cm)・塩ビ線・LCRメータ ほか", 13, False, GRAY),
])
pagenum(s, 10)

# ============================================================== 11. 実験3-1 結果
s = add_slide()
content_header(s, "実験3-1：結果 ― 理論値 vs 測定値", GREEN, kicker="EXPERIMENT 3 / INDUCTOR")
rows = [["項目", "値"],
        ["外径 D [m]", "【要記入】"],
        ["巻数 N", "20"],
        ["ソレノイド長 ℓ [m]", "【要記入】"],
        ["係数 K", "【要記入】"],
        ["理論値 L [H]", "【要記入】"],
        ["測定値 L（1kHz）[H]", "【要記入】"],
        ["誤差 [%]", "【要記入】"]]
tbl = s.shapes.add_table(len(rows), 2, Inches(0.9), Inches(1.55), Inches(7.2), Inches(4.7)).table
tbl.columns[0].width = Inches(4.4)
tbl.columns[1].width = Inches(2.8)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = tbl.cell(r, c)
        cell.text = val
        para = cell.text_frame.paragraphs[0]
        para.font.size = Pt(14)
        para.font.name = FONT
        para.font.bold = (r == 0)
        para.font.color.rgb = WHITE if r == 0 else BLACK
        cell.fill.solid()
        cell.fill.fore_color.rgb = GREEN if r == 0 else (WHITE if r % 2 else LIGHT)
textbox(s, Inches(8.4), Inches(1.55), Inches(4.4), Inches(4.7),
        [("メモ", 16, True, GREEN, 0),
         ("・誤差=|測定−理論|/理論×100", 14, False, BLACK, 0),
         ("・浮遊インダクタンスを差引", 14, False, BLACK, 0),
         ("・巻きの密着度で L が変化", 14, False, BLACK, 0)], space_after=8)
placeholder_note(s)
pagenum(s, 11)

# ============================================================== 12. 実験3-2 トロイダル
s = add_slide()
content_header(s, "実験3-2：トロイダルコア ― 巻数 N とインダクタンス", GREEN, kicker="EXPERIMENT 3 / INDUCTOR")
rows = [["コアに通した回数 N", "測定 L [μH]"],
        ["0（線のみ）", "【要記入】"],
        ["1", "【要記入】"],
        ["2", "【要記入】"],
        ["3", "【要記入】"],
        ["4", "【要記入】"],
        ["5", "【要記入】"]]
tbl = s.shapes.add_table(len(rows), 2, Inches(0.9), Inches(1.55), Inches(6.0), Inches(4.7)).table
tbl.columns[0].width = Inches(3.6)
tbl.columns[1].width = Inches(2.4)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = tbl.cell(r, c)
        cell.text = val
        para = cell.text_frame.paragraphs[0]
        para.font.size = Pt(15)
        para.font.name = FONT
        para.font.bold = (r == 0)
        para.font.color.rgb = WHITE if r == 0 else BLACK
        cell.fill.solid()
        cell.fill.fore_color.rgb = GREEN if r == 0 else (WHITE if r % 2 else LIGHT)
textbox(s, Inches(7.3), Inches(1.55), Inches(5.5), Inches(4.7),
        [("グラフ枠", 16, True, GREEN, 0),
         ("【要差込】N（横軸）vs L（縦軸）の散布図/折れ線", 14, False, GRAY, 0),
         ("", 8, False, BLACK, 0),
         ("予想：L ∝ N²（巻数の2乗に比例）", 15, True, NAVY, 0),
         ("→ データが N² に乗るか確認", 14, False, BLACK, 0)], space_after=8)
rect(s, Inches(7.4), Inches(2.6), Inches(5.2), Inches(3.4), RGBColor(0xEE, 0xF3, 0xEE),
     line=RGBColor(0xBB, 0xCC, 0xBB))
textbox(s, Inches(7.4), Inches(4.1), Inches(5.2), Inches(0.6),
        [("（ここにグラフ画像を貼付）", 13, False, GRAY, 0)], align=PP_ALIGN.CENTER)
pagenum(s, 12)

# ============================================================== 13. 実験3 考察
s = add_slide()
content_header(s, "実験3：考察 ― 理論との差とトロイダルコイル", GREEN, kicker="EXPERIMENT 3 / DISCUSSION")
body_box(s, [
    bullet("ボビンコイル（実験3-1）", 19, True, GREEN),
    bullet("測定値と理論値の差の要因：巻きの密着度・実効ソレノイド長 ℓ の読み取り・端効果・浮遊成分", 17, False, BLACK, 1),
    plain("", 6, False, BLACK),
    bullet("トロイダルコイル（実験3-2）", 19, True, GREEN),
    bullet("磁束がコア内に閉じ込められ漏れが少ない → 同じ巻数でも空心より大きな L が得られる", 17, False, BLACK, 1),
    bullet("L は巻数 N の2乗に比例（L ∝ N²）。測定データが N² 則に従うか検証する", 17, False, BLACK, 1),
    bullet("高透磁率コアでインダクタンスを稼げる＝小型化に有利。外部磁界の影響も受けにくい", 17, False, BLACK, 1),
    plain("", 8, False, BLACK),
    plain("→ 空心ソレノイドとトロイダルの比較から「磁心の役割」を確認できた【データで補強】", 16, True, NAVY),
])
pagenum(s, 13)

# ============================================================== 14. 全体考察・まとめ
s = add_slide()
content_header(s, "全体考察・まとめ", NAVY)
body_box(s, [
    bullet("抵抗 R：値は材料の抵抗率 ρ と形状（長さ L・断面積 S）で決まる（R = ρL/S）", 18, True, RED),
    bullet("コンデンサ C：C = εS/d。面積を大きく・間隔を小さく・誘電率を高くすると容量増加", 18, True, TEAL),
    bullet("インダクタ L：巻数・形状・磁心で決まる。トロイダルは漏れが少なく L ∝ N²", 18, True, GREEN),
    plain("", 10, False, BLACK),
    bullet("共通して学んだこと", 18, True, NAVY),
    bullet("「自作」することで各素子の値を決める物理量（ρ, ε, μ／形状）を体感的に理解できた", 16, False, BLACK, 1),
    bullet("理論値と測定値の差から、製作精度（密着度・短絡・端効果）と測定の注意点を把握した", 16, False, BLACK, 1),
    bullet("LCRメータでは浮遊インダクタンスの補正が必要であることを確認した", 16, False, BLACK, 1),
])
pagenum(s, 14)

# ============================================================== 15. 参考文献
s = add_slide()
content_header(s, "参考文献", NAVY)
body_box(s, [
    plain("[1] 実験テキスト「第3章 受動部品の種類と製作」", 17, False, BLACK),
    plain("[2] 配布資料「受動部品の種類と製作」（講義スライド）", 17, False, BLACK),
    plain("[3] 「LCRメータの使い方」（Keysight U1731C/U1732C/U1733C, DE-5000, Proster PST077）", 17, False, BLACK),
    plain("[4] 【要追記】参考にした書籍・Webページ（抵抗器・コンデンサの分類で WEB検索した出典）", 17, False, GRAY),
], space_after=14)
pagenum(s, 15)

prs.save("受動部品の種類と製作_実験レポート.pptx")
print("saved:", "受動部品の種類と製作_実験レポート.pptx", "slides:", len(prs.slides._sldIdLst))
