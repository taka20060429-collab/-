# -*- coding: utf-8 -*-
"""受動部品の種類と製作 — 実験レポート（文書/Word）生成スクリプト
過去レポート(東京理科大・受動部品の種類と制作)の構成を踏襲。黒一色・A4縦。
目的/原理/方法は教科書準拠で記述、結果は本人の写真+データ、
考察・まとめは本人の実験データに合わせて記述。C/Lの測定値は【要記入】。
"""
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

IMGDIR = os.path.join(os.path.dirname(__file__), "img")

BODY_FONT = "ＭＳ 明朝"      # 本文（明朝）
HEAD_FONT = "ＭＳ ゴシック"  # 見出し（ゴシック）
ASCII_FONT = "Times New Roman"

doc = Document()

# ---- 既定スタイル ----
style = doc.styles["Normal"]
style.font.name = ASCII_FONT
style.font.size = Pt(10.5)
style.font.color.rgb = RGBColor(0, 0, 0)
style._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
pf = style.paragraph_format
pf.line_spacing = 1.15
pf.space_after = Pt(0)

# ---- ページ設定 A4 ----
sec = doc.sections[0]
sec.page_width = Mm(210)
sec.page_height = Mm(297)
sec.top_margin = Mm(25)
sec.bottom_margin = Mm(25)
sec.left_margin = Mm(25)
sec.right_margin = Mm(25)


def set_run_font(run, ascii_font=ASCII_FONT, ea_font=BODY_FONT, size=None, bold=False):
    run.font.name = ascii_font
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    rf.set(qn("w:ascii"), ascii_font)
    rf.set(qn("w:hAnsi"), ascii_font)
    rf.set(qn("w:eastAsia"), ea_font)
    if size:
        run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)


def para(text="", size=10.5, bold=False, align=None, ea=BODY_FONT, ascii_f=ASCII_FONT,
         space_before=0, space_after=4, indent=None, line=1.15):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if indent is not None:
        pf.first_line_indent = Pt(indent)
    if text:
        r = p.add_run(text)
        set_run_font(r, ascii_f, ea, size, bold)
    return p


def heading(num, text, level=1):
    sizes = {1: 14, 2: 12}
    sb = {1: 14, 2: 8}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(sb[level])
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(f"{num}　{text}")
    set_run_font(r, HEAD_FONT, HEAD_FONT, sizes[level], bold=True)
    return p


def body(text, indent=10.5):
    """本文段落（1字下げ）。"""
    return para(text, size=10.5, space_after=4, indent=indent, line=1.3)


def add_figure(path, caption, width_mm=120):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    if os.path.exists(path):
        run.add_picture(path, width=Mm(width_mm))
    else:
        rr = p.add_run(f"[画像なし: {path}]"); set_run_font(rr)
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(10)
    cr = cap.add_run(caption)
    set_run_font(cr, ASCII_FONT, BODY_FONT, 9.5, bold=False)


def todo(text):
    """【要記入】等の赤字注記（最終版では黒に。ここでは視認用に太字）。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Mm(4)
    r = p.add_run(text)
    set_run_font(r, ASCII_FONT, BODY_FONT, 10.5, bold=True)
    return p


def page_number_footer():
    footer = sec.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
    fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
    set_run_font(run, ASCII_FONT, BODY_FONT, 10)


# =================================================================== 表紙
for _ in range(3):
    para("", space_after=2)
para("電気基礎実験", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
     ea=HEAD_FONT, ascii_f=HEAD_FONT, space_after=10)
para("受動部品の種類と製作", size=22, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
     ea=HEAD_FONT, ascii_f=HEAD_FONT, space_after=60)
para("学籍番号：【要記入】　　氏名：【要記入】", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
para("実験日：2026年6月10日", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
para("提出日：【要記入】", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
doc.add_page_break()

# =================================================================== 1 実験目的
heading("1", "実験目的")
body("抵抗器、コンデンサ、インダクタンスなどの基本素子について調べて理解を深め、さらに実際に"
     "製作することで、どのような特性のものが実現できるかを調べることが本実験の目的である。")

# =================================================================== 2 原理
heading("2", "原理")

heading("2-1", "抵抗器（resistor）", level=2)
body("抵抗器（通称、単に抵抗と呼ばれる）は電気回路において非常によく用いられる部品である。"
     "抵抗値はΩ（オーム）という単位で呼ばれ、1000Ωは1kΩ（キロオーム）、1000kΩ（1000000Ω）は"
     "1MΩ（メグオーム）のように、大きな抵抗値になった場合はk（キロ、10³）やM（メグ、10⁶）の"
     "補助単位をΩにつけて呼ぶのが一般的である。")
body("抵抗器には固定抵抗器と可変抵抗器の2種類があり、また抵抗材料の組成の差により、炭素系、"
     "金属系、半導体系などに分類される。固定抵抗器の代表例として、円筒状のセラミックの上に炭素の"
     "薄い皮膜を作り螺旋状の溝を切って抵抗値を出す炭素皮膜抵抗器、炭素粉末と樹脂を混合・成形した"
     "炭素系コンポジション抵抗器、温度係数・電流雑音が小さく安定性に優れる金属皮膜抵抗器、"
     "マンガニン線やニクロム線を巻き付けた巻線抵抗器などがある。")
body("抵抗器の抵抗値はカラーコード表（JIS）によって読み取ることができる。第1・第2数字をa、b、"
     "ゼロの数をnとすると、抵抗値は ab×10ⁿ [Ω] で表される。また、連続的に印加できる電力の最大値"
     "を定格電力といい、1/16W、1/8W、1/4W、… のように多数の種類がある。")

heading("2-2", "コンデンサ（capacitor）", level=2)
body("コンデンサは抵抗器とともに電気回路内で多く用いられる部品であり、電荷を蓄える働きをもつ。"
     "その容量はF（ファラッド）という単位で表され、実際にはμF（10⁻⁶F）やpF（10⁻¹²F）の補助単位が"
     "よく用いられる。種類としては、比較的大きな容量が得られる電解コンデンサ、小型で周波数特性の"
     "良いセラミックコンデンサ、ペーパーコンデンサ、経時変化が小さく高精度なマイカコンデンサ、"
     "プラスチックフィルムコンデンサなどがある。")
body("平行平板コンデンサの静電容量Cは、誘電率をε（=ε₀εᵣ、ε₀は真空の誘電率 8.854×10⁻¹² F/m、"
     "εᵣは比誘電率）、電極の表面積をS [m²]、電極間の間隔をd [m] とすると、次式で与えられる。")
para("　　C ＝ ε S ／ d", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
body("この式より、静電容量Cは電極面積Sに比例し、電極間隔dに反比例することがわかる。")

heading("2-3", "インダクタンス（inductor）", level=2)
body("電磁誘導作用を利用したコイルやトランス、およびその性質をもつものをインダクタンス部品と呼ぶ。"
     "コイルは構造的に空心コイルと磁心入りコイルに大別され、形状からはソレノイド形・トロイダル形・"
     "スパイラル形などに分類される。空心コイルは無歪・非飽和・高安定などの特長をもつが、特性を"
     "良くしようとすると寸法が大きくなる。一方、磁心入りコイルは高透磁率・低損失の磁心材料を"
     "用いることで、大きなインダクタンスを小型で実現できる。")
body("ソレノイドコイルのインダクタンスLは、ボビンの外径をD、巻数をN、ソレノイド長をℓとすると、"
     "一般に次式で与えられる。")
para("　　L ＝ K・D²N² × 10⁻⁷　　,　　K ＝ 1 ／ { 1 ＋ 0.45(D／ℓ) ＋ 0.005(D／ℓ)² }",
     size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
body("トロイダル形は磁束がコア内に閉じ込められて漏れが少なく、同じ巻数でも空心より大きな"
     "インダクタンスが得られる。インダクタンスは巻数Nの2乗に比例する。")

# =================================================================== 3 実験方法
heading("3", "実験方法")

heading("3-1", "抵抗器の製作", level=2)
body("① 厚紙上に鉛筆（4B以上）を用いて 20mm×5mm の長方形を描き、塗りつぶした。塗り方を変えた"
     "ものを3種類用意し、それぞれの両端の抵抗値をテスタで測定した。")
body("② マンガニン線をベークライト棒に、線同士をクロスさせないように巻き付け、両端をハンダで"
     "固定して抵抗値をテスタで測定した。")

heading("3-2", "コンデンサの製作", level=2)
body("① アルミ箔5枚を重ね、1辺の端から1cmの位置をホチキスで縦にとめ、端に導線もホチキスでとめた"
     "（被膜は紙ヤスリで除去）。同じものをもう1組作製した。")
body("② アルミ箔とグラフ用紙（誘電体）がサンドイッチ状になるように重ね、2本の銅線間の抵抗値を"
     "テスタで測定して、2組のアルミ箔が接触していないことを確認した。")
body("③ のり巻きを作るように巻き込み、セロテープで固定した。LCRメータで周波数1kHzのときの容量を"
     "測定した。")

heading("3-3", "インダクタの製作", level=2)
body("① エナメル線の端を5cmほど出してボビンの端をセロテープで固定し、ボビンにエナメル線を単層"
     "密着で20回巻いた。反対側も5cmほど出して固定し、両端を紙ヤスリで磨いてエナメルを除去した。")
body("② LCRメータで周波数1kHzのときのインダクタンスを測定した。")

# =================================================================== 4 実験結果
heading("4", "実験結果")

heading("4-1", "抵抗器の製作実験", level=2)
body("鉛筆で塗り方を変えて描いた3つの長方形①②③を図4.1に示す。それぞれの両端の抵抗値をテスタで"
     "測定したところ、① 2100 Ω、② 2000 Ω、③ 1600 kΩ であった。")
todo("※②・③の数値は写真の手書きが不明瞭なため要確認（③は1000 kΩの可能性あり）。"
     "また、各長方形の塗り方（濃く塗った／薄く塗った／標準）の対応も確認してください。")
add_figure(os.path.join(IMGDIR, "fig_resistor.jpg"),
           "図4.1　鉛筆で塗りつぶした3つの長方形と測定値", width_mm=120)
todo("（マンガニン線抵抗を製作した場合）測定した抵抗値：【要記入】Ω　— 製作した抵抗器の写真があれば追加します。")

heading("4-2", "コンデンサの製作実験", level=2)
body("アルミ箔とグラフ用紙を重ねた段階で2本の銅線間の抵抗値をテスタで測定したところ、2組のアルミ箔"
     "が接触していないことを確認できた。製作し終えたコンデンサを図4.2に示す。LCRメータで周波数1kHz"
     "のときの容量を測定した。")
todo("測定した静電容量：【要記入】nF（複数回測定した場合はその値も）")
add_figure(os.path.join(IMGDIR, "fig_capacitor.jpg"),
           "図4.2　製作したコンデンサ", width_mm=120)

heading("4-3", "インダクタの製作実験", level=2)
body("ボビンにエナメル線を単層密着で20回巻いてインダクタを製作した。製作したインダクタを図4.3に"
     "示す。LCRメータで周波数1kHzのときのインダクタンスを測定した。")
todo("測定したインダクタンス：【要記入】µH（巻数 N＝20、ボビン外径 D＝【要記入】mm、"
     "ソレノイド長 ℓ＝【要記入】mm）")
add_figure(os.path.join(IMGDIR, "fig_inductor.jpg"),
           "図4.3　製作したインダクタ（ボビンコイル）", width_mm=120)

# =================================================================== 5 考察
heading("5", "考察")

heading("5-1", "抵抗器の製作についての考察", level=2)
body("(1) 長方形の大きさを同じにしたまま、高抵抗・低抵抗を得るには長方形をどのように塗ればよいか"
     "を考察した。一般に、ある導体の長さをL [m]、断面積をS [m²]、電気抵抗率をρ [Ω・m] とすると、"
     "電気抵抗R [Ω] は次式で表される。")
para("　　R ＝ ρ L ／ S", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
body("本実験では黒鉛の電気抵抗率ρと長方形の長さ20mmが固定されているため、抵抗Rを小さくするには"
     "断面積S（≒筆跡の厚さ）を大きく、抵抗Rを大きくするには断面積Sを小さくすればよいことがわかる。"
     "すなわち、濃く厚く塗りつぶすと炭素（黒鉛）が紙面に多く付着して電流が流れやすくなり抵抗値が"
     "小さくなり、薄く塗ると炭素の付着が少なく電流が流れづらいため抵抗値が大きくなる。"
     "測定結果（③が他より著しく大きい）も、この考えと整合する。")
todo("※どの長方形をどう塗ったかを確認のうえ、「濃く塗った②が低抵抗・薄く塗った③が高抵抗」"
     "のように測定値と対応づけて記述します。")
body("(2) 鉛筆の芯の電気抵抗率について考察した。上式より ρ ＝ R S ／ L である。長方形②は長さ20mm、"
     "断面積は筆跡の厚さ h [m]×幅5.0×10⁻³ m と考えられるので、測定値 R＝2000 Ω を用いると鉛筆の芯の"
     "電気抵抗率を見積もることができる。黒鉛の電気抵抗率の文献値（参考文献[5]）と比較して、"
     "筆跡の厚さhを評価する。")
todo("※確定した抵抗値・幅・想定する黒鉛含有率を用いて、ρおよび筆跡の厚さhを具体的に計算して"
     "記入します（過去レポートと同じ計算手順）。")

heading("5-2", "コンデンサの製作についての考察", level=2)
body("(1) コンデンサの製作の上で大変だったことを振り返った。アルミ箔を重ねる際やホチキスでとめる"
     "際に、アルミ箔やグラフ用紙が破れたりずれたりしないよう慎重に作業する必要があった。また、"
     "静電容量を大きくするために全体を強く巻く際や測定の際に、アルミ箔にしわが寄ったり、一度とめた"
     "銅線がホチキスの針から抜けてしまいそうになったりした。")
body("(2) コンデンサの容量の理論値を C＝εS／d より計算し、実験結果4-2の測定値と比較する。"
     "電極の表面積S、電極間隔d（グラフ用紙の厚さ）、紙の比誘電率εᵣを用いて理論値を求め、"
     "測定値との差について考察する。容量を大きくするには面積Sを大きく・間隔dを小さく・比誘電率εᵣ"
     "の大きい誘電体を用いればよく、小さくするにはその逆にすればよい。巻く強さによって極板間距離d"
     "が変化するため、静電容量の大小に影響が出ると考えられる。")
todo("※確定した測定容量とS・dの値を用いて、理論値と実測値を具体的に比較して記入します。")

heading("5-3", "インダクタの製作についての考察", level=2)
body("ソレノイドコイルのインダクタンスの理論値を L＝K・D²N²×10⁻⁷ より計算し、実験結果4-3の測定値"
     "と比較する。巻数N＝20、ボビン外径D、ソレノイド長ℓを代入して理論値を求め、測定値との差の要因"
     "（巻きの密着度、ソレノイド長の読み取り、端効果、プローブの浮遊インダクタンスなど）について"
     "考察する。なお、トロイダルコアを用いると磁束の漏れが少なくなるため、同じ巻数でも空心より"
     "大きなインダクタンスが得られ、インダクタンスは巻数Nの2乗に比例する。")
todo("※確定した測定インダクタンスとD・ℓの値を用いて、理論値と実測値を具体的に比較して記入します。")

# =================================================================== 6 まとめ
heading("6", "まとめ")
body("電気回路の基本素子である抵抗器、コンデンサ、インダクタンスについて、それぞれの種類や性質を"
     "学んだ上で、実際に製作する過程を通して、各部品がどのようにして各々の特性を実現しているのかを"
     "理解した。抵抗器では塗り方（断面積）によって抵抗値が変化すること、コンデンサでは電極面積や"
     "極板間距離によって静電容量が変化すること、インダクタでは巻数や磁心によってインダクタンスが"
     "変化することを、自作と測定を通して確認した。同じ基本素子であっても使用用途によって適切な部品"
     "の種類が異なるため、今後電気回路を製作する際には、その都度使用する部品の特性を考慮して適切に"
     "見極められるようにしたい。")

# =================================================================== 参考文献
heading("参考文献", "", level=1)
refs = [
    "[1] 前期電気基礎実験 実験指導書、東京理科大学工学部電気工学科（3章 受動部品）",
    "[2] 前期電気基礎実験 講義資料「受動部品の種類と製作」、東京理科大学工学部電気工学科",
    "[3] 前期電気基礎実験 補足資料「LCRメータの使い方」、東京理科大学工学部電気工学科",
    "[4] 日本筆記具工業会、鉛筆お役立ち情報「鉛筆の種類」　http://www.jwima.org/pencil/04shurui/04shurui.html",
    "[5] 富士黒鉛工業株式会社「黒鉛の特性」　https://www.fujikokuen.co.jp/about",
    "[6] 株式会社Y.E.I「比誘電率表」　http://www.yei-jp.com/tech-infor/dielectric/dielectric02.html",
]
for r in refs:
    p = para(r, size=10, space_after=3)
    p.paragraph_format.left_indent = Mm(8)
    p.paragraph_format.first_line_indent = Mm(-8)

page_number_footer()

out = os.path.join(os.path.dirname(__file__), "受動部品の種類と製作_実験レポート.docx")
doc.save(out)
print("saved:", out)
