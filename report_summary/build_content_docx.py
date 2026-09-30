# -*- coding: utf-8 -*-
"""レポート内容の改善点（自己点検）生成スクリプト。A4縦・黒一色・約4ページ。
内容（有効数字・単位・グラフ・誤差評価・測定条件・仮定・考察の深さ）の改善点を
改善前→改善後→理由の対比で示す。"""
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

IMGDIR = os.path.join(os.path.dirname(__file__), "img")
BODY_FONT = "ＭＳ 明朝"; HEAD_FONT = "ＭＳ ゴシック"; ASCII_FONT = "Times New Roman"

doc = Document()
s = doc.styles["Normal"]
s.font.name = ASCII_FONT; s.font.size = Pt(10.5); s.font.color.rgb = RGBColor(0, 0, 0)
s._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
s.paragraph_format.line_spacing = 1.15; s.paragraph_format.space_after = Pt(0)
sec = doc.sections[0]
sec.page_width = Mm(210); sec.page_height = Mm(297)
sec.top_margin = Mm(22); sec.bottom_margin = Mm(20); sec.left_margin = Mm(24); sec.right_margin = Mm(24)


def srf(run, af=ASCII_FONT, ea=BODY_FONT, size=None, bold=False):
    run.font.name = af
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    rf.set(qn("w:ascii"), af); rf.set(qn("w:hAnsi"), af); rf.set(qn("w:eastAsia"), ea)
    if size:
        run.font.size = Pt(size)
    run.font.bold = bold; run.font.color.rgb = RGBColor(0, 0, 0)


def para(text="", size=10.5, bold=False, align=None, ea=BODY_FONT, af=ASCII_FONT, sb=0, sa=4, indent=None, line=1.3):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format; pf.space_before = Pt(sb); pf.space_after = Pt(sa); pf.line_spacing = line
    if indent is not None:
        pf.first_line_indent = Pt(indent)
    if text:
        r = p.add_run(text); srf(r, af, ea, size, bold)
    return p


def heading(num, text, level=1):
    sizes = {1: 13.5, 2: 11.5}; sb = {1: 12, 2: 8}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(sb[level]); p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(f"{num}　{text}" if text else num); srf(r, HEAD_FONT, HEAD_FONT, sizes[level], True)
    return p


def body(text, indent=10.5):
    return para(text, size=10.5, sa=5, indent=indent, line=1.35)


def ba(before, after, why):
    for lbl, txt in [("改善前：", before), ("改善後：", after), ("理由：", why)]:
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.left_indent = Mm(4); p.paragraph_format.line_spacing = 1.32
        r1 = p.add_run(lbl); srf(r1, HEAD_FONT, HEAD_FONT, 10.5, True)
        r2 = p.add_run(txt); srf(r2, ASCII_FONT, BODY_FONT, 10.5, False)


def add_table(headers, rows, colw, fs=9.5):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers)); t.style = "Table Grid"; t.alignment = 1
    for j, h in enumerate(headers):
        c = t.rows[0].cells[j]; c.text = ""; rr = c.paragraphs[0].add_run(h)
        srf(rr, HEAD_FONT, BODY_FONT, fs, True); c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = t.rows[i + 1].cells[j]; c.text = ""; rr = c.paragraphs[0].add_run(v)
            srf(rr, ASCII_FONT, BODY_FONT, fs, False)
    for r in t.rows:
        for j, w in enumerate(colw):
            r.cells[j].width = Mm(w)
    return t


def figure(path, cap, w=160):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    if os.path.exists(path):
        run.add_picture(path, width=Mm(w))
    cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER; cp.paragraph_format.space_after = Pt(8)
    cr = cp.add_run(cap); srf(cr, ASCII_FONT, BODY_FONT, 9.5, False)


def footer_pn():
    p = sec.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; run = p.add_run()
    a = OxmlElement("w:fldChar"); a.set(qn("w:fldCharType"), "begin")
    b = OxmlElement("w:instrText"); b.set(qn("xml:space"), "preserve"); b.text = "PAGE"
    c = OxmlElement("w:fldChar"); c.set(qn("w:fldCharType"), "end")
    run._r.append(a); run._r.append(b); run._r.append(c); srf(run, ASCII_FONT, BODY_FONT, 10)


C = WD_ALIGN_PARAGRAPH.CENTER
# ===== タイトル =====
para("レポート内容の改善点（自己点検）", size=19, bold=True, align=C, ea=HEAD_FONT, af=HEAD_FONT, sb=6, sa=6)
para("―「受動部品の種類と製作」ほか実験レポートの内容面で、さらに良くできる点 ―", size=11.5, align=C, sa=10)
para("作成者：4326091　森木崇允　／　作成日：2026年9月30日", size=10.5, align=C, sa=12)

# ===== 1 はじめに =====
heading("1", "はじめに")
body("本書は、作成した実験レポートについて、内容（データ処理・図表・考察）の面でさらに改善できる点を、"
     "改善前・改善後の具体例とともに自己点検としてまとめたものである。いずれも完成版レポートを土台に、"
     "精度と説得力をさらに高めるための細かな改善点である。")

# ===== 2 改善点の一覧 =====
heading("2", "改善点の一覧")
add_table(
    ["項目", "改善前", "改善後"],
    [["有効数字", "桁数が不統一（2100Ω, 20kΩ, ≒2.7nF）", "測定確度に合わせ2桁に統一（2.1×10³Ω 等）"],
     ["単位表記", "kΩ／Ω／指数が混在", "SI接頭辞で統一（2.1kΩ／20kΩ／1.0MΩ）"],
     ["グラフ", "表のみ。線形軸だと小さい値が潰れる", "対数目盛の棒グラフ＋数値ラベル"],
     ["誤差評価", "「約8倍」等の定性表現・1回測定", "相対誤差[%]＋複数回測定の平均±ばらつき"],
     ["測定条件", "機器・周波数・回数・室温の記載不足", "使用機器(DE-5000)・1kHz・回数・室温を明記"],
     ["理論値の仮定", "S・dを仮定値のまま計算", "S・dを実測、εᵣは出典明記、途中式に単位"],
     ["考察の深さ", "誤差要因を列挙するのみ", "系統誤差／偶然誤差に分類し支配要因を推定"],
     ["結果の説明(NW)", "記入シート写真のみ", "主要数値(IP設計・サブネット)を本文で説明"]],
    colw=[24, 62, 66])
para("", sa=2)
figure(os.path.join(IMGDIR, "graph_beforeafter.png"),
       "図1　グラフの改善例（左：線形目盛で①②が潰れる／右：対数目盛で3点とも判読可）", w=160)

# ===== 3 各改善の詳細 =====
heading("3", "各改善の詳細（改善前 → 改善後）")

heading("3-1", "有効数字の統一", level=2)
ba("測定値・計算値の有効数字がばらばらである。例：「① 2100 Ω」「h≒1.9×10⁻⁵ m」「C≒2.7 nF」。"
   "測定器の確度を超えた桁を書いている箇所がある。",
   "測定器の確度（例：±(1.0%＋数字)）を踏まえ、有効数字を2桁に統一する。"
   "「R₁ ＝ 2.1×10³ Ω」「h ≈ 1.9×10⁻⁵ m」「C ≈ 2.7 nF」のように桁を揃え、「（有効数字2桁）」と根拠を添える。",
   "測定精度を超えた桁は物理的に意味を持たず、有効数字を揃えるのは実験レポートの基本作法だから。")

heading("3-2", "単位・量記号の表記統一", level=2)
ba("「20kΩ」「1000kΩ」「2100Ω」のように接頭辞と桁が混在し、桁の比較がしづらい。",
   "桁に応じてSI接頭辞を統一する（2.1 kΩ／20 kΩ／1.0 MΩ）。数値と単位の間の空白なども統一する。",
   "単位系の一貫性は可読性と正確さに直結するから。")

heading("3-3", "グラフの活用と見やすさ", level=2)
ba("抵抗値を表だけで示しており、塗り方と抵抗の関係が視覚的に伝わりにくい。仮にグラフ化しても、"
   "線形目盛では①(2.1kΩ)・②(20kΩ)が③(1.0MΩ)に潰れて読めない（図1左）。",
   "縦軸を対数目盛にした棒グラフにし、各棒に数値ラベルを付す。桁の異なる3点を同時に判読でき、"
   "塗りが濃いほど抵抗が小さい傾向も一目で分かる（図1右）。",
   "桁差の大きいデータは対数軸が適切で、視認性が大幅に向上するから。")

heading("3-4", "誤差の定量評価", level=2)
ba("理論値と実測値の差を「約8倍」「妥当」など定性的に述べるにとどまり、測定も1回のみである。",
   "相対誤差[%]を計算して明記する（相対誤差 ＝ |実測−理論| ／ 理論 × 100）。さらに各量を3回測定し、"
   "平均±標準偏差の形で偶然誤差の大きさを示す。",
   "定量的な誤差評価があってはじめて、考察の主張に説得力が生まれるから。")

heading("3-5", "測定条件の明記", level=2)
ba("「LCRメータで測定した」とあるが、機器の型番・測定周波数・測定回数・室温などの条件が本文に不足している。",
   "使用機器（例：DE-5000）、測定周波数1kHz、測定回数、室温・湿度などの条件を「実験方法」または"
   "「結果」に明記する。",
   "測定条件の記載は、実験の再現性を確保するために不可欠だから。")

heading("3-6", "理論値計算における仮定の扱い", level=2)
ba("コンデンサの理論値計算で、電極面積S・電極間隔d・比誘電率εᵣを仮定値のまま用いており、"
   "途中式や単位換算もやや簡略である。",
   "Sはアルミ箔の重なり面積を実測し、dは紙厚をノギスやマイクロメータで測定、εᵣは出典を明記する。"
   "途中式は単位を付けて段階的に示す。",
   "仮定に頼らず実測に基づけば、理論値との比較の信頼性が高まるから。")

heading("3-7", "考察における誤差要因の分析", level=2)
ba("誤差要因を「巻きの密着度」「端効果」などと列挙するにとどまっている。",
   "誤差を系統誤差（測定器の確度、dの過大評価など）と偶然誤差（接触状態のばらつきなど）に分類し、"
   "どちらが支配的かを見積もったうえで、具体的な改善策（強く均一に巻く等）を併記する。",
   "要因を分類・定量化することで、考察がより深く論理的になるから。")

heading("3-8", "結果数値の本文説明（ネットワーク実習）", level=2)
ba("結果を記入シートの写真のみで示しており、本文に数値の意味づけがない。",
   "主要な結果を本文で簡潔に説明する。例：「自分(192.168.111.132)とパートナー(192.168.111.100)は"
   "同一サブネット(/24)にあり、ホスト部のみが異なるためpingが通った」「有線・無線ともデフォルト"
   "ゲートウェイは192.168.11.1で共通だった」など。",
   "数値を示すだけでなくその意味を説明することで、理解が読み手に伝わるから。")

# ===== 4 まとめ =====
heading("4", "まとめ")
body("以上は、いずれも完成版レポートを土台にした細かな改善点であるが、①有効数字と単位の統一、"
     "②図表（特にグラフ）の見やすさ、③誤差の定量評価、④測定条件と計算の仮定の明示、という4つの観点は、"
     "実験レポートの精度と説得力を大きく左右する重要な要素である。次回以降のレポートでは、これらを"
     "最初から意識して作成し、データの示し方と考察の質をさらに高めていきたい。")

footer_pn()
out = os.path.join(os.path.dirname(__file__), "レポート内容の改善点_自己点検.docx")
doc.save(out); print("saved:", out)
