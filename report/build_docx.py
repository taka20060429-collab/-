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


def add_table(headers, rows, caption=None, col_widths=None, font_size=9.5):
    """簡易な罫線付き表。headers: list[str], rows: list[list[str]]"""
    ncol = len(headers)
    tbl = doc.add_table(rows=1 + len(rows), cols=ncol)
    tbl.style = "Table Grid"
    tbl.alignment = 1  # center
    # header
    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = ""
        rr = cell.paragraphs[0].add_run(h)
        set_run_font(rr, ASCII_FONT, BODY_FONT, font_size, bold=True)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.rows[i + 1].cells[j]
            cell.text = ""
            rr = cell.paragraphs[0].add_run(val)
            set_run_font(rr, ASCII_FONT, BODY_FONT, font_size, bold=False)
    if col_widths:
        for row in tbl.rows:
            for j, w in enumerate(col_widths):
                row.cells[j].width = Mm(w)
    if caption:
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_before = Pt(2)
        cap.paragraph_format.space_after = Pt(8)
        cr = cap.add_run(caption)
        set_run_font(cr, ASCII_FONT, BODY_FONT, 9.5, bold=False)
    return tbl


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
para("学籍番号：4326091　　氏名：森木崇允", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
para("実験日：2026年6月10日", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
para("提出日：2026年6月15日", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
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
     "1MΩ（メグオーム）のように、大きな抵抗値になった場合はk（キロ、10³を表す）やM（メグ、10⁶を"
     "表す）のような補助単位をΩにつけて呼ぶのが一般的である。抵抗器には固定抵抗器と可変抵抗器の"
     "2種類があり、また抵抗材料の組成の差により、炭素系、金属系、半導体系などに分類される。"
     "抵抗器の分類を表1に示す。")
para("表1　抵抗器の分類", size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=2)
add_table(["大分類", "材料系", "主な種類"],
          [["固定抵抗器", "炭素系", "炭素皮膜形／炭素系コンポジション形"],
           ["", "金属系", "金属皮膜形（薄膜形）／金属酸化物形／厚膜形／巻線形"],
           ["", "半導体系", "非直線形"],
           ["可変抵抗器", "炭素系", "コンポジション形（皮膜形・ソリッド形）"],
           ["", "金属系", "金属皮膜形（薄膜形）／巻線形"],
           ["", "半導体系", "変換素子・能動素子"]],
          col_widths=[24, 22, 100])
body("(1) 固定抵抗器")
body("① 炭素皮膜抵抗器　円筒状のセラミックの上に炭素の薄い皮膜を作り、そこに螺旋状の溝を切って"
     "抵抗値を出し、両端にリード線を出したものである。抵抗値は皮膜の厚さと溝の切り方によって"
     "決定されるが、高抵抗を得るために皮膜を薄くすると温度特性が悪化する。")
body("② 炭素系コンポジション抵抗器　皮膜形とソリッド形がある。前者は炭素の粉末と樹脂結合剤を混合"
     "してセラミックの上に塗布・焼き付けて皮膜を形成し、後者は炭素粉末・樹脂・無機質充填剤の混合物"
     "を加圧・加熱成形して作る。ソリッド形は皮膜形に比べ周波数特性は劣るが、断線がなく量産性・"
     "小型化に優れており広く用いられている。ただし、両者に共通して電流雑音などは他の抵抗器より"
     "大きい。")
body("③ 金属皮膜抵抗器　ガラスやセラミックなどの上に薄い金属皮膜を蒸着・メッキなどにより形成した"
     "もので、温度係数・電流雑音が小さく安定性にも優れている。多少価格は高い。")
body("④ 巻線抵抗器　セラミックなどのボビンにエナメル被覆されたマンガニン線やニクロム線などを巻き"
     "付けてあるもので、精度・安定性は比較的良く、大電流を流すことができるものが多いが周波数特性"
     "は悪い。")
body("(2) 可変抵抗器")
body("回転軸の回転により軸に連結した摺動子が抵抗体表面を摺動し、端子1-2間と2-3間の抵抗値が変化"
     "する構造をもつ。回転角と抵抗値の関係を抵抗変化特性といい、A形・B形の2種類がよく用いられる。"
     "一定の回路定数に調整した後に半固定化する可変抵抗器は、半固定形やトリマ形と呼ばれ、小型のもの"
     "が多い。")
body("① 炭素系可変抵抗器　抵抗体は皮膜形とソリッド形があり、馬蹄形をした抵抗体と回転軸・軸受・"
     "摺動子などの部品を組み合わせて完成する。最も一般的に使われるのは皮膜形であるが、摺動雑音・"
     "回転寿命・負荷特性の点ではソリッド形の方が優れている。")
body("② 金属皮膜可変抵抗　抵抗体は金属皮膜の固定抵抗器と同じものが用いられる。価格は高いが、"
     "分解能・摺動雑音・接触抵抗・回転寿命などはいずれも炭素系可変抵抗器よりも優れている。")
body("③ 巻線形可変抵抗器　樹脂などでできた巻心にニクロムやマンガニンなどを巻き、環状に湾曲させて"
     "抵抗体としたものである。精度と安定度はよいが、分解能・摺動雑音などは他のものより劣る。")
body("④ その他の可変抵抗器　有効回転角度が360°未満のもののほかに、ネジの機構を応用した多回転形の"
     "ポテンショメータと呼ばれる可変抵抗器もあり、精密な抵抗調整を行いたい場合に用いられる。")
body("(3) 抵抗器の見方")
body("① カラーコード表（JIS）　抵抗値が直接抵抗器の上に数値で印刷されている場合以外は、カラーコード"
     "表に従う。第1・第2数字をa、b、ゼロの数をnとすると、抵抗値は ab×10ⁿ [Ω] となる。例えば "
     "a=2（赤色）、b=5（緑色）、n=3（橙色）であれば、25×10³ Ω＝25kΩ である。金属皮膜などの高精度"
     "抵抗は第3数字まであり、有効数字が3桁になる。カラーコード表を表2に示す。")
para("表2　カラーコード表", size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=2)
add_table(["色", "数字・ゼロの数", "誤差"],
          [["黒", "0", ""], ["茶", "1", "1%"], ["赤", "2", "2%"], ["橙", "3", ""],
           ["黄", "4", ""], ["緑", "5", ""], ["青", "6", ""], ["紫", "7", ""],
           ["灰", "8", ""], ["白", "9", ""], ["金", "n=-1", "5%"], ["銀", "n=-2", "10%"]],
          col_widths=[20, 40, 24])
body("② 定格電力　その抵抗器に連続的に印加できる電力の最大値を指すが、周囲温度が40℃以上の場合に"
     "は制限しなければならない。定格電力としては 1/16W、1/8W、1/4W、1/2W、1W、1.5W、2W、… のように"
     "多数の種類があるが、大きな電力を許容できる抵抗器は外形も大きくなる。")

heading("2-2", "コンデンサ（capacitor）", level=2)
body("コンデンサは抵抗器とともに電気回路内で多く用いられる部品であり、電荷を蓄える働きをもつ。"
     "その容量はF（ファラッド）という単位で表され、実際にはμF（マイクロファラッド、10⁻⁶F）や"
     "pF（ピコファラッド、10⁻¹²F）のように補助単位μ、pをつけたものがよく用いられる。以下に代表的な"
     "コンデンサの種類と特徴を示す。")
body("① 電解コンデンサ　対向電極が電解液のものを電解コンデンサと呼ぶ。数μF〜数1000μFの比較的大きな"
     "容量が必要なときに用いられる。周波数特性は他の種類に比べてあまり良くないため、電源回路など"
     "比較的低い周波数の信号を扱う際に有効である。周囲温度や経年変化に伴う電解液の変質により静電"
     "容量の変化・漏れ電流の増加が発生し、真の静電容量が表示値から数％〜十数％異なることもある。"
     "近年は電解液の代わりに二酸化マンガンやタンタルなどの酸化物を用いた固体電解コンデンサも用い"
     "られ、温度係数・周波数特性が改善されている。")
body("② セラミックコンデンサ　誘電率を大きくとれるので小型にでき、数pF〜1μFの範囲で用いられる。"
     "周波数特性は概して良いが、誘電率の温度係数が大きいため周囲温度の影響を受けやすいものが多い。"
     "容量は第1数字・第2数字・指数用の第3数字で表示され、単位はpFである。例えば473は "
     "47×10³ pF＝0.047μF、102は10×10² pF＝1000pF＝0.001μF を表す。")
body("③ ペーパーコンデンサ　0.001μF〜1μFの静電容量が得られる。2枚のアルミニウムなどの金属箔を"
     "交互に重ねて巻き込み、含浸液を含ませて外装する。紙本来の誘電的性質に含浸液の性質が付加される"
     "ため温度係数はあまり良くない。数100Vの耐圧が必要な回路系で用いられる。")
body("④ マイカコンデンサ　マイカのへき開性を利用して25〜5μmの厚さにしたものを誘電体として用いる。"
     "経時変化がなく、温度係数も小さく、絶縁抵抗も10⁵ MΩと高い値を保つので、高信頼度通信機用など"
     "高い精度が必要な回路系で用いられる。静電容量は数pF〜0.1μF程度で、かなり高い耐圧を持つものも"
     "ある。他のコンデンサに比べ多少価格は高い。")
body("⑤ プラスチックフィルムコンデンサ　各種のプラスチックフィルムを誘電体として金属箔とともに"
     "巻き込んだものである。中でもスチロールフィルムコンデンサ（通称スチコン）は高周波特性が良好で"
     "絶縁抵抗も高く、温度係数も比較的小さい。ポリエチレンテレフタレートコンデンサは電気的特性は"
     "若干劣るが、フィルムの機械的・熱的安定性に優れ、小型大容量化が可能である。")
body("平行平板コンデンサの静電容量Cは、誘電率をε（=ε₀εᵣ、ε₀は真空の誘電率 8.854×10⁻¹² F/m、"
     "εᵣは比誘電率）、電極の表面積をS [m²]、電極間の間隔をd [m] とすると、次式で与えられる。")
para("　　C ＝ ε S ／ d", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
body("この式より、静電容量Cは電極面積Sおよび誘電率εに比例し、電極間隔dに反比例することがわかる。")

heading("2-3", "インダクタンス（inductor）", level=2)
body("電磁誘導作用を利用したコイルやトランス、およびその性質をもつものをインダクタンス部品と呼ぶ。")
body("(1) コイル")
body("コイルを形状から分類すると、ソレノイド形（空心）、トロイダル形（空心）、スパイラル形、"
     "ソレノイド形（磁心入り）、トロイダル形（磁心入り）、つぼ形などの6種類程度が挙げられ、"
     "構造的には空心コイルと磁心入りコイルに大別される。空心コイルは無歪・非飽和・高安定などの"
     "特長をもつが、特性を良くしよう（高いQを実現しよう）とするとどうしても寸法が大きくなるため、"
     "回路を小型化したい場合には不向きである。一方、磁心入りコイルは高透磁率・低損失の磁心材料を"
     "用いて大きなインダクタンスやQを小型でも実現でき、磁心を出し入れできるようにすればインダクタ"
     "ンスを可変にできるという特長をもつ。ただし、磁心による歪電圧の発生・飽和特性・直流重畳特性"
     "などがあり注意が必要である。")
body("ソレノイドコイルのインダクタンスLは、ボビンの外径をD、巻数をN、ソレノイド長をℓとすると、"
     "一般に次式で与えられる。")
para("　　L ＝ K・D²N² × 10⁻⁷　　,　　K ＝ 1 ／ { 1 ＋ 0.45(D／ℓ) ＋ 0.005(D／ℓ)² }",
     size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
body("また、トロイダル形は磁束がコア内に閉じ込められて漏れが少なく、同じ巻数でも空心より大きな"
     "インダクタンスが得られる。インダクタンスは巻数Nの2乗に比例する。")
body("(2) トランス")
body("トランスの主な機能は、①電圧・電流・インピーダンスの変換、②極性反転、③回路内の電位分離、"
     "④平衡回路・不平衡回路の変換、⑤電力分配 などがあり、これらの機能を単独あるいは組み合わせて、"
     "目的に応じた多種類のトランスが使われている。")
body("a. 低周波トランス　音声周波数（約30Hz）から数10kHzまでの周波数帯で用いられるもので、"
     "インピーダンスの変換や電圧・電流変換などを目的として用いられる。")
body("b. 中間周波トランス　無線受信信号などの中間周波増幅回路と呼ばれる回路系で用いられるもので、"
     "混信などを防ぐために容量Cと組み合わせて周波数選択性をもたせてある。")
body("c. 線路形トランス　平衡線路内のインピーダンス不整合を取り除くために設けられたトランスである。")
body("d. パルストランス　サイリスタなどの素子に高圧のパルスを供給したい場合に用いられる。")
body("このほかにも、電力分配を目的としたハイブリッドトランスや、高電圧発生を目的としたフライバック"
     "トランスなどがある。")

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

# =================================================================== 4 実験結果
heading("4", "実験結果")

heading("4-1", "抵抗器の製作実験", level=2)
body("鉛筆で塗り方を変えて描いた3つの長方形①②③を図4.1に示す。①が最も濃く、③が最も薄く塗ったもの"
     "である。それぞれの両端の抵抗値をテスタで測定したところ、① 2100 Ω、② 20 kΩ（20000 Ω）、"
     "③ 1000 kΩ であった。濃く塗った長方形ほど抵抗値が小さく、薄く塗った長方形ほど抵抗値が大きい"
     "という結果が得られた。")
add_figure(os.path.join(IMGDIR, "fig_resistor.jpg"),
           "図4.1　鉛筆で塗りつぶした3つの長方形（上から①②③）と測定値", width_mm=120)
body("次に、ベークライト棒にマンガニン線1mを巻き付けた抵抗器を製作した。両端をハンダで固定し、"
     "テスタで抵抗値を測定したところ 17 Ω であった。製作した抵抗器を図4.2に示す。")
add_figure(os.path.join(IMGDIR, "fig_manganin.jpg"),
           "図4.2　ベークライト棒にマンガニン線を巻き付けた抵抗器", width_mm=110)

heading("4-2", "コンデンサの製作実験", level=2)
body("アルミ箔とグラフ用紙を重ねた段階で2本の銅線間の抵抗値をテスタで測定したところ、2組のアルミ箔"
     "が接触していないことを確認できた。製作し終えたコンデンサを図4.3に示す。LCRメータで周波数1kHz"
     "のときの容量を測定したところ、23 nF であった。")
add_figure(os.path.join(IMGDIR, "fig_capacitor.jpg"),
           "図4.3　製作したコンデンサ", width_mm=120)

heading("4-3", "インダクタの製作実験", level=2)
body("今回、インダクタの製作実験は行わなかった。")

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
     "小さくなり、薄く塗ると炭素の付着が少なく電流が流れづらいため抵抗値が大きくなる。実際、最も"
     "濃く塗った①は2100Ωと最も小さく、最も薄く塗った③は1000kΩと最も大きくなっており、この考えと"
     "整合する結果が得られた。")
body("(2) 鉛筆の芯の電気抵抗率および筆跡の厚さについて考察した。上式より ρ ＝ R S ／ L であり、"
     "長方形は長さ L＝20×10⁻³ m、断面積は筆跡の厚さ h [m]×幅 5.0×10⁻³ m と考えられる。ここで、鉛筆"
     "の芯の黒鉛含有率はHB〜2Bで約70%であることと、黒鉛の電気抵抗率の文献値（参考文献[5]）"
     "1.375×10⁻³ Ω・m を用いると、芯の電気抵抗率の理論値は ρ ＝ 1.375×10⁻³ × 0.7 ＝ 9.625×10⁻² Ω・m "
     "となる。これを用いて、長方形②（R＝20×10³ Ω）の筆跡の厚さ h を求めると次のようになる。")
para("　　h ＝ ρ L ／ ( R × 5.0×10⁻³ ) ＝ ( 9.625×10⁻² × 20×10⁻³ ) ／ ( 20×10³ × 5.0×10⁻³ ) ≒ 1.9×10⁻⁵ [m]",
     size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
body("得られた筆跡の厚さ 1.9×10⁻⁵ m は、コンデンサの実験で誘電体として用いたグラフ用紙の厚さ"
     "（約4.0×10⁻⁵ m）と同程度のオーダーであり、鉛筆で描いた黒鉛層の厚さとして妥当な値であると考えた。")
body("(3) マンガニン線抵抗器は17Ωと小さな抵抗値が得られた。マンガニンは抵抗温度係数が非常に小さく"
     "安定であるため、低抵抗で精度の求められる用途に適している。一方、鉛筆抵抗は塗り方によって"
     "kΩ〜MΩの広い範囲の抵抗値が簡便に得られるが、接触状態や塗りむらの影響を受けやすいと考えられる。")

heading("5-2", "コンデンサの製作についての考察", level=2)
body("(1) コンデンサの製作の上で大変だったことを振り返った。アルミ箔を重ねる際やホチキスでとめる"
     "際に、アルミ箔やグラフ用紙が破れたりずれたりしないよう慎重に作業する必要があった。また、"
     "静電容量を大きくするために全体を強く巻く際や測定の際に、アルミ箔にしわが寄ったり、一度とめた"
     "銅線がホチキスの針から抜けてしまいそうになったりした。")
body("(2) コンデンサの容量の理論値を C＝εS／d より計算し、実験結果4-2の測定値23nFと比較した。"
     "誘電体の紙の比誘電率を εᵣ≒3、真空の誘電率を ε₀＝8.854×10⁻¹² F/m、アルミ箔1対あたりの対向"
     "面積を S≒1.0×10⁻² m²（100mm×100mm）、電極間隔（グラフ用紙の厚さ）を d≒1.0×10⁻⁴ m と仮定"
     "すると、1対あたりの容量の理論値は次のようになる。")
para("　　C ＝ εᵣ ε₀ S ／ d ＝ 3 × 8.854×10⁻¹² × 1.0×10⁻² ／ 1.0×10⁻⁴ ≒ 2.7 [nF]",
     size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
body("本実験では複数枚のアルミ箔を重ねて巻き込んでいるため、実効的な対向面積は1対の場合より大きく"
     "なり、それらが並列に接続されたとみなせる。測定値23nFは1対の理論値2.7nFの約8倍であり、これは"
     "重ね合わせと巻き込みによって実効面積が増加したためと考えられる。容量を大きくするには面積Sを"
     "大きく・間隔dを小さく・比誘電率εᵣの大きい誘電体を用いればよく、小さくするにはその逆にすれば"
     "よい。なお、巻く強さによって極板間距離dが変化するため、強く巻くほど容量が大きくなると考えられる。")
todo("※S・d・εᵣは代表的な仮定値です。実際に測定した対向面積・グラフ用紙の厚さがあれば、その値で"
     "理論値を計算し直すとより正確に比較できます。")

# =================================================================== 6 まとめ
heading("6", "まとめ")
body("電気回路の基本素子である抵抗器、コンデンサ、インダクタンスについて、それぞれの種類や性質を"
     "学んだ上で、実際に製作する過程を通して、各部品がどのようにして各々の特性を実現しているのかを"
     "理解した。特に、抵抗器では塗り方（断面積）によって抵抗値が大きく変化すること、コンデンサでは"
     "電極面積や極板間距離によって静電容量が変化することを、自作と測定を通して確認した。インダクタンス"
     "についても、巻数や磁心の有無によってインダクタンスが決まることを原理的に理解した。"
     "また、実際に手を動かしてみて、マンガニン線抵抗器の抵抗値が17Ωと予想していたよりもかなり低く、"
     "正直なところ驚いた。一方で、鉛筆の塗り方を変えるだけで抵抗値がkΩからMΩまで大きく変化したのは"
     "面白く、身近な材料でも工夫次第で部品の特性を作り込めることを実感した。コンデンサについても、"
     "巻く強さによって容量が変わりそうだと作りながら感じ、製作の丁寧さが特性に直結することを学んだ。"
     "同じ基本素子"
     "であっても使用用途によって適切な部品"
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
