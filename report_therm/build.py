# -*- coding: utf-8 -*-
"""14. サーミスタの実験 — 報告書（Word）の生成。 python3 build.py
グラフと表は TA 資料「レポート作成時の注意点」に合わせる：
  グラフ…タイトル・補助線・グラフエリアの枠線なし，プロットエリアの枠線あり，文字と線は黒，本文と同じフォント，
         軸ラベル（単位つき），目盛あり，図題は下．
  表……表題は上，中央揃え，罫線，長い表は二段，有効数字をそろえる，空欄は斜線．"""
import math
import os
import subprocess
import sys

import numpy as np
import openpyxl
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm
from openpyxl.chart import Reference, Series
from openpyxl.chart.axis import ChartLines  # noqa: F401
from openpyxl.chart.marker import Marker  # noqa: F401
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "report_optics"))
import common as C  # noqa: E402
from common import body, eq, heading, item, nxt  # noqa: E402

C.MATH[0] = "omml"   # 数式は Word の数式で書く

META = {"kai": None, "date": None, "submit": None, "han": "D班", "name": "森木崇允", "grade": "1",
        "student_id": "4326091",
        "coworkers": [("石井真人", "4326007"), ("木元喜紹", "4326029"), ("陶山稜平", "4326051"),
                      ("林田知子", "4326071")]}
IMG = os.path.join(HERE, "img")

# ================================================================ データ
ws_in = openpyxl.load_workbook(os.path.join(HERE, "data", "thermistor.xlsx"), data_only=True).active
DATA = [(ws_in[f"B{r}"].value, ws_in[f"C{r}"].value) for r in range(3, 24)]
TC = np.array([t for t, _ in DATA], float)          # 温度 [℃]
R = np.array([r for _, r in DATA], float)           # 抵抗 [Ω]
TK = TC + 273.15
X = 1 / TK
Y = np.log(R)
n = len(X)
Xm, Ym, X2m = X.mean(), Y.mean(), (X ** 2).mean()
B_LS = ((X * Y).sum() - n * Xm * Ym) / (n * (X2m - Xm ** 2))
A_LS = Ym - B_LS * Xm
R2 = np.corrcoef(X, Y)[0, 1] ** 2
fitR = lambda x: math.exp(A_LS + B_LS * x)  # noqa: E731

# グラフの直線上から読み取る 2 点（1/T を目盛のきりのよい値にとり，R はグラフから 3 桁で読む）
P1X, P2X = 3.18e-3, 3.38e-3
P1R, P2R = round(fitR(P1X), 1), round(fitR(P2X), 1)
B_2P = 2.3 * (math.log10(P2R) - math.log10(P1R)) / (P2X - P1X)
B_2P_exact = math.log(P2R / P1R) / (P2X - P1X)
B_END = math.log(R[0] / R[-1]) / (X[0] - X[-1])     # 測定値の両端 2 点をそのまま使った場合
diffp = (B_2P - B_LS) / B_LS * 100
# 最初の 2 点を除いた場合
Bx, Ax = np.polyfit(X[2:], Y[2:], 1)
res = Y - (A_LS + B_LS * X)
ALPHA25 = -B_LS / 298.15 ** 2
DE = 2 * B_LS * 8.617e-5                            # B = ΔE/2k より [eV]
R25 = fitR(1 / 298.15)


# ================================================================ TA 資料の書き方に合わせたグラフ
def ta_chart(xt, yt, w=15, h=10):
    """Excel の既定の見た目の散布図（マーカーのみ）。"""
    ch = C.base_chart(xt, yt, w=w, h=h)
    ch.legend = None
    return ch


def ch_TR():
    ws = C.new_sheet("T-R")
    ws.append(["温度 T [℃]", "抵抗 R [Ω]"])
    for t, r in DATA:
        ws.append([t, r])
    ch = ta_chart("温度 T [℃]", "抵抗 R [Ω]")
    s = Series(Reference(ws, min_col=2, min_row=2, max_row=1 + n), Reference(ws, min_col=1, min_row=2, max_row=1 + n),
               title="測定値")
    s.marker = C.marker("circle", 5, C.BLUE)
    C.line(s, none=True)
    ch.series.append(s)
    ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = 20, 45, 5
    ch.y_axis.scaling.min, ch.y_axis.scaling.max, ch.y_axis.majorUnit = 14, 28, 2
    return C.native(ws, ch, "D2")


def ch_logR():
    """横軸は 1/T を 10⁻³ K⁻¹ 単位で表す（目盛を 3.20 のように読みやすくする）。"""
    ws = C.new_sheet("1T-R")
    ws.append(["1/T [×10^-3 K^-1]", "R [Ω]"])
    for x, r in zip(X, R):
        ws.append([float(x) * 1e3, float(r)])
    ch = ta_chart("温度の逆数 1/T [×10⁻³ K⁻¹]", "抵抗 R [Ω]")
    s = Series(Reference(ws, min_col=2, min_row=2, max_row=1 + n), Reference(ws, min_col=1, min_row=2, max_row=1 + n),
               title="測定値")
    s.marker = C.marker("circle", 5, C.BLUE)
    C.line(s, none=True)
    ch.series.append(s)
    ch.y_axis.scaling.logBase = 10
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 10, 100
    ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = 3.15, 3.40, 0.05
    ch.x_axis.number_format = "0.00"
    return C.native(ws, ch, "G2")


# ================================================================ 表（二段組，空欄は斜線）
def diag(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    d = OxmlElement("w:tl2br")
    for k, v in (("w:val", "single"), ("w:sz", "4"), ("w:space", "0"), ("w:color", "000000")):
        d.set(qn(k), v)
    b.append(d)
    tcPr.append(b)


def shade(cell, color="F2F2F2"):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), color)
    tcPr.append(sh)


def two_column_table(caption, num, headers, rows, widths):
    half = (len(rows) + 1) // 2
    left, right = rows[:half], rows[half:]
    gap = [""]
    hdr = headers + gap + headers
    body_rows = []
    for i in range(half):
        r = right[i] if i < len(right) else [None] * len(headers)
        body_rows.append(left[i] + gap + r)
    t = C.table(caption, num, hdr, [[("" if v is None else v) for v in r] for r in body_rows],
                widths + [4] + widths, size=9.5)
    k = len(headers)
    for j in range(len(hdr)):
        if j != k:
            shade(t.rows[0].cells[j])
    for i, r in enumerate(body_rows):
        for j, v in enumerate(r):
            if v is None:
                diag(t.rows[i + 1].cells[j])
    # 二段の間の列は罫線なし
    for row in t.rows:
        c = row.cells[k]
        tcPr = c._tc.get_or_add_tcPr()
        b = OxmlElement("w:tcBorders")
        for side in ("top", "bottom"):
            e = OxmlElement(f"w:{side}")
            e.set(qn("w:val"), "nil")
            b.append(e)
        tcPr.append(b)
    return t


# ================================================================ 文書
C.new_doc(os.path.join(HERE, "img_eq"))
C.cover(META, "サーミスタの実験")

heading("１　実験目的")
body("サーミスタは，温度によって電気抵抗が大きく変わる半導体の素子であり，その性質を利用して温度センサとして"
     "広く使われている．本実験の目的は次の2つである．")
item("・サーミスタの温度を変えて抵抗を測定し，その変化の仕方が白金などの普通の金属とはまったく異なることを理解する．")
item("・サーミスタの特性を表すサーミスタ定数 B を，測定結果から決定する．")

heading("２　原理")
heading("2.1　サーミスタの種類", 2)
f1 = nxt("fig")
body(f"サーミスタはもともと「熱に敏感な半導体」をまとめた呼び名である．現在は，鉄（Fe）・コバルト（Co）・"
     f"ニッケル（Ni）・マンガン（Mn）・銅（Cu）などの遷移金属の酸化物を焼き固めたセラミックの半導体が多く使われる．"
     f"温度による抵抗の変わり方から，次の3種類に分けられる（図{f1}）．")
item("・NTC サーミスタ（負の温度係数，Negative Temperature Coefficient）：温度が上がると抵抗がなめらかに小さくなる．"
     "一般に「サーミスタ」というとこれを指し，本実験で用いる D21A もこの種類である．", left=0)
item("・PTC サーミスタ（正の温度係数，ポジスタ）：ある温度を超えると抵抗が急に大きくなる．", left=0)
item("・CTR サーミスタ（クリテジスタ）：負の温度係数をもち，特定の温度で抵抗が急に小さくなる．", left=0)
body("これに対して，白金などの金属は温度が上がると抵抗がゆるやかに大きくなる（白金の抵抗温度係数は"
     "約 ＋0.0039 K⁻¹）．温度が上がると原子の熱振動が激しくなり，電子が散乱されやすくなるためである．")
C.figure(os.path.join(IMG, "fig_RT.png"), "各種材料における抵抗の温度特性（指導書[1]より）", f1, 70)
heading("2.2　NTC サーミスタの抵抗と温度の関係", 2)
body("サーミスタの電気伝導は，バンド理論だけでは説明できない複雑なものである．基本的には，同じ金属元素でも価数の"
     "異なるイオン（例えば Mn³⁺ と Mn⁴⁺）が結晶の中に混ざっており，その間を電子が飛び移る（ホッピングする）ことで"
     "電流が流れると考えられている．電子が飛び移るにはエネルギーの山を越える必要があり，温度が高いほど山を越えられる"
     "電子が増えるので，抵抗は小さくなる．このため，抵抗 R と絶対温度 T の間には近似的に")
e1 = eq(r"R=R_0\exp\left\{B\left(\dfrac{1}{T}-\dfrac{1}{T_0}\right)\right\}")
body("の関係がある．ここで R_{0}，R［Ω］はそれぞれ温度 T_{0}，T［K］での抵抗であり，B［K］をサーミスタ定数という．"
     "サーミスタの活性化エネルギーを ΔE，ボルツマン定数を k とすると B＝ΔE/2k であり，B の値は材料の組成や"
     "作り方によって決まる．")
body(f"式({e1})の両辺の自然対数をとると")
e2 = eq(r"\ln R=\left(\ln R_0-\dfrac{B}{T_0}\right)+B\cdot\dfrac{1}{T}")
body("となる．右辺の第1項は定数なので，A＝ln R_{0}－B/T_{0}，Y＝ln R，X＝1/T とおくと")
e3 = eq(r"Y=A+BX")
body("と X の1次式になる．したがって，横軸に 1/T，縦軸に R を対数目盛でとったグラフ（片対数グラフ）では測定点が"
     "直線に並び，その傾きからサーミスタ定数 B を求めることができる．また，温度による抵抗の変わりやすさを表す"
     f"抵抗温度係数は，式({e1})を T で微分して")
e_a = eq(r"\alpha=\dfrac{1}{R}\dfrac{dR}{dT}=-\dfrac{B}{T^2}")
body("となる．")
heading("2.3　サーミスタの応用", 2)
body("サーミスタは抵抗の温度係数が大きいことを利用して，体温計，エアコンや冷蔵庫の温度センサ，電池の温度監視などに"
     "広く使われている．また，気体や液体が触れるとその種類や量によって奪われる熱が変わり，サーミスタの温度が変わる"
     "ので，これを抵抗の変化として測ることで，ガスの分析，真空計，湿度計にも応用できる．さらに，電子回路の中で"
     "トランジスタなどの温度による特性の変化を打ち消す素子（バイアス安定化）としても使われている．")

heading("３　使用器具")
t_eq = nxt("tab")
C.table("使用器具", t_eq, ["器具", "台数"],
        [["電気炉", "1台"], ["ヒータ電源（直流電源）", "1台"], ["デジタルボルトメータ", "1台"], ["熱電対温度計", "1本"],
         ["サーミスタ（D21A）", "1個"]], [60, 25])
body("サーミスタ D21A の使用温度範囲は －70〜70 ℃，抵抗の標準値は 20 Ω，自己加熱による抵抗変化は 2 %以内，"
     "許容動作電流は 30 mA，許容動作電圧は 0.6 V である．")

heading("４　実験方法")
f2 = nxt("fig")
item(f"【１】　図{f2}のように装置を接続した．電気炉は，真鍮のパイプにシースヒータ H を巻き，水を入れた容器 C の中に"
     f"入れた構造になっている．使用温度の上限が 70 ℃なので，水を使用した．水が多すぎると温度が上がるのに時間が"
     f"かかり，少なすぎるとヒータが水面から出て危険なので，適量を入れた．ヒータの中のサーミスタ Th の端子に"
     f"デジタルボルトメータを接続し，熱電対温度計 T をサーミスタに接触させた．加熱電源は電源を切った状態でヒータに"
     f"接続し，誤った配線でサーミスタに電圧がかからないよう，電源を入れる前に TA に配線を確認してもらった．", left=0)
item("【２】　デジタルボルトメータを抵抗測定にして電源を入れ，温度計の指示が十分に安定するまで約5分待ってから，"
     "最初の温度と抵抗を記録した．次に，ヒータ電源のつまみが0の位置にあることを確かめてから電源を入れ，"
     "温度が毎分 1 ℃程度で上がるように電流を少しずつ増やした．加熱中は攪拌器 K を上下に動かして水温を"
     "均一にし，温度が 1 ℃上がるごとに抵抗を読み取った．", left=0)
C.figure(os.path.join(IMG, "fig_app.png"), "実験装置（指導書[1]より）", f2, 110)

heading("５　実験結果")
t1 = nxt("tab")
f3, f4 = nxt("fig"), nxt("fig")
body(f"温度 T と抵抗 R の測定結果を表{t1}に示す．温度は 21.7 ℃から 41.7 ℃まで 1 ℃ごとに，{n}点測定した．"
     f"また，温度と抵抗の関係を図{f3}に，温度の逆数 1/T（T は絶対温度）と抵抗の関係を片対数グラフにして図{f4}に示す．")
rows = [[f"{t:.1f}", f"{tk:.2f}", f"{x * 1e3:.4f}", f"{r:.3f}"] for t, tk, x, r in zip(TC, TK, X, R)]
two_column_table("温度と抵抗の関係", t1, ["温度 T [℃]", "温度 T [K]", "1/T [10^{-3} K^{-1}]", "抵抗 R [Ω]"],
                 rows, [18, 19, 20, 18])
C.chart_figure(ch_TR(), "温度 T と抵抗 R の関係", f3, 120)
body(f"図{f3}より，温度が上がると抵抗は小さくなり，{TC[0]:.1f} ℃の{R[0]:.2f} Ω から{TC[-1]:.1f} ℃の{R[-1]:.2f} Ω まで，"
     f"20 ℃の上昇で約{(1 - R[-1] / R[0]) * 100:.0f} %も減少した．グラフはわずかに下に凸の曲線で，温度が高いほど"
     f"1 ℃あたりの抵抗の減り方は小さくなっている．これは式({e1})のように抵抗が温度に対して指数関数的に変化する"
     f"ためである．温度が上がると抵抗が下がるという，白金などの金属とは逆の性質（負の温度係数）がはっきり確かめられた．")
C.chart_figure(ch_logR(), "温度の逆数 1/T と抵抗 R の関係（片対数グラフ）", f4, 120)
body(f"図{f4}より，片対数グラフでは測定点がほぼ1本の直線上に並んだ（ln R と 1/T の決定係数 R²＝{R2:.4f}）．"
     f"したがって，サーミスタの抵抗は式({e1})に従い，ln R は 1/T に比例して変化することが確かめられた．"
     f"直線の傾きが正であるのは，温度が高い（1/T が小さい）ほど抵抗が小さいことに対応している．")

heading("６　考察")
heading("6.1　近似式によるサーミスタ定数 B の決定（考察課題①）", 2)
body(f"まず，近似式を導く．直線上の2点 (R_{{1}}, 1/T_{{1}})，(R_{{2}}, 1/T_{{2}}) について式({e1})を書くと")
eq(r"R_1=R_0\exp\left\{B\left(\dfrac{1}{T_1}-\dfrac{1}{T_0}\right)\right\},\qquad"
   r"R_2=R_0\exp\left\{B\left(\dfrac{1}{T_2}-\dfrac{1}{T_0}\right)\right\}", numbered=False)
body("である．辺々割って R_{0} と T_{0} を消去すると")
eq(r"\dfrac{R_1}{R_2}=\exp\left\{B\left(\dfrac{1}{T_1}-\dfrac{1}{T_2}\right)\right\}", numbered=False)
body("となり，両辺の自然対数をとって B について解くと")
eq(r"B=\dfrac{\ln R_1-\ln R_2}{1/T_1-1/T_2}", numbered=False)
body("を得る．常用対数との関係 ln x＝(ln 10)·log x≒2.303 log x を用いると")
e_2p = eq(r"B=\dfrac{2.3\,(\log R_1-\log R_2)}{1/T_1-1/T_2}")
body(f"となり，指導書の近似式が導かれた．係数 2.3 は ln 10＝2.3026… を丸めた値である．")
body(f"次に，図{f4}の測定点の並びに沿う直線を考え，その直線上から2点を読み取った．なるべく離れた2点を選ぶと読み取りの誤差の影響が小さくなるので，"
     f"1/T＝{P1X * 1e3:.2f}×10⁻³ K⁻¹ と{P2X * 1e3:.2f}×10⁻³ K⁻¹ の点をとると，抵抗はそれぞれ R_{{1}}＝{P1R} Ω，"
     f"R_{{2}}＝{P2R} Ω であった．これを式({e_2p})に代入すると")
eq(rf"B=\dfrac{{2.3\,(\log {P1R}-\log {P2R})}}{{{P1X * 1e3:.2f}\times10^{{-3}}-{P2X * 1e3:.2f}\times10^{{-3}}}}"
   rf"=\dfrac{{2.3\times({math.log10(P1R):.4f}-{math.log10(P2R):.4f})}}{{{(P1X - P2X) * 1e3:.2f}\times10^{{-3}}}}"
   rf"\simeq{B_2P:.0f}\ \mathrm{{K}}", numbered=False)
body(f"となった．")
heading("6.2　最小二乗法によるサーミスタ定数 B の決定（考察課題②）", 2)
body(f"X＝1/T，Y＝ln R として，指導書の補足の式(5)")
e_ls = eq(r"B=\dfrac{\sum_{i=1}^{n}X_iY_i-n\overline{X}\cdot\overline{Y}}{n\left\{\overline{X^2}-\left(\overline{X}\right)^2\right\}}")
body(f"により，{n}点（n＝{n}）すべてを用いて B を計算した．各量は")
eq(rf"\overline{{X}}={Xm * 1e3:.5f}\times10^{{-3}}\ \mathrm{{K^{{-1}}}},\quad \overline{{Y}}={Ym:.5f},\quad"
   rf"\overline{{X^2}}={X2m * 1e6:.6f}\times10^{{-6}}\ \mathrm{{K^{{-2}}}},\quad \sum X_iY_i={(X * Y).sum():.6f}\ \mathrm{{K^{{-1}}}}",
   numbered=False, size=11)
body("であり，これを代入すると")
eq(rf"B={B_LS:.0f}\ \mathrm{{K}},\qquad A=\overline{{Y}}-B\overline{{X}}={A_LS:.3f}", numbered=False)
body("となった．")
heading("6.3　2つの求め方の比較検討（考察課題③）", 2)
t2 = nxt("tab")
body(f"6.1と6.2で求めた B を表{t2}にまとめる．")
C.table("サーミスタ定数 B の比較", t2, ["求め方", "サーミスタ定数 B [K]"],
        [["近似式（グラフの直線上の2点）", f"{B_2P:.0f}"], ["最小二乗法（全21点）", f"{B_LS:.0f}"],
         ["（参考）測定値の両端の2点をそのまま使用", f"{B_END:.0f}"]], [80, 45])
body(f"近似式で求めた B は{B_2P:.0f} K，最小二乗法で求めた B は{B_LS:.0f} K で，その差は{abs(diffp):.1f} %と"
     f"ほぼ一致した．この差の主な原因は，グラフから抵抗を3桁（0.1 Ω 単位）で読み取ったことである．例えば R_{{1}} を"
     f" 0.1 Ω 読み違えると，B は約{0.1 / P1R / (P2X - P1X):.0f} K 変わる．なお，ln 10＝2.3026 を 2.3 に丸めた"
     f"ことによる差は{B_2P_exact - B_2P:.0f} K（約0.1 %）であり，ほとんど影響しない．")
body("2つの求め方を比べると，近似式による方法は計算が簡単で，グラフさえ描けばすぐに B の見当がつくという利点が"
     "ある．しかし，どの直線を引くか，どの2点を読むかは人によって異なり，結果に個人差が出る．特に，直線を引かずに"
     f"測定点そのものを2点選ぶと，その2点の測定誤差がそのまま B に入ってしまう．実際に，測定値の両端の2点"
     f"（{TC[0]:.1f} ℃と{TC[-1]:.1f} ℃）をそのまま使うと B＝{B_END:.0f} K となり，最小二乗法の値より"
     f"{abs(B_END - B_LS) / B_LS * 100:.0f} %小さくなった．これは，最初の2点（{TC[0]:.1f} ℃と{TC[1]:.1f} ℃）が"
     f"直線から下にずれていたためである．")
body("一方，最小二乗法は，全ての測定点からの残差の2乗和が最小になるように直線を決めるため，誰が計算しても同じ値に"
     "なり，個々の点の誤差が平均されて結果への影響が小さくなる．したがって，B を決める方法としては最小二乗法の方が"
     "信頼性が高いと考えられる．ただし，最小二乗法も外れた点の影響は受けるので，グラフで測定点と直線を見比べ，"
     f"外れた点がないか確かめることが大切である．試しに直線から外れていた最初の2点を除いて計算すると，"
     f"B＝{Bx:.0f} K となり，全点の場合と{abs(Bx - B_LS) / B_LS * 100:.1f} %の差であった．")
body(f"最初の2点が直線から外れた原因としては，加熱を始めた直後で水温がまだ均一になっておらず，熱電対の示す温度と"
     f"サーミスタの実際の温度にずれがあったことが考えられる．温度を上げながら測定しているので，温度計とサーミスタの"
     f"熱の伝わり方の違い（応答の遅れ）も誤差の原因になる．")
heading("6.4　最小二乗法による(4)式の導出（考察課題④）", 2)
body(f"指導書の(3)式（本報告書の式({e3})）Y＝A＋BX に対して，残差の2乗和")
eq(r"S(A,B)=\sum_{i=1}^{n}\left(Y_i-A-BX_i\right)^2", numbered=False)
body("を最小にする A，B を求める．S を A，B で偏微分して0とおくと")
eq(r"\dfrac{\partial S}{\partial A}=-2\sum\left(Y_i-A-BX_i\right)=0,\qquad"
   r"\dfrac{\partial S}{\partial B}=-2\sum X_i\left(Y_i-A-BX_i\right)=0", numbered=False)
body("第1式を n で割ると Ȳ－A－BX̄＝0 より A＝Ȳ－BX̄ である．これを第2式に代入すると")
eq(r"\sum X_i\left\{\left(Y_i-\overline{Y}\right)-B\left(X_i-\overline{X}\right)\right\}=0", numbered=False)
body("となる．ここで Σ(Y_{i}－Ȳ)＝0，Σ(X_{i}－X̄)＝0 なので，X_{i} を (X_{i}－X̄) に置き換えても和は変わらず")
eq(r"\sum\left(X_i-\overline{X}\right)\left(Y_i-\overline{Y}\right)=B\sum\left(X_i-\overline{X}\right)^2", numbered=False)
body("よって")
eq(r"B=\dfrac{\sum_{i=1}^{n}\left(X_i-\overline{X}\right)\left(Y_i-\overline{Y}\right)}{\sum_{i=1}^{n}\left(X_i-\overline{X}\right)^2}",
   numbered=False)
body("が得られ，指導書の(4)式が導かれた．この分子と分母を展開すると，補足に示されているように(5)式になる．")
heading("6.5　サーミスタと金属の比較", 2)
body(f"最小二乗法で求めた B を用いて，25 ℃（298.15 K）での抵抗温度係数を式({e_a})から計算すると")
eq(rf"\alpha=-\dfrac{{{B_LS:.0f}}}{{298.15^2}}\simeq{ALPHA25:.4f}\ \mathrm{{K^{{-1}}}}", numbered=False)
body(f"となり，温度が 1 ℃上がると抵抗が約{abs(ALPHA25) * 100:.1f} %減ることになる．白金の抵抗温度係数"
     f"（約 ＋0.0039 K⁻¹）と比べると，符号が逆で，大きさは約{abs(ALPHA25) / 0.0039:.0f}倍である．このように"
     f"小さな温度変化でも抵抗が大きく変わることが，サーミスタが温度センサとして優れている理由である．"
     f"また，B＝ΔE/2k より活性化エネルギーは ΔE＝2kB≒{DE:.2f} eV となった．")
body(f"なお，求めた式から 25 ℃での抵抗を計算すると約{R25:.1f} Ω となり，D21A の標準値 20 Ω より大きかった．"
     f"抵抗の標準値は決められた基準温度での値であり，その温度が 25 ℃と異なることや，測定に使ったリード線や端子の"
     f"抵抗がサーミスタの抵抗に加わって測定されていることが考えられる．")

heading("７　まとめ")
body(f"サーミスタの温度を 21.7 ℃から 41.7 ℃まで上げながら抵抗を測定し，温度が上がると抵抗が指数関数的に"
     f"小さくなることを確かめた．片対数グラフでは 1/T と R が直線関係になり，その傾きから求めたサーミスタ定数は，"
     f"近似式で{B_2P:.0f} K，最小二乗法で{B_LS:.0f} K となり，ほぼ一致した．25 ℃での抵抗温度係数は約"
     f"{ALPHA25:.3f} K⁻¹ で，白金とは符号が逆で大きさが約{abs(ALPHA25) / 0.0039:.0f}倍であった．"
     f"B の決定には，すべての測定点を使い個人差が出ない最小二乗法の方が適していると考えられる．")

C.references([
    "[1] 東京理科大学 工学部 電気工学科，“物理学実験 14．サーミスタの実験”，2026，pp.96–100．",
    "[2] “NTCサーミスタとは？動作原理”，株式会社村田製作所，"
    "https://www.murata.com/ja-jp/products/thermistor/ntc/overview/basic/about，閲覧日：2026年11月12日．",
])

# 表がページをまたがないようにする
for t in C.doc.tables:
    for r in t.rows[:-1]:
        for c in r.cells:
            for p in c.paragraphs:
                p.paragraph_format.keep_with_next = True

out = os.path.join(HERE, "サーミスタの実験_報告書.docx")
C.save(out)
print("saved", out)
print("B_LS", B_LS, "A", A_LS, "R2", R2, "2pt", P1R, P2R, B_2P, B_2P_exact, "end", B_END, "excl", Bx,
      "alpha", ALPHA25, "dE", DE, "R25", R25)
