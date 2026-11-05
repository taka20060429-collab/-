# -*- coding: utf-8 -*-
"""12. 光センサの実験 ＋ 13. ホール素子の実験 — 報告書（Word）とデータ用 Excel の生成。 python3 build.py"""
import io
import math
import os
import sys

import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.chart import Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.trendline import Trendline
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "report_optics"))
import common as C  # noqa: E402
from common import body, eq, heading, item, nxt, table  # noqa: E402

META = {"kai": None, "date": None, "submit": None, "han": "D班", "name": "森木崇允", "grade": "1",
        "student_id": "4326091",
        "coworkers": [("石井真人", "4326007"), ("木元喜紹", "4326029"), ("陶山稜平", "4326051"),
                      ("林田知子", "4326071")]}
IMG = os.path.join(HERE, "img")
fig = lambda name: os.path.join(IMG, name)  # noqa: E731

wb_in = openpyxl.load_workbook(os.path.join(HERE, "data", "sensor.xlsx"), data_only=True)
s1, s2 = wb_in["Sheet1"], wb_in["Sheet2"]
col = lambda ws, c, r0, r1: [ws[f"{c}{r}"].value for r in range(r0, r1 + 1)]  # noqa: E731

# ================================================================ 光センサのデータ
LV = col(s2, "B", 4, 19)                     # 光源電圧 [V]
LI = col(s2, "C", 4, 19)                     # 光源電流 [mA]
LE = col(s2, "E", 4, 19)                     # 照度 [lx]
LID = col(s2, "I", 4, 19)                    # ダイオード電流 [mA]
LW = [v * i / 1000 for v, i in zip(LV, LI)]  # 光源電力 [W]
DX = col(s2, "L", 4, 24)                     # 移動距離 [cm]
DID = col(s2, "M", 4, 24)                    # ダイオード電流 [mA]
VD = col(s2, "Q", 5, 15)                     # 順方向 Vd [V]
F_ON = col(s2, "R", 5, 15)                   # 順方向・照射時 Id [mA]
F_OFF = col(s2, "W", 5, 15)                  # 順方向・無照射時
R_VD = [-0.1, -0.5, -1.0]
R_ON = col(s2, "T", 5, 7)                    # 逆方向・照射時
R_OFF = col(s2, "Y", 5, 7)                   # 逆方向・無照射時

# 照度 E とダイオード電流 Id の関係：2 次の多項式で近似（実験(2)(iv)）
EC = np.polyfit(LID, LE, 2)
E_of = lambda i: float(np.polyval(EC, i))  # noqa: E731
DE = [E_of(i) for i in DID]
DY = [1 / math.sqrt(e) for e in DE]
xa, ya = np.array(DX, float), np.array(DY)
n_ = len(xa)
K_ = (n_ * (xa @ ya) - xa.sum() * ya.sum()) / (n_ * (xa @ xa) - xa.sum() ** 2)
M_ = ((xa @ xa) * ya.sum() - (xa @ ya) * xa.sum()) / (n_ * (xa @ xa) - xa.sum() ** 2)
A_ = 1 / K_ ** 2
R0 = M_ / K_
R2_E = np.corrcoef(xa, ya)[0, 1] ** 2
yi = 1 / np.sqrt(np.array(DID))
Ki, Mi = np.polyfit(xa, yi, 1)
R0_i = Mi / Ki
R2_i = np.corrcoef(xa, yi)[0, 1] ** 2
RATIO = [(e / i) for e, i in zip(LE, LID) if i > 0 and e > 0]

# ================================================================ ホール素子のデータ
IX0, VH0 = s1["B4"].value, s1["B5"].value    # [mA], [mV]
H0 = 1.67e4
KH = H0 / VH0                                 # [A/m per mV]
CI = col(s1, "D", 5, 13)                      # コイル電流 [A]
CV = col(s1, "E", 5, 13)                      # Vh [mV]
CH = [KH * v for v in CV]
AH = sum(i * h for i, h in zip(CI, CH)) / sum(i * i for i in CI)
PX = col(s1, "L", 5, 23)                      # x [cm]
PV1 = col(s1, "M", 5, 23)
PV4 = col(s1, "P", 5, 23)
PH1 = [KH * v for v in PV1]
PH4 = [KH * v for v in PV4]
CENTER = s1["B12"].value
SIX = col(s1, "T", 4, 11)
SVH = col(s1, "U", 4, 11)
SK = sum(i * v for i, v in zip(SIX, SVH)) / sum(i * i for i in SIX)
H2A = AH * 2.0


def half_width(xs, hs):
    hmax = max(hs)
    left = right = None
    for (x1, h1), (x2, h2) in zip(zip(xs, hs), zip(xs[1:], hs[1:])):
        if h1 < hmax / 2 <= h2 and left is None:
            left = x1 + (hmax / 2 - h1) * (x2 - x1) / (h2 - h1)
        if h1 >= hmax / 2 > h2:
            right = x1 + (h1 - hmax / 2) * (x2 - x1) / (h1 - h2)
    return left, right


def plateau(xs, hs, frac=0.95):
    hmax = max(hs)
    inside = [x for x, h in zip(xs, hs) if h >= frac * hmax]
    return min(inside), max(inside)


# ================================================================ Excel グラフ
def scatter(ws, cols, names, xt, yt, colors=None, lines=False, trend=None, h=9, legend=True, syms=None):
    ch = C.base_chart(xt, yt, h=h)
    colors = colors or [C.BLUE, C.ORANGE, C.GRAY, C.YELLOW]
    syms = syms or ["circle", "square", "triangle", "diamond"]
    for k, ((xc, yc, r0, r1), name) in enumerate(zip(cols, names)):
        s = Series(Reference(ws, min_col=yc, min_row=r0, max_row=r1),
                   Reference(ws, min_col=xc, min_row=r0, max_row=r1), title=name)
        s.marker = C.marker(syms[k], 6, colors[k])
        if lines:
            C.line(s, colors[k], width=12700)
        else:
            C.line(s, none=True)
        if trend and trend[k]:
            s.trendline = trend[k]
        ch.series.append(s)
    if not legend:
        ch.legend = None
    return ch


def sheet(rows, title="data"):
    ws = C.new_sheet(title)
    for r in rows:
        ws.append(list(r))
    return ws


def ch_WE():
    ws = sheet([("W", "E")] + list(zip(LW, LE)))
    ch = scatter(ws, [(1, 2, 2, 17)], ["照度"], "光源電力 W [W]", "照度 E [lx]", lines=True, legend=False)
    ch.x_axis.scaling.min = 0
    ch.y_axis.scaling.min = 0
    return C.native(ws, ch, "E2")


def ch_WId():
    ws = sheet([("W", "Id")] + list(zip(LW, LID)))
    ch = scatter(ws, [(1, 2, 2, 17)], ["Id"], "光源電力 W [W]", "ダイオード電流 Id [mA]", lines=True, legend=False)
    ch.x_axis.scaling.min = 0
    return C.native(ws, ch, "E2")


def ch_EId():
    ws = sheet([("E", "Id")] + list(zip(LE, LID)))
    ch = scatter(ws, [(1, 2, 2, 17)], ["測定値"], "照度 E [lx]", "ダイオード電流 Id [mA]", lines=False, legend=False)
    ch.x_axis.scaling.min = 0
    return C.native(ws, ch, "E2")


def ch_IdE():
    ws = sheet([("Id", "E")] + list(zip(LID, LE)))
    ch = scatter(ws, [(1, 2, 2, 17)], ["測定値"], "ダイオード電流 Id [mA]", "照度 E [lx]",
                 trend=[Trendline(trendlineType="poly", order=2, dispEq=False, dispRSqr=False)])
    ch.y_axis.scaling.min = 0
    return C.native(ws, ch, "E2")


def ch_xE():
    ws = sheet([("x", "E")] + list(zip(DX, DE)))
    ch = scatter(ws, [(1, 2, 2, 22)], ["照度"], "移動距離 x [cm]", "照度 E [lx]", lines=True, legend=False)
    ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = 0, 10, 1
    ch.y_axis.scaling.min = 0
    return C.native(ws, ch, "E2")


def ch_xy():
    rows = [("x", "y", "", "x", "fit")] + [(x, y, None, None, None) for x, y in zip(DX, DY)]
    for j, x in enumerate((0, 10)):
        rows[1 + j] = rows[1 + j][:3] + (x, K_ * x + M_)
    ws = sheet(rows)
    ch = scatter(ws, [(1, 2, 2, 22), (4, 5, 2, 3)], ["測定値", f"実験式 y＝{K_:.4f}x＋{M_:.4f}"],
                 "移動距離 x [cm]", "1/√E [lx^-1/2]")
    ch.series[1].marker = Marker(symbol="none")
    C.line(ch.series[1], C.ORANGE, width=19050)
    ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = 0, 10, 1
    ch.y_axis.scaling.min = 0
    return C.native(ws, ch, "G2")


def ch_IV():
    rows = [("Vd", "照射", "無照射")]
    for v, a, b in zip(VD, F_ON, F_OFF):
        rows.append((v, a, b))
    for v, a, b in zip(R_VD, R_ON, R_OFF):
        rows.append((v, a, b))
    rows.sort(key=lambda r: r[0] if isinstance(r[0], (int, float)) else -99)
    ws = sheet(rows)
    n = len(rows)
    ch = scatter(ws, [(1, 2, 2, n), (1, 3, 2, n)], ["照射時（光源電力 4 W）", "無照射時"],
                 "ダイオード電圧 Vd [V]", "ダイオード電流 Id [mA]", lines=True)
    for ax in (ch.x_axis, ch.y_axis):
        ax.tickLblPos = "low"
    return C.native(ws, ch, "E2")


def ch_IH():
    rows = [("I", "Vh", "H", "", "I", "fit")] + [(i, v, h, None, None, None) for i, v, h in zip(CI, CV, CH)]
    for j, i in enumerate((0, 4)):
        rows[1 + j] = rows[1 + j][:4] + (i, AH * i)
    ws = sheet(rows)
    ch = scatter(ws, [(1, 3, 2, 10), (5, 6, 2, 3)], ["測定値", f"実験式 H＝{AH:.0f}I"],
                 "コイル電流 I [A]", "磁場 H [A/m]")
    ch.series[1].marker = Marker(symbol="none")
    C.line(ch.series[1], C.ORANGE, width=19050)
    ch.x_axis.scaling.min, ch.x_axis.scaling.max = 0, 4
    ch.y_axis.scaling.min = 0
    return C.native(ws, ch, "H2")


def ch_IVh():
    ws = sheet([("I", "Vh")] + list(zip(CI, CV)))
    ch = scatter(ws, [(1, 2, 2, 10)], ["Vh"], "コイル電流 I [A]", "ホール電圧 Vh [mV]", lines=True, legend=False)
    ch.x_axis.scaling.min, ch.x_axis.scaling.max = 0, 4
    ch.y_axis.scaling.min = 0
    return C.native(ws, ch, "E2")


def ch_xH():
    ws = sheet([("x", "H1", "H4")] + list(zip(PX, PH1, PH4)))
    ch = scatter(ws, [(1, 3, 2, 20), (1, 2, 2, 20)], ["コイル電流 4 A", "コイル電流 1 A"],
                 "コイル中心からの距離 x [cm]", "磁場 H [A/m]", lines=True)
    ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = -20, 20, 4
    ch.y_axis.scaling.min = 0
    for ax in (ch.x_axis, ch.y_axis):
        ax.tickLblPos = "low"
    return C.native(ws, ch, "E2")


def ch_ixVh():
    rows = [("ix", "Vh", "", "ix", "fit")] + [(i, v, None, None, None) for i, v in zip(SIX, SVH)]
    for j, i in enumerate((0, 8)):
        rows[1 + j] = rows[1 + j][:3] + (i, SK * i)
    ws = sheet(rows)
    ch = scatter(ws, [(1, 2, 2, 9), (4, 5, 2, 3)],
                 [f"測定値（I＝2 A，H＝{H2A:.3g} A/m）", f"比例直線 Vh＝{SK:.3f}ix"],
                 "試料電流 ix [mA]", "ホール電圧 Vh [mV]")
    ch.series[1].marker = Marker(symbol="none")
    C.line(ch.series[1], C.ORANGE, width=19050)
    ch.x_axis.scaling.min, ch.x_axis.scaling.max = 0, 8
    ch.y_axis.scaling.min = 0
    return C.native(ws, ch, "G2")


# ================================================================ 文書
C.new_doc(os.path.join(HERE, "img_eq"))
C.cover(META, "光センサの実験\nホール素子の実験")


def title(t):
    C.para(t, size=14, bold=True, align=C.CENTER, ea=C.HEAD, latin=C.HEAD, sa=10)


# ################################################################ 光センサ
C.set_part("")
title("光センサの実験")
heading("1．目的")
body("光を電気信号に変える素子であるフォトダイオードに光を当てる実験を行い，フォトダイオードの性質を理解する．"
     "また，光源からの距離と照度の関係を調べ，最小二乗法などのデータ処理の方法についても考える．")

heading("2．原理")
heading("2.1　光と電気の変換", 2)
body("光を使って情報を伝えるには，光を出す部分と光を受ける部分が必要であり，現在その多くに半導体の固体素子が"
     "使われている．光を電気に変える現象には，光が当たると起電力が生じる光起電力効果，光の量に応じて電気の"
     "流れやすさが変わる光伝導効果，光のエネルギーで電子が物質の外に飛び出す光電子放出がある．光起電力効果を"
     "使う素子には本実験のフォトダイオードのほか，フォトトランジスタや太陽電池があり，光伝導効果を使う素子には"
     "CdS セルなどがある．")
heading("2.2　半導体のエネルギーバンド", 2)
f_band = nxt("fig")
body(f"物質中の電子がとれるエネルギーは，電子が存在できる許容帯と，存在できない禁制帯（バンドギャップ）に"
     f"分かれている（図{f_band}）．電子はエネルギーの低い許容帯から順に詰まっていく．金属では許容帯の途中まで"
     f"電子が詰まっているため，少しのエネルギーで電子が動けるようになり電流がよく流れる．絶縁体では，電子で"
     f"埋まった価電子帯と空の伝導帯の間に大きなバンドギャップがあり，電子は伝導帯に上がれないので電流は流れない．"
     f"半導体ではバンドギャップが小さいため，熱などのエネルギーで電子が伝導帯に上がり（励起），伝導帯の自由電子と"
     f"価電子帯に残った正孔（電子の抜けた穴，正の電荷として振る舞う）がわずかな電流を運ぶ．")
C.figure(fig("ps_band.png"), "金属，（真性）半導体，絶縁体のエネルギーバンド図（指導書[1]より）", f_band, 110)
f_pn = nxt("fig")
body(f"4価のシリコン（Si）に5価のリン（P）を加えると電子が1個余り，これが自由電子となる n 型半導体ができる．"
     f"3価のアルミニウム（Al）やボロン（B）を加えると電子が1個足りず，正孔が電流を運ぶ p 型半導体になる．"
     f"このとき，バンドギャップ内にそれぞれドナー準位，アクセプタ準位が生じる（図{f_pn}(a)）．")
heading("2.3　pn 接合とフォトダイオード", 2)
body(f"p 型と n 型を接合すると，電子は n 型から p 型へ，正孔は p 型から n 型へ拡散する．接合部には n 型側に"
     f"正のドナーイオン，p 型側に負のアクセプタイオンが残り，n 型から p 型へ向かう電界が生じて拡散が止まる．"
     f"このときの電位差を拡散電位 V_{{D}} といい，エネルギーバンドは図{f_pn}(b)のように傾く．p 型に正の電圧を"
     f"かけると障壁が低くなって電流が流れ（順方向），逆向きにかけると障壁が高くなって電流はほとんど流れない"
     f"（逆方向）．これがダイオードの整流性である．")
C.figure(fig("ps_pn.png"), "(a) p 型半導体と n 型半導体，(b) pn 接合のエネルギーバンド図（指導書[1]より）", f_pn, 120)
f_pd = nxt("fig")
body(f"フォトダイオードも pn 接合ダイオードであるが，接合部にバンドギャップ以上のエネルギー hν をもつ光が"
     f"当たると電子・正孔対が生まれ，電子は n 型へ，正孔は p 型へ移動して電流（光電流 i_{{C}}）が生じる．"
     f"図{f_pd}のように電圧をかけ，接合部の電圧を V，流れる電流を i とすると，光がないときは")
e_d = eq(r"i=i_{\mathrm{s}}\left\{\exp\left(\dfrac{eV}{k_{\mathrm{B}}T}\right)-1\right\}")
body("であり，光が当たると光電流 i_{C} が n 型から p 型へ流れるので")
e_p = eq(r"i=i_{\mathrm{s}}\left\{\exp\left(\dfrac{eV}{k_{\mathrm{B}}T}\right)-1\right\}-i_{\mathrm{C}}")
body("となる．ここで i_{s} は逆方向飽和電流，e は電気素量，k_{B} はボルツマン定数，T は絶対温度である．"
     "電流を流さないとき（i＝0）の開放電圧と，両端を短絡したとき（V＝0）の短絡電流は")
e_oc = eq(r"V_{\mathrm{OC}}=\dfrac{k_{\mathrm{B}}T}{e}\ln\left(1+\dfrac{i_{\mathrm{C}}}{i_{\mathrm{s}}}\right),"
          r"\qquad i_{\mathrm{SC}}=-i_{\mathrm{C}}")
body("となる．すなわち，開放電圧は光の量の対数に比例し，短絡電流は光の量に比例する．本実験ではフォトダイオードの"
     "両端に電流計をつなぎ，ほぼ短絡電流を測るので，ダイオード電流は光の量に比例すると期待される．")
C.figure(fig("ps_pd.png"), "フォトダイオード（指導書[1]より）", f_pd, 70)
heading("2.4　光の量を表す単位", 2)
f_vis = nxt("fig")
body(f"光が1秒間に運ぶエネルギーを放射束［W］という．しかし人の目の感じ方は波長によって異なり，図{f_vis}の"
     f"比視感度曲線のように 555 nm（黄緑）付近で最も敏感である．目の感じ方で重みをつけた光の量を光束"
     f"［lm（ルーメン）］といい，波長 555 nm の光 1 W は 680 lm に相当する．")
C.figure(fig("ps_vis.png"), "比視感度曲線（指導書[1]より）", f_vis, 80)
body("光源からある方向の微小な立体角 dΩ に出る光束を dF とすると，I＝dF/dΩ をその方向の光度［cd（カンデラ）］"
     "という．また，面に当たる単位面積あたりの光束 E＝dF/dS を照度［lx（ルクス）＝lm/m²］という．光度 I の"
     "点光源から距離 r の位置では dΩ＝dS/r² なので")
e_inv = eq(r"E=\dfrac{I}{r^2}")
body("となり，照度は距離の2乗に反比例する（逆2乗則）．")
heading("2.5　距離と照度の関係の直線化と最小二乗法", 2)
f_pos = nxt("fig")
body(f"図{f_pos}のように，光源を最も近づけたときの光源と素子の距離を r_{{0}}，そこからの移動距離を x とすると，"
     f"素子までの距離は x＋r_{{0}} であり，a を定数として")
e_E = eq(r"E=\dfrac{a}{(x+r_0)^2}")
body("となる．これを変形すると")
e_lin = eq(r"y=\dfrac{1}{\sqrt{E}}=\dfrac{1}{\sqrt{a}}x+\dfrac{r_0}{\sqrt{a}}=kx+m")
body("となり，y＝1/√E は x の1次関数になる．l 組のデータ (x_{i}, y_{i}) から，残差の2乗和 "
     "S＝Σ{y_{i}－(kx_{i}＋m)}² を最小にする k，m を求めると（最小二乗法）")
e_ls = eq(r"k=\dfrac{l\sum x_iy_i-\sum x_i\sum y_i}{l\sum x_i^2-\left(\sum x_i\right)^2},\qquad"
          r"m=\dfrac{\sum x_i^2\sum y_i-\sum x_iy_i\sum x_i}{l\sum x_i^2-\left(\sum x_i\right)^2}")
body("となり，a＝1/k²，r_{0}＝m/k が求まる．")
C.figure(fig("ps_pos.png"), "光源と受光素子の位置（指導書[1]より）", f_pos, 105)

heading("3．使用器具")
body("光源および光検出装置（光源：定格 12 V 5 W のハロゲンランプ，前方に直径 5 mm のスリット）1台，照度計1台，"
     "直流電源（18 V 1 A）2台，フォトダイオード（Si）1個，直流電圧計・電流計パネル（500 mA，20 V）1台，"
     "デジタルマルチメータ2台．")

heading("4．実験方法")
heading("(1)　光源電力 W に対する照度 E の測定", 2)
f_c1 = nxt("fig")
body(f"光源を最も近い位置（目盛 300）に置き，図{f_c1}のように直流電源・電圧計・電流計をつないだ．照度計の"
     f"受光部の中心をスリットの中心にぴったり合わせ，測定中は動かさないようにした．光源電圧 V を 15 V から"
     f"1 V ずつ 0 V まで変え，光源電流 I と照度 E を測定し，光源電力 W＝VI を計算した．")
C.figure(fig("ps_c1.png"), "光源電力の変化に対する照度の測定回路（指導書[1]より）", f_c1, 105)
heading("(2)　光源電力 W に対するフォトダイオード電流 Id の測定", 2)
f_c2 = nxt("fig")
body(f"図{f_c2}のように照度計の代わりにフォトダイオードを入れ，その両端に電流計をつないだ．光源電圧を 0 V から"
     f"15 V まで 1 V ずつ変え，光源電流 I とダイオード電流 Id を測定した．(1) の結果と合わせて，照度 E と"
     f"ダイオード電流 Id の関係を求めた．")
C.figure(fig("ps_c2.png"), "照度の変化に対するフォトダイオード電流の測定回路（指導書[1]より）", f_c2, 105)
heading("(3)　移動距離 x に対する照度 E の測定", 2)
body("光源電圧を 12 V（光源電力約 4 W）とし，光源を目盛 300 から 0.5 cm ずつ遠ざけて，目盛 200（10 cm）まで"
     "各位置でダイオード電流 Id を測定した．Id を (2) で求めた関係を使って照度 E に換算し，y＝1/√E と x の"
     "関係を調べ，最小二乗法で a，r_{0} を求めた．")
heading("(4)　ダイオード電流 Id－ダイオード電圧 Vd 特性の測定", 2)
f_c3 = nxt("fig")
body(f"図{f_c3}のようにフォトダイオードに直流電源・電流計・電圧計をつなぎ，光源を目盛 300 に置いた．"
     f"光を当てないとき（無照射時）と光源電力 4 W で光を当てたとき（照射時）について，順方向にダイオード電圧 Vd を"
     f"0 V から 0.1 V ずつ 1.0 V まで変えてダイオード電流 Id を測定した．次に電源の向きを逆にして，逆方向にも"
     f"同様に測定した．電圧が 0.5 V を超えると電流が急に増えるので，微調整つまみで少しずつ電圧を上げた．")
C.figure(fig("ps_c3.png"), "ダイオード電流－ダイオード電圧特性の測定回路（指導書[1]より）", f_c3, 105)

heading("5．実験結果")
heading("(1)　光源電力 W に対する照度 E", 2)
t1 = nxt("tab")
fW = nxt("fig")
body(f"光源電圧・電流・電力と照度を表{t1}に，光源電力と照度の関係を図{fW}に示す．")
table("光源電力 W と照度 E", t1, ["光源電圧 V [V]", "光源電流 I [mA]", "光源電力 W [W]", "照度 E [lx]"],
      [[f"{v}", f"{i}", f"{w:.3f}", f"{e}"] for v, i, w, e in zip(LV, LI, LW, LE)], [32, 32, 32, 32], size=9)
C.chart_figure(ch_WE(), "光源電力 W と照度 E の関係", fW, 120)
body(f"図{fW}より，光源電力が約 0.5 W 以下では照度はほぼ0で，それを超えると照度は電力とともに急に大きくなり，"
     f"比例ではなく下に凸の曲線になった．ハロゲンランプは電力が小さいとフィラメントの温度が低く，出る光の"
     f"ほとんどが目に見えない赤外線になる．電力を上げてフィラメントの温度が上がると，目に見える光の割合が急に"
     f"増えるため，照度は電力に比例するよりも速く増えたと考えられる．なお，光源電圧 9 V から 10 V の間で照度が"
     f"{LE[9]} lx から{LE[10]} lx へと約{LE[10] / LE[9]:.1f}倍に跳ね上がっており，前後の増え方と比べて不自然に大きい．")
heading("(2)　照度 E に対するフォトダイオード電流 Id", 2)
t2 = nxt("tab")
fWI, fEI, fIE = nxt("fig"), nxt("fig"), nxt("fig")
body(f"光源電力に対するダイオード電流と，(1) の照度を合わせて表{t2}に示す．光源電力とダイオード電流の関係を"
     f"図{fWI}に，照度とダイオード電流の関係を図{fEI}と図{fIE}に示す．")
table("光源電力 W，照度 E とダイオード電流 Id", t2, ["光源電圧 V [V]", "光源電力 W [W]", "照度 E [lx]",
                                                 "ダイオード電流 Id [mA]"],
      [[f"{v}", f"{w:.3f}", f"{e}", f"{i:.3f}"] for v, w, e, i in zip(LV, LW, LE, LID)], [30, 32, 30, 38], size=9)
C.chart_figure(ch_WId(), "光源電力 W とダイオード電流 Id の関係", fWI, 115)
C.chart_figure(ch_EId(), "照度 E とダイオード電流 Id の関係", fEI, 115)
C.chart_figure(ch_IdE(), "ダイオード電流 Id と照度 E の関係（曲線は2次の多項式近似）", fIE, 115)
body(f"図{fWI}より，ダイオード電流も光源電力とともに下に凸の形で増え，照度と同じような変化をした．図{fEI}より，"
     f"ダイオード電流は照度とともに増えたが，原点を通る直線にはならず，照度が大きいほど傾きが小さくなった．"
     f"ダイオード電流 1 mA あたりの照度 E/Id は，光源電圧 6 V で約{LE[6] / LID[6]:.0f} lx/mA，15 V で約"
     f"{LE[15] / LID[15]:.0f} lx/mA と，光が強いほど大きくなった．このため，(3) で使う Id から E への換算には，"
     f"図{fIE}の2次の多項式近似")
e_conv = eq(rf"E={EC[0]:.0f}\,I_{{\mathrm{{d}}}}^2+{EC[1]:.0f}\,I_{{\mathrm{{d}}}}{EC[2]:+.1f}\qquad"
            rf"(E\ \mathrm{{[lx]}},\ I_{{\mathrm{{d}}}}\ \mathrm{{[mA]}})")
body("を用いた．")
heading("(3)　移動距離 x に対する照度 E", 2)
t3 = nxt("tab")
fxE, fxy = nxt("fig"), nxt("fig")
body(f"各移動距離でのダイオード電流と，式({e_conv})で換算した照度，y＝1/√E を表{t3}に示す．移動距離と照度の"
     f"関係を図{fxE}に，移動距離と y の関係を図{fxy}に示す．")
table("移動距離 x，ダイオード電流 Id，照度 E と 1/√E", t3,
      ["移動距離 x [cm]", "ダイオード電流 Id [mA]", "照度 E [lx]", "1/√E [lx^{-1/2}]"],
      [[f"{x:.1f}", f"{i:.3f}", f"{e:.0f}", f"{y:.4f}"] for x, i, e, y in zip(DX, DID, DE, DY)], [30, 38, 28, 34],
      size=9)
C.chart_figure(ch_xE(), "移動距離 x と照度 E の関係", fxE, 115)
body(f"図{fxE}より，光源を遠ざけると照度は急に小さくなり，最初の 2 cm で約{DE[4] / DE[0] * 100:.0f} %まで"
     f"下がった．その後はゆるやかに減り，距離の2乗に反比例するような曲線になった．")
C.chart_figure(ch_xy(), "移動距離 x と 1/√E の関係（直線は最小二乗法による実験式）", fxy, 120)
body(f"21組のデータ（l＝21）から式({e_ls})で k，m を求めると")
eq(rf"k={K_:.4f}\ \mathrm{{cm^{{-1}}lx^{{-1/2}}}},\qquad m={M_:.5f}\ \mathrm{{lx^{{-1/2}}}}", numbered=False)
body("となり，実験式は")
eq(rf"y=\dfrac{{1}}{{\sqrt{{E}}}}={K_:.4f}\,x+{M_:.5f}", numbered=False)
body("である．よって")
eq(rf"a=\dfrac{{1}}{{k^2}}={A_:.0f}\ \mathrm{{lx\cdot cm^2}},\qquad r_0=\dfrac{{m}}{{k}}={R0:.2f}\ \mathrm{{cm}}",
   numbered=False)
body(f"となった．図{fxy}より，y はほぼ x に比例して増えており（決定係数 R²＝{R2_E:.3f}），照度が距離の2乗に"
     f"反比例することが確かめられた．ただし，x が大きいところ（8 cm 以降）では測定点が直線より上にずれ，"
     f"全体としてわずかに下に凸の曲線になった．")
heading("(4)　ダイオード電流 Id－ダイオード電圧 Vd 特性", 2)
t4 = nxt("tab")
fIV = nxt("fig")
body(f"ダイオード電圧とダイオード電流の測定結果を表{t4}に，そのグラフを図{fIV}に示す．逆方向の電圧は負の値で表した．")
rows4 = [[f"{v:.1f}", f"{a:.3f}", f"{b:.3f}"] for v, a, b in zip(R_VD[::-1], R_ON[::-1], R_OFF[::-1])] + \
        [[f"{v:.1f}", f"{a:.3f}", f"{b:.3f}"] for v, a, b in zip(VD, F_ON, F_OFF)]
table("ダイオード電圧 Vd とダイオード電流 Id", t4, ["ダイオード電圧 Vd [V]", "照射時 Id [mA]", "無照射時 Id [mA]"],
      rows4, [40, 40, 40], size=9)
C.chart_figure(ch_IV(), "ダイオード電流－ダイオード電圧特性", fIV, 125)
body(f"無照射時は，逆方向から順方向 0.4 V まではほとんど電流が流れず（約{abs(F_OFF[0]):.3f} mA），0.5 V を"
     f"超えると電流が急に増えた．これは式({e_d})の通常のダイオードの整流特性である．照射時は，逆方向から"
     f"0.4 V まで約{abs(F_ON[0]):.2f} mA の負の電流がほぼ一定に流れた．これが光によって生じた光電流 i_{{C}} で，"
     f"電圧によらずほぼ一定であることがわかる．0.5 V 以下では，照射時の曲線は無照射時の曲線を光電流の分だけ"
     f"下にずらした形になっており，式({e_p})と一致する．順方向の電圧を大きくすると両者の差は小さくなり，"
     f"0.8 V 以上ではほぼ重なった．"
     f"また，照射時に電流が0になる電圧（開放電圧）は約 0.6 V であった．")

# ################################################################ ホール素子
C.doc.add_page_break()
C.set_part("")
title("ホール素子の実験")
heading("1．目的")
body("磁場を測るセンサとして広く使われているホール素子について，磁場に応じて電圧が生じる仕組み（ホール効果）を"
     "理解する．また，多層ソレノイドコイルの中の磁場の分布を測定し，電流と磁場の関係についての理解を深める．")
heading("2．原理")
heading("2.1　ホール効果", 2)
f_h = nxt("fig")
body(f"物質に電流を流しながら，それに垂直に磁場を加えると，電流と磁場の両方に垂直な向きに電圧が生じる．"
     f"これをホール効果といい，生じる電圧をホール電圧という．図{f_h}のように電流を x 方向，磁場を z 方向とする．"
     f"電流を運ぶ粒子（キャリア）が電子のとき，磁場によるローレンツ力で電子は y 方向に曲げられ，上面が負，"
     f"下面が正に帯電して y 方向の電場（ホール電場）ができる．十分時間がたつと，磁場による力と電場による力が"
     f"つり合う．キャリアの電荷を q，速度を v，磁束密度を B，ホール電場を E とすると，ローレンツ力は")
e_F = eq(r"\boldsymbol{F}=q(\boldsymbol{E}+\boldsymbol{v}\times\boldsymbol{B})")
body("であり，y 方向のつり合い F_{y}＝0 と v_{z}＝0 から E_{y}＝v_{x}B_{z} が得られる．キャリア密度を n とすると"
     "電流密度は J_{x}＝nqv_{x} である．素子の幅を w，厚さを d，電流を i_{x} とすると V_{h}＝E_{y}w，"
     "i_{x}＝J_{x}wd より")
e_vh = eq(r"V_h=\dfrac{1}{nq}\dfrac{i_xB_z}{d}=R_H\dfrac{i_xB_z}{d},\qquad R_H=\dfrac{1}{nq}")
body("となる．R_{H} をホール係数といい，キャリアが電子なら負，正孔なら正になるので，キャリアの種類を判定できる．"
     "式からわかるように，ホール電圧は試料電流と磁束密度の積に比例する．したがって，試料電流を一定にすれば"
     "ホール電圧は磁場に比例し，既知の磁場でのホール電圧がわかれば，ホール電圧から未知の磁場を求められる．"
     "金属ではキャリア密度が大きいため R_{H} が小さく，ホール電圧は μV 程度しかない．そのため磁場センサには"
     "キャリア密度の小さい半導体（Ge，InAs，GaAs など）が使われる．")
C.figure(fig("h_hall.png"), "ホール効果（指導書[2]より）", f_h, 85)
heading("2.2　多層ソレノイドコイルの中の磁場", 2)
f_coil = nxt("fig")
body(f"図{f_coil}のように，長さ 2l，単位長さあたりの巻数 n，内半径 R_{{1}}，外半径 R_{{2}} の多層ソレノイドに"
     f"電流 I を流したとき，中心 O から軸上で距離 x の点 P の磁場は")
e_coil = eq(r"H=\dfrac{nI}{2(R_2-R_1)}\left\{(l+x)\ln\dfrac{R_2+\sqrt{R_2^2+(l+x)^2}}{R_1+\sqrt{R_1^2+(l+x)^2}}"
            r"+(l-x)\ln\dfrac{R_2+\sqrt{R_2^2+(l-x)^2}}{R_1+\sqrt{R_1^2+(l-x)^2}}\right\}", size=11)
body("［A/m］で与えられる．x を決めれば H はコイル電流 I に比例する．中心（x＝0）では")
e_c0 = eq(r"H_O=\dfrac{nI}{R_2-R_1}\ln\dfrac{R_2+\sqrt{R_2^2+l^2}}{R_1+\sqrt{R_1^2+l^2}}")
body("となる．コイルの長さが半径より十分長ければ，中心付近の磁場はほぼ一様で，端に近づくと小さくなる．")
C.figure(fig("h_coil.png"), "多層ソレノイドコイルの断面図（指導書[2]より）", f_coil, 85)
heading("2.3　磁場と磁束密度の単位", 2)
body("SI 単位系では磁場 H を［A/m］，磁束密度 B を［T］（＝Wb/m²）で表し，B＝μH の関係がある．真空の透磁率は"
     "μ_{0}＝4π×10⁻⁷ H/m で，空気中でもほぼ同じである．CGS 単位系では H を［Oe］，B を［G］で表し，"
     "1 T＝10⁴ G，1 A/m＝4π×10⁻³ Oe の関係がある．")
heading("3．使用器具")
body("多層ソレノイドコイル1台，デジタルボルトメータ1台，直流電流計（30 A）1台，直流電流計（300 mA）1台，"
     "直流電源（20 V 10 A）1台，直流電源（18 V 1 A）1台，ホール素子1枚，永久磁石（1.67×10⁴ A/m）1台．")
heading("4．実験方法")
heading("4.1　実験の準備", 2)
f_w, f_lead, f_comp = nxt("fig"), nxt("fig"), nxt("fig")
body(f"図{f_w}のように配線した．ホール素子はコイルの長さ方向に動くレールの上にあり，位置はコイルの外の 1 mm 目盛の"
     f"スケールで読む．ホール素子の電極は図{f_lead}のとおりである．磁場が0でも，材料の不均一さや電極の位置の"
     f"ずれによってわずかな電圧（オフセット電圧）が生じるので，図{f_comp}の補償回路を使い，コイルに電流を"
     f"流さない状態で試料電流 i_{{x}} を流し，20 kΩ の可変抵抗でホール電圧を0に合わせた（零点調整）．"
     f"i_{{x}} を変えたときは，そのたびに零点調整をやり直した．")
C.figure(fig("h_wire.png"), "実験装置の配線図（指導書[2]より）", f_w, 110)
C.figure(fig("h_lead.png"), "ホール素子の電極リード（指導書[2]より）", f_lead, 75)
C.figure(fig("h_comp.png"), "電流・電圧リードの結線図（指導書[2]より）", f_comp, 90)
heading("4.2　実験手順", 2)
f_mag = nxt("fig")
for s_ in [f"(1)　ホール電圧 V_{{h}} と磁場 H の関係：ホール素子をコイルから出して i_{{x}}＝1 mA で零点調整をした．"
           f"図{f_mag}のように素子を永久磁石（H_{{0}}＝1.67×10⁴ A/m）の中央に置き，試料電流 i_{{x0}} と"
           f"ホール電圧 V_{{h0}} を記録した．",
           "(2)　コイル中心での測定：再び零点調整をし，ホール素子がコイルの左端と右端にあるときの目盛 x_{L}，x_{R} から"
           "中心の目盛 (x_{L}＋x_{R})/2 を求めて素子を中心に置いた．コイル電流 I を 0 A から 0.5 A ずつ 4 A まで変え，"
           "ホール電圧 V_{h} を測定した．",
           "(3)　磁場の分布：コイル電流 1 A と 4 A について，素子をコイルの左端から右端まで 2 cm ずつ動かし，"
           "コイル中心からの距離 x に対するホール電圧を測定した．",
           "(4)　試料電流とホール電圧：コイル電流を 2 A とし，素子をコイル中心に置いて，試料電流 i_{x} を 1 mA から"
           "8 mA まで 1 mA ずつ変えてホール電圧を測定した．i_{x} を変えるたびに，素子をコイルから出して零点調整をした．"]:
    item(s_, left=0)
C.figure(fig("h_mag.png"), "ホール素子の取り付け方（指導書[2]より）", f_mag, 75)

heading("5．実験結果")
heading("(1)　ホール電圧 Vh と磁場 H の関係", 2)
t5 = nxt("tab")
body(f"永久磁石での測定結果を表{t5}に示す．")
table("永久磁石によるホール電圧と比例定数", t5, ["項目", "値"],
      [["永久磁石の磁場 H_{0}", "1.67×10⁴ A/m"], ["試料電流 i_{x0}", f"{IX0:.3f} mA"],
       ["ホール電圧 V_{h0}", f"{VH0:.2f} mV"], ["比例定数 k＝H_{0}/V_{h0}", f"{KH:.0f} (A/m)/mV"]], [60, 50])
body("ホール電圧は磁場に比例するので，以下ではホール電圧 V_{h}［mV］を次の式で磁場 H［A/m］に換算した．")
e_k = eq(rf"H=kV_h={KH:.0f}\,V_h")
heading("(2)　コイル中心でのコイル電流とホール電圧・磁場", 2)
t6 = nxt("tab")
fIVh, fIH = nxt("fig"), nxt("fig")
body(f"コイル中心の目盛は{CENTER} cm であった．結果を表{t6}に，コイル電流とホール電圧の関係を図{fIVh}に，"
     f"磁場の関係を図{fIH}に示す．")
table("コイル電流 I，ホール電圧 Vh と磁場 H（コイル中心）", t6, ["コイル電流 I [A]", "ホール電圧 V_{h} [mV]",
                                                       "磁場 H [A/m]"],
      [[f"{i:.1f}", f"{v:.3f}", f"{h:.0f}"] for i, v, h in zip(CI, CV, CH)], [40, 40, 40], size=9)
C.chart_figure(ch_IVh(), "コイル電流 I とホール電圧 Vh の関係", fIVh, 115)
C.chart_figure(ch_IH(), "コイル電流 I と磁場 H の関係（直線は最小二乗法による実験式）", fIH, 115)
body("最小二乗法で H＝aI の比例定数を求めると")
eq(rf"a=\dfrac{{\sum I_kH_k}}{{\sum I_k^2}}={AH:.0f}\ \mathrm{{(A/m)/A}}", numbered=False)
body(f"となり，実験式は H＝{AH:.0f}I［A/m］となった．図{fIVh}・図{fIH}より，ホール電圧と磁場はコイル電流に"
     f"比例して増えており，測定点は実験式の直線によく乗った．式({e_c0})のとおり，コイルの中の磁場は電流に"
     f"比例することが確かめられた．4 A のときの磁場は{CH[-1]:.0f} A/m で，磁束密度にすると"
     f"μ_{{0}}H＝{4e-7 * math.pi * CH[-1] * 1000:.1f} mT である．")
heading("(3)　コイル中心からの距離 x に対する磁場 H", 2)
t7 = nxt("tab")
fxH = nxt("fig")
body(f"結果を表{t7}と図{fxH}に示す．")
table("コイル中心からの距離 x，ホール電圧 Vh と磁場 H", t7,
      ["距離 x [cm]", "1 A：V_{h} [mV]", "1 A：H [A/m]", "4 A：V_{h} [mV]", "4 A：H [A/m]"],
      [[f"{x}", f"{a:.3f}", f"{h1:.0f}", f"{b:.3f}", f"{h4:.0f}"] for x, a, h1, b, h4 in zip(PX, PV1, PH1, PV4, PH4)],
      [26, 28, 28, 28, 28], size=9)
C.chart_figure(ch_xH(), "コイル中心からの距離 x と磁場 H の関係", fxH, 125)
p1, p2 = plateau(PX, PH4)
l4, r4 = half_width(PX, PH4)
ratio = [h4 / h1 for h1, h4 in zip(PH1, PH4)]
body(f"どちらの電流でも，磁場は中心付近（x＝{p1}〜{p2} cm）でほぼ一定（最大値の95 %以上）で，コイルの端に"
     f"近づくにつれて急に小さくなった．x＝±18 cm での磁場は中心の約{PH4[0] / max(PH4) * 100:.0f}〜"
     f"{PH4[-1] / max(PH4) * 100:.0f} %であった．コイルの中央部では一様な磁場が得られるが，コイルの長さが有限で"
     f"あるため，端では磁力線が外に広がって磁場が弱くなることがわかる．また，各位置で 4 A のときの磁場は"
     f" 1 A のときの{min(ratio):.2f}〜{max(ratio):.2f}倍で，ほぼ4倍になっており，コイルのどの位置でも磁場は"
     f"電流に比例している．分布は x＝0 に対してほぼ左右対称であった．")
heading("(4)　一定磁場における試料電流 ix とホール電圧 Vh", 2)
t8 = nxt("tab")
fix = nxt("fig")
body(f"コイル電流 2 A（(2) の実験式より H＝{AH:.0f}×2＝{H2A:.0f} A/m）での結果を表{t8}と図{fix}に示す．")
table("試料電流 ix とホール電圧 Vh（コイル電流 2 A）", t8, ["試料電流 i_{x} [mA]", "ホール電圧 V_{h} [mV]"],
      [[f"{i}", f"{v:.3f}"] for i, v in zip(SIX, SVH)], [45, 45], size=9)
C.chart_figure(ch_ixVh(), "試料電流 ix とホール電圧 Vh の関係（磁場 H 一定）", fix, 115)
body(f"図{fix}より，磁場が一定でも試料電流を大きくするとホール電圧はそれに比例して大きくなった．原点を通る直線で"
     f"近似すると Vh＝{SK:.3f}i_{{x}}［mV］（i_{{x}} は mA）となり，式({e_vh})のとおりホール電圧は試料電流と"
     f"磁場の積に比例することが確かめられた．")

# ################################################################ 共通：考察・まとめ
C.doc.add_page_break()
C.set_part("")
heading("6．考察")
heading("6.1　光センサの実験について", 2)
body(f"(a) 照度とダイオード電流の関係：2.3で述べたように，短絡に近い状態で測ったダイオード電流は光の量に比例する"
     f"はずである．しかし実験(2)では，照度あたりのダイオード電流（Id/E）は光源電力が大きいほど小さくなった．"
     f"この原因は，照度計とフォトダイオードが「光の量」を別の基準で測っていることにあると考えられる．照度計は"
     f"人の目の感じ方（比視感度，555 nm で最大）に合わせた量を測るのに対し，Si フォトダイオードは約1100 nm までの"
     f"赤外線にも強く反応する．ハロゲンランプは電力が小さいときはフィラメントの温度が低く，赤外線の割合が大きい．"
     f"そのため，弱い光ほど照度計には暗く見えるのにダイオードには電流が流れやすく，Id/E が大きくなった．"
     f"電力を上げると可視光の割合が増えるので，照度は電力以上の割合で増え，Id/E は小さくなる．")
body(f"(b) 逆2乗則と最小二乗法：実験(3)で照度に換算した値を使うと，y＝1/√E は x とほぼ直線関係になり"
     f"（R²＝{R2_E:.3f}），r_{{0}}＝{R0:.2f} cm となった．しかし x の大きいところでは直線から上にずれた．"
     f"この原因は，換算に使った式({e_conv})にあると考えられる．式({e_conv})は実験(2)で光源電力を変えながら"
     f"求めたものであり，(a) で述べたように光の色（赤外線の割合）の変化も含んだ関係になっている．一方，実験(3)では"
     f"光源電力を 4 W に固定して距離だけを変えたので，光の色は変わらず，ダイオード電流は素子に届く光の量に"
     f"そのまま比例するはずである．そこで y＝1/√Id を使って同じように最小二乗法を行うと，R²＝{R2_i:.3f} と"
     f"より直線に近くなり，r_{{0}}＝{R0_i:.1f} cm となった．このように，同じ測定値でも換算に使う関係式によって"
     f"パラメータの値が大きく変わる．最小二乗法は与えられた式に最もよく合うパラメータを機械的に求める方法なので，"
     f"換算や近似の仕方が実験の条件に合っているかを，グラフで確かめることが大切である．")
body("(c) Id－Vd 特性：照射時の逆方向の電流は電圧によらずほぼ一定（約 0.56〜0.57 mA）であり，これは光によって"
     "生じた光電流である．光電流の大きさは光の量で決まり電圧にはほとんどよらないため，逆方向に電圧をかけて"
     "使えば，光の量に比例した電流をとり出すことができる．これがフォトダイオードを光センサとして使うときの"
     "基本的な使い方である．一方，順方向で電流が0になる電圧（開放電圧）は約 0.6 V であり，このとき素子は電池と"
     "して働いている．これが太陽電池の原理である．")
heading("6.2　ホール素子の実験について", 2)
body(f"(a) ホール素子による磁場測定：実験(2)(4)から，ホール電圧はコイル電流（磁場）にも試料電流にも比例した．"
     f"したがって，ホール素子で正しく磁場を測るには試料電流を一定に保つことが重要である．例えば試料電流が 1 %"
     f"変わるとホール電圧も 1 %変わり，そのまま磁場の誤差になる．永久磁石で比例定数 k を決めたときの試料電流が"
     f" {IX0} mA であったことも，測定の前提として記録しておく必要がある．")
body(f"(b) 零点調整の必要性：磁場が0でも，ホール素子には電極の位置のずれなどによるオフセット電圧が生じる．"
     f"この電圧は試料電流に比例して変わるため，試料電流を変えるたびに零点調整をやり直す必要があった．実験(4)の"
     f"結果が原点を通る直線によく乗ったのは，毎回零点調整を行ったためである．")
body(f"(c) コイルの磁場分布：コイルの中央部では磁場がほぼ一様であったが，端では大きく下がった．一様な磁場を"
     f"使いたい場合は，コイルの中央部（本実験では中心から±{p2} cm 程度）を使う必要がある．また，ホール電圧から"
     f"求めた磁場はすべて永久磁石1点での比例定数 k に基づいているため，k の誤差（永久磁石の磁場の値や，そのときの"
     f"素子の置き方）がそのまま全体の磁場の値に影響する．")

heading("7．まとめ")
body(f"光センサの実験では，光源電力とともに照度とフォトダイオード電流が増えることを確かめ，ダイオード電流から"
     f"照度への換算式を求めた．光源を遠ざけたときの照度は距離の2乗にほぼ反比例し，最小二乗法から a＝{A_:.0f} "
     f"lx·cm²，r_{{0}}＝{R0:.2f} cm を得た．Id－Vd 特性では，光を当てると電圧によらずほぼ一定の光電流が流れ，"
     f"ダイオード特性が光電流の分だけずれることを確かめた．")
body(f"ホール素子の実験では，永久磁石を用いてホール電圧と磁場の比例定数 k＝{KH:.0f} (A/m)/mV を求め，コイル中心の"
     f"磁場がコイル電流に比例すること（H＝{AH:.0f}I）を確かめた．コイルの中の磁場は中央でほぼ一様で端に近づくと"
     f"小さくなり，ホール電圧は試料電流にも比例することがわかった．")
C.references([
    "[1] 物理学実験 指導書「12．光センサの実験」，2026．",
    "[2] 物理学実験 指導書「13．ホール素子の実験」，2026．",
])

# 表がページをまたがないように，最終行以外の段落を「次の段落と分離しない」にする
for t in C.doc.tables:
    for r in t.rows[:-1]:
        for c in r.cells:
            for p in c.paragraphs:
                p.paragraph_format.keep_with_next = True
    for r in t.rows:
        trPr = r._tr.get_or_add_trPr()
        el = C.OxmlElement("w:cantSplit")
        trPr.append(el)

out = os.path.join(HERE, "光センサとホール素子の実験_報告書.docx")
C.save(out)
print("saved", out)
print("EC", EC, "k m", K_, M_, "a r0", A_, R0, R2_E, "Id-based", R0_i, R2_i)
print("KH", KH, "AH", AH, "SK", SK, "H2A", H2A, "plateau", plateau(PX, PH4), half_width(PX, PH4))
