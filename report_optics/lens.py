# -*- coding: utf-8 -*-
"""5. 薄レンズの焦点距離の測定 — 報告書（Word）生成。 python3 lens.py"""
import math
import os
import statistics as st
import sys

import matplotlib.pyplot as plt
import numpy as np
from openpyxl.chart import Reference, Series
from openpyxl.chart.marker import Marker

import common as C
from common import body, eq, heading, item, nxt, para, sci, table, tsci

HERE = os.path.dirname(os.path.abspath(__file__))
IMG = os.path.join(HERE, "img_lens")
os.makedirs(IMG, exist_ok=True)

META = {"kai": None, "date": None, "submit": None, "han": "D班", "name": "森木崇允", "grade": "1",
        "student_id": "4326091",
        "coworkers": [("石井真人", "4326007"), ("木元喜紹", "4326029"), ("陶山稜平", "4326051"),
                      ("林田知子", "4326071")]}

# ---------------------------------------------------------------- 測定値（実験ノートの Excel より）
T_CONVEX = 0.495     # 凸レンズの厚さ t [cm]
T_CONCAVE = 0.200    # 凹レンズの厚さ t' [cm]
L_T = 21.0           # T字棒の長さ l [cm]
CONVEX = [(22.370, 41.600, 62.620, 104.620), (22.555, 32.000, 53.300, 153.810),
          (22.400, 55.310, 76.705, 102.000), (22.430, 81.015, 102.340, 116.200)]   # x1..x4 [cm]
BESSEL = [(47.200, 96.600, 22.420, 119.905), (51.620, 78.570, 22.500, 106.500),
          (50.900, 80.320, 22.455, 107.490), (54.200, 82.510, 22.500, 108.400)]   # x, x', y, y' [cm]
CONCAVE = [(89.700, 102.050, 114.000), (87.700, 107.025, 155.200),
           (87.230, 106.600, 178.070), (83.400, 102.610, 171.900)]                # x1, x2, x3 [cm]
H1_RAW, H2_RAW, H_ZERO = 0.346, 0.060, 0.011   # レンズメータの読み [mm]（第一面，第二面，平面ガラス）
A_LM = 10.0                                    # レンズメータの a [mm]

# ---------------------------------------------------------------- 計算
cv = []
for x1, x2, x3, x4 in CONVEX:
    a = abs(x3 - x1) - T_CONVEX / 2
    b = abs(x4 - x2) - T_CONVEX / 2
    cv.append(dict(x=(x1, x2, x3, x4), a=a, b=b, ia=1 / a, ib=1 / b, inv=1 / a + 1 / b, f=1 / (1 / a + 1 / b)))
bs = []
for x, xp, y, yp in BESSEL:
    D = abs(yp - y) + L_T
    d = abs(xp - x)
    bs.append(dict(x=(x, xp, y, yp), D=D, d=d, f=(D * D - d * d) / (4 * D)))
cc = []
for x1, x2, x3 in CONCAVE:
    a = abs(x2 - x1) - T_CONCAVE / 2
    b = abs(x3 - x1) - T_CONCAVE / 2
    cc.append(dict(x=(x1, x2, x3), a=a, b=b, ia=1 / a, ib=1 / b, inv=1 / a - 1 / b, f=1 / (1 / a - 1 / b)))


def stats(v):
    return st.mean(v), st.stdev(v)


f1, s1 = stats([r["f"] for r in cv])
f2, s2 = stats([r["f"] for r in bs])
f3, s3 = stats([r["f"] for r in cc])
h1, h2 = H1_RAW - H_ZERO, H2_RAW - H_ZERO
R1 = 0.5 * (A_LM ** 2 / h1 + h1)
R2 = 0.5 * (A_LM ** 2 / h2 + h2)
n_lens = (1 / (f1 * 10)) / (1 / R1 + 1 / R2) + 1
n_bes = (1 / (f2 * 10)) / (1 / R1 + 1 / R2) + 1
# 読み h が 0.005 mm ずれたときの影響（感度）
def n_of(h1_, h2_, f_):
    r1 = 0.5 * (A_LM ** 2 / h1_ + h1_)
    r2 = 0.5 * (A_LM ** 2 / h2_ + h2_)
    return (1 / (f_ * 10)) / (1 / r1 + 1 / r2) + 1
dn_h1 = n_of(h1 + 0.005, h2, f1) - n_lens
dn_h2 = n_of(h1, h2 + 0.005, f1) - n_lens
dn_f = n_of(h1, h2, f1 + s1) - n_lens
# 凹レンズ：a' と b' が 0.5 cm ずれたときの f' の変化
ac, bc = st.mean([r["a"] for r in cc]), st.mean([r["b"] for r in cc])
df_da = 1 / (1 / (ac + 0.5) - 1 / bc) - 1 / (1 / ac - 1 / bc)
df_db = 1 / (1 / ac - 1 / (bc + 0.5)) - 1 / (1 / ac - 1 / bc)
N_BK7 = 1.5168
F_NOM = 24.0     # レンズの公称の焦点距離 [cm]
err = lambda v: (v - F_NOM) / F_NOM * 100


# ---------------------------------------------------------------- 模式図（自作）
def _arrow(ax, x, y0, y1, **kw):
    ax.annotate("", xy=(x, y1), xytext=(x, y0), arrowprops=dict(arrowstyle="-|>", lw=1.4, color="black"), **kw)


def _lens(ax, x, h=1.6, concave=False):
    from matplotlib.patches import Ellipse, Polygon
    if concave:
        ax.add_patch(Polygon([[x - 0.18, h], [x + 0.18, h], [x + 0.05, 0], [x + 0.18, -h], [x - 0.18, -h],
                              [x - 0.05, 0]], closed=True, fill=False, lw=1.2))
    else:
        ax.add_patch(Ellipse((x, 0), 0.35, 2 * h, fill=False, lw=1.2))


def fig_convex(path):
    fig, ax = plt.subplots(figsize=(6.2, 2.6))
    f, a = 2.0, 3.2
    b = 1 / (1 / f - 1 / a)
    ho = 0.9
    hi = -ho * b / a
    ax.axhline(0, color="black", lw=0.7)
    _lens(ax, 0)
    _arrow(ax, -a, 0, ho)
    _arrow(ax, b, 0, hi)
    ax.plot([-a, 0, b], [ho, ho, hi], color="black", lw=0.9)
    ax.plot([-a, b], [ho, hi], color="black", lw=0.9)
    for xf in (-f, f):
        ax.plot([xf], [0], "o", color="black", ms=3)
    ax.text(f, -0.25, "F", ha="center", va="top", fontsize=9)
    ax.text(-f, -0.25, "F", ha="center", va="top", fontsize=9)
    ax.text(-a, ho + 0.12, "物体 A", ha="center", fontsize=9)
    ax.text(b, hi - 0.15, "像 B", ha="center", va="top", fontsize=9)
    ax.text(0.05, 1.75, "凸レンズ L（光心 O）", ha="center", fontsize=9)
    for x0, x1, y, t in [(-a, 0, -1.25, "a"), (0, b, -1.25, "b"), (0, f, 0.45, "f")]:
        ax.annotate("", xy=(x1, y), xytext=(x0, y), arrowprops=dict(arrowstyle="<->", lw=0.8))
        ax.text((x0 + x1) / 2, y + 0.07, t, ha="center", va="bottom", fontsize=10, style="italic")
    ax.set_xlim(-a - 0.6, b + 0.6); ax.set_ylim(-1.9, 2.0); ax.axis("off")
    fig.tight_layout(); fig.savefig(path, dpi=220); plt.close(fig)


def fig_bessel(path):
    fig, axs = plt.subplots(2, 1, figsize=(6.2, 3.3), sharex=True)
    D, f = 10.0, 2.0
    d = (D * D - 4 * D * f) ** 0.5
    a1 = (D - d) / 2
    for ax, xl, lab in [(axs[0], a1, "O"), (axs[1], D - a1, "O′")]:
        ax.axhline(0, color="black", lw=0.7)
        _arrow(ax, 0, 0, 0.8)
        ax.plot([D, D], [-1.3, 1.3], color="black", lw=1.6)
        _lens(ax, xl, h=1.1)
        ax.text(xl, 1.2, f"L（光心 {lab}）", ha="center", fontsize=8.5)
        ax.text(0, 0.95, "C（物体）", ha="center", fontsize=8.5)
        ax.text(D, 1.4, "S（スクリーン）", ha="center", fontsize=8.5)
        ax.set_ylim(-1.6, 1.9); ax.axis("off")
    axs[0].annotate("", xy=(a1, -1.0), xytext=(0, -1.0), arrowprops=dict(arrowstyle="<->", lw=0.8))
    axs[0].text(a1 / 2, -0.95, "a", ha="center", va="bottom", fontsize=10, style="italic")
    axs[0].annotate("", xy=(D, -1.0), xytext=(a1, -1.0), arrowprops=dict(arrowstyle="<->", lw=0.8))
    axs[0].text((a1 + D) / 2, -0.95, "b", ha="center", va="bottom", fontsize=10, style="italic")
    axs[1].annotate("", xy=(D - a1, -0.8), xytext=(a1, -0.8), arrowprops=dict(arrowstyle="<->", lw=0.8))
    axs[1].text(D / 2, -0.75, "d（レンズの移動距離）", ha="center", va="bottom", fontsize=9)
    axs[1].plot([a1, a1], [-1.1, 1.1], color="0.6", lw=0.8, ls=":")
    axs[1].annotate("", xy=(D, -1.45), xytext=(0, -1.45), arrowprops=dict(arrowstyle="<->", lw=0.8))
    axs[1].text(D / 2, -1.4, "D（物体とスクリーンの距離）", ha="center", va="bottom", fontsize=9)
    axs[0].set_xlim(-0.8, D + 0.8)
    fig.tight_layout(); fig.savefig(path, dpi=220); plt.close(fig)


def fig_concave(path):
    fig, ax = plt.subplots(figsize=(6.2, 2.3))
    ax.axhline(0, color="black", lw=0.7)
    _arrow(ax, 0, 0, 0.7)
    _lens(ax, 2.0, h=1.2)
    _lens(ax, 5.0, h=1.0, concave=True)
    xb, xbp = 6.5, 9.5
    ax.plot([xb, xb], [-0.9, 0.2], color="0.5", lw=1.0, ls="--")
    ax.plot([xbp, xbp], [-1.2, 0.2], color="black", lw=1.6)
    ax.text(0, 0.8, "C（物体）", ha="center", fontsize=8.5)
    ax.text(2.0, 1.3, "凸レンズ L", ha="center", fontsize=8.5)
    ax.text(5.0, 1.1, "凹レンズ L′（光心 O′）", ha="center", fontsize=8.5)
    ax.text(xb, 0.3, "B（L だけの像）", ha="center", fontsize=8.5)
    ax.text(xbp, 0.3, "B′（L′ を入れた後の像）", ha="center", fontsize=8.5)
    for x0, x1, y, t in [(5.0, xb, -1.15, "a′"), (5.0, xbp, -1.55, "b′")]:
        ax.annotate("", xy=(x1, y), xytext=(5.0, y), arrowprops=dict(arrowstyle="<->", lw=0.8))
        ax.text((x0 + x1) / 2, y + 0.06, t, ha="center", va="bottom", fontsize=10, style="italic")
    ax.set_xlim(-0.8, 11.2); ax.set_ylim(-1.8, 1.6); ax.axis("off")
    fig.tight_layout(); fig.savefig(path, dpi=220); plt.close(fig)


# ---------------------------------------------------------------- Excel グラフ
def chart_inverse():
    """1/a と 1/b の関係（凸レンズ）：測定点とレンズの公式 1/b = 1/f − 1/a。"""
    ws = C.new_sheet("1a_1b")
    ws.append(["1/a [cm⁻¹]", "1/b [cm⁻¹]（測定）", "", "1/a [cm⁻¹]", "1/b [cm⁻¹]（1/b = 1/f − 1/a）"])
    for r in cv:
        ws.append([r["ia"], r["ib"]])
    vmax = max(max(r["ia"] for r in cv), max(r["ib"] for r in cv), 1 / f1)
    top = math.ceil(vmax / 0.005) * 0.005 + 0.005
    xs = np.linspace(0, top, 21)
    for i, x in enumerate(xs):
        ws.cell(row=2 + i, column=4, value=float(x))
        ws.cell(row=2 + i, column=5, value=float(1 / f1 - x))
    ch = C.base_chart("1/a [cm⁻¹]", "1/b [cm⁻¹]")
    s_l = Series(Reference(ws, min_col=5, min_row=2, max_row=22), Reference(ws, min_col=4, min_row=2, max_row=22),
                 title=f"レンズの公式（f = {f1:.2f} cm）")
    s_l.marker = Marker(symbol="none"); C.line(s_l, C.ORANGE)
    s_p = Series(Reference(ws, min_col=2, min_row=2, max_row=5), Reference(ws, min_col=1, min_row=2, max_row=5),
                 title="測定値")
    s_p.marker = C.marker("circle", 7, C.BLUE); C.line(s_p, none=True)
    ch.series.append(s_l); ch.series.append(s_p)
    for ax in (ch.x_axis, ch.y_axis):
        ax.scaling.min, ax.scaling.max, ax.majorUnit = 0, top, 0.005
    return C.native(ws, ch, "G2")


def chart_compare():
    """測定ごとの焦点距離（3 つの方法）。"""
    ws = C.new_sheet("f")
    ws.append(["回", "凸（レンズの公式）f [cm]", "凸（ベッセル）f [cm]", "凹 f′ [cm]"])
    for i in range(4):
        ws.append([i + 1, cv[i]["f"], bs[i]["f"], cc[i]["f"]])
    ch = C.base_chart("測定回", "焦点距離 [cm]", h=9)
    for col, color, name in [(2, C.BLUE, "凸レンズ（レンズの公式）"), (3, C.ORANGE, "凸レンズ（ベッセルの方法）"),
                             (4, C.GRAY, "凹レンズ")]:
        s = Series(Reference(ws, min_col=col, min_row=2, max_row=5), Reference(ws, min_col=1, min_row=2, max_row=5),
                   title=name)
        s.marker = C.marker("circle", 7, color); C.line(s, color, width=19050)
        ch.series.append(s)
    ch.x_axis.scaling.min, ch.x_axis.scaling.max, ch.x_axis.majorUnit = 0.5, 4.5, 1
    allf = [r["f"] for r in cv + bs + cc]
    ch.y_axis.scaling.min = math.floor(min(allf)) - 1
    ch.y_axis.scaling.max = math.ceil(max(allf)) + 1
    ch.y_axis.majorUnit = 1
    return C.native(ws, ch, "G2")


# ================================================================ 文書
STANDALONE = __name__ == "__main__"
if STANDALONE:
    C.new_doc(os.path.join(IMG, "eq"))
fig_convex(os.path.join(IMG, "fig_convex.png"))
fig_bessel(os.path.join(IMG, "fig_bessel.png"))
fig_concave(os.path.join(IMG, "fig_concave.png"))
if STANDALONE:
    C.cover(META, "薄レンズの焦点距離の測定")
heading("A．レンズの焦点距離の測定")

heading("A.1　目的")
body("薄い凸レンズおよび凹レンズの焦点距離を測定し，レンズから物体および像までの距離の間に成り立つ"
     "レンズの公式を確かめる．凸レンズについてはレンズの公式による方法とベッセルの方法の2通りで焦点距離を"
     "求めて比較する．また，ダイヤル・レンズメータでレンズ面の曲率半径を測定し，レンズに用いられている"
     "ガラスの屈折率を推定する．")

heading("A.2　理論")
heading("(1)　薄い凸レンズの公式", 2)
body("厚さが無視できるレンズを薄レンズという．レンズの2つの球面の曲率半径をR_{1}，R_{2}とし，光軸から"
     "高さhの位置を通る光を考える．この位置のレンズは，2つの面がわずかに傾いた小さなプリズムとみなせる．"
     "2つの面の接線がなす角（プリズムの頂角）θは，hが小さいとき")
n_th = eq(r"\theta=h\left(\dfrac{1}{R_1}+\dfrac{1}{R_2}\right)")
body("と表される[1]．頂角の小さいプリズムに小さな入射角で入った光は，ガラスの屈折率をnとして")
n_dl = eq(r"\delta=(n-1)\,\theta")
body(f"だけ曲げられる（この関係の導出は A.5 の課題(1)で示す）．式({n_th})，({n_dl})から，光の曲がる角δは"
     "hに比例する．物体上の点Aからレンズまでの距離をa，像の点Bまでの距離をbとすると，小さい角の近似で"
     "δ≈h/a＋h/b となるので，")
n_lm = eq(r"\dfrac{1}{a}+\dfrac{1}{b}=(n-1)\left(\dfrac{1}{R_1}+\dfrac{1}{R_2}\right)")
body("が得られる．右辺はhによらないので，Aから出てレンズを通った光はすべて1点Bに集まり，BはAの実像と"
     "なる．物体が無限遠にある（平行光線が入る）ときの像の位置が焦点Fであり，レンズから焦点までの距離fを"
     "焦点距離という．a→∞として")
n_fn = eq(r"\dfrac{1}{f}=(n-1)\left(\dfrac{1}{R_1}+\dfrac{1}{R_2}\right)")
body("となり，これを用いると式(" + str(n_lm) + ")は")
n_lf = eq(r"\dfrac{1}{a}+\dfrac{1}{b}=\dfrac{1}{f}")
f_cv = nxt("fig")
body(f"と書ける．これを薄い凸レンズの公式という[1][2]．a，b，fの関係を図{f_cv}に示す．像の倍率は "
     "m＝b/a である．")
C.figure(os.path.join(IMG, "fig_convex.png"), "薄い凸レンズによる結像（模式図）", f_cv, 120)

heading("(2)　ベッセルの方法", 2)
f_bs = nxt("fig")
body("物体とスクリーンの距離Dを一定に保ってレンズを動かすと，スクリーン上に像ができる位置は2か所ある．"
     "一方の位置でレンズから物体・像までの距離がa，bであれば，もう一方の位置ではそれがb，aと入れ替わる"
     f"（図{f_bs}）．2つの位置の間隔をdとすると a＋b＝D，b－a＝d であり，a＝(D－d)/2，b＝(D＋d)/2 を"
     f"式({n_lf})に代入すると")
n_bs = eq(r"f=\dfrac{D^2-d^2}{4D}")
body("が得られる（ベッセルの式）[1]．この方法では，D と d がいずれもレンズの位置の「差」から求まるため，"
     "レンズの光心の位置や厚さを知らなくても焦点距離を求められる．なお，像が2か所にできるためには "
     "D＞4f である必要がある．")
C.figure(os.path.join(IMG, "fig_bessel.png"), "ベッセルの方法におけるレンズの2つの位置（模式図）", f_bs, 120)

heading("(3)　薄い凹レンズの公式", 2)
body("凹レンズを通った光は，ある点Bから出たかのように広がる（虚像）．物体までの距離をa，虚像までの距離をb"
     "とすると，凸レンズと同様にして")
n_cc = eq(r"\dfrac{1}{b}-\dfrac{1}{a}=\dfrac{1}{f}")
f_cc = nxt("fig")
body("が成り立つ（fは凹レンズの焦点距離）[1]．凹レンズだけでは実像ができないため，本実験では凸レンズLで"
     f"つくった実像Bを凹レンズL′に対する物体（虚物体）として用いる（図{f_cc}）．L′の光心O′からBまでの距離を"
     "a′，L′を入れた後に像ができる位置B′までの距離をb′とすると，")
n_cc2 = eq(r"\dfrac{1}{f'}=\dfrac{1}{a'}-\dfrac{1}{b'}")
body("から凹レンズの焦点距離f′が求まる．ここでa′，b′，f′はすべて正の量としている．")
C.figure(os.path.join(IMG, "fig_concave.png"), "凹レンズの焦点距離の測定（模式図）", f_cc, 120)

heading("(4)　レンズのガラスの屈折率", 2)
body("ダイヤル・レンズメータは，間隔2aの2本の固定端子の間にある可動端子の押し込み量（読み）hから，レンズ面の"
     "曲率半径Rを求める器具である．球面の幾何学的な関係から")
n_rh = eq(r"R=\dfrac{1}{2}\left(\dfrac{a^2}{h}+h\right)")
body(f"となる（本器では a＝10 mm，固定端子の半径は0とした）[1]．R_{{1}}，R_{{2}}を測定し，式({n_fn})を変形した")
n_n = eq(r"n=\dfrac{1/f}{\,1/R_1+1/R_2\,}+1")
body("に焦点距離fとともに代入すると，ガラスの屈折率nが推定できる．一般にガラスの屈折率は1.4〜1.9である[1]．")

# ---------------------------------------------------------------- 3 実験方法
heading("A.3　実験手順")
heading("(1)　凸レンズの焦点距離の測定", 2)
heading("(a)　レンズの公式による方法", 3)
body("光学台の上に光源，目盛りを刻んだすりガラス板C（物体），凸レンズL，スクリーンSを並べ，高さを調節して"
     "光軸をそろえた．L とS を動かしてCの像をS上に結ばせたのち，T字棒を用いて各位置の目盛りを読んだ．"
     "T字棒をCとLの間に入れ，横棒の一端をCに触れさせたときの目盛りをx_{1}，他端をLに触れさせたときの目盛りを"
     "x_{2}とした．同様にLとSの間でLとSに触れさせたときの目盛りをx_{3}，x_{4}とした．レンズの厚さtをノギスで"
     "測り，a＝|x_{3}－x_{1}|－t/2，b＝|x_{4}－x_{2}|－t/2 として，式(" + str(n_lf) + ")からfを求めた．"
     "測定は4回行った．")
heading("(b)　ベッセルの方法", 3)
body("CとSを(a)で求めたfの4倍以上離して固定した．LをCからSに向かって動かし，S上に像ができる位置のレンズ台の"
     "目盛りをx，さらにSに近づけて再び像ができる位置の目盛りをx′とし，d＝|x′－x|とした．T字棒の左端をCに"
     "触れさせたときの目盛りをy，右端をSに触れさせたときの目盛りをy′とし，T字棒の長さlを加えて "
     "D＝|y′－y|＋l とした．式(" + str(n_bs) + ")からfを求めた．測定は4回行った．")
heading("(2)　凹レンズの焦点距離の測定", 2)
body("凸レンズLだけでCの像をつくり，そのときのスクリーン左側に触れさせたT字棒の目盛りをx_{2}とした．次に"
     "CとLを動かさずにLと像の間に凹レンズL′を入れ，スクリーンを動かして新たな像の位置B′を探した．この状態で"
     "L′の位置に対応する目盛りx_{1}と，B′の位置の目盛りx_{3}を読んだ．凹レンズの厚さt′を測り，"
     "a′＝|x_{2}－x_{1}|－t′/2，b′＝|x_{3}－x_{1}|－t′/2 として式(" + str(n_cc2) + ")からf′を求めた．測定は4回"
     "行った．")
heading("(3)　レンズのガラスの屈折率の測定", 2)
body("はじめに平面ガラスにダイヤル・レンズメータを当て，ゼロ点の読みを求めた．次に，レンズに傷をつけないよう"
     "静かにレンズメータを押し当て，凸レンズの第一面と第二面の中心付近で読みを求めた．ゼロ点の読みを"
     "差し引いた値をhとし，式(" + str(n_rh) + ")，(" + str(n_n) + ")から屈折率を求めた．")

heading("(4)　使用器具", 2)
t_eq = nxt("tab")
body(f"使用した器具を表{t_eq}に示す．")
table("使用器具", t_eq, ["器具", "数量", "備考"],
      [["光学台", "1台", "副尺つきの台を含む"], ["光源（白熱電球とレンズ）", "1個", ""],
       ["目盛りを刻んだすりガラス板 C", "1枚", "物体として使用"], ["凸レンズ L", "1個", f"厚さ t＝{T_CONVEX:.3f} cm"],
       ["凹レンズ L′", "1個", f"厚さ t′＝{T_CONCAVE:.3f} cm"], ["スクリーン S", "1枚", ""],
       ["T字棒", "1本", f"長さ l＝{L_T:.1f} cm"], ["ノギス", "1本", "レンズの厚さの測定"],
       ["ダイヤル・レンズメータ", "1台", "a＝10 mm"], ["平面ガラス", "1枚", "レンズメータのゼロ点の確認"]],
      [62, 18, 60], aligns=[C.LEFT, C.CENTER, C.LEFT])

# ---------------------------------------------------------------- 5 実験結果
heading("A.4　実験結果と考察")
heading("(1)(a)　レンズの公式による凸レンズの焦点距離", 2)
t1 = nxt("tab")
body(f"凸レンズの厚さは t＝{T_CONVEX:.3f} cm であった．各回の測定値と計算結果を表{t1}に示す．")
table("レンズの公式による凸レンズの焦点距離", t1,
      ["回", "x_{1} [cm]", "x_{2} [cm]", "x_{3} [cm]", "x_{4} [cm]", "a [cm]", "b [cm]", "1/a [cm^{-1}]",
       "1/b [cm^{-1}]", "f [cm]"],
      [[i + 1, *[f"{v:.3f}" for v in r["x"]], f"{r['a']:.2f}", f"{r['b']:.2f}", f"{r['ia']:.5f}",
        f"{r['ib']:.5f}", f"{r['f']:.2f}"] for i, r in enumerate(cv)],
      [8, 17, 17, 17, 15, 14, 14, 20, 20, 14], size=8.5)
r = cv[0]
body("1回目を例に計算過程を示す．")
eq(rf"a=|{r['x'][2]:.3f}-{r['x'][0]:.3f}|-\dfrac{{{T_CONVEX:.3f}}}{{2}}={r['a']:.3f}\ \mathrm{{cm}},\quad"
   rf"b=|{r['x'][3]:.3f}-{r['x'][1]:.3f}|-\dfrac{{{T_CONVEX:.3f}}}{{2}}={r['b']:.3f}\ \mathrm{{cm}}",
   numbered=False, size=11)
eq(rf"\dfrac{{1}}{{f}}=\dfrac{{1}}{{{r['a']:.3f}}}+\dfrac{{1}}{{{r['b']:.3f}}}={r['inv']:.5f}\ \mathrm{{cm^{{-1}}}},"
   rf"\quad f={r['f']:.2f}\ \mathrm{{cm}}", numbered=False, size=11)
body(f"4回の平均は f＝{f1:.2f} cm，標準偏差（標本標準偏差）は{s1:.2f} cm であった．よって，レンズの公式による"
     f"凸レンズの焦点距離は f＝({f1:.2f}±{s1:.2f}) cm である．レンズの公称の焦点距離 {F_NOM:.1f} cm に対する"
     f"相対誤差は")
eq(rf"\dfrac{{|{f1:.2f}-{F_NOM:.1f}|}}{{{F_NOM:.1f}}}\times100={abs(err(f1)):.1f}\ \%", numbered=False, size=11)
body("であった．")
f_inv = nxt("fig")
body(f"また，各回の1/aと1/bを図{f_inv}に示す．直線は，求めた f を用いたレンズの公式 1/b＝1/f－1/a を表す．")
C.chart_figure(chart_inverse(), "凸レンズにおける 1/a と 1/b の関係", f_inv, 120)

heading("(1)(b)　ベッセルの方法による凸レンズの焦点距離", 2)
t2 = nxt("tab")
body(f"T字棒の長さは l＝{L_T:.1f} cm である．測定値と計算結果を表{t2}に示す．")
table("ベッセルの方法による凸レンズの焦点距離", t2,
      ["回", "x [cm]", "x′ [cm]", "y [cm]", "y′ [cm]", "D [cm]", "d [cm]", "f [cm]"],
      [[i + 1, *[f"{v:.3f}" for v in r["x"]], f"{r['D']:.3f}", f"{r['d']:.3f}", f"{r['f']:.2f}"]
       for i, r in enumerate(bs)], [10, 20, 20, 20, 20, 20, 18, 16], size=9)
r = bs[0]
body("1回目を例に計算過程を示す．")
eq(rf"D=|{r['x'][3]:.3f}-{r['x'][2]:.3f}|+{L_T:.1f}={r['D']:.3f}\ \mathrm{{cm}},\quad"
   rf"d=|{r['x'][1]:.3f}-{r['x'][0]:.3f}|={r['d']:.3f}\ \mathrm{{cm}}", numbered=False, size=11)
eq(rf"f=\dfrac{{{r['D']:.3f}^2-{r['d']:.3f}^2}}{{4\times{r['D']:.3f}}}={r['f']:.2f}\ \mathrm{{cm}}",
   numbered=False, size=11)
body(f"4回の平均は f＝{f2:.2f} cm，標準偏差は{s2:.2f} cm であった．よって，ベッセルの方法による凸レンズの"
     f"焦点距離は f＝({f2:.2f}±{s2:.2f}) cm であり，公称値 {F_NOM:.1f} cm に対する相対誤差は"
     f"{abs(err(f2)):.1f} %であった．")

heading("(2)　凹レンズの焦点距離", 2)
t3 = nxt("tab")
body(f"凹レンズの厚さは t′＝{T_CONCAVE:.3f} cm であった．測定値と計算結果を表{t3}に示す．")
table("凹レンズの焦点距離", t3,
      ["回", "x_{1} [cm]", "x_{2} [cm]", "x_{3} [cm]", "a′ [cm]", "b′ [cm]", "1/a′ [cm^{-1}]",
       "1/b′ [cm^{-1}]", "f′ [cm]"],
      [[i + 1, *[f"{v:.3f}" for v in r["x"]], f"{r['a']:.2f}", f"{r['b']:.2f}", f"{r['ia']:.5f}",
        f"{r['ib']:.5f}", f"{r['f']:.2f}"] for i, r in enumerate(cc)],
      [8, 18, 18, 18, 15, 15, 21, 21, 15], size=8.5)
r = cc[0]
body("1回目を例に計算過程を示す．")
eq(rf"a'=|{r['x'][1]:.3f}-{r['x'][0]:.3f}|-\dfrac{{{T_CONCAVE:.3f}}}{{2}}={r['a']:.3f}\ \mathrm{{cm}},\quad"
   rf"b'=|{r['x'][2]:.3f}-{r['x'][0]:.3f}|-\dfrac{{{T_CONCAVE:.3f}}}{{2}}={r['b']:.3f}\ \mathrm{{cm}}",
   numbered=False, size=11)
eq(rf"\dfrac{{1}}{{f'}}=\dfrac{{1}}{{{r['a']:.3f}}}-\dfrac{{1}}{{{r['b']:.3f}}}={r['inv']:.5f}\ \mathrm{{cm^{{-1}}}},"
   rf"\quad f'={r['f']:.2f}\ \mathrm{{cm}}", numbered=False, size=11)
body(f"4回の平均は f′＝{f3:.2f} cm，標準偏差は{s3:.2f} cm であった．よって，凹レンズの焦点距離は "
     f"f′＝({f3:.1f}±{s3:.1f}) cm であり，公称値 {F_NOM:.1f} cm に対する相対誤差は{abs(err(f3)):.1f} %であった．")
f_cmp = nxt("fig")
body(f"3つの測定で得た各回の焦点距離を図{f_cmp}に示す．")
C.chart_figure(chart_compare(), "各回の測定で得た焦点距離", f_cmp, 120)

heading("(3)　凸レンズのガラスの屈折率", 2)
t4 = nxt("tab")
body(f"ダイヤル・レンズメータの読みを表{t4}に示す．平面ガラスでの読み（ゼロ点）は{H_ZERO:.3f} mm であり，"
     "これを各面の読みから差し引いた値をhとした．")
table("ダイヤル・レンズメータの読みと曲率半径（凸レンズ）", t4,
      ["", "読み [mm]", "h（ゼロ点補正後）[mm]", "曲率半径 R [mm]"],
      [["第一面", f"{H1_RAW:.3f}", f"{h1:.3f}", f"{R1:.1f}"], ["第二面", f"{H2_RAW:.3f}", f"{h2:.3f}", f"{R2:.0f}"],
       ["平面ガラス（ゼロ点）", f"{H_ZERO:.3f}", "—", "—"]], [42, 30, 42, 36])
eq(rf"R_1=\dfrac{{1}}{{2}}\left(\dfrac{{10^2}}{{{h1:.3f}}}+{h1:.3f}\right)={R1:.1f}\ \mathrm{{mm}},\quad"
   rf"R_2=\dfrac{{1}}{{2}}\left(\dfrac{{10^2}}{{{h2:.3f}}}+{h2:.3f}\right)={R2:.0f}\ \mathrm{{mm}}",
   numbered=False, size=11)
body(f"A.4(1)(a)の f＝{f1:.2f} cm＝{f1 * 10:.1f} mm を式(" + str(n_n) + ")に代入すると，")
eq(rf"n=\dfrac{{1/{f1 * 10:.1f}}}{{1/{R1:.1f}+1/{R2:.0f}}}+1={n_lens:.3f}", numbered=False, size=11.5)
body(f"となった．ベッセルの方法による f＝{f2:.2f} cm を用いると n＝{n_bes:.3f} である．")
t5 = nxt("tab")
body(f"以上の結果を表{t5}にまとめる．")
table("実験結果のまとめ", t5, ["測定量", "結果"],
      [["凸レンズの焦点距離（レンズの公式）", f"({f1:.2f}±{s1:.2f}) cm"],
       ["凸レンズの焦点距離（ベッセルの方法）", f"({f2:.2f}±{s2:.2f}) cm"],
       ["凹レンズの焦点距離", f"({f3:.1f}±{s3:.1f}) cm"],
       ["凸レンズのガラスの屈折率", f"{n_lens:.3f}（ベッセルの値を用いると{n_bes:.3f}）"]], [78, 72],
      aligns=[C.LEFT, C.CENTER])

# ---------------------------------------------------------------- 6 考察
heading("(4)　考察", 2)
heading("(a)　2つの方法による凸レンズの焦点距離の比較", 3)
diff = (f2 - f1) / f1 * 100
closer = "ベッセルの方法" if abs(err(f2)) < abs(err(f1)) else "レンズの公式による方法"
inv_spread = (max(r["inv"] for r in cv) - min(r["inv"] for r in cv)) / st.mean([r["inv"] for r in cv]) * 100
body(f"凸レンズの焦点距離は，レンズの公式による方法で ({f1:.2f}±{s1:.2f}) cm，ベッセルの方法で "
     f"({f2:.2f}±{s2:.2f}) cm となり，両者の差は{abs(f2 - f1):.2f} cm（{diff:+.1f} %）であった．公称値 {F_NOM:.1f} cm と"
     f"比べると，レンズの公式による方法は{err(f1):+.1f} %，ベッセルの方法は{err(f2):+.1f} %で，{closer}の方が公称値に"
     f"近かった．また，レンズの公式による測定では，a を {min(r['a'] for r in cv):.1f}〜{max(r['a'] for r in cv):.1f} cm と"
     f"大きく変えたにもかかわらず，1/a＋1/b の値の違いは最大でも{inv_spread:.1f} %にとどまり，図{f_inv}の測定点はいずれも"
     f"レンズの公式の直線 1/b＝1/f－1/a の近くにある．物体の位置によらず 1/a＋1/b が一定になるというレンズの公式が"
     f"成り立っていることが確かめられた．")
def outlier(rows):
    m = st.mean([r["f"] for r in rows])
    k = max(range(len(rows)), key=lambda i: abs(rows[i]["f"] - m))
    rest = st.mean([r["f"] for i, r in enumerate(rows) if i != k])
    return k, rest
k1, rest1 = outlier(cv)
k2, rest2 = outlier(bs)
body(f"各回のばらつきを見ると，レンズの公式による方法では{k1 + 1}回目（f＝{cv[k1]['f']:.2f} cm）が，ベッセルの方法では"
     f"{k2 + 1}回目（f＝{bs[k2]['f']:.2f} cm）が平均から最も離れていた．これらを除いた3回の平均はそれぞれ"
     f" {rest1:.2f} cm，{rest2:.2f} cm である．どちらの方法でも，スクリーン上の像がもっともはっきり見えるレンズ（または"
     f"スクリーン）の位置を目で判断している．像のピントが合って見える範囲には幅があるため，その判断の差がa，b"
     f"やdに直接影響し，回ごとのばらつきの主な原因になったと考えられる．")
same_side = all(r["f"] > F_NOM for r in cv) or all(r["f"] < F_NOM for r in cv)
body("一方，レンズの公式による方法では，T字棒をレンズの「中心」に触れさせて光心の位置を決め，レンズの厚さの半分を"
     "差し引くという補正をしている．実際のレンズでは光心（主点）の位置が厚さの中央とは限らないため，この方法には"
     "系統的な誤差が残る可能性がある．"
     + (f"4回の測定値がすべて公称値より{'大きい' if cv[0]['f'] > F_NOM else '小さい'}側にあることは，偶然誤差だけでなく"
        "このような系統誤差が含まれていることを示唆している．" if same_side else "")
     + "これに対してベッセルの方法は，レンズの位置の差 d と，物体とスクリーンの距離 D だけから f を求めるので，"
     "光心の位置や厚さの影響を受けない．ただしベッセルの方法でも，D が T 字棒の長さ l を含む複数の読みの和で"
     "あるため，読みの誤差が積み重なりやすい．2つの方法の結果が互いに標準偏差程度の範囲で一致したことから，"
     "どちらの方法も凸レンズの焦点距離を十分な精度で求められると考えられる．"
     if abs(f2 - f1) < s1 + s2 else
     "一方，レンズの公式による方法では，T字棒をレンズの「中心」に触れさせて光心の位置を決め，レンズの厚さの半分を"
     "差し引くという補正をしている．実際のレンズでは光心（主点）の位置が厚さの中央とは限らないため，この方法には"
     "系統的な誤差が残る可能性がある．これに対してベッセルの方法は，レンズの位置の差 d と，物体とスクリーンの距離"
     f" D だけから f を求めるので，光心の位置や厚さの影響を受けない．したがって，今回の{abs(f2 - f1):.2f} cm の差には，"
     "レンズの公式による方法の系統誤差と，像の位置の判断のばらつきの両方が含まれていると考えられる．")

heading("(b)　凹レンズの焦点距離の精度", 3)
body(f"凹レンズの焦点距離は ({f3:.1f}±{s3:.1f}) cm で，標準偏差は凸レンズの場合"
     f"{'より大きかった' if s3 > max(s1, s2) else 'と同程度であった'}．表{t3}を見ると，"
     f"b′ は {min(r['b'] for r in cc):.1f}〜{max(r['b'] for r in cc):.1f} cm と回によって大きく変わっている．"
     f"しかし式(" + str(n_cc2) + f")からf′がa′とb′のどちらに敏感かを調べると，平均的な値（a′＝{ac:.1f} cm，"
     f"b′＝{bc:.1f} cm）で a′ が0.5 cm 大きくなると f′ は {df_da:+.1f} cm 変わるのに対し，b′ が0.5 cm 大きく"
     f"なっても f′ は {df_db:+.2f} cm しか変わらない．つまり f′ の精度は，b′ よりも a′，すなわち凸レンズだけで"
     f"できた像Bの位置をどれだけ正確に決められるかで決まる．凹レンズの測定では，凸レンズの像の位置と凹レンズを"
     f"入れた後の像の位置の2回，ピントの判断が必要であり，そのばらつきが f′ のばらつきの主な原因と考えられる．"
     f"精度を上げるには，像のピントが合う範囲の両端を読んでその中央を像の位置とするなどの工夫が有効である．")

heading("(c)　ガラスの屈折率についての検討", 3)
body(f"凸レンズのガラスの屈折率は n＝{n_lens:.3f} となり，一般的なガラスの範囲（1.4〜1.9）[1] に入った．"
     f"光学ガラスとしてよく使われるクラウンガラス（BK7，n≈{N_BK7}）[3] と比べると差は"
     f"{(n_lens - N_BK7) / N_BK7 * 100:+.1f} % であり，本実験のレンズもクラウンガラス系のガラスでできていると"
     f"考えられる．")
body(f"この推定値がどの測定に左右されるかを調べた．レンズメータの読み h_{{1}} が0.005 mm 大きいと n は"
     f"{dn_h1:+.3f}，h_{{2}} が0.005 mm 大きいと{dn_h2:+.3f}，焦点距離が標準偏差1つ分（{s1:.2f} cm）大きいと"
     f"{dn_f:+.4f} 変化する．"
     + ("焦点距離のばらつきの影響は小さく，屈折率の精度はレンズメータの読みで決まることが分かる．"
        if abs(dn_f) < 0.5 * abs(dn_h1) else
        "焦点距離のばらつきとレンズメータの読みの誤差は，屈折率に同じ程度の影響を与えることが分かる．")
     + f"第二面はほぼ平面（R_{{2}}≈{R2 / 1000:.1f} m）で読みが {h2:.3f} mm と小さいため，R_{{2}} 自体の"
     f"相対誤差は非常に大きい．しかし h が a に比べて小さいとき，式(" + str(n_rh) + f")は 1/R≈2h/a^{{2}} と"
     f"近似でき，1/R は h に比例する．そのため読みの絶対誤差が n に与える影響は，第一面と第二面で同じ大きさに"
     f"なる．とくにゼロ点の読みは両面の h に同じ向きに効くので，ゼロ点が0.005 mm ずれると n は約"
     f"{2 * dn_h1:+.3f} 変わる．したがって，ゼロ点補正を丁寧に行い，各面を複数の場所で測定して平均することが，"
     f"屈折率の精度を上げるうえで重要である．なお，ここで得た n は光の屈折から直接測ったものではなく，焦点距離"
     f"と曲率半径の測定値から式(" + str(n_n) + ")を通じて間接的に求めたものである．")

heading("A.5　課題")
para("(1)　δ＝(i－r)＋(i′－r′)＝(n－1)(r＋r′)，r＋r′＝θ の導出", sa=2, keep=True)
body("頂角θのプリズムの第一面に入射角iで入った光が屈折角rで屈折し，第二面に入射角r′で達して屈折角i′で出ていく"
     "とする．第一面で光が曲がる角は i－r，第二面で曲がる角は i′－r′ なので，全体のふれの角は "
     "δ＝(i－r)＋(i′－r′) である．プリズム内の光線と2つの面の法線がつくる三角形を考えると，2つの法線のなす角は"
     "頂角θに等しく，三角形の内角の関係から r＋r′＝θ が成り立つ．入射角・屈折角が小さいとき，屈折の法則 "
     "sin i＝n sin r，sin i′＝n sin r′ は i≈nr，i′≈nr′ と近似できるので，")
eq(r"\delta=(nr-r)+(nr'-r')=(n-1)(r+r')=(n-1)\,\theta", numbered=False, size=11.5)
body("となり，式(" + str(n_dl) + ")が得られる．")
para("(2)　f＝30 cm の凸レンズの前方 5，10，15 cm に置いた物体の像", sa=2, keep=True)
rows = []
for a_ in (5, 10, 15):
    b_ = a_ * 30 / (30 - a_)
    rows.append([f"{a_}", f"{b_:.0f}", f"{b_ / a_:.1f}"])
body("物体が焦点より内側にあるとき，像はレンズの前方（物体と同じ側）にできる正立の虚像になる．虚像の位置を"
     "レンズの前方にbとすると，課題(4)で示すように 1/a－1/b＝1/f，倍率は m＝b/a＝f/(f－a) である．"
     "これを用いた結果を次に示す．")
t6 = nxt("tab")
table("f＝30 cm の凸レンズによる虚像", t6, ["物体の位置 a [cm]", "虚像の位置 b [cm]（レンズの前方）", "倍率 m"],
      rows, [44, 62, 30])
body("物体を焦点に近づけるほど，虚像は遠くにでき，倍率も大きくなる．これが虫眼鏡（ルーペ）の原理である．")
para("(3)　式(" + str(n_cc) + ")の導出", sa=2, keep=True)
body("凹レンズでは，光軸から高さhの位置のプリズムが凸レンズとは逆向き（頂角が光軸と反対側）になるため，光は"
     "光軸から遠ざかる向きに δ＝(n－1)h(1/R_{1}＋1/R_{2}) だけ曲げられる．物体Aから距離aの位置で高さhのところに"
     "入った光が，レンズの前方の距離bの点Bから出たように見えるとすると，Aから来た光線の傾きは h/a，Bから出た"
     "ように見える光線の傾きは h/b であり，ふれの角は δ＝h/b－h/a となる．よって")
eq(r"\dfrac{1}{b}-\dfrac{1}{a}=(n-1)\left(\dfrac{1}{R_1}+\dfrac{1}{R_2}\right)", numbered=False, size=11.5)
body("が得られる．平行光線（a→∞）が入ったとき光が出てくるように見える点が焦点であり，その距離を f とすると "
     "1/f＝(n－1)(1/R_{1}＋1/R_{2}) となるので，1/b－1/a＝1/f が成り立つ．")
para("(4)　凸レンズによる虚像の公式と倍率", sa=2, keep=True)
body("物体が焦点の内側（a＜f）にあるとき，レンズを出た光は広がり，レンズの前方の距離bの点から出たように見える．"
     "高さhを通る光の傾きの変化を考えると，入射光の傾きは h/a（光軸から離れる向き），出射光は前方の虚像点から"
     "出たように見えるので傾き h/b（離れる向き）である．凸レンズは光を光軸側へ δ＝h/f だけ曲げるので "
     "h/a－h/b＝h/f，すなわち 1/a－1/b＝1/f となる．これを b について解くと b＝af/(f－a) であり，物体と虚像は"
     "光軸の同じ側にあるので像は正立し，倍率は")
eq(r"m=\dfrac{b}{a}=\dfrac{f}{f-a}=\dfrac{f+b}{f}", numbered=False, size=11.5)
body("となる（最後の変形は 1/a＝1/b＋1/f＝(f＋b)/(bf) から a＝bf/(f＋b) を用いた）．")
para("(5)　ズームレンズの原理", sa=2, keep=True)
body("ズームレンズは，複数のレンズ群の間隔を変えることで，像の位置（ピント）を保ったまま全体の焦点距離を連続的"
     "に変えられるレンズである．焦点距離 f_{1}，f_{2} の2枚の薄レンズを間隔 e で並べたときの合成焦点距離 F は "
     "1/F＝1/f_{1}＋1/f_{2}－e/(f_{1}f_{2}) で表され，e を変えると F が変わる[2]．実際のズームレンズでは，焦点距離を"
     "変える変倍群と，それに伴うピントのずれを補正する補正群を，カムなどで連動して動かしている．")
para("(6)　アッベ数", sa=2, keep=True)
body("ガラスの屈折率は光の波長によって異なり（分散），そのためにレンズの焦点距離が色によって変わる色収差が"
     "生じる．アッベ数 ν_{d} は分散の小ささを表す量で，")
eq(r"\nu_d=\dfrac{n_d-1}{n_F-n_C}", numbered=False, size=11.5)
body("で定義される[3]．ここで n_{d}，n_{F}，n_{C} はそれぞれ波長587.6 nm（黄色），486.1 nm（青），656.3 nm（赤）"
     "の光に対する屈折率である．アッベ数が大きいほど分散が小さく色収差が出にくい．クラウンガラスはアッベ数が"
     "大きく，フリントガラスは小さいので，両者を組み合わせた色消しレンズ（アクロマート）では色収差を打ち消す"
     "ことができる．そのため光学レンズの設計では，屈折率とともにアッベ数が重要になる．")

heading("A.6　まとめ")
body(f"薄い凸レンズの焦点距離をレンズの公式とベッセルの方法の2通りで測定し，それぞれ {f1:.2f} cm，{f2:.2f} cm を"
     f"得た．1/a と 1/b の関係から，レンズの公式が成り立つことを確かめられた．凹レンズの焦点距離は {f3:.1f} cm，"
     f"凸レンズのガラスの屈折率は {n_lens:.3f} となり，クラウンガラスに近い値であった．どの測定でも，像のピントが"
     f"合う位置の判断が結果を左右しており，とくに凹レンズでは像の位置の決め方を工夫する必要があると感じた．"
     f"また，同じ量を異なる方法で測って比べることで，測定の偶然誤差と系統誤差を区別して考えられることが分かった．")

C.references([
    "[1] 東京理科大学，“物理学実験 実験指導書 5．薄レンズの焦点距離の測定，”2026．",
    "[2] E. Hecht, “Optics,” 5th ed., Pearson, 2017.",
    "[3] 国立天文台 編，“理科年表 2024，”丸善出版，2023．",
])

if STANDALONE:
    out = os.path.join(HERE, "薄レンズの焦点距離の測定_報告書.docx")
    C.save(out)
    print("saved:", out)
print(f"f1={f1:.3f}±{s1:.3f} f2={f2:.3f}±{s2:.3f} f3={f3:.3f}±{s3:.3f} n={n_lens:.4f} nB={n_bes:.4f}")
print(f"dn_h1={dn_h1:.4f} dn_h2={dn_h2:.4f} dn_f={dn_f:.5f} df_da={df_da:.3f} df_db={df_db:.3f}")
