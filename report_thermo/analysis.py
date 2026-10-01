# -*- coding: utf-8 -*-
"""熱電対の検定実験 — 計算と図の作成。
data モジュールを受け取り、報告書に載せる数値・表・図をすべて計算する。"""
import math
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
import numpy as np

JP_FONT = "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"
fm.fontManager.addfont(JP_FONT)
plt.rcParams["font.family"] = fm.FontProperties(fname=JP_FONT).get_name()
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["mathtext.fontset"] = "dejavuserif"

MELT = {"Sn": 231.97, "Pb": 327.52, "Zn": 419.58}           # 融点 [℃]（テキスト記載値）
NAME_JP = {"Sn": "スズ", "Pb": "鉛", "Zn": "亜鉛"}
ORDER = ["Sn", "Pb", "Zn"]
PLATEAU_THR = 0.02                                             # 冷却時プラトー判定のしきい値 [mV/15 s]
HEAT_THR = 0.05                                                # 加熱時に上昇が鈍った区間の判定 [mV/15 s]

# ---- K熱電対の基準熱起電力（NIST ITS-90, 0〜1372 ℃, E[mV], t[℃]） ----
_C = [-0.176004136860E-01, 0.389212049750E-01, 0.185587700320E-04, -0.994575928740E-07,
      0.318409457190E-09, -0.560728448890E-12, 0.560750590590E-15, -0.320207200030E-18,
      0.971511471520E-22, -0.121047212750E-25]
_A0, _A1, _A2 = 0.118597600000E+00, -0.118343200000E-03, 0.126968600000E+03


def emf_std(t):
    """基準熱起電力 [mV]（基準接点 0 ℃）。"""
    return sum(c * t ** i for i, c in enumerate(_C)) + _A0 * math.exp(_A1 * (t - _A2) ** 2)


def seebeck_std(t):
    """基準のゼーベック係数 dE/dt [mV/℃]。"""
    d = sum(i * c * t ** (i - 1) for i, c in enumerate(_C) if i > 0)
    return d + _A0 * math.exp(_A1 * (t - _A2) ** 2) * 2 * _A1 * (t - _A2)


def tol_class2(t):
    """K熱電対 クラス2 の許容差 [℃]（JIS C 1602 / IEC 60584-1）: ±2.5 ℃ または ±0.0075|t| の大きい方。"""
    return max(2.5, 0.0075 * abs(t))


# ---------------------------------------------------------------- データ判定
def split_run(D, m):
    """(冷却データ, 加熱データ, 最大点の添字) を返す。RAW があれば最大点で分割する。"""
    raw = getattr(D, "RAW", None) or {}
    if raw.get(m):
        E = list(raw[m])
        ip = max(range(len(E)), key=lambda i: E[i])
        return E[ip:], E[:ip + 1], ip
    return list(D.COOLING[m]), None, None


def is_complete(D):
    raw = getattr(D, "RAW", None) or {}
    cool = getattr(D, "COOLING", None) or {}
    return (D.ROOM_TEMP is not None and D.ROOM_EMF is not None and D.ZERO_EMF is not None
            and all(raw.get(m) or cool.get(m) for m in ORDER))


# ---------------------------------------------------------------- プラトー判定
def find_plateau(E, dt, manual=None, thr=PLATEAU_THR):
    """融点に対応する区間を返す: (i0, i1) は両端を含む添字。"""
    if manual is not None:
        i0, i1 = int(round(manual[0] / dt)), int(round(manual[1] / dt))
        return i0, i1
    best = (0, 0)
    start = 0
    for i in range(1, len(E)):
        if abs(E[i] - E[i - 1]) > thr:
            start = i
        elif i - start > best[1] - best[0]:
            best = (start, i)
    return best


def plateau_stats(E, i0, i1):
    v = np.array(E[i0:i1 + 1], dtype=float)
    n = len(v)
    mean = v.mean()
    sd = v.std(ddof=1) if n > 1 else 0.0
    return {"values": v.tolist(), "n": n, "mean": mean, "sd": sd, "se": sd / math.sqrt(n)}


# ---------------------------------------------------------------- 最小二乗法（テキストの式）
def lsq(x, y):
    x = np.asarray(x, float); y = np.asarray(y, float); N = len(x)
    m = lambda arr: float(np.mean(arr))
    xb, yb = m(x), m(y)
    x2b, x3b, x4b = m(x ** 2), m(x ** 3), m(x ** 4)
    xyb, x2yb = m(x * y), m(x ** 2 * y)
    Sx = x2b - xb ** 2
    Sxy = xyb - xb * yb
    Sx2 = x4b - x2b ** 2
    Sx2y = x2yb - x2b * yb
    Sxx2 = x3b - xb * x2b
    # 1次式
    a1 = Sxy / Sx
    b1 = yb - a1 * xb
    # 2次式
    Dq = Sx * Sx2 - Sxx2 ** 2
    a2 = (Sx * Sx2y - Sxy * Sxx2) / Dq
    b2 = (Sxy * Sx2 - Sxx2 * Sx2y) / Dq
    c2 = yb - a2 * x2b - b2 * xb
    # 検算（numpy）
    p1 = np.polyfit(x, y, 1); p2 = np.polyfit(x, y, 2)
    assert np.allclose([a1, b1], p1, rtol=1e-6, atol=1e-9)
    assert np.allclose([a2, b2, c2], p2, rtol=1e-5, atol=1e-8)
    f1 = a1 * x + b1
    f2 = a2 * x ** 2 + b2 * x + c2
    rows = [{"x": xi, "y": yi, "x2": xi ** 2, "x3": xi ** 3, "x4": xi ** 4,
             "xy": xi * yi, "x2y": xi ** 2 * yi} for xi, yi in zip(x, y)]
    return {
        "N": N, "rows": rows,
        "xb": xb, "yb": yb, "x2b": x2b, "x3b": x3b, "x4b": x4b, "xyb": xyb, "x2yb": x2yb,
        "Sx": Sx, "Sxy": Sxy, "Sx2": Sx2, "Sx2y": Sx2y, "Sxx2": Sxx2, "Dq": Dq,
        "a1": a1, "b1": b1, "a2": a2, "b2": b2, "c2": c2,
        "res1": (y - f1).tolist(), "res2": (y - f2).tolist(),
        "rss1": float(np.sum((y - f1) ** 2)), "rss2": float(np.sum((y - f2) ** 2)),
        "sum_res1": float(np.sum(y - f1)),
        "rss_mean": float(np.sum((y - yb) ** 2)),   # 水平線 y = ȳ（残差和 0 だが不適）
    }


def inv_lin(E, a, b):
    return (E - b) / a


def inv_quad(E, a, b, c):
    disc = b * b - 4 * a * (c - E)
    if a == 0:
        return (E - c) / b
    r1 = (-b + math.sqrt(disc)) / (2 * a)
    r2 = (-b - math.sqrt(disc)) / (2 * a)
    return r1 if -50 <= r1 <= 1500 else r2


# ---------------------------------------------------------------- 全体の計算
def compute(D):
    R = {"cool": {}}
    for m in ORDER:
        E, H, ip = split_run(D, m)
        i0, i1 = find_plateau(E, D.DT, (D.PLATEAU or {}).get(m))
        st = plateau_stats(E, i0, i1)
        R["cool"][m] = {"E": E, "t": [i * D.DT for i in range(len(E))], "i0": i0, "i1": i1,
                        "t0": i0 * D.DT, "t1": i1 * D.DT, **st,
                        "Emin": min(E), "Emax": max(E),
                        "supercool": _supercool(E, i0),
                        "t_peak": None if ip is None else ip * D.DT, "E_peak": max(E)}
        if H and len(H) > 3:
            h0, h1 = find_plateau(H, D.DT, thr=HEAT_THR)
            hs = plateau_stats(H, h0, h1)
            R["cool"][m]["heat"] = {"t0": h0 * D.DT, "t1": h1 * D.DT, **hs}
    pts = [("0 ℃（氷水）", 0.0, D.ZERO_EMF), ("室温", float(D.ROOM_TEMP), D.ROOM_EMF)]
    pts += [(f"{NAME_JP[m]}（{m}）の融点", MELT[m], R["cool"][m]["mean"]) for m in ORDER]
    R["points"] = pts
    x = [p[1] for p in pts]; y = [p[2] for p in pts]
    L = lsq(x, y)
    R["lsq"] = L
    comp = []
    for label, t, Em in pts:
        Es = emf_std(t)
        dE = Em - Es
        S = seebeck_std(t)
        comp.append({"label": label, "t": t, "E": Em, "Es": Es, "dE": dE, "dT": dE / S,
                     "dE0": dE - D.ZERO_EMF, "tol": tol_class2(t),
                     "rel": (dE / Es * 100) if abs(Es) > 0.5 else None,
                     "f1": L["a1"] * t + L["b1"], "f2": L["a2"] * t ** 2 + L["b2"] * t + L["c2"],
                     "t_lin": inv_lin(Em, L["a1"], L["b1"]),
                     "t_quad": inv_quad(Em, L["a2"], L["b2"], L["c2"])})
    R["comp"] = comp
    # 近似式と基準曲線の差（0〜450 ℃ の範囲）
    ts = np.linspace(0, 450, 451)
    std = np.array([emf_std(t) for t in ts])
    d1 = (L["a1"] * ts + L["b1"]) - std
    d2 = (L["a2"] * ts ** 2 + L["b2"] * ts + L["c2"]) - std
    sstd = np.array([seebeck_std(t) for t in ts])
    R["dev"] = {"max1_mV": float(np.max(np.abs(d1))), "max2_mV": float(np.max(np.abs(d2))),
                "max1_C": float(np.max(np.abs(d1 / sstd))), "max2_C": float(np.max(np.abs(d2 / sstd)))}
    # 基準曲線そのものを同じ5温度で2次近似したときの係数（曲率の比較用）
    pstd = np.polyfit(x, [emf_std(t) for t in x], 2)
    R["std_quad"] = pstd.tolist()
    # ゼーベック係数
    R["seebeck"] = {
        "lin": L["a1"],
        "quad": {t: 2 * L["a2"] * t + L["b2"] for t in (0, 200, 400)},
        "std": {t: seebeck_std(t) for t in (0, 200, 400)},
        "std_avg": (emf_std(419.58) - emf_std(0)) / 419.58,
    }
    return R


def _supercool(E, i0):
    """プラトー直前に起電力が一度下がってから上がる（過冷却→再輝）かを判定。"""
    if i0 < 2:
        return None
    seg = E[:i0 + 1]
    j = int(np.argmin(seg))
    if j < i0 and E[i0] - seg[j] > 0.05:
        return {"t_min": j, "E_min": seg[j], "rise": E[i0] - seg[j]}
    return None


# ---------------------------------------------------------------- 図
def _ax_style(ax):
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(direction="in", labelsize=9)


def fig_cooling(R, m, D, path):
    c = R["cool"][m]
    t = np.array(c["t"]); E = np.array(c["E"])
    fig, ax = plt.subplots(figsize=(6.0, 3.6))
    ax.axvspan(c["t0"], c["t1"], color="0.88", lw=0, label="融点とみなした区間")
    ax.plot(t, E, "-", color="black", lw=0.8)
    ax.plot(t, E, "o", ms=3.2, mfc="white", mec="black", mew=0.8, label="測定値")
    ax.axhline(c["mean"], color="black", ls="--", lw=0.8)
    ax.text(t[0] + (t[-1] - t[0]) * 0.02, c["mean"], f"E = {c['mean']:.3f} mV", va="bottom", ha="left", fontsize=9)
    pad = (E.max() - E.min()) * 0.08
    ax.set_ylim(E.min() - pad, E.max() + pad)
    ax.set_xlim(0, t[-1] + D.DT)
    ax.set_xlabel("時間 [s]", fontsize=10); ax.set_ylabel("熱起電力 [mV]", fontsize=10)
    ax.legend(fontsize=8, loc="upper right", framealpha=1.0, edgecolor="0.6")
    _ax_style(ax)
    fig.tight_layout(); fig.savefig(path, dpi=200); plt.close(fig)


def fig_calibration(R, path):
    L = R["lsq"]
    x = np.array([p[1] for p in R["points"]]); y = np.array([p[2] for p in R["points"]])
    ts = np.linspace(-10, 450, 300)
    fig, ax = plt.subplots(figsize=(6.0, 4.0))
    ax.plot(ts, L["a1"] * ts + L["b1"], "-", color="black", lw=1.0, label="近似直線（1次式）")
    ax.plot(ts, L["a2"] * ts ** 2 + L["b2"] * ts + L["c2"], "--", color="0.35", lw=1.2, label="近似曲線（2次式）")
    ax.plot(x, y, "o", ms=6, mfc="white", mec="black", mew=1.2, label="測定点", zorder=5)
    ax.set_xlim(-10, 450); ax.set_ylim(-1, 20)
    ax.set_xlabel("温度 t [℃]", fontsize=10); ax.set_ylabel("熱起電力 E [mV]", fontsize=10)
    ax.legend(fontsize=9, frameon=False, loc="upper left")
    _ax_style(ax)
    fig.tight_layout(); fig.savefig(path, dpi=200); plt.close(fig)


def fig_deviation(R, path):
    """基準熱起電力からのずれ（温度換算）を、測定点・1次式・2次式で比較。"""
    L = R["lsq"]
    ts = np.linspace(0, 450, 451)
    std = np.array([emf_std(t) for t in ts]); s = np.array([seebeck_std(t) for t in ts])
    d1 = ((L["a1"] * ts + L["b1"]) - std) / s
    d2 = ((L["a2"] * ts ** 2 + L["b2"] * ts + L["c2"]) - std) / s
    fig, ax = plt.subplots(figsize=(6.0, 3.6))
    ax.axhline(0, color="0.6", lw=0.8)
    ax.plot(ts, d1, "-", color="black", lw=1.0, label="近似直線（1次式）")
    ax.plot(ts, d2, "--", color="0.35", lw=1.2, label="近似曲線（2次式）")
    ax.plot([c["t"] for c in R["comp"]], [c["dT"] for c in R["comp"]], "o", ms=6,
            mfc="white", mec="black", mew=1.2, label="測定点", zorder=5)
    ax.set_xlim(0, 450)
    ax.set_xlabel("温度 t [℃]", fontsize=10); ax.set_ylabel("基準からのずれ（温度換算）[℃]", fontsize=10)
    ax.legend(fontsize=9, frameon=False)
    _ax_style(ax)
    fig.tight_layout(); fig.savefig(path, dpi=200); plt.close(fig)


# ---------------------------------------------------------------- 原理の模式図（データ不要）
def fig_principle_circuit(path):
    fig, ax = plt.subplots(figsize=(6.2, 2.6))
    ax.set_xlim(0, 10); ax.set_ylim(0, 4.2); ax.axis("off")
    # 測温接点
    ax.plot([1.2], [2.1], "o", color="black", ms=7)
    ax.text(1.2, 1.45, "測温接点\n温度 T", ha="center", va="top", fontsize=9)
    # 素線A(+), B(-)
    ax.plot([1.2, 5.0], [2.1, 3.0], color="black", lw=1.6)
    ax.plot([1.2, 5.0], [2.1, 1.2], color="black", lw=1.6, ls=(0, (5, 2)))
    ax.text(3.0, 2.85, "金属A（クロメル, +）", fontsize=8.5, ha="center", rotation=13)
    ax.text(3.0, 1.15, "金属B（アルメル, −）", fontsize=8.5, ha="center", rotation=-13)
    # 基準接点ボックス
    ax.add_patch(plt.Rectangle((4.7, 0.7), 1.2, 2.8, fill=False, lw=1.0, ls=":"))
    ax.text(5.3, 3.65, "基準接点\n$T_0$ = 0 ℃（氷水）", ha="center", va="bottom", fontsize=8.5)
    ax.plot([5.0], [3.0], "s", color="black", ms=4); ax.plot([5.0], [1.2], "s", color="black", ms=4)
    # 銅線
    ax.plot([5.0, 8.0, 8.0], [3.0, 3.0, 2.45], color="0.45", lw=1.4)
    ax.plot([5.0, 8.0, 8.0], [1.2, 1.2, 1.75], color="0.45", lw=1.4)
    ax.text(6.6, 3.1, "銅線（金属C）", fontsize=8.5, ha="center")
    ax.add_patch(plt.Circle((8.0, 2.1), 0.38, fill=False, lw=1.3))
    ax.text(8.0, 2.1, "V", ha="center", va="center", fontsize=11)
    ax.text(8.55, 2.1, "電圧計\n（熱起電力 E）", fontsize=8.5, va="center")
    fig.tight_layout(); fig.savefig(path, dpi=220); plt.close(fig)


def fig_principle_cooling(path):
    """冷却曲線の模式図（過冷却あり／なし）。"""
    t = np.linspace(0, 100, 600)

    def curve(t, sc):
        Tm = 50.0
        y = np.where(t < 30, Tm + 25 * np.exp(-t / 14) - 25 * np.exp(-30 / 14) + 0, 0)
        y = np.where(t < 30, Tm + (25 - 25 * (t / 30) ** 0.8), y)
        if sc:
            y = np.where((t >= 30) & (t < 38), Tm - 6 * np.sin(np.pi * (t - 30) / 16), y)
            y = np.where((t >= 38) & (t < 42), Tm - 6 + 6 * (t - 38) / 4, y)
            y = np.where((t >= 42) & (t < 70), Tm, y)
        else:
            y = np.where((t >= 30) & (t < 70), Tm, y)
        y = np.where(t >= 70, Tm - 22 * (1 - np.exp(-(t - 70) / 18)), y)
        return y
    fig, ax = plt.subplots(figsize=(6.0, 3.2))
    ax.plot(t, curve(t, False), color="black", lw=1.3, label="過冷却なし")
    ax.plot(t, curve(t, True), color="0.4", lw=1.1, ls="--", label="過冷却あり")
    ax.axhline(50, color="0.6", lw=0.7, ls=":")
    ax.text(101, 50, "融点", va="center", fontsize=9)
    ax.annotate("液相のみ", (12, 64), fontsize=9)
    ax.annotate("固相と液相の共存\n（温度一定）", (46, 52.5), fontsize=9, ha="center")
    ax.annotate("固相のみ", (84, 38), fontsize=9)
    ax.annotate("過冷却", xy=(37, 44.5), xytext=(18, 40), fontsize=9,
                arrowprops=dict(arrowstyle="->", lw=0.8))
    ax.set_xticks([]); ax.set_yticks([])
    ax.set_xlabel("時間", fontsize=10); ax.set_ylabel("温度（熱起電力）", fontsize=10)
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(fontsize=8.5, frameon=False, loc="upper right")
    ax.set_xlim(0, 108)
    fig.tight_layout(); fig.savefig(path, dpi=220); plt.close(fig)


def fig_setup(path):
    """実験装置の結線（ブロック図）。"""
    fig, ax = plt.subplots(figsize=(6.4, 3.0))
    ax.set_xlim(0, 12); ax.set_ylim(0, 5.4); ax.axis("off")

    def box(x, y, w, h, text, ls="-"):
        ax.add_patch(plt.Rectangle((x, y), w, h, fill=False, lw=1.1, ls=ls))
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=8.5)
    box(0.2, 3.6, 1.6, 1.0, "コンセント\n（AC100 V）")
    box(2.4, 3.6, 1.8, 1.0, "スライダック\n（130 V, 10 A）")
    box(4.8, 3.6, 1.8, 1.0, "交流電流計\n（10 A レンジ）")
    box(7.3, 2.6, 2.2, 2.4, "")
    ax.text(8.4, 4.75, "電気炉", ha="center", fontsize=8.5)
    box(7.6, 2.85, 1.6, 1.6, "タンマン管\n＋試料\n＋保護管", ls="--")
    box(7.3, 0.2, 2.2, 1.4, "デュワービン\n（氷水 0 ℃）")
    box(10.1, 0.25, 1.7, 1.3, "デジタル\nボルトメータ")
    for x0, x1 in [(1.8, 2.4), (4.2, 4.8)]:
        ax.annotate("", xy=(x1, 4.1), xytext=(x0, 4.1), arrowprops=dict(arrowstyle="-", lw=1.0))
    ax.annotate("", xy=(7.3, 4.1), xytext=(6.6, 4.1), arrowprops=dict(arrowstyle="-", lw=1.0))
    ax.text(6.95, 4.25, "ヒータ", fontsize=7.5, ha="center")
    ax.plot([8.25, 8.25], [2.85, 1.6], color="black", lw=1.4)
    ax.plot([8.55, 8.55], [2.85, 1.6], color="black", lw=1.4, ls=(0, (4, 2)))
    ax.text(8.75, 2.2, "熱電対\n（＋クロメル／−アルメル）", fontsize=7.5, va="center")
    ax.plot([9.5, 10.1], [0.9, 0.9], color="0.45", lw=1.3)
    ax.text(9.8, 1.0, "銅線", fontsize=7.5, ha="center", va="bottom")
    fig.tight_layout(); fig.savefig(path, dpi=220); plt.close(fig)


def make_static_figs(imgdir):
    os.makedirs(imgdir, exist_ok=True)
    fig_principle_circuit(os.path.join(imgdir, "fig_circuit.png"))
    fig_principle_cooling(os.path.join(imgdir, "fig_cool_schematic.png"))
    fig_setup(os.path.join(imgdir, "fig_setup.png"))


def make_data_figs(R, D, imgdir):
    for m in ORDER:
        fig_cooling(R, m, D, os.path.join(imgdir, f"fig_cool_{m}.png"))
    fig_calibration(R, os.path.join(imgdir, "fig_calibration.png"))
    fig_deviation(R, os.path.join(imgdir, "fig_deviation.png"))
