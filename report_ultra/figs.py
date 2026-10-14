"""模式図（ブラッグ反射）を描く。"""
import os
import sys

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "report_optics"))
import common  # noqa: E402,F401  （日本語フォントの設定）
import matplotlib.pyplot as plt  # noqa: E402

OUT = os.path.join(HERE, "img")


def bragg():
    fig, ax = plt.subplots(figsize=(5.2, 3.0))
    d = 1.0
    for j in range(3):
        yv = -j * d
        ax.plot(np.arange(-2, 2.01, 1.0), [yv] * 5, "ko", ms=6)
        ax.plot([-2.6, 2.6], [yv, yv], color="0.7", lw=0.7)
    th = np.radians(30)
    for j in range(2):
        yv = -j * d
        P = np.array([0, yv])
        A = P + 3.0 * np.array([-np.cos(th), np.sin(th)])
        B = P + 3.0 * np.array([np.cos(th), np.sin(th)])
        ax.annotate("", P, A, arrowprops=dict(arrowstyle="->", lw=1.1))
        ax.annotate("", B, P, arrowprops=dict(arrowstyle="->", lw=1.1))
    t = np.linspace(np.pi - th, np.pi, 20)
    ax.plot(0.8 * np.cos(t), 0.8 * np.sin(t), "k", lw=0.8)
    ax.text(-1.15, 0.12, r"$\theta$", fontsize=11)
    ax.annotate("", (2.3, -1), (2.3, 0), arrowprops=dict(arrowstyle="<->", lw=0.9))
    ax.text(2.4, -0.55, r"$d$", fontsize=11)
    ax.text(-0.1, -1.55, r"経路差 $2d\sin\theta$", fontsize=10, ha="center")
    ax.set_xlim(-2.9, 2.9)
    ax.set_ylim(-2.3, 1.7)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, "fig_bragg.png"), dpi=250)
    plt.close(fig)


bragg()
