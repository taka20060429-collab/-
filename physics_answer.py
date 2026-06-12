import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch
from matplotlib import rcParams
import matplotlib.font_manager as fm

# IPAGothic font
font_path = '/usr/share/fonts/truetype/fonts-japanese-gothic.ttf'
jp_font = fm.FontProperties(fname=font_path)
jp_font_bold = fm.FontProperties(fname=font_path, weight='bold')

rcParams['mathtext.fontset'] = 'cm'

fig = plt.figure(figsize=(8.27, 11.69))  # A4
fig.patch.set_facecolor('white')

ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis('off')

def jp(ax, x, y, text, size=11, color='black', ha='left', va='top', bold=False):
    fp = jp_font_bold if bold else jp_font
    ax.text(x, y, text, fontproperties=fp, fontsize=size,
            color=color, ha=ha, va=va, transform=ax.transAxes)

def math(ax, x, y, text, size=13, color='black', ha='left', va='top'):
    ax.text(x, y, text, fontsize=size, color=color,
            ha=ha, va=va, transform=ax.transAxes,
            usetex=False)

# Title
jp(ax, 0.5, 0.97, '物理学1（電気工学科）模範解答', size=15, ha='center', bold=True)
jp(ax, 0.5, 0.935, '質量mの質点を鉛直上向きに初速v0で投げ上げた（速度に比例する空気抵抗あり）', size=10, ha='center', color='#444444')

# Divider
ax.axhline(y=0.915, xmin=0.05, xmax=0.95, color='#333333', linewidth=1.5)

y = 0.895

# ── Section 1: 座標設定 ──
rect = FancyBboxPatch((0.04, y-0.005), 0.92, 0.032,
                      boxstyle="round,pad=0.005", linewidth=0,
                      facecolor='#1a4a7a', edgecolor='none')
ax.add_patch(rect)
jp(ax, 0.06, y+0.021, '【1】 座標設定', size=11, color='white', bold=True)
y -= 0.015

jp(ax, 0.06, y, '・ 鉛直上向きを正の方向とする', size=10.5)
y -= 0.028
jp(ax, 0.06, y, '・ 投げ上げた点を原点  x = 0  とする', size=10.5)
y -= 0.028
jp(ax, 0.06, y, '・ 初期条件：  t = 0  で  v = v0 > 0', size=10.5)
y -= 0.035

# ── Section 2: 運動方程式 ──
rect2 = FancyBboxPatch((0.04, y-0.005), 0.92, 0.032,
                       boxstyle="round,pad=0.005", linewidth=0,
                       facecolor='#1a4a7a', edgecolor='none')
ax.add_patch(rect2)
jp(ax, 0.06, y+0.021, '【2】 力の定義と運動方程式', size=11, color='white', bold=True)
y -= 0.015

jp(ax, 0.06, y, '作用する力：', size=10.5)
y -= 0.028
jp(ax, 0.08, y, '① 重力：大きさ mg，下向き（負方向）→  −mg', size=10.5)
y -= 0.028
jp(ax, 0.08, y, '② 空気抵抗：速度 v に比例，運動と逆向きに作用 →  −kv', size=10.5)
y -= 0.03

jp(ax, 0.06, y, 'Newton の運動の第2法則  F = ma  を適用すると：', size=10.5)
y -= 0.04

# Boxed equation
eq_box = FancyBboxPatch((0.25, y-0.012), 0.50, 0.038,
                        boxstyle="round,pad=0.008", linewidth=1.5,
                        facecolor='#f0f4ff', edgecolor='#1a4a7a')
ax.add_patch(eq_box)
ax.text(0.50, y+0.007, r'$m\,\dfrac{dv}{dt} = -mg - kv$',
        fontsize=14, ha='center', va='center', transform=ax.transAxes)
y -= 0.058

# ── Section 3: 速度 ──
rect3 = FancyBboxPatch((0.04, y-0.005), 0.92, 0.032,
                       boxstyle="round,pad=0.005", linewidth=0,
                       facecolor='#1a4a7a', edgecolor='none')
ax.add_patch(rect3)
jp(ax, 0.06, y+0.021, '【3】 速度 v(t) の導出（変数分離法）', size=11, color='white', bold=True)
y -= 0.018

ax.text(0.12, y, r'$\dfrac{dv}{mg + kv} = -\dfrac{dt}{m}$',
        fontsize=12, ha='left', va='top', transform=ax.transAxes)
y -= 0.048

jp(ax, 0.06, y, '両辺を積分し，初期条件  t = 0，v = v0  を代入：', size=10.5)
y -= 0.042

eq_box2 = FancyBboxPatch((0.10, y-0.015), 0.80, 0.045,
                         boxstyle="round,pad=0.008", linewidth=1.5,
                         facecolor='#fff8e0', edgecolor='#c8860a')
ax.add_patch(eq_box2)
ax.text(0.50, y+0.008, r'$v(t) = \left(v_0 + \dfrac{mg}{k}\right)e^{-\frac{k}{m}t} - \dfrac{mg}{k}$',
        fontsize=14, ha='center', va='center', transform=ax.transAxes)
y -= 0.065

# ── Section 4: 位置 ──
rect4 = FancyBboxPatch((0.04, y-0.005), 0.92, 0.032,
                       boxstyle="round,pad=0.005", linewidth=0,
                       facecolor='#1a4a7a', edgecolor='none')
ax.add_patch(rect4)
jp(ax, 0.06, y+0.021, '【4】 位置 x(t) の導出', size=11, color='white', bold=True)
y -= 0.018

jp(ax, 0.06, y, 'v(t) = dx/dt を積分し，t = 0 で x = 0 を代入：', size=10.5)
y -= 0.042

eq_box3 = FancyBboxPatch((0.05, y-0.015), 0.90, 0.045,
                         boxstyle="round,pad=0.008", linewidth=1.5,
                         facecolor='#fff8e0', edgecolor='#c8860a')
ax.add_patch(eq_box3)
ax.text(0.50, y+0.008,
        r'$x(t) = \dfrac{m}{k}\!\left(v_0 + \dfrac{mg}{k}\right)\!\left(1 - e^{-\frac{k}{m}t}\right) - \dfrac{mg}{k}\,t$',
        fontsize=13, ha='center', va='center', transform=ax.transAxes)
y -= 0.065

# ── Section 5: 物理的考察 ──
rect5 = FancyBboxPatch((0.04, y-0.005), 0.92, 0.032,
                       boxstyle="round,pad=0.005", linewidth=0,
                       facecolor='#1a4a7a', edgecolor='none')
ax.add_patch(rect5)
jp(ax, 0.06, y+0.021, '【5】 物理的考察', size=11, color='white', bold=True)
y -= 0.018

# Sub: 最高点
jp(ax, 0.06, y, '① 最高点に達する時刻  （v(t1) = 0 とおく）', size=10.5, bold=True)
y -= 0.035

ax.text(0.30, y,
        r'$t_1 = \dfrac{m}{k}\ln\!\left(1 + \dfrac{kv_0}{mg}\right)$',
        fontsize=12, ha='left', va='top', transform=ax.transAxes)
y -= 0.048

jp(ax, 0.08, y, '→ 空気抵抗なし（t1 = v0/g）より小さい：早く止まる', size=10.5, color='#1a4a7a')
y -= 0.033

# Sub: 終端速度
jp(ax, 0.06, y, '② 終端速度（terminal velocity）', size=10.5, bold=True)
y -= 0.028
jp(ax, 0.08, y, '落下し続けると  t → ∞  で：', size=10.5)
y -= 0.032

ax.text(0.35, y, r'$v \;\longrightarrow\; -\,\dfrac{mg}{k}$',
        fontsize=12, ha='left', va='top', transform=ax.transAxes)
y -= 0.042
jp(ax, 0.08, y, '→ 重力と空気抵抗が釣り合い，一定速度で落下する', size=10.5, color='#1a4a7a')
y -= 0.035

# Comparison table
jp(ax, 0.06, y, '③ 空気抵抗あり／なしの比較', size=10.5, bold=True)
y -= 0.028

table_top = y
col = [0.06, 0.35, 0.68]
row_h = 0.028

headers = ['項目', '空気抵抗あり', '空気抵抗なし']
rows = [
    ['最高点到達時刻', 'v0/g より小さい', 'v0/g'],
    ['最高点の高さ', 'v0²/2g より低い', 'v0²/2g'],
    ['落下速度', '終端速度 mg/k に収束', '無限に増加'],
]

for ci, h in enumerate(headers):
    rect_h = patches.Rectangle((col[ci]-0.01, table_top-row_h+0.003), 0.32, row_h,
                                facecolor='#2c6fad', edgecolor='white', linewidth=0.5)
    ax.add_patch(rect_h)
    jp(ax, col[ci]+0.01, table_top-0.003, h, size=9.5, color='white', bold=True)

for ri, row in enumerate(rows):
    bg = '#f5f8ff' if ri % 2 == 0 else 'white'
    ry = table_top - row_h * (ri + 1)
    for ci, cell in enumerate(row):
        rect_c = patches.Rectangle((col[ci]-0.01, ry-row_h+0.003), 0.32, row_h,
                                   facecolor=bg, edgecolor='#cccccc', linewidth=0.5)
        ax.add_patch(rect_c)
        jp(ax, col[ci]+0.01, ry-0.003, cell, size=9.5)

y = table_top - row_h * (len(rows) + 1) - 0.015

# Footer
ax.axhline(y=y, xmin=0.05, xmax=0.95, color='#333333', linewidth=1.0)
y -= 0.02
jp(ax, 0.5, y, '物理学1（電気工学科）　中間試験対策　模範解答', size=9,
   ha='center', color='#666666')

plt.savefig('/home/user/-/physics_model_answer.pdf',
            format='pdf', bbox_inches='tight', dpi=200,
            facecolor='white')
plt.close()
print("PDF generated successfully.")
