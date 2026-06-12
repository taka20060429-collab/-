import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib import rcParams
import matplotlib.font_manager as fm

font_path = '/usr/share/fonts/truetype/fonts-japanese-gothic.ttf'
jp_font = fm.FontProperties(fname=font_path)
jp_font_bold = fm.FontProperties(fname=font_path, weight='bold')
rcParams['mathtext.fontset'] = 'cm'

ML = 0.05   # left margin
MB = 0.04   # bottom margin

class R:
    def __init__(self, path):
        self.pdf = PdfPages(path)
        self._np()

    def _np(self):
        self.fig = plt.figure(figsize=(8.27, 11.69))
        self.fig.patch.set_facecolor('white')
        self.ax = self.fig.add_axes([0, 0, 1, 1])
        self.ax.set_xlim(0, 1); self.ax.set_ylim(0, 1); self.ax.axis('off')
        self.y = 0.96

    def _sp(self):
        self.pdf.savefig(self.fig, bbox_inches='tight')
        plt.close(self.fig)

    def _chk(self, h):
        if self.y - h < MB:
            self._sp(); self._np()

    def title(self, text):
        self._chk(0.08)
        self.ax.text(0.5, self.y, text, fontproperties=jp_font_bold, fontsize=14,
                     color='#1a3a5c', ha='center', va='top', transform=self.ax.transAxes)
        self.y -= 0.046
        self.ax.axhline(y=self.y, xmin=0.04, xmax=0.96, color='#1a3a5c', linewidth=2)
        self.y -= 0.022

    def wk(self, text):
        self._chk(0.09)
        self.y -= 0.014
        h = 0.040
        self.ax.add_patch(FancyBboxPatch((0.03, self.y-h), 0.94, h,
            boxstyle='round,pad=0.005', lw=0, facecolor='#1a3a5c',
            transform=self.ax.transAxes))
        self.ax.text(0.06, self.y-h/2, text, fontproperties=jp_font_bold, fontsize=12,
                     color='white', ha='left', va='center', transform=self.ax.transAxes)
        self.y -= h + 0.020

    def ph(self, text):
        self._chk(0.065)
        self.y -= 0.010
        h = 0.034
        self.ax.add_patch(FancyBboxPatch((0.04, self.y-h), 0.92, h,
            boxstyle='round,pad=0.004', lw=0, facecolor='#2c6fad',
            transform=self.ax.transAxes))
        self.ax.text(0.07, self.y-h/2, text, fontproperties=jp_font_bold, fontsize=10.5,
                     color='white', ha='left', va='center', transform=self.ax.transAxes)
        self.y -= h + 0.018

    def t(self, text, ind=0, sz=10.5, col='#111111', bold=False):
        self._chk(0.036)
        fp = jp_font_bold if bold else jp_font
        self.ax.text(ML+0.02+ind, self.y, text, fontproperties=fp, fontsize=sz,
                     color=col, ha='left', va='top', transform=self.ax.transAxes)
        self.y -= 0.032

    def eq(self, expr, sz=13, h=0.070, box=True, bg='#fff8e0', edge='#c8860a', wide=False):
        tot = h + 0.022
        self._chk(tot)
        xl = 0.06 if wide else 0.10
        xw = 0.88 if wide else 0.80
        if box:
            self.ax.add_patch(FancyBboxPatch((xl, self.y-h+0.004), xw, h-0.008,
                boxstyle='round,pad=0.008', lw=1.2, facecolor=bg, edgecolor=edge,
                transform=self.ax.transAxes))
        self.ax.text(0.50, self.y-h/2, expr, fontsize=sz,
                     ha='center', va='center', transform=self.ax.transAxes)
        self.y -= tot

    def ans(self, expr, sz=13, h=0.070, wide=False):
        self.eq(expr, sz=sz, h=h, box=True, bg='#eaf4ff', edge='#1a4a7a', wide=wide)

    def sp(self, h=0.012):
        self.y -= h

    def done(self):
        self._sp(); self.pdf.close()


r = R('/home/user/-/physics_summary.pdf')

# ── TITLE ─────────────────────────────────────────────────────────────────────
r.title('物理学1（電気工学科）　§6〜§8週　解答まとめ')

# ── §6週目 ────────────────────────────────────────────────────────────────────
r.wk('§6週目　仕事')

r.ph('問題1：W = ∫F·dr の一般的な定義の導出')
r.t('力Fが一定・直線運動なら  W = F × d（スカラー）')
r.t('一般化：経路を微小区間 dr に分割すると、そのときの仕事は')
r.t('"力の進行方向成分 × 距離" = ドット積')
r.eq(r"$dW = \mathbf{F}\cdot d\mathbf{r} = |\mathbf{F}||d\mathbf{r}|\cos\theta$",
     sz=13, h=0.062)
r.t('全経路で足し合わせると（積分）：')
r.ans(r"$W = \int_A^B \mathbf{F}(\mathbf{r})\cdot d\mathbf{r}$",
      sz=14, h=0.072)
r.sp()

r.ph('問題2(1)：クーロン力の仕事')
r.t('点Oに電荷 q 固定。無限遠から点A（距離R）へ電荷 q\' を移動させる仕事。')
r.t('クーロン力の動径成分：')
r.eq(r"$F_r = \dfrac{qq'}{4\pi\varepsilon_0}\dfrac{1}{r^2}$", sz=13, h=0.075)
r.t('無限遠 → 点A への仕事：')
r.eq(r"$W = \int_\infty^R \dfrac{qq'}{4\pi\varepsilon_0 r^2}dr"
     r" = \dfrac{qq'}{4\pi\varepsilon_0}\!\left[-\dfrac{1}{r}\right]_\infty^R$",
     sz=12, h=0.078, wide=True)
r.ans(r"$W = \dfrac{qq'}{4\pi\varepsilon_0 R}$", sz=14, h=0.072)
r.sp()

r.ph('問題2(2)：隕石の運動エネルギー')
r.t('力学的エネルギー保存則（無限遠基準：KE=0、PE=0）：')
r.eq(r"$0 + 0 = KE - \dfrac{GMm}{R}"
     r"\quad\Rightarrow\quad KE = \dfrac{GMm}{R}$",
     sz=12, h=0.075, wide=True)
r.t('代入：G=6.674×10^-11,  M=5.97×10^24 kg,  m=5 kg,  R=6.38×10^6 m')
r.eq(r"$KE = \dfrac{6.674\times10^{-11}\times5.97\times10^{24}\times5}"
     r"{6.38\times10^6}$", sz=12, h=0.085)
r.ans(r"$KE \approx 3.12\times10^8 \;\mathrm{J} \approx 312\;\mathrm{MJ}$",
      sz=13, h=0.062)
r.sp()

r.ph('問題3：なぜ KE = (1/2)mv² か')
r.t('【1次元の場合】', bold=True)
r.eq(r"$W = \int F\,dx = \int m\dfrac{dv}{dt}dx = \int mv\,dv = \dfrac{1}{2}mv^2$",
     sz=12, h=0.075, wide=True)
r.t('【3次元の場合】', bold=True)
r.eq(r"$W = \int \mathbf{F}\cdot d\mathbf{r}"
     r" = \int m\dfrac{d\mathbf{v}}{dt}\cdot\mathbf{v}\,dt$", sz=12, h=0.072)
r.t('　d(v)/dt · v = (1/2)d(v²)/dt  を使うと：')
r.ans(r"$W = \dfrac{1}{2}mv_B^2 - \dfrac{1}{2}mv_A^2"
      r"$", sz=12, h=0.065, wide=True)

# ── §7週目 ────────────────────────────────────────────────────────────────────
r.wk('§7週目　保存力と位置エネルギー')
r.t('【判定条件（2D）】∂Fy/∂x = ∂Fx/∂y　なら保存力')
r.t('【判定条件（3D）】∇×F = 0　なら保存力')
r.sp(0.010)

r.ph('(1)  F = (xy², x^2y)')
r.eq(r"$\dfrac{\partial(x^2y)}{\partial x}=2xy"
     r"\;=\;\dfrac{\partial(xy^2)}{\partial y}=2xy"
     r"\quad\Rightarrow\quad$", sz=11, h=0.068, wide=True)
r.t('-∂U/∂x = xy²　→　積分すると：')
r.ans(r"$U = -\dfrac{x^2y^2}{2} + C$", sz=14, h=0.072)
r.sp()

r.ph('(2)  F = (axy, by²)')
r.eq(r"$\dfrac{\partial(by^2)}{\partial x}=0,\quad"
     r"\dfrac{\partial(axy)}{\partial y}=ax"
     r"\quad a=0$", sz=11, h=0.068, wide=True)
r.ans(r"$U = -\dfrac{by^3}{3} + C$", sz=14, h=0.072)
r.sp()

r.ph('(3)  F = (axy, bx^2)')
r.eq(r"$\dfrac{\partial(bx^2)}{\partial x}=2bx,\quad"
     r"\dfrac{\partial(axy)}{\partial y}=ax"
     r"\quad a=2b$", sz=11, h=0.068, wide=True)
r.ans(r"$U = -bx^2y + C$", sz=14, h=0.072)
r.sp()

r.ph('(4)  F = (a₁₁x+a₁₂y+a₁₃z,  a₂₁x+a₂₂y+a₂₃z,  a₃₁x+a₃₂y+a₃₃z)')
r.t('∇×F = 0 の条件：a₂₁ = a₁₂，a₃₂ = a₂₃，a₁₃ = a₃₁')
r.t('→ 係数行列が対称行列（aᵢⱼ = aⱼᵢ）のとき保存力', ind=0.02)
r.ans(r"$U = -\dfrac{1}{2}(a_{11}x^2+a_{22}y^2+a_{33}z^2)"
      r" - a_{12}xy - a_{13}xz - a_{23}yz + C$", sz=11, h=0.072, wide=True)
r.sp()

r.ph('(5)  F = A(x, y, z)/r³,   r = √(x^2+y²+z²)')
r.eq(r"$\dfrac{\partial}{\partial x}\dfrac{Ay}{r^3}"
     r"=-\dfrac{3Axy}{r^5}"
     r"=\dfrac{\partial}{\partial y}\dfrac{Ax}{r^3}"
     r"\quad\Rightarrow\quad$", sz=11, h=0.072, wide=True)
r.t('∇(1/r) = −r/r³ を使うと：  F = −A∇(1/r)')
r.ans(r"$U = \dfrac{A}{r} + C$", sz=14, h=0.072)
r.t('　クーロン力：A = qq\'/(4πε₀)　→  U = qq\'/(4πε₀r)', ind=0.02, sz=10)
r.t('　万有引力　：A = −Gm₁m₂　→  U = −Gm₁m₂/r　（A < 0 なので引力）', ind=0.02, sz=10)
r.sp()

r.ph('(6)  F = (−krx, −kry, −krz),   r = √(x^2+y²+z²)')
r.eq(r"$\dfrac{\partial(-kry)}{\partial x}=-\dfrac{kxy}{r}"
     r"=\dfrac{\partial(-krx)}{\partial y}"
     r"\quad\Rightarrow\quad$", sz=11, h=0.068, wide=True)
r.t('∇(r³) = 3r(x,y,z) を使うと：  F = −(k/3)∇(r³)')
r.ans(r"$U = \dfrac{k}{3}r^3 + C = \dfrac{k}{3}(x^2+y^2+z^2)^{3/2} + C$",
      sz=12, h=0.068, wide=True)
r.t('（距離の2乗に比例して増す，中心方向の引力）', col='#555555', sz=10)

# ── §8週目 ────────────────────────────────────────────────────────────────────
r.wk('§8週目　エネルギー保存則・仕事率・エネルギー積分')

r.ph('問題1：エネルギー積分の基本（仕事・エネルギー定理の導出）')
r.t('運動方程式の両辺に速度 v をかける：')
r.eq(r"$m\dfrac{d\mathbf{v}}{dt}\cdot\mathbf{v} = \mathbf{F}\cdot\mathbf{v}$",
     sz=13, h=0.068)
r.t('左辺は運動エネルギーの時間微分，右辺は仕事率 P：')
r.eq(r"$\dfrac{d}{dt}\!\left(\dfrac{1}{2}mv^2\right) = P = \mathbf{F}\cdot\mathbf{v}$",
     sz=13, h=0.078)
r.t('時刻 tA → tB で積分（仕事・エネルギー定理）：')
r.ans(r"$\dfrac{1}{2}mv_B^2 - \dfrac{1}{2}mv_A^2 = W"
      r"$", sz=12, h=0.068, wide=True)
r.sp()

r.ph('問題3：空気抵抗がある場合のエネルギー変化')
r.t('空気抵抗：f = −kv（速度に比例，逆向き）')
r.t('運動方程式：')
r.eq(r"$m\dfrac{d\mathbf{v}}{dt} = \mathbf{F}_{ext} - k\mathbf{v}$", sz=13, h=0.068)
r.t('両辺に v をかけてエネルギー積分：')
r.eq(r"$\dfrac{d}{dt}\!\left(\dfrac{1}{2}mv^2\right)"
     r" = \mathbf{F}_{ext}\cdot\mathbf{v} - kv^2$", sz=13, h=0.078)
r.t('kv² > 0 なので，空気抵抗によってエネルギーは常に減少（散逸）する。')
r.t('経路 A → B で積分：')
r.ans(r"$\dfrac{1}{2}mv_B^2 - \dfrac{1}{2}mv_A^2 = W_{ext}"
      r" - \int_A^B kv\,ds$", sz=12, h=0.078, wide=True)
r.t('∫kv ds > 0 なので，力学的エネルギーは保存されない。',
    col='#1a4a7a', bold=True)

r.done()
print("Done!")
