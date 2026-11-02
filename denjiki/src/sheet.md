---
title: 電磁気学基礎 中間試験 直前確認シート
---

<p class="byline">試験30分前に上から順に見る用。型の記号（A〜N）は対策ブックと同じ。</p>

## 0. 最初の1分でやること

1. **10問全部に目を通して、型（A〜N）を書き込む。** 問1〜7はマークシート（数値だけ）、問8〜10は記述（途中式が点になる）。
2. **単位に丸をつける。** cm、mm、g、μ、n、μC/cm²、比誘電率、誘電率。
3. **電卓の確認。** 角度モード（DEG/RAD）、指数の入力（EXP / ×10ˣ）。

定数：$k = \dfrac{1}{4\pi\varepsilon_0} = 9\times10^9$、$\varepsilon_0 = 8.854\times10^{-12}$ F/m、$e = 1.602\times10^{-19}$ C

| 接頭辞・換算 | 値 |
|---|---|
| m（ミリ）、μ（マイクロ）、n（ナノ）、p（ピコ） | $10^{-3}$、$10^{-6}$、$10^{-9}$、$10^{-12}$ |
| 1 cm、1 mm、1 g | $10^{-2}$ m、$10^{-3}$ m、$10^{-3}$ kg |
| 1 μC/cm² | $10^{-6}$ C ÷ $10^{-4}$ m² $= 10^{-2}$ C/m² |

## 1. 計算問題（マークシート）

::: card
### A. 点電荷の電場の大きさ

**見分け方：**「〜C の電荷が座標 (…) にあるとき、(…) における電界の大きさ」「〜cm 離れた点の電界」

1. 差をとる：（求める点）−（電荷の位置）
2. $r^2 = \Delta x^2 + \Delta y^2 + \Delta z^2$（平方根はとらなくていい）
3. $E = \dfrac{kQ}{r^2}$　［V/m］

**例（小3-2）：** $1.5\times10^{-8}$ C が $(1,2,3)$、点 $(2,-3,1)$ → 差 $(1,-5,-2)$、$r^2 = 30$ → $E = 9\times10^9\times1.5\times10^{-8}/30 = 4.5$ V/m

> ミス：差の符号（$-3-2 = -5$）、cm を m に直し忘れ
:::

::: card
### B. 複数の点電荷による電位

**見分け方：**「3つの点電荷が…の位置にある。原点O（点P）の電位を求めよ」

1. 各電荷から**求める点まで**の距離 $r_i$ を出す（原点なら $\sqrt{x^2+y^2+z^2}$）
2. $V = k\left(\dfrac{Q_1}{r_1} + \dfrac{Q_2}{r_2} + \dfrac{Q_3}{r_3}\right)$　［V］

**例（小3-3）：** $r_1 = 3$、$r_2 = \sqrt2$、$r_3 = \sqrt{13}$ → $V \approx 23.8$ V

> ミス：負の電荷は**マイナスのまま**入れる。電位はスカラー（向きなし）。2025年は**原点ではなく点P**だった
:::

::: card
### C. クーロン力（一直線上・媒質中）

**見分け方：**「〜に働くクーロン力を求めよ」「力の向き」「誘電率 ε の媒質中」

1. 力の大きさを1組ずつ：$F = \dfrac{kQ_1Q_2}{r^2}$（媒質中は $F = \dfrac{Q_1Q_2}{4\pi\varepsilon r^2}$）
2. 向きを図で決める：同符号 → 斥力、異符号 → 引力
3. 問題文の「正の向き」に合わせて ± をつけて足す

| 媒質の書き方 | 使い方 |
|---|---|
| 「誘電率 $\varepsilon = 4.2$ F/m」 | そのまま $4\pi\varepsilon r^2$ に入れる |
| 「比誘電率 2.5」 | $\varepsilon = 2.5\,\varepsilon_0$ にしてから入れる |

**例（小1-1）：** 電子と陽子が $10^{-8}$ cm $= 10^{-10}$ m → $F \approx 2.3\times10^{-8}$ N、**引力**

**文字式（小1-2）：** $Q_1, Q_2, Q_3$ が間隔 $a$、正は $Q_3 \to Q_1$ の向き

$$F_1 = \frac{kQ_1}{a^2}\left(Q_2 + \frac{Q_3}{4}\right),\quad F_2 = \frac{kQ_2}{a^2}(Q_3 - Q_1),\quad F_3 = -\frac{kQ_3}{a^2}\left(Q_2 + \frac{Q_1}{4}\right)$$

> 検算：$F_1 + F_2 + F_3 = 0$。$Q_1$ と $Q_3$ は距離 $2a$ なので $\frac14$
:::

::: card
### D. 電荷を置き換えても力が同じになる位置

**見分け方：**「A球の代わりに C を…どの点に置けば、B球に作用する力が等しくなるか」

$$\frac{kQ_AQ_B}{r^2} = \frac{kQ_CQ_B}{x^2}\quad\Rightarrow\quad x = r\sqrt{\frac{Q_C}{Q_A}}$$

**例（小2-2）：** $r = 3$ m、$Q_C = 4Q$ → $x = 3\times2 = 6$ m（B から、A があった側）
:::

::: card
### E. 糸でつるした2物体

**見分け方：**「質量 〜g」「絶縁糸で同一の点からつるした」「鉛直に対する傾き θ」

1. 間隔：$r = 2L\sin\theta$（$L$：糸の長さ）
2. つり合い：縦 $T\cos\theta = mg$、横 $T\sin\theta = F$ → **$F = mg\tan\theta$**
3. $Q = r\sqrt{\dfrac{F}{k}}$

**例（小1-3）：** 3 g、0.4 m、35° → $r \approx 0.459$ m、$F \approx 2.06\times10^{-2}$ N → $Q \approx 6.9\times10^{-7}$ C

> ミス：$\sin$ ではなく **$\tan$**。g → **kg**（3 g = $3\times10^{-3}$ kg）。$\pi/10$ rad なら電卓を RAD か、18° に直す
:::

::: card
### G. 逆算（まだ中間試験に出ていない ★）

| 問題文 | 解き方 | 例 |
|---|---|---|
| 等しい2電荷、距離 $r$、反発力 $F$ → 電荷と符号 | $Q = r\sqrt{F/k}$、反発なら**同符号**、引力なら**異符号** | 小1-4：$Q \approx 8.2\times10^{-7}$ C、同符号 |
| 比 $Q_1:Q_2 = 5:2$、距離、力 | $Q_1 = 5x$、$Q_2 = 2x$ → $10x^2 = Fr^2/k$ | 小2-1：$Q_1 \approx -2.1\times10^{-5}$、$Q_2 \approx -8.5\times10^{-6}$ C（負の電荷） |
| ある点の $E$ と $V$ → $Q$ と $r$ | $r = V/E$、$Q = Vr/k$ | 小3-4：$r = 1.44$ m、$Q \approx 5.8\times10^{-8}$ C |
| 球面上の $E$、半径 $r$ → 中心の電荷 | $Q = Er^2/k$ | 小4-3：$Q \approx 1.3\times10^{-6}$ C |
| 閉曲面から出る本数 $N$ → 電荷 | 合計 $= N\varepsilon_0$、知っている電荷を引く | 小6-4：$Q \approx 3.0\times10^{-5}$ C |

> 逆算の答えには**符号**を必ず書く
:::

::: card
### I. 電気力線の本数

**見分け方：**「電気力線の数」「閉曲面から出る（入る）」「油中」

$$N = \frac{Q_\text{合計}}{\varepsilon}\qquad(\text{真空：}\varepsilon_0、\text{比誘電率 }\varepsilon_r\text{ の油：}\varepsilon_r\varepsilon_0)$$

**例（小4-2）：** $4.0 - 2.5 + 6.0 = 7.5$ μC → $N \approx 8.5\times10^5$ 本　**例（小5-4）：** $3\times10^{-6}/(2.5\varepsilon_0) \approx 1.4\times10^5$ 本

> 合計が正なら**出る**、負なら**入る**。単位は「本」
:::

::: card
### M. 導体球・薄い球殻（数値）

**見分け方：**「表面電荷密度 σ の導体球」「薄い導体球殻が 〜C に帯電」

| 場所 | 電界 | 電位 |
|---|---|---|
| 外側 $r > R$ | $kQ/r^2$ | $kQ/r$ |
| 内側 $r < R$ | **0** | **$kQ/R$**（表面と同じ） |

1. σ が与えられたら $Q = \sigma\times4\pi R^2$（σ は C/m²、R は m に直す）
2. 外側の点は中心に点電荷 $Q$ があるのと同じ（球対称なので厳密に成り立つ）

**例（小4-4）：** σ = 3 μC/cm² = 0.03 C/m²、R = 2 cm → $Q \approx 1.5\times10^{-4}$ C、15 cm で $E \approx 6.0\times10^7$ V/m、$V \approx 9.0\times10^6$ V

> 内側の**電位は 0 ではない**（電界が 0 なだけ）
:::

## 2. 記述問題（途中式が点になる）

> **書き出しの型：**「（対称性）より電界は〜向きで、〜上で大きさ一定。半径 $r$ の〜をガウス面とすると、ガウスの定理より $E\times$（面積）$= Q_\text{in}/\varepsilon_0$」→ $E$ → 「無限遠を基準として」$V = -\int E\,dr$

::: card
### F. 2つの点電荷による電場（ベクトル・成分）

**見分け方：**「x-y平面の原点に $Q_1$、…に $Q_2$。点Pにおける電界を $\vec i$, $\vec j$ で」「x軸方向の成分」

1. 各電荷から P へのベクトル（P − 電荷）を書く
2. それぞれ $\dfrac{kQ}{(\text{長さ})^3}$ を掛ける
3. $\vec i$ どうし、$\vec j$ どうしを足す

$Q_1$ が原点、$Q_2$ が $(a, 0)$ のとき：

$$E_x = k\left\{\frac{Q_1x}{(x^2+y^2)^{3/2}} + \frac{Q_2(x-a)}{\{(x-a)^2+y^2\}^{3/2}}\right\},\quad E_y = k\left\{\frac{Q_1y}{(x^2+y^2)^{3/2}} + \frac{Q_2\,y}{\{(x-a)^2+y^2\}^{3/2}}\right\}$$

$Q_2$ が $(0, a)$ なら、$(x-a)$ の代わりに $(y-a)$ が入る（2024年）。

**数値例（小3-5）：** $\vec E \approx 1.8\times10^2\,\vec i + 3.7\times10^2\,\vec j$ V/m　**（小6-3）：** P が $Q_1$ の真上なので $Q_1$ の $E_x = 0$ → $E_x \approx -9.1\times10^2$ V/m
:::

::: card
### H. 対称な配置の合力（★）

**見分け方：**「y軸上の $(0, \pm a)$ に $+Q$、x軸上の点に $+2Q$」

1. 2つの力は大きさが同じ：$\dfrac{k\cdot Q\cdot2Q}{x^2+a^2}$
2. y成分は打ち消し合う。x成分は $\cos\phi = \dfrac{x}{\sqrt{x^2+a^2}}$ を掛けて**2倍**

$$F = \frac{4kQ^2x}{(x^2+a^2)^{3/2}} = \frac{Q^2x}{\pi\varepsilon_0(x^2+a^2)^{3/2}}\quad(+x\text{ 向き})$$
:::

::: card
### J. 無限に長い円筒（円筒中に一様に分布、単位長さあたり $Q$）

ガウス面：中心軸が同じ、**半径 $r$、長さ $l$ の円筒**。側面だけ電束が通る（底面は 0）。

| 聞かれること | 式の立て方 | 答え |
|---|---|---|
| 外側 $E_o$（$r > a$） | $E\cdot2\pi rl = \dfrac{Ql}{\varepsilon_0}$ | $\dfrac{Q}{2\pi\varepsilon_0 r}$ |
| 内側の単位長さあたりの電荷 | 断面積の比 $\dfrac{\pi r^2}{\pi a^2}$ | $Q\dfrac{r^2}{a^2}$ |
| 内側 $E_i$（$r < a$） | $E\cdot2\pi rl = \dfrac{Ql}{\varepsilon_0}\dfrac{r^2}{a^2}$ | $\dfrac{Qr}{2\pi\varepsilon_0 a^2}$ |
| 外側2点の電位差（$r_2 > r_1 > a$） | $\displaystyle\int_{r_1}^{r_2}\frac{Q}{2\pi\varepsilon_0 r}dr$ | $\dfrac{Q}{2\pi\varepsilon_0}\ln\dfrac{r_2}{r_1}$ |

> 内側は $r/a$ ではなく **$r^2/a^2$**（面積の比）。無限に長いので「電位」ではなく「**電位差**」を聞かれる
:::

::: card
### K. 平行平板

**(a) 面密度 $\pm\sigma$、間隔 $d$（小5-2）**

1. 1枚の平面が作る電界は $\dfrac{\sigma}{2\varepsilon_0}$（型L）
2. 板の間：$+\sigma$ からは離れる向き、$-\sigma$ へは向かう向きで**同じ向き** → 足す。外側：逆向き → 0

$$E = \frac{\sigma}{2\varepsilon_0} + \frac{\sigma}{2\varepsilon_0} = \frac{\sigma}{\varepsilon_0}\quad(+\to-),\qquad V = Ed = \frac{\sigma d}{\varepsilon_0}$$

**(b) 面積 $S$、間隔 $d$、電位差 $V$（小6-2）：** $E = \dfrac{V}{d}$、$Q = \sigma S = \varepsilon_0 ES = \dfrac{\varepsilon_0 SV}{d}$（各板 $+Q$ と $-Q$）

> 数値で出たら：電界に $d$ は**使わない**、$d$ は電位差だけ
:::

::: card
### L. 無限に広い平面1枚（小6-1）

ガウス面：平面を**垂直に貫く円柱**（底面積 $A$）。電束は上下の底面から出る（側面は 0）。

$$E\cdot A + E\cdot A = \frac{\sigma A}{\varepsilon_0}\quad\Rightarrow\quad E = \frac{\sigma}{2\varepsilon_0}$$

向き：$\sigma > 0$ なら両側とも平面から**離れる**、$\sigma < 0$ なら平面に**向かう**。

> 平面1枚は **2で割る**。平行平板の間は割らない
:::

::: card
### M（記述）. 薄い導体球殻（小5-3）

ガウス面：半径 $r$ の同心球面。

- 外側：$E_o\cdot4\pi r^2 = \dfrac{Q}{\varepsilon_0}$ → $E_o = \dfrac{Q}{4\pi\varepsilon_0 r^2}$、$V_o = -\displaystyle\int_\infty^r E_o\,dr = \dfrac{Q}{4\pi\varepsilon_0 r}$
- 内側：ガウス面内の電荷 0 → $E_i = 0$、$V_i = V(a) = \dfrac{Q}{4\pi\varepsilon_0 a}$（電界 0 なので電位は変わらない）
:::

::: card
### N. 同心の導体球（毎年 問10）

**まず導体2の電荷を見る。** 導体2の内側の面には必ず $-Q$、外側の面には残り（導体2の電荷 $+ Q$）。

ガウス面：半径 $r$ の同心球面 → $E\cdot4\pi r^2 = \dfrac{Q_\text{in}}{\varepsilon_0}$

| | 導体2が $-Q$（小6-5, 2022, 2024） | 導体2が 0（小4-1, 2025） |
|---|---|---|
| 外側の面の電荷 | 0 | $+Q$ |
| $E_o$（$r > c$） | 0 | $\dfrac{Q}{4\pi\varepsilon_0 r^2}$ |
| $V_o$（$r > c$） | 0 | $\dfrac{Q}{4\pi\varepsilon_0 r}$ |
| $E_i$（$a < r < b$） | $\dfrac{Q}{4\pi\varepsilon_0 r^2}$ | $\dfrac{Q}{4\pi\varepsilon_0 r^2}$ |
| $V_i$（$a < r < b$） | $\dfrac{Q}{4\pi\varepsilon_0}\left(\dfrac1r - \dfrac1b\right)$ | $\dfrac{Q}{4\pi\varepsilon_0}\left(\dfrac1r - \dfrac1b + \dfrac1c\right)$ |
| $V_2$ | **0** | $\dfrac{Q}{4\pi\varepsilon_0 c}$ |
| $V_1$ | $\dfrac{Q}{4\pi\varepsilon_0}\left(\dfrac1a - \dfrac1b\right)$ | $\dfrac{Q}{4\pi\varepsilon_0}\left(\dfrac1a - \dfrac1b + \dfrac1c\right)$ |

**積分の書き方（導体2が 0 のとき）：**

$$V_2 = -\int_\infty^c\frac{Q}{4\pi\varepsilon_0 r^2}dr = \frac{Q}{4\pi\varepsilon_0 c},\qquad V_1 = V_2 - \int_b^a\frac{Q}{4\pi\varepsilon_0 r^2}dr = \frac{Q}{4\pi\varepsilon_0}\left(\frac1c + \frac1a - \frac1b\right)$$

> $b < r < c$ は導体の中で電界 0 → $V(b) = V(c)$。$V_2$ の式に入るのは **$c$**（$b$ ではない）。導体2が $-Q$ なら、外側の電界が 0 なので「無限遠を基準として」$V_2 = 0$
:::

## 3. 見直しチェック（解き終わったら）

- [ ] 単位換算：cm → m、g → kg、μ・n、μC/cm² → C/m²（$\times10^{-2}$）
- [ ] 誘電率 $\varepsilon$（そのまま）と比誘電率 $\varepsilon_r$（$\times\varepsilon_0$）
- [ ] 電界は $r^2$、電位は $r$ で割った
- [ ] 電位を求める点：原点か点Pか
- [ ] 負の電荷の符号、力の向き、出る/入る、同符号/異符号
- [ ] 糸のつり合いは $\tan$
- [ ] 平面1枚は $\sigma/2\varepsilon_0$、平行平板の間は $\sigma/\varepsilon_0$
- [ ] 円筒の内側は $r^2/a^2$
- [ ] 球殻・導体の内側：電界 0、電位は表面と同じ
- [ ] 同心球：導体2の電荷を確認、$V_2$ は $c$ で書く
- [ ] 答えに単位（N、V/m、V、C、本）
- [ ] 記述：ガウス面の設定 → ガウスの定理の式 → 積分の式 を書いた
