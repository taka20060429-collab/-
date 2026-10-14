"""GL240 の CSV を読み込む。"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")


def load(name):
    raw = open(os.path.join(DATA, name), "rb").read()
    for enc in ("utf-8-sig", "cp932"):
        try:
            t = raw.decode(enc)
            break
        except UnicodeDecodeError:
            pass
    rows = []
    started = False
    for line in t.splitlines():
        p = [x.strip() for x in line.split(",")]
        if p[0] == "NO.":
            started = True
            continue
        if started:
            try:
                rows.append((int(p[0]), float(p[3]), float(p[4])))
            except (ValueError, IndexError):
                pass
    return rows


def angle(rows, start, move):
    """注意点の式(2)：角度 = 移動開始角度 + カウンタ値 × (移動角度 / 総カウンタ数)。"""
    n = len(rows)
    return [start + r[0] * move / n for r in rows]
