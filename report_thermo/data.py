# -*- coding: utf-8 -*-
"""熱電対の検定実験 — 測定データ入力ファイル。
None の項目は報告書上で【要記入】になる。測定値が全部そろうと、表・冷却曲線・
最小二乗法の計算過程・考察の数値が自動で作成される。"""

META = {
    "kai": None,                 # 第何回（例: 10）
    "date": None,                # 実験年月日（例: "2026年10月7日"）
    "submit": None,              # 報告書提出日
    "han": None,                 # 班名
    "name": "森木崇允",
    "grade": "1",
    "student_id": "4326091",
    "coworkers": [],             # [("氏名", "学籍番号"), ...]
    "room_thermometer": None,    # 室温の測定に使った温度計（任意）
}

# 実験1: 室温 [℃] と室温での起電力 [mV]
ROOM_TEMP = None
ROOM_EMF = None

# 実験2: 0 ℃（氷水）での起電力 [mV]
ZERO_EMF = None

# 実験3: 冷却時の起電力 [mV]（冷却の記録開始を 0 s とし、15 s 間隔で並べる）
DT = 15.0
COOLING = {
    "Sn": None,   # [14.270, 14.247, ...]
    "Pb": None,
    "Zn": None,
}

# 融点に対応する起電力を求める区間（プラトー）を手動指定する場合は (開始時刻 s, 終了時刻 s)。
# None なら隣り合う測定値の変化量から自動で判定する。
PLATEAU = {"Sn": None, "Pb": None, "Zn": None}
