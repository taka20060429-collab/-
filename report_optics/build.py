# -*- coding: utf-8 -*-
"""過去レポートと同じ形（2つの実験を1冊にまとめた報告書）を作る。 python3 build.py"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import common as C

META = {"kai": None, "date": None, "submit": None, "han": "D班", "name": "森木崇允", "grade": "1",
        "student_id": "4326091",
        "coworkers": [("石井真人", "4326007"), ("木元喜紹", "4326029"), ("陶山稜平", "4326051"),
                      ("林田知子", "4326071")]}

C.new_doc(os.path.join(HERE, "img_eq"))
C.cover(META, "レンズの焦点距離の測定とレーザー光の回折と干渉")
C.set_part("A")
import lens  # noqa: E402,F401  （A．レンズの焦点距離の測定を書き込む）
C.doc.add_page_break()
C.set_part("B")
import laser  # noqa: E402,F401  （B．レーザー光の回折と干渉の実験を書き込む）
out = os.path.join(HERE, "レンズの焦点距離の測定とレーザー光の回折と干渉_報告書.docx")
C.save(out)
print("saved:", out)
