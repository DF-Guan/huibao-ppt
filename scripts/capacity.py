# -*- coding: utf-8 -*-
"""capacity — 每个版式能"装得下多少",超限提前预警。

容量以"排版单位"计:一个汉字或一个英文单词 ≈ 1。
每个版式对关键字段标注容量上限;spec 超限时 build 前报警并建议换版式。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pptx_kit as K

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

def _u(s):
    return K._cjk_units(s)


# 每版式关键字段的容量上限(排版单位)
CAPACITY = {
    "cover": {"title": 18, "subtitle": 30, "lead": 60, "kicker": 40, "ticker": 50},
    "toc": {"item_name": 8, "item_desc": 22, "count": (1, 4)},
    "stats": {"card_label": 8, "card_sub": 16, "count": (2, 4), "ticker": 40},
    "compare": {"num": 6, "label": 16, "note_text": 60},
    "items": {"item_title": 16, "item_desc": 30, "count": (1, 4), "ticker": 50},
    "thanks": {"line": 40, "footnote": 40},
    "chart": {"title": 30, "insights": 120, "big_label": 24},
    "photo_gallery": {"title": 30, "caption": 20, "count": (1, 6)},
    "logo_wall": {"title": 34, "lead": 40, "body": 80},
    "org_chart": {"title": 34, "body": 90},
    "scene": {"title": 40, "body": 90, "lead": 24},
    "timeline": {"title": 30, "item_title": 12, "item_text": 24, "count": (2, 6), "body": 80},
}


def check_slide(idx, sd):
    """检查单页 spec 是否超容量,返回警告列表。"""
    stype = sd.get("type")
    cap = CAPACITY.get(stype)
    if not cap:
        return []
    warns = []

    def field_warn(field, text, key):
        lim = cap.get(key)
        if lim and text and _u(text) > lim:
            warns.append(f"page{idx} [{stype}] {key} 超容量: {_u(text):.0f} > {lim}")

    if stype == "cover":
        field_warn("title", sd.get("title"), "title")
        field_warn("subtitle", sd.get("subtitle"), "subtitle")
        field_warn("lead", sd.get("lead"), "lead")
    elif stype in ("toc", "stats", "items", "photo_gallery", "timeline"):
        items = sd.get("items") or sd.get("cards") or sd.get("photos") or []
        lo, hi = cap.get("count", (0, 99))
        if not lo <= len(items) <= hi:
            warns.append(f"page{idx} [{stype}] 条目数 {len(items)} 超出 {lo}-{hi}")
        for it in items:
            field_warn("name", it.get("name") or it.get("title") or it.get("label"), "item_name" if stype == "toc" else ("item_title" if stype in ("items", "timeline") else "card_label"))
            field_warn("desc", it.get("desc") or it.get("sub") or it.get("text"), "item_desc" if stype == "toc" else ("item_text" if stype == "timeline" else "card_sub"))
        if sd.get("title"):
            field_warn("title", sd.get("title"), "title")
    elif stype in ("chart", "photo_gallery", "logo_wall", "org_chart", "scene"):
        field_warn("title", sd.get("title"), "title")
        field_warn("insights", sd.get("insights"), "insights")
        field_warn("body", sd.get("body"), "body")
        field_warn("lead", sd.get("lead"), "lead")
    return warns


def check_deck(spec):
    warns = []
    for i, sd in enumerate(spec.get("slides", []), 1):
        warns.extend(check_slide(i, sd))
    return warns


if __name__ == "__main__":
    import json
    with open(sys.argv[1], encoding="utf-8") as f:
        spec = json.load(f)
    ws = check_deck(spec)
    for w in ws:
        print("CAPACITY:", w)
    print(f"共 {len(ws)} 处容量提醒" if ws else "容量自检通过")

CAPACITY.update({
    "cover_type": {"title": 12, "subtitle": 24, "body": 60, "kicker": 30, "ticker": 40},
    "big_quote": {"quote": 60, "by": 16, "note_text": 40},
    "section": {"title": 12, "subtitle": 40, "kicker": 24},
    "shot_stack": {"title": 24, "item_title": 10, "item_text": 26, "count": (2, 4), "body": 60},
    "overlap": {"card_title": 12, "card_text": 40, "count": (2, 4)},
    "free": {},
})

CAPACITY.update({
    "ed_cover": {"title": 10, "standfirst": 70, "kicker": 24, "issue": 24},
    "ed_split": {"numeral": 3, "title": 18, "quote": 50, "body": 130, "numeral_label": 20},
    "ed_report": {"title": 14, "focus_big": 5, "focus_label": 10, "focus_text": 60, "footnote": 40},
})

CAPACITY.update({"ed_figures": {"num": 6, "label": 12, "sub": 22, "count": (2, 4), "title": 20, "body": 90}})

CAPACITY.update({"table": {"col_head": 10, "cell": 14, "max_cols": 6, "max_rows": 8}})
