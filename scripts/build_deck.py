# -*- coding: utf-8 -*-
"""build_deck — 「汇报大师 PPT」页面生成器 v2(风格驱动 · 实战版式 · 口播备注)。

用法:
    python build_deck.py <deck-spec.json> <输出.pptx> [--style brutal|dark_premium]

spec 顶层:
    { "style": "brutal", "interactive": true, "transition": "push",
      "slides": [ {"type": "...", "note": "口播稿", ...} ] }

每页可带 "note" 字段,自动写入演讲者备注。
风格定义见 styles/<name>.json;版式契约见 references/page-archetypes.md。
"""
import json
import os
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

import pptx_kit as K
from pptx_kit import (C, IN, MONO, YAHEI, add_chart, add_notes, add_transition,

                      check_capacity, connector, link, logo_wall, org_node,
                      photo_contain, photo_cover, rect, runs_text, shape_text,
                      snap_to_scale, text)
try:
    import capacity as _cap
except Exception:
    _cap = None

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TRANSITION = "push"


# ---------- 风格加载 ----------

def load_style(name):
    path = os.path.join(SKILL_DIR, "styles", name + ".json")
    if not os.path.exists(path):
        raise FileNotFoundError(f"未知风格: {name}(styles/ 下应有 {name}.json)")
    with open(path, encoding="utf-8") as f:
        s = json.load(f)
    s["_c"] = {k: C(v) for k, v in s["colors"].items()}
    return s


def col(style, key):
    return style["_c"][key]


# ============================================================
# 版式渲染器 —— 全部由风格令牌驱动
# ============================================================

def _bg(slide, style):
    rect(slide, 0, 0, K.PAGE_W, K.PAGE_H, fill=col(style, "bg"))


def _depth_card(slide, style, x, y, w, h, fill=None, rot=None):
    """按 decor.depth 画卡片:hard_shadow / flat。"""
    c = col(style, "card") if fill is None else fill
    if style["decor"]["depth"] == "hard_shadow":
        off = style["geometry"]["shadow_off"]
        rect(slide, x + off, y + off, w, h, fill=col(style, "ink"), rot=rot)
        return rect(slide, x, y, w, h, fill=c, line=col(style, "ink"),
                    lw=style["geometry"]["border_w"], rot=rot)
    return rect(slide, x, y, w, h, fill=c, line=col(style, "line"),
                lw=style["geometry"]["border_w"], rot=rot)


def _header(slide, style, title):
    """页眉:brutal 用强调色牌,dark 用 kicker+大标题+发丝线。"""
    if style["decor"]["depth"] == "hard_shadow":
        w = 0.62 + 0.42 * K._cjk_units(title)
        hd = rect(slide, 0.6, 0.5, w, 0.55, fill=col(style, "accent"))
        shape_text(hd, title, 17, col(style, "on_ink"), bold=True)
        rect(slide, 0.6, 1.14, w, 0.05, fill=col(style, "ink"))
    else:
        text(slide, 0.85, 0.55, 9.0, 0.3, title.upper() if title.isascii() else title,
             size=11, color=col(style, "accent"), bold=True, tracking=2.2)


def _h1(slide, style, t, y=0.92, size=30):
    size = snap_to_scale(size, style["type_scale"])
    text(slide, 0.8, y, 11.7, size / 72 * 1.7, t, size=size, color=col(style, "ink"),
         bold=True, spacing=1.12)


def _hair(slide, style, y, x=0.85, w=11.6):
    rect(slide, x, y, w, 0.016, fill=col(style, "line"))


def _goldbar(slide, style, x, y, w=1.3):
    rect(slide, x, y, w, 0.045, fill=col(style, "accent"))


def _note(slide, spec):
    if spec.get("note"):
        add_notes(slide, spec["note"])


# ---------- 已有版式(风格化) ----------

def r_cover(slide, spec, ctx):
    style = ctx["style"]
    _bg(slide, style)
    if spec.get("bg_image"):
        photo_cover(slide, spec["bg_image"], 0, 0, K.PAGE_W, K.PAGE_H)
        ov = rect(slide, 0, 0, K.PAGE_W, K.PAGE_H, fill=col(style, "bg"))
        K.set_alpha(ov, spec.get("bg_dim", 60))
    rect(slide, 0, 0, K.PAGE_W, 1.05, fill=col(style, "ink"))
    text(slide, 0.6, 0, 9.0, 1.05, spec.get("kicker", ""), size=13, color=col(style, "on_ink"),
         bold=True, font=MONO, anchor=MSO_ANCHOR.MIDDLE, tracking=2)
    rect(slide, 12.35, 0.32, 0.42, 0.42, fill=col(style, "signal"), shape=MSO_SHAPE.OVAL)
    title = spec["title"]
    hl_w = min(0.62 * K._cjk_units(title) + 0.5, 12.0)
    rect(slide, 0.78, 1.95, hl_w, 1.1, fill=col(style, "accent"), rot=-1)
    text(slide, 0.9, 1.88, 12.0, 1.25, title, size=44, color=col(style, "ink"), bold=True, rot=-1)
    if spec.get("subtitle"):
        bar = rect(slide, 0.82, 3.3, min(0.34 * K._cjk_units(spec["subtitle"]) + 0.7, 11.5), 0.52, fill=col(style, "ink"))
        shape_text(bar, spec["subtitle"], 11.5, col(style, "on_ink"), bold=True)
    if spec.get("lead"):
        _depth_card(slide, style, 0.82, 4.15, 8.1, 1.55, rot=0.8)
        text(slide, 1.05, 4.35, 7.65, 1.15, spec["lead"], size=13, color=col(style, "ink"), bold=True,
             spacing=1.4, rot=0.8)
    if spec.get("badge"):
        star = rect(slide, 10.95, 1.75, 1.5, 1.5, fill=col(style, "signal"), line=col(style, "ink"), lw=2,
                    shape=MSO_SHAPE.STAR_8_POINT, rot=15)
        shape_text(star, spec["badge"], 20, col(style, "on_ink"), bold=True, font=MONO)
    rect(slide, 0, 6.9, K.PAGE_W, 0.6, fill=col(style, "ink"))
    if spec.get("ticker"):
        text(slide, 0, 6.9, K.PAGE_W, 0.6, spec["ticker"], size=11, color=col(style, "on_ink"),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=MONO, tracking=2)
    _note(slide, spec)


def r_toc(slide, spec, ctx):
    style = ctx["style"]
    _bg(slide, style)
    hd = rect(slide, 0.6, 0.55, 4.3, 0.62, fill=col(style, "ink"))
    shape_text(hd, spec.get("title", "目录 CONTENTS"), 17, col(style, "on_ink"), bold=True)
    rect(slide, 5.05, 0.67, 0.38, 0.38, fill=col(style, "signal"))
    if ctx.get("interactive"):
        text(slide, 8.7, 0.7, 3.8, 0.3, "放映模式 · 点击卡片跳转", size=10,
             color=col(style, "faint"), align=PP_ALIGN.RIGHT)
    items = spec["items"]
    if not 1 <= len(items) <= 4:
        raise ValueError(f"toc 仅支持 1-4 个条目, 收到 {len(items)} 个")
    pos = [(0.75, 1.62), (6.85, 1.62), (0.75, 4.12), (6.85, 4.12)][: len(items)]
    links = ctx.setdefault("toc_links", [])
    ACC = [col(style, "accent"), col(style, "accent2"), col(style, "accent3"), col(style, "accent")]
    for i, ((cx, cy), item) in enumerate(zip(pos, items)):
        card = _depth_card(slide, style, cx, cy, 5.75, 2.15)
        n1 = text(slide, cx + 0.22, cy + 0.12, 1.4, 0.65, item["num"], size=32,
                  color=col(style, "ink"), bold=True, font=MONO)
        n2 = text(slide, cx + 0.22, cy + 0.82, 4.2, 0.5, item["name"], size=19,
                  color=col(style, "ink"), bold=True)
        n3 = text(slide, cx + 0.22, cy + 1.38, 4.6, 0.32, item.get("desc", ""), size=10.5,
                  color=col(style, "muted"))
        rect(slide, cx + 0.22, cy + 1.82, 1.4, 0.12, fill=ACC[i % 4])
        text(slide, cx + 4.95, cy + 0.14, 0.6, 0.55, "→", size=22, color=col(style, "ink"), bold=True)
        links.append((card, n1, n2, n3))
    _note(slide, spec)


def r_stats(slide, spec, ctx):
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec["title"])
    cards = spec["cards"]
    if not 2 <= len(cards) <= 4:
        raise ValueError(f"stats 仅支持 2-4 张数据卡, 收到 {len(cards)} 张")
    n = len(cards)
    w = (11.83 - (n - 1) * 0.3) / n
    num_size = snap_to_scale(56 if n <= 3 else 42, style["type_scale"])
    ACC = [col(style, "accent"), col(style, "accent2"), col(style, "accent3"), col(style, "accent")]
    for i, cd in enumerate(cards):
        x = 0.75 + i * (w + 0.3)
        _depth_card(slide, style, x, 1.6, w, 3.35)
        rect(slide, x + 0.24, 2.75, min(w - 1.0, 0.52 * len(cd["num"]) + 0.4), 0.4, fill=ACC[i % 4])
        runs_text(slide, x + 0.24, 1.85, w - 0.4, 1.2,
                  [(cd["num"], num_size, col(style, "ink"), True, MONO),
                   (" " + cd.get("unit", ""), 16, col(style, "ink"), True, YAHEI)])
        lb = rect(slide, x + 0.24, 3.35, min(0.34 * K._cjk_units(cd["label"]) + 0.6, w - 0.48), 0.5,
                  fill=col(style, "ink"))
        shape_text(lb, cd["label"], 13, col(style, "on_ink"), bold=True)
        if cd.get("sub"):
            text(slide, x + 0.24, 4.05, w - 0.48, 0.4, cd["sub"], size=11, color=col(style, "muted"))
    if spec.get("ticker"):
        rect(slide, 0, 6.55, K.PAGE_W, 0.55, fill=col(style, "ink"))
        text(slide, 0, 6.55, K.PAGE_W, 0.55, spec["ticker"], size=13, color=col(style, "on_ink"),
             bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=MONO)
    _note(slide, spec)


def r_compare(slide, spec, ctx):
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec["title"])
    before, after = spec["before"], spec["after"]
    rect(slide, 0.9, 1.8, 4.5, 3.5, fill=col(style, "ink"))
    text(slide, 1.15, 2.05, 2.0, 0.35, before.get("tag", "BEFORE"), size=12,
         color=col(style, "faint"), font=MONO, tracking=2)
    text(slide, 1.1, 2.6, 4.1, 1.2, before["num"], size=54, color=col(style, "on_ink"), bold=True, font=MONO)
    rect(slide, 2.6, 2.7, 0.09, 1.15, fill=col(style, "signal"), rot=45)
    rect(slide, 2.6, 2.7, 0.09, 1.15, fill=col(style, "signal"), rot=-45)
    text(slide, 1.15, 4.35, 4.0, 0.4, before.get("label", ""), size=12, color=col(style, "faint"))
    text(slide, 5.65, 2.95, 1.6, 0.9, "→", size=48, color=col(style, "ink"), bold=True, align=PP_ALIGN.CENTER)
    _depth_card(slide, style, 7.25, 1.8, 4.5, 3.5, fill=col(style, "accent"))
    text(slide, 7.5, 2.05, 2.0, 0.35, after.get("tag", "NOW"), size=12, color=col(style, "on_ink"),
         font=MONO, tracking=2, bold=True)
    text(slide, 7.45, 2.6, 4.1, 1.2, after["num"], size=54, color=col(style, "on_ink"), bold=True, font=MONO)
    text(slide, 7.5, 4.35, 4.0, 0.4, after.get("label", ""), size=12, color=col(style, "on_ink"), bold=True)
    if spec.get("stamp"):
        circ = rect(slide, 10.75, 1.35, 1.15, 1.15, fill=col(style, "accent2"), line=col(style, "ink"), lw=2,
                    shape=MSO_SHAPE.OVAL, rot=10)
        shape_text(circ, spec["stamp"], 10, col(style, "ink"), bold=True)
    if spec.get("note_text"):
        text(slide, 0.9, 5.75, 11.5, 0.4, spec["note_text"], size=12, color=col(style, "muted"), align=PP_ALIGN.CENTER)
    _note(slide, spec)


def r_items(slide, spec, ctx):
    style = ctx["style"]
    _bg(slide, style)
    rect(slide, 0, 0.7, K.PAGE_W, 0.52, fill=col(style, "accent"))
    rect(slide, 0, 1.22, K.PAGE_W, 0.04, fill=col(style, "ink"))
    text(slide, 0.6, 0.7, 8.0, 0.52, spec["title"], size=15, color=col(style, "on_ink"), bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
    items = spec["items"]
    if not 1 <= len(items) <= 4:
        raise ValueError(f"items 仅支持 1-4 条横带, 收到 {len(items)} 条")
    n = len(items)
    h, gap = min(1.8, 4.4 / n), 0.3
    total = n * h + (n - 1) * gap
    y0 = 1.55 + max(0, (4.6 - total) / 2)
    ACC = [col(style, "accent"), col(style, "accent2"), col(style, "accent3"), col(style, "accent")]
    for i, item in enumerate(items):
        y = y0 + i * (h + gap)
        _depth_card(slide, style, 0.75, y, 11.8, h)
        sq = rect(slide, 0.75, y, 0.75, min(0.75, h), fill=col(style, "ink"))
        shape_text(sq, item["num"], 18, col(style, "on_ink"), bold=True, font=MONO)
        text(slide, 1.75, y + 0.22, 9.6, 0.55, item["title"], size=19, color=col(style, "ink"), bold=True)
        if item.get("desc"):
            text(slide, 1.75, y + 0.9, 9.6, 0.4, item["desc"], size=11.5, color=col(style, "muted"))
        rect(slide, 12.05, y, 0.5, h, fill=ACC[i % 4])
    if spec.get("ticker"):
        rect(slide, 0.75, 6.35, 11.8, 0.55, fill=col(style, "ink"))
        text(slide, 0.75, 6.35, 11.8, 0.55, spec["ticker"], size=12, color=col(style, "on_ink"),
             bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    _note(slide, spec)


def r_thanks(slide, spec, ctx):
    style = ctx["style"]
    rect(slide, 0, 0, K.PAGE_W, K.PAGE_H, fill=col(style, "ink"))
    text(slide, 1.0, 2.0, 11.0, 1.6, spec.get("big", "THANKS"), size=84, color=col(style, "accent"),
         bold=True, font=MONO, tracking=4)
    text(slide, 1.05, 3.85, 9.5, 0.6, spec.get("line", "谢谢观看 · 敬请指正"), size=18,
         color=col(style, "on_ink"), bold=True)
    rect(slide, 1.05, 4.55, 3.0, 0.12, fill=col(style, "accent"))
    if spec.get("badge"):
        star = rect(slide, 10.8, 4.3, 1.35, 1.35, fill=col(style, "signal"), shape=MSO_SHAPE.STAR_8_POINT, rot=18)
        shape_text(star, spec["badge"], 18, col(style, "on_ink"), bold=True, font=MONO)
    back = rect(slide, 1.05, 5.35, 1.7, 0.52, fill=col(style, "accent3"))
    shape_text(back, "← 目录", 13, col(style, "ink"), bold=True)
    if spec.get("footnote"):
        text(slide, 1.05, 6.6, 8.0, 0.3, spec["footnote"], size=10, color=col(style, "faint"),
             font=MONO, tracking=2)
    _note(slide, spec)
    return back


# ---------- 新增实战版式 ----------

def r_chart(slide, spec, ctx):
    """原生图表页:图表 + 结论。kind: bar/pie/doughnut。"""
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec.get("kicker", spec["title"]))
    _h1(slide, style, spec["title"])
    _hair(slide, style, 1.72)
    ch_spec = spec["chart"]
    hl = ch_spec.get("highlight_idx")
    add_chart(slide, ch_spec.get("kind", "bar"), ch_spec["categories"], ch_spec.get("series", "值"),
              ch_spec["values"], 0.85, 2.1, 7.2, 4.4,
              highlight_idx=hl, highlight_color=col(style, "accent"),
              base_color=col(style, "panel"), legend=ch_spec.get("legend", False),
              number_format=ch_spec.get("number_format"))
    if spec.get("big_num"):
        runs_text(slide, 8.5, 2.3, 4.0, 1.1, [(spec["big_num"], 36, col(style, "accent"), True, MONO)])
        text(slide, 8.5, 3.4, 4.0, 0.4, spec.get("big_label", ""), size=11, color=col(style, "muted"))
    if spec.get("insights"):
        y = 3.9 if spec.get("big_num") else 2.3
        _hair(slide, style, y, x=8.5, w=4.0)
        text(slide, 8.5, y + 0.2, 4.0, 2.4, spec["insights"], size=11.5, color=col(style, "muted"), spacing=1.6)
    if spec.get("source"):
        text(slide, 0.85, 6.75, 7.0, 0.3, spec["source"], size=9.5, color=col(style, "faint"))
    _note(slide, spec)


def r_photo_gallery(slide, spec, ctx):
    """照片墙:2-6 张图完整显示(contain),每张带标题/说明。"""
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec.get("kicker", spec["title"]))
    _h1(slide, style, spec["title"])
    _hair(slide, style, 1.72)
    photos = spec["photos"]
    n = len(photos)
    if not 1 <= n <= 6:
        raise ValueError(f"photo_gallery 支持 1-6 张, 收到 {n} 张")
    cols = 3 if n > 2 else n
    rows = -(-n // cols)
    gap = 0.3
    w = (11.6 - (cols - 1) * gap) / cols
    h = 3.6 if rows == 1 else 2.15
    y0 = 2.1
    for i, ph in enumerate(photos):
        cidx, ridx = i % cols, i // cols
        x = 0.85 + cidx * (w + gap)
        y = y0 + ridx * (h + 0.75)
        photo_contain(slide, ph["image"], x, y, w, h, bg=col(style, "panel"))
        if ph.get("caption"):
            text(slide, x, y + h + 0.05, w, 0.3, ph["caption"], size=10, color=col(style, "muted"))
    _note(slide, spec)


def r_logo_wall(slide, spec, ctx):
    """logo 墙页:自动换行防溢出。"""
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec.get("kicker", spec["title"]))
    _h1(slide, style, spec["title"])
    _hair(slide, style, 1.72)
    logo_wall(slide, spec["logos"], 0.85, 2.15, 11.6, 3.4, per_row=spec.get("per_row", 5),
              box_fill=col(style, "card"), box_line=col(style, "line"))
    if spec.get("lead"):
        text(slide, 0.85, 5.75, 11.6, 0.5, spec["lead"], size=13, color=col(style, "ink"), bold=True)
    if spec.get("body"):
        text(slide, 0.85, 6.3, 11.6, 0.8, spec["body"], size=11, color=col(style, "muted"), spacing=1.5)
    _note(slide, spec)


def r_org_chart(slide, spec, ctx):
    """组织架构图:nodes + 层级连接。"""
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec.get("kicker", spec["title"]))
    _h1(slide, style, spec["title"])
    _hair(slide, style, 1.72)
    top = spec["top"]
    org_node(slide, 5.42, 2.05, 2.5, 0.72, top["label"], col(style, "accent"), col(style, "on_ink"))
    mids = spec.get("middle", [])
    subs = spec.get("sub", [])
    ny = 3.3
    mw = 2.5
    for i, m in enumerate(mids):
        x = 0.9 + i * ((11.6 - mw) / max(len(mids) - 1, 1)) if len(mids) > 1 else 5.42
        org_node(slide, x, ny, mw, 0.72, m["label"], col(style, "panel"), col(style, "ink"), sub=m.get("sub"))
        connector(slide, x + mw / 2, 2.77, 0.016, ny - 2.77, col(style, "accent"))
    if len(mids) > 1:
        xs = [0.9 + i * ((11.6 - mw) / (len(mids) - 1)) + mw / 2 for i in range(len(mids))]
        connector(slide, min(xs), 3.0, max(xs) - min(xs), 0.016, col(style, "accent"))
    connector(slide, 6.67, 2.77, 0.016, 0.53, col(style, "accent"))
    sy = 4.6
    sw = 2.1
    parent_idx = spec.get("sub_parent", 1)
    px = (0.9 + parent_idx * ((11.6 - mw) / max(len(mids) - 1, 1)) + mw / 2) if len(mids) > 1 else 6.67
    for i, sb in enumerate(subs):
        x = 0.9 + i * (sw + 0.55)
        org_node(slide, x, sy, sw, 0.66, sb, col(style, "panel"), col(style, "ink"))
    if subs:
        xs2 = [0.9 + i * (sw + 0.55) + sw / 2 for i in range(len(subs))]
        connector(slide, px, 4.02, 0.016, 0.2, col(style, "line"))
        connector(slide, min(xs2), 4.22, max(xs2) - min(xs2), 0.014, col(style, "line"))
        for xx in xs2:
            connector(slide, xx, 4.22, 0.014, sy - 4.22, col(style, "line"))
    if spec.get("body"):
        text(slide, 0.9, 5.7, 11.5, 0.5, spec["body"], size=11, color=col(style, "muted"), spacing=1.5)
    _note(slide, spec)


def r_scene(slide, spec, ctx):
    """场景大图页:左/右大图 + 文字区。"""
    style = ctx["style"]
    _bg(slide, style)
    side = spec.get("side", "right")
    img = spec["image"]
    if side == "right":
        photo_cover(slide, img, 7.0, 0, 6.333, K.PAGE_H)
        tx = 0.85
    else:
        photo_cover(slide, img, 0, 0, 6.333, K.PAGE_H)
        tx = 7.4
    text(slide, tx, 0.6, 5.5, 0.3, spec.get("kicker", ""), size=11, color=col(style, "accent"),
         bold=True, tracking=2.2)
    text(slide, tx, 1.6, 5.6, 2.4, spec["title"], size=32, color=col(style, "ink"), bold=True, spacing=1.2)
    _goldbar(slide, style, tx + 0.03, 4.1)
    if spec.get("body"):
        text(slide, tx, 4.4, 5.6, 2.0, spec["body"], size=12.5, color=col(style, "muted"), spacing=1.55)
    if spec.get("lead"):
        text(slide, tx, 6.6, 5.6, 0.4, spec["lead"], size=11.5, color=col(style, "accent2"), bold=True)
    _note(slide, spec)


def r_timeline(slide, spec, ctx):
    """时间线:3-6 个里程碑横排。"""
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec.get("kicker", spec["title"]))
    _h1(slide, style, spec["title"])
    _hair(slide, style, 1.72)
    items = spec["items"]
    n = len(items)
    if not 2 <= n <= 6:
        raise ValueError(f"timeline 支持 2-6 个节点, 收到 {n} 个")
    y = 3.4
    x0, x1 = 1.0, 12.3
    connector(slide, x0, y, x1 - x0, 0.03, col(style, "accent"))
    for i, it in enumerate(items):
        x = x0 + i * ((x1 - x0) / (n - 1)) if n > 1 else (x0 + x1) / 2
        rect(slide, x - 0.14, y - 0.14, 0.28, 0.28, fill=col(style, "accent"), shape=MSO_SHAPE.OVAL)
        text(slide, x - 1.0, y - 0.85, 2.0, 0.4, it.get("date", ""), size=13, color=col(style, "accent2"),
             bold=True, align=PP_ALIGN.CENTER, font=MONO)
        text(slide, x - 1.05, y + 0.3, 2.1, 0.4, it["title"], size=12.5, color=col(style, "ink"),
             bold=True, align=PP_ALIGN.CENTER)
        if it.get("text"):
            text(slide, x - 1.05, y + 0.7, 2.1, 1.2, it["text"], size=9.5, color=col(style, "muted"),
                 align=PP_ALIGN.CENTER, spacing=1.3)
    if spec.get("body"):
        text(slide, 0.85, 6.0, 11.6, 0.6, spec["body"], size=11, color=col(style, "muted"), spacing=1.5)
    _note(slide, spec)


RENDERERS = {
    "cover": r_cover, "toc": r_toc, "stats": r_stats, "compare": r_compare,
    "items": r_items, "thanks": r_thanks,
    "chart": r_chart, "photo_gallery": r_photo_gallery, "logo_wall": r_logo_wall,
    "org_chart": r_org_chart, "scene": r_scene, "timeline": r_timeline,
}


# ---------- 主编排 ----------

def build(spec_path, out_path, style_override=None):
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)
    slides_spec = spec.get("slides")
    if not slides_spec:
        raise ValueError("spec 缺少 slides 数组")
    style_name = style_override or spec.get("style", "brutal")
    style = load_style(style_name)
    interactive = spec.get("interactive", True)
    transition = spec.get("transition", TRANSITION)

    prs = Presentation()
    prs.slide_width, prs.slide_height = IN(K.PAGE_W), IN(K.PAGE_H)
    blank = prs.slide_layouts[6]

    ctx = {"interactive": interactive, "style": style}
    back_shapes = []
    toc_page_idx = None
    for sd in slides_spec:
        stype = sd.get("type")
        if stype not in RENDERERS:
            raise ValueError(f"未知页面类型: {stype!r}, 支持: {sorted(RENDERERS)}")
        slide = prs.slides.add_slide(blank)
        if stype == "toc":
            toc_page_idx = len(prs.slides._sldIdLst) - 1
        ret = RENDERERS[stype](slide, sd, ctx)
        if stype == "thanks" and ret is not None:
            back_shapes.append(ret)

    pages = list(prs.slides)
    if interactive and toc_page_idx is not None:
        toc_links = ctx.get("toc_links", [])
        for i, shapes in enumerate(toc_links):
            target_idx = min(toc_page_idx + 1 + i, len(pages) - 1)
            for sh in shapes:
                link(sh, pages[target_idx])
        for idx, slide in enumerate(pages):
            stype = slides_spec[idx]["type"]
            if stype in ("cover", "toc"):
                continue
            if stype == "thanks":
                link(back_shapes[0], pages[toc_page_idx])
            else:
                box = rect(slide, 12.15, 0.55, 0.85, 0.42, fill=col(style, "ink"))
                shape_text(box, "目录", 11, col(style, "on_ink"), bold=True)
                link(box, pages[toc_page_idx])
    for slide in pages:
        add_transition(slide, transition)

    warnings = []
    if _cap is not None:
        for w in _cap.check_deck(spec):
            warnings.append("CAPACITY " + w)
    for i, slide in enumerate(pages, 1):
        warnings.extend(K.lint_slide(slide, i))
        warnings.extend(K.check_text_alignment(slide, i))

    prs.save(out_path)
    print(f"saved: {out_path} | slides: {len(pages)} | style: {style_name} | interactive: {interactive}")
    for w in warnings:
        print("LINT:", w)
    if not warnings:
        print("LINT: 通过")



# ============================================================
# v3 艺术版式(字体配对 / 艺术字 / 艺术排版)
# ============================================================
from pptx_kit import FONT_PAIRS, set_font_pair, text_art


def _fonts(style, role):
    return FONT_PAIRS.get(style["name"], FONT_PAIRS["brutal"])[role]


def r_cover_type(slide, spec, ctx):
    """大字封面:标题即画面,无图片。靠字阶/字距/艺术字撑住。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _fonts(style, "heading")
    bl, be = _fonts(style, "body")
    text(slide, 0.85, 0.65, 9.0, 0.3, spec.get("kicker", ""), size=11, color=col(style, "accent"),
         bold=True, tracking=3)
    title = spec["title"]
    ts = spec.get("size", 60)
    if spec.get("stroke"):
        text_art(slide, 0.8, 1.7, 11.8, 2.6, title, ts, col(style, "ink"), latin=hl, ea=he,
                 bold=True, tracking=1, stroke=spec["stroke"], stroke_w=spec.get("stroke_w", 1.2),
                 spacing=spec.get("spacing", 1.1))
    elif spec.get("gradient"):
        text_art(slide, 0.8, 1.7, 11.8, 2.6, title, ts, col(style, "ink"), latin=hl, ea=he,
                 bold=True, tracking=1, gradient=spec["gradient"], spacing=spec.get("spacing", 1.1))
    else:
        text_art(slide, 0.8, 1.7, 11.8, 2.6, title, ts, col(style, "ink"), latin=hl, ea=he,
                 bold=True, tracking=1, spacing=spec.get("spacing", 1.1))
    if spec.get("subtitle"):
        text(slide, 0.85, 4.35, 11.0, 0.5, spec["subtitle"], size=16, color=col(style, "muted"))
    rect(slide, 0.88, 5.15, 1.5, 0.045, fill=col(style, "accent"))
    if spec.get("body"):
        text(slide, 0.85, 5.5, 11.0, 1.0, spec["body"], size=12.5, color=col(style, "muted"), spacing=1.6)
    if spec.get("ticker"):
        rect(slide, 0, 6.9, K.PAGE_W, 0.6, fill=col(style, "ink"))
        text(slide, 0, 6.9, K.PAGE_W, 0.6, spec["ticker"], size=11, color=col(style, "on_ink"),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=MONO, tracking=2)
    _note(slide, spec)


def r_big_quote(slide, spec, ctx):
    """金句页:巨型引号 + 衬线斜体引文。"""
    style = ctx["style"]
    _bg(slide, style)
    ql, qe = _fonts(style, "quote")
    text(slide, 0.7, 0.8, 2.5, 2.9, "“", size=150, color=col(style, "line"),
         bold=True, font=ql)
    if spec.get("quote"):
        text_art(slide, 1.5, 2.5, 10.5, 2.8, spec["quote"], spec.get("size", 30),
                 col(style, "ink"), latin=ql, ea=qe, bold=False, italic=True,
                 spacing=1.5, anchor=MSO_ANCHOR.TOP)
    if spec.get("by"):
        rect(slide, 1.55, 5.55, 0.8, 0.04, fill=col(style, "accent"))
        text(slide, 1.55, 5.75, 9.0, 0.4, spec["by"], size=12, color=col(style, "muted"))
    if spec.get("note_text"):
        text(slide, 1.55, 6.3, 10.5, 0.5, spec["note_text"], size=10.5, color=col(style, "faint"))
    _note(slide, spec)


def r_section(slide, spec, ctx):
    """章节页:超大幽灵章节号 + 章节题。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _fonts(style, "heading")
    ghost_color = col(style, "panel") if style["decor"]["depth"] != "hard_shadow" else C("EFEfE0")
    text(slide, 6.6, 1.1, 6.4, 3.2, spec.get("number", "01"), size=180, color=ghost_color,
         bold=True, align=PP_ALIGN.RIGHT, font=MONO, tracking=2)
    text(slide, 0.9, 2.35, 6.2, 0.35, spec.get("kicker", ""), size=11, color=col(style, "accent"),
         bold=True, tracking=3)
    text_art(slide, 0.85, 2.85, 6.6, 1.8, spec["title"], 34, col(style, "ink"), latin=hl, ea=he,
             bold=True, spacing=1.15)
    rect(slide, 0.88, 4.85, 1.3, 0.045, fill=col(style, "accent"))
    if spec.get("subtitle"):
        text(slide, 0.9, 5.15, 6.4, 0.9, spec["subtitle"], size=13, color=col(style, "muted"), spacing=1.5)
    _note(slide, spec)


def r_shot_stack(slide, spec, ctx):
    """编号图组:2-4 张图按宽高比分配宽度,自动编号,可带说明。"""
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec.get("kicker", spec["title"]))
    _h1(slide, style, spec["title"])
    _hair(slide, style, 1.72)
    items = spec["items"]
    if not 2 <= len(items) <= 4:
        raise ValueError(f"shot_stack 支持 2-4 张, 收到 {len(items)} 张")
    gap = 0.3
    w = (11.6 - gap * (len(items) - 1)) / len(items)
    ACC = [col(style, "accent"), col(style, "accent2"), col(style, "accent3"), col(style, "accent")]
    for i, it in enumerate(items):
        x = 0.85 + i * (w + gap)
        _depth_card(slide, style, x, 2.05, w, 3.5)
        photo_contain(slide, it["image"], x + 0.1, 2.15, w - 0.2, 2.6, bg=col(style, "panel"), pad=0.02)
        sq = rect(slide, x, 2.05, 0.5, 0.5, fill=ACC[i % 4])
        shape_text(sq, f"{i+1}", 14, col(style, "ink"), bold=True, font=MONO)
        if it.get("title"):
            text(slide, x + 0.15, 4.85, w - 0.3, 0.35, it["title"], size=12, color=col(style, "ink"), bold=True)
        if it.get("text"):
            text(slide, x + 0.15, 5.2, w - 0.3, 0.7, it["text"], size=10.5, color=col(style, "muted"), spacing=1.3)
    if spec.get("body"):
        text(slide, 0.85, 6.15, 11.6, 0.7, spec["body"], size=11, color=col(style, "muted"), spacing=1.5)
    _note(slide, spec)


def r_overlap(slide, spec, ctx):
    """错位叠卡:2-4 张卡片层叠错位、微旋转,制造艺术张力。"""
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec.get("kicker", spec["title"]))
    _h1(slide, style, spec["title"])
    _hair(slide, style, 1.72)
    cards = spec["cards"]
    if not 2 <= len(cards) <= 4:
        raise ValueError(f"overlap 支持 2-4 张卡, 收到 {len(cards)} 张")
    ACC = [col(style, "accent"), col(style, "accent2"), col(style, "accent3"), col(style, "accent")]
    n = len(cards)
    cw = 4.9 if n <= 3 else 3.6
    chh = 4.0
    x0 = 0.85
    for i, cd in enumerate(cards):
        cx = x0 + i * (cw * 0.62)
        rot = spec.get("rot", [-1.5, 1, -0.8, 1.2][i % 4])
        _depth_card(slide, style, cx, 2.35 + i * 0.12, cw, chh, rot=rot)
        rect(slide, cx, 2.35 + i * 0.12, cw, 0.5, fill=ACC[i % 4], rot=rot)
        text(slide, cx + 0.2, 2.35 + i * 0.12, cw - 0.4, 0.5, cd.get("title", ""), size=13.5,
             color=col(style, "on_ink"), bold=True, anchor=MSO_ANCHOR.MIDDLE, rot=rot)
        if cd.get("big"):
            nl, ne = _fonts(style, "numeral")
            text_art(slide, cx + 0.22, 3.05 + i * 0.12, cw - 0.44, 0.95, cd["big"], 34,
                     col(style, "ink"), latin=nl, ea=ne, bold=True, rot=rot)
        if cd.get("text"):
            text(slide, cx + 0.22, 4.15 + i * 0.12, cw - 0.44, 2.0, cd["text"], size=10.5,
                 color=col(style, "muted"), spacing=1.45, rot=rot)
    _note(slide, spec)


def r_free(slide, spec, ctx):
    """自由版:spec 直接给元素清单,手工艺术排版。"""
    style = ctx["style"]
    _bg(slide, style)
    for el in spec.get("elements", []):
        t = el.get("t", "text")
        x, y = el["x"], el["y"]
        w, h = el.get("w", 2), el.get("h", 0.5)
        if t == "text":
            role = el.get("font", "body")
            latin, ea = _fonts(style, role)
            text_art(slide, x, y, w, h, el["text"], el.get("size", 13), col(style, el.get("color", "ink")),
                     latin=latin, ea=ea, bold=el.get("bold", False), italic=el.get("italic", False),
                     tracking=el.get("tracking"), stroke=el.get("stroke"), stroke_w=el.get("stroke_w", 1.0),
                     gradient=el.get("gradient"), align=el.get("align", PP_ALIGN.LEFT),
                     anchor=el.get("valign", MSO_ANCHOR.TOP), spacing=el.get("spacing", 1.0),
                     rot=el.get("rot"), vert=el.get("vert", False))
        elif t == "rect":
            rect(slide, x, y, w, h, fill=col(style, el.get("fill", "card")),
                 line=col(style, el.get("line")) if el.get("line") else None,
                 lw=el.get("lw", style["geometry"]["border_w"]), rot=el.get("rot"),
                 shape=MSO_SHAPE.OVAL if el.get("shape") == "ellipse" else
                 (MSO_SHAPE.ROUNDED_RECTANGLE if el.get("shape") == "round" else MSO_SHAPE.RECTANGLE),
                 radius=el.get("radius"))
        elif t == "image":
            if el.get("fit", "cover") == "cover":
                photo_cover(slide, el["path"], x, y, w, h)
            else:
                photo_contain(slide, el["path"], x, y, w, h, bg=col(style, el.get("bg", "panel")), pad=el.get("pad", 0.04))
        elif t == "line":
            rect(slide, x, y, max(w, 0.012), max(h, 0.012), fill=col(style, el.get("color", "line")))
    _note(slide, spec)


RENDERERS.update({
    "cover_type": r_cover_type,
    "big_quote": r_big_quote,
    "section": r_section,
    "shot_stack": r_shot_stack,
    "overlap": r_overlap,
    "free": r_free,
})



# ============================================================
# v3.2 杂志编辑风版式(审美课:editorial + brutalist 论文提炼)
# 高级感 = 字体对撞 + 结构即装饰 + 反色焦点 + 克制用色
# ============================================================

def _serif(style):
    return _fonts(style, "heading")


def r_ed_cover(slide, spec, ctx):
    """杂志封面:报头信息行 + 超大衬线标题 + 导语 + 底部元信息栏。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _serif(style)
    ql, qe = _fonts(style, "quote")
    rect(slide, 0.85, 0.8, 11.6, 0.014, fill=col(style, "ink"))
    text(slide, 0.85, 1.0, 6.0, 0.3, spec.get("kicker", ""), size=11, color=col(style, "accent"),
         bold=True, tracking=3)
    text(slide, 7.0, 1.0, 5.45, 0.3, spec.get("issue", ""), size=11, color=col(style, "muted"),
         align=PP_ALIGN.RIGHT, tracking=1)
    rect(slide, 0.85, 1.38, 11.6, 0.014, fill=col(style, "ink"))
    # 超大衬线标题(允许跨过一条发丝线)
    text_art(slide, 0.8, 2.0, 11.9, 2.5, spec["title"], spec.get("size", 54),
             col(style, "ink"), latin=hl, ea=he, bold=True, spacing=spec.get("spacing", 1.08),
             tracking=spec.get("tracking", 0.5))
    rect(slide, 0.85, 4.72, 11.6, 0.008, fill=col(style, "line"))
    if spec.get("standfirst"):
        text_art(slide, 0.85, 4.95, 10.8, 1.2, spec["standfirst"], 15,
                 col(style, "muted"), latin=ql, ea=qe, bold=False, italic=True, spacing=1.55)
    # 底部元信息三栏
    metas = spec.get("meta", [])
    if metas:
        rect(slide, 0.85, 6.35, 11.6, 0.014, fill=col(style, "ink"))
        for i, m in enumerate(metas[:3]):
            x = 0.85 + i * 3.95
            text(slide, x, 6.55, 3.6, 0.3, m.get("label", ""), size=9.5, color=col(style, "faint"), tracking=1.5)
            text(slide, x, 6.85, 3.6, 0.35, m.get("value", ""), size=12, color=col(style, "ink"), bold=True)
            if i > 0:
                rect(slide, x - 0.2, 6.55, 0.01, 0.65, fill=col(style, "line"))
    _note(slide, spec)


def r_ed_split(slide, spec, ctx):
    """不对称分栏:超大首字数字 + 竖发丝线 + 引文 + 正文。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _serif(style)
    ql, qe = _fonts(style, "quote")
    nl, ne = _fonts(style, "numeral")
    text(slide, 0.85, 0.6, 6.0, 0.3, spec.get("kicker", ""), size=11, color=col(style, "accent"),
         bold=True, tracking=3)
    # 左栏:超大数字(首字下沉)
    text_art(slide, 0.7, 1.35, 3.9, 2.4, spec.get("numeral", "01"), spec.get("numeral_size", 130),
             col(style, "accent"), latin=nl, ea=ne, bold=True, spacing=1.0)
    if spec.get("numeral_label"):
        text(slide, 0.85, 3.95, 3.5, 0.6, spec["numeral_label"], size=12, color=col(style, "muted"), spacing=1.4)
    rect(slide, 4.75, 1.2, 0.014, 5.5, fill=col(style, "ink"))
    # 右栏:标题 + 引文 + 正文
    text_art(slide, 5.15, 1.15, 7.3, 1.6, spec["title"], 30, col(style, "ink"),
             latin=hl, ea=he, bold=True, spacing=1.15)
    if spec.get("quote"):
        rect(slide, 5.15, 2.85, 0.05, 1.05, fill=col(style, "accent"))
        text_art(slide, 5.42, 2.85, 6.9, 1.1, spec["quote"], 14, col(style, "muted"),
                 latin=ql, ea=qe, bold=False, italic=True, spacing=1.55)
    if spec.get("body"):
        rect(slide, 5.15, 4.25, 7.3, 0.008, fill=col(style, "line"))
        text(slide, 5.15, 4.45, 7.3, 2.2, spec["body"], size=12, color=col(style, "ink"), spacing=1.7)
    _note(slide, spec)


def r_ed_report(slide, spec, ctx):
    """报纸报道页:报头黑条 + 反色焦点格 + 多栏正文。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _serif(style)
    # 报头
    rect(slide, 0, 0, K.PAGE_W, 1.15, fill=col(style, "ink"))
    text(slide, 0.85, 0.22, 7.0, 0.3, spec.get("kicker", ""), size=10.5, color=col(style, "accent2"),
         bold=True, tracking=3)
    text_art(slide, 0.82, 0.5, 9.0, 0.55, spec["title"], 24, col(style, "on_ink"),
             latin=hl, ea=he, bold=True)
    text(slide, 10.0, 0.5, 2.45, 0.4, spec.get("date", ""), size=11, color=col(style, "on_ink"),
         align=PP_ALIGN.RIGHT, font=MONO)
    # 反色焦点格(左) + 多栏正文(右)
    rect(slide, 0.85, 1.55, 4.5, 4.9, fill=col(style, "ink"))
    if spec.get("focus_big"):
        nl, ne = _fonts(style, "numeral")
        text_art(slide, 1.15, 1.95, 3.9, 1.5, spec["focus_big"], 52, col(style, "accent"),
                 latin=nl, ea=ne, bold=True)
    if spec.get("focus_label"):
        text(slide, 1.15, 3.55, 3.9, 0.4, spec["focus_label"], size=13, color=col(style, "on_ink"), bold=True)
    if spec.get("focus_text"):
        text(slide, 1.15, 4.05, 3.9, 2.1, spec["focus_text"], size=11, color=C("C9C2B4"), spacing=1.6)
    # 多栏
    cols = spec.get("columns", [])
    n = len(cols)
    if n:
        cw = (7.15 - (n - 1) * 0.3) / n
        for i, c in enumerate(cols):
            x = 5.65 + i * (cw + 0.3)
            text(slide, x, 1.55, cw, 0.35, c.get("head", ""), size=10.5, color=col(style, "accent"),
                 bold=True, tracking=1.5)
            rect(slide, x, 1.92, cw, 0.014, fill=col(style, "ink"))
            text(slide, x, 2.1, cw, 4.2, c.get("body", ""), size=11, color=col(style, "ink"), spacing=1.65)
    if spec.get("footnote"):
        rect(slide, 0.85, 6.7, 11.6, 0.008, fill=col(style, "line"))
        text(slide, 0.85, 6.85, 11.6, 0.3, spec["footnote"], size=9.5, color=col(style, "faint"))
    _note(slide, spec)


RENDERERS.update({
    "ed_cover": r_ed_cover,
    "ed_split": r_ed_split,
    "ed_report": r_ed_report,
})



# ============================================================
# v3.4 杂志风精修(用户认证四页后的挑剔式升级)
# 1. 封面大标题横穿发丝线 + 左缘书脊竖排
# 2. 章节页正文首字下沉(drop cap)
# 3. 新增 ed_figures:报纸行情式数据行(无卡片,发丝线分隔)
# ============================================================

def _drop_cap(slide, style, x, y, ch, size=42):
    """首字下沉:正文首字符放大成装饰字(accent 色),正文从其右侧接排。"""
    hl, he = _fonts(style, "heading")
    text_art(slide, x, y, 1.0, 0.95, ch, size, col(style, "accent"), latin=hl, ea=he,
             bold=True, spacing=1.0)


def r_ed_figures(slide, spec, ctx):
    """报纸行情式数据行:双线报界 + 发丝线分隔的大数字(无卡片)。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _serif(style)
    nl, ne = _fonts(style, "numeral")
    text(slide, 0.85, 0.6, 6.0, 0.3, spec.get("kicker", ""), size=11, color=col(style, "accent"),
         bold=True, tracking=3)
    text_art(slide, 0.8, 0.95, 11.7, 0.9, spec["title"], 30, col(style, "ink"),
             latin=hl, ea=he, bold=True)
    rect(slide, 0.85, 1.85, 11.6, 0.016, fill=col(style, "ink"))
    rect(slide, 0.85, 1.92, 11.6, 0.006, fill=col(style, "ink"))
    figs = spec["figures"]
    if not 2 <= len(figs) <= 4:
        raise ValueError(f"ed_figures 支持 2-4 个数字, 收到 {len(figs)} 个")
    n = len(figs)
    col_w = 11.6 / n

    def _em(t):
        u = 0.0
        for ch in t:
            o = ord(ch)
            if o > 0x2E7F:
                u += 1.0
            elif ch == " ":
                u += 0.28
            elif ch.isdigit():
                u += 0.56
            elif ch.isupper():
                u += 0.72
            elif ch.islower():
                u += 0.5
            else:
                u += 0.33
        return u

    for i, f in enumerate(figs):
        x = 0.85 + i * col_w
        # 精确宽度自适应:按 em 宽估算,长数字自动缩号防撞邻列
        avail_em = (col_w - 0.4) * 72 / 46
        num_size = 46 if _em(f["num"]) <= avail_em else max(28, int(46 * avail_em / _em(f["num"])))
        text_art(slide, x, 2.35, col_w - 0.35, 1.15, f["num"], num_size, col(style, "ink"),
                 latin=nl, ea=ne, bold=True)
        text(slide, x, 3.58, col_w - 0.45, 0.35, f["label"], size=12.5, color=col(style, "ink"), bold=True)
        if f.get("sub"):
            text(slide, x, 3.95, col_w - 0.5, 0.75, f["sub"], size=10.5, color=col(style, "muted"), spacing=1.4)
        if i > 0:
            rect(slide, x - 0.28, 2.35, 0.008, 2.3, fill=col(style, "line"))
    rect(slide, 0.85, 5.7, 11.6, 0.008, fill=col(style, "line"))
    if spec.get("body"):
        text(slide, 0.85, 5.95, 11.6, 0.9, spec["body"], size=11.5, color=col(style, "ink"), spacing=1.6)
    if spec.get("source"):
        text(slide, 0.85, 6.9, 11.6, 0.3, spec["source"], size=9.5, color=col(style, "faint"))
    _note(slide, spec)


def _patch_ed_cover(slide, spec):
    """ed_cover 精修:标题横穿发丝线 + 左缘书脊竖排。"""
    style = None  # 由调用方传入 ctx,这里直接用 spec 里的颜色令牌名
    return None


def r_ed_cover_v2(slide, spec, ctx):
    """杂志封面 v2:书脊 + 横穿发丝线的大标题 + 首字下沉导语。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _serif(style)
    ql, qe = _fonts(style, "quote")
    # 左缘书脊(竖排)
    if spec.get("spine"):
        text_art(slide, 0.18, 1.0, 0.5, 5.5, spec["spine"], 10, col(style, "faint"),
                 latin=hl, ea=he, bold=False, tracking=2, vert=True, spacing=1.6)
    rect(slide, 0.85, 0.8, 11.6, 0.014, fill=col(style, "ink"))
    text(slide, 0.85, 1.0, 6.0, 0.3, spec.get("kicker", ""), size=11, color=col(style, "accent"),
         bold=True, tracking=3)
    text(slide, 7.0, 1.0, 5.45, 0.3, spec.get("issue", ""), size=11, color=col(style, "muted"),
         align=PP_ALIGN.RIGHT, tracking=1)
    rect(slide, 0.85, 1.38, 11.6, 0.014, fill=col(style, "ink"))
    # 大标题 + 横穿其上的发丝线(杂志手法:标题跨过规则线)
    ty = 2.0
    text_art(slide, 0.8, ty, 11.9, 2.5, spec["title"], spec.get("size", 54),
             col(style, "ink"), latin=hl, ea=he, bold=True, spacing=spec.get("spacing", 1.08),
             tracking=spec.get("tracking", 0.5))
    rect(slide, 0.85, ty + spec.get("cross_at", 1.15), 11.6, 0.01, fill=col(style, "line"))
    rect(slide, 0.85, 4.72, 11.6, 0.008, fill=col(style, "line"))
    # 导语:首字下沉
    sf = spec.get("standfirst", "")
    if sf:
        _drop_cap(slide, style, 0.83, 4.98, sf[0], 34)
        text_art(slide, 1.5, 4.98, 9.9, 1.2, sf[1:], 15, col(style, "muted"),
                 latin=ql, ea=qe, bold=False, italic=True, spacing=1.55)
    metas = spec.get("meta", [])
    if metas:
        rect(slide, 0.85, 6.35, 11.6, 0.014, fill=col(style, "ink"))
        for i, m in enumerate(metas[:3]):
            x = 0.85 + i * 3.95
            text(slide, x, 6.55, 3.6, 0.3, m.get("label", ""), size=9.5, color=col(style, "faint"), tracking=1.5)
            text(slide, x, 6.85, 3.6, 0.35, m.get("value", ""), size=12, color=col(style, "ink"), bold=True)
            if i > 0:
                rect(slide, x - 0.2, 6.55, 0.01, 0.65, fill=col(style, "line"))
    _note(slide, spec)


def r_ed_split_v2(slide, spec, ctx):
    """不对称分栏章页 v2:超大数字跨线 + 首字下沉正文 + 大引文。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _serif(style)
    ql, qe = _fonts(style, "quote")
    nl, ne = _fonts(style, "numeral")
    text(slide, 0.85, 0.6, 6.0, 0.3, spec.get("kicker", ""), size=11, color=col(style, "accent"),
         bold=True, tracking=3)
    # 左栏:超大数字,横穿一条规则线
    text_art(slide, 0.7, 1.3, 3.9, 2.5, spec.get("numeral", "01"), spec.get("numeral_size", 130),
             col(style, "accent"), latin=nl, ea=ne, bold=True, spacing=1.0)
    rect(slide, 0.85, 3.35, 3.4, 0.012, fill=col(style, "ink"))
    if spec.get("numeral_label"):
        text(slide, 0.85, 3.55, 3.5, 0.8, spec["numeral_label"], size=12, color=col(style, "muted"), spacing=1.45)
    rect(slide, 4.75, 1.2, 0.014, 5.5, fill=col(style, "ink"))
    # 右栏
    text_art(slide, 5.15, 1.1, 7.3, 1.6, spec["title"], 30, col(style, "ink"),
             latin=hl, ea=he, bold=True, spacing=1.15)
    if spec.get("quote"):
        rect(slide, 5.15, 2.8, 0.05, 1.2, fill=col(style, "accent"))
        text_art(slide, 5.42, 2.75, 6.9, 1.35, spec["quote"], 15.5, col(style, "muted"),
                 latin=ql, ea=qe, bold=False, italic=True, spacing=1.55)
    if spec.get("body"):
        rect(slide, 5.15, 4.35, 7.3, 0.008, fill=col(style, "line"))
        # 首字下沉 + 接排
        _drop_cap(slide, style, 5.13, 4.55, spec["body"][0], 34)
        text(slide, 5.85, 4.55, 6.6, 2.1, spec["body"][1:], size=12, color=col(style, "ink"), spacing=1.7)
    _note(slide, spec)


# 覆盖旧版(保留旧名以兼容已有 spec)
RENDERERS.update({
    "ed_cover": r_ed_cover_v2,
    "ed_split": r_ed_split_v2,
    "ed_figures": r_ed_figures,
})



# ============================================================
# v3.5 杂志风再精修(用户再评审:1满分/2·5合格/3·4·6待打磨)
# 3号页:头条数字制(一个主角+三个配角),替代四数字平铺
# 4/6号页:报名双线报头 + 不等宽分栏 + 首栏首字下沉 + 焦点数字放大
# ============================================================

def r_ed_figures_v2(slide, spec, ctx):
    """报纸行情式数据页 v2:头条数字(左大) + 配角数字(右列),主次分明。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _serif(style)
    nl, ne = _fonts(style, "numeral")

    def _em(t):
        u = 0.0
        for ch in t:
            o = ord(ch)
            if o > 0x2E7F:
                u += 1.0
            elif ch == " ":
                u += 0.28
            elif ch.isdigit():
                u += 0.56
            elif ch.isupper():
                u += 0.72
            elif ch.islower():
                u += 0.5
            else:
                u += 0.33
        return u

    # 报头
    rect(slide, 0.85, 0.72, 0.09, 0.09, fill=col(style, "accent"))
    text(slide, 1.05, 0.62, 6.0, 0.3, spec.get("kicker", ""), size=11, color=col(style, "accent"),
         bold=True, tracking=3)
    text_art(slide, 0.8, 0.95, 11.7, 0.9, spec["title"], 30, col(style, "ink"),
             latin=hl, ea=he, bold=True)
    rect(slide, 0.85, 1.85, 11.6, 0.016, fill=col(style, "ink"))
    rect(slide, 0.85, 1.92, 11.6, 0.006, fill=col(style, "ink"))

    figs = spec["figures"]
    if not 2 <= len(figs) <= 4:
        raise ValueError(f"ed_figures 支持 2-4 个数字, 收到 {len(figs)} 个")
    lead, rest = figs[0], figs[1:]

    # 头条数字(左栏,约 4.6in)
    lx = 0.85
    avail_em = 4.3 * 72 / 60
    lead_size = 60 if _em(lead["num"]) <= avail_em else max(34, int(60 * avail_em / _em(lead["num"])))
    text_art(slide, lx, 2.3, 4.5, 1.3, lead["num"], lead_size, col(style, "ink"),
             latin=nl, ea=ne, bold=True)
    text(slide, lx, 3.75, 4.2, 0.4, lead["label"], size=14, color=col(style, "ink"), bold=True)
    if lead.get("sub"):
        text(slide, lx, 4.18, 4.2, 0.8, lead["sub"], size=11, color=col(style, "muted"), spacing=1.45)
    rect(slide, 5.25, 2.3, 0.014, 2.9, fill=col(style, "ink"))

    # 配角数字(右列,横排)
    n = len(rest)
    rw = (7.2 - (n - 1) * 0.3) / n
    for i, f in enumerate(rest):
        x = 5.65 + i * (rw + 0.3)
        avail = (rw - 0.1) * 72 / 34
        sz = 34 if _em(f["num"]) <= avail else max(22, int(34 * avail / _em(f["num"])))
        text_art(slide, x, 2.35, rw, 0.85, f["num"], sz, col(style, "ink"),
                 latin=nl, ea=ne, bold=True)
        text(slide, x, 3.25, rw, 0.35, f["label"], size=12, color=col(style, "ink"), bold=True)
        if f.get("sub"):
            text(slide, x, 3.62, rw - 0.1, 0.9, f["sub"], size=10.5, color=col(style, "muted"), spacing=1.4)
        if i > 0:
            rect(slide, x - 0.16, 2.35, 0.008, 2.1, fill=col(style, "line"))

    rect(slide, 0.85, 5.55, 11.6, 0.008, fill=col(style, "line"))
    if spec.get("body"):
        text(slide, 0.85, 5.8, 11.6, 0.9, spec["body"], size=11.5, color=col(style, "ink"), spacing=1.6)
    if spec.get("source"):
        text(slide, 0.85, 6.9, 11.6, 0.3, spec["source"], size=9.5, color=col(style, "faint"))
    _note(slide, spec)


def r_ed_report_v2(slide, spec, ctx):
    """报纸报道页 v2:报名双线 + 不等宽分栏(首栏宽,带首字下沉)+ 焦点数字放大。"""
    style = ctx["style"]
    _bg(slide, style)
    hl, he = _serif(style)
    nl, ne = _fonts(style, "numeral")
    # 报名(黑条 + 双规则线)
    rect(slide, 0, 0, K.PAGE_W, 1.05, fill=col(style, "ink"))
    text(slide, 0.85, 0.2, 7.0, 0.28, spec.get("kicker", ""), size=10.5, color=col(style, "accent2"),
         bold=True, tracking=3)
    text_art(slide, 0.82, 0.46, 9.6, 0.5, spec["title"], 23, col(style, "on_ink"),
             latin=hl, ea=he, bold=True)
    text(slide, 10.4, 0.5, 2.05, 0.4, spec.get("date", ""), size=11, color=col(style, "on_ink"),
         align=PP_ALIGN.RIGHT, font=MONO)
    rect(slide, 0, 1.05, K.PAGE_W, 0.03, fill=col(style, "ink"))
    rect(slide, 0, 1.12, K.PAGE_W, 0.012, fill=col(style, "ink"))
    # 反色焦点格
    rect(slide, 0.85, 1.5, 4.4, 4.95, fill=col(style, "ink"))
    fb = spec.get("focus_big", "")
    if fb:
        avail = 3.9 * 72 / 64
        fsz = 64 if _em_pub(fb) <= avail else max(34, int(64 * avail / _em_pub(fb)))
        text_art(slide, 1.15, 1.9, 3.9, 1.5, fb, fsz, col(style, "accent"),
                 latin=nl, ea=ne, bold=True)
    if spec.get("focus_label"):
        text(slide, 1.15, 3.6, 3.8, 0.4, spec["focus_label"], size=13, color=col(style, "on_ink"), bold=True)
    if spec.get("focus_text"):
        text(slide, 1.15, 4.1, 3.8, 2.1, spec["focus_text"], size=11, color=C("C9C2B4"), spacing=1.6)
    # 不等宽分栏:首栏 1.5 倍宽,带首字下沉
    cols = spec.get("columns", [])
    n = len(cols)
    if n:
        total_w = 7.15
        weights = [1.5] + [1.0] * (n - 1)
        wsum = sum(weights)
        widths = [total_w * wt / wsum for wt in weights]
        x = 5.65
        for i, c in enumerate(cols):
            cw = widths[i] - 0.3
            text(slide, x, 1.55, cw, 0.35, c.get("head", ""), size=10.5, color=col(style, "accent"),
                 bold=True, tracking=1.5)
            rect(slide, x, 1.92, cw, 0.016, fill=col(style, "ink"))
            body = c.get("body", "")
            if i == 0 and body:
                _drop_cap(slide, style, x - 0.02, 2.12, body[0], 30)
                text(slide, x + 0.62, 2.12, cw - 0.62, 4.2, body[1:], size=11,
                     color=col(style, "ink"), spacing=1.65)
            else:
                text(slide, x, 2.12, cw, 4.2, body, size=11, color=col(style, "ink"), spacing=1.65)
            x += widths[i]
    if spec.get("footnote"):
        rect(slide, 0.85, 6.7, 11.6, 0.008, fill=col(style, "line"))
        text(slide, 0.85, 6.85, 11.6, 0.3, spec["footnote"], size=9.5, color=col(style, "faint"))
    _note(slide, spec)


def _em_pub(t):
    u = 0.0
    for ch in t:
        o = ord(ch)
        if o > 0x2E7F:
            u += 1.0
        elif ch == " ":
            u += 0.28
        elif ch.isdigit():
            u += 0.56
        elif ch.isupper():
            u += 0.72
        elif ch.islower():
            u += 0.5
        else:
            u += 0.33
    return u


RENDERERS.update({
    "ed_figures": r_ed_figures_v2,
    "ed_report": r_ed_report_v2,
})



# ============================================================
# v4.1 table 原生表格版式(三风格各自的表格语言)
# ============================================================

def _cell_border(cell, color_hex, w_pt=1.0, edges="LRTB"):
    """给单元格加边框(XML 注入;python-pptx 无边框 API)。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for e in reversed(edges):  # 逆序插到最前,保证 lnL/lnR/lnT/lnB 顺序
        tag = {"L": "a:lnL", "R": "a:lnR", "T": "a:lnT", "B": "a:lnB"}[e]
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = tcPr.makeelement(qn(tag), {"w": str(int(w_pt * 12700)), "cap": "flat"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        clr = fill.makeelement(qn("a:srgbClr"), {"val": color_hex.lstrip("#")})
        fill.append(clr)
        ln.append(fill)
        tcPr.insert(0, ln)


def _cell_text(cell, s, size, color, bold=False, align=PP_ALIGN.LEFT, font=YAHEI):
    tf = cell.text_frame
    tf.word_wrap = True
    cell.margin_left = cell.margin_right = IN(0.1)
    cell.margin_top = cell.margin_bottom = IN(0.04)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = s
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    K._set_font(r, font)


def r_table(slide, spec, ctx):
    """原生表格:可编辑。三风格表格语言:
    brutal=粗黑框+黄表头 / dark=细线+金表头 / editorial=横发丝线+墨表头(无竖线)。"""
    style = ctx["style"]
    _bg(slide, style)
    _header(slide, style, spec.get("kicker", spec["title"]))
    _h1(slide, style, spec["title"])
    _hair(slide, style, 1.72) if style["name"] != "brutal" else rect(slide, 0.6, 1.14, 11.6, 0.05, fill=col(style, "ink"))
    columns = spec["columns"]
    rows = spec["rows"]
    if not columns or not rows:
        raise ValueError("table 需要 columns 与 rows")
    hl = spec.get("highlight_col")
    n_rows, n_cols = len(rows) + 1, len(columns)
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, IN(0.85), IN(2.05), IN(11.6), IN(0.5 * n_rows))
    tbl = tbl_shape.table
    tbl.first_row = False       # 关默认样式,全手工
    tbl.horz_banding = False
    bw = style["geometry"]["border_w"]
    border_w = max(bw, 1.5) if style["name"] == "brutal" else 0.75
    ink_hex = "1B2432" if style["name"] == "dark_premium" else "111111"
    line_hex = "33405A" if style["name"] == "dark_premium" else ("D8D2C4" if style["name"] == "editorial" else "111111")
    for ci, cname in enumerate(columns):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        if style["name"] == "editorial":
            cell.fill.background()
        else:
            cell.fill.fore_color.rgb = col(style, "accent")
        _cell_text(cell, cname, 12.5, col(style, "on_ink") if style["name"] != "editorial" else col(style, "ink"),
                   bold=True, align=PP_ALIGN.CENTER)
        _cell_border(cell, line_hex if style["name"] != "editorial" else ink_hex, border_w, "TB")
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = col(style, "card") if style["name"] != "editorial" else col(style, "bg")
            is_hl = hl is not None and ci == hl
            _cell_text(cell, str(val), 11.5,
                       col(style, "accent") if is_hl else col(style, "ink"),
                       bold=is_hl or ci == 0,
                       align=PP_ALIGN.CENTER if ci == 0 or is_hl else PP_ALIGN.LEFT)
            _cell_border(cell, line_hex, border_w,
                         "TB" if style["name"] == "editorial" else "LRTB")
    if spec.get("note_row"):
        text(slide, 0.85, 2.2 + 0.5 * n_rows + 0.25, 11.6, 0.5, spec["note_row"],
             size=10.5, color=col(style, "muted"))
    _note(slide, spec)


RENDERERS.update({"table": r_table})


if __name__ == "__main__":
    argv = sys.argv[1:]
    style_arg = None
    if "--style" in argv:
        i = argv.index("--style")
        style_arg = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    if len(argv) != 2:
        print(__doc__)
        sys.exit(1)
    build(argv[0], argv[1], style_arg)
