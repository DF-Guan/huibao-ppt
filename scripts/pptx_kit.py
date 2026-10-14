# -*- coding: utf-8 -*-
"""pptx_kit — 「汇报大师 PPT」原生渲染引擎(原创)。

职责:把 python-pptx 包装成一套声明式图元,并注入其原生 API 不直接支持的
高级能力(填充透明度、虚线、字距、竖排、切换动效、页内跳转链接)。

设计约束:
- 所有函数只依赖 python-pptx 与标准库;
- 坐标单位一律为英寸(inch),页面画布 13.333 x 7.5(16:9);
- 失败立即抛错,不静默吞错。
"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

IN = Inches

# 常用字体(中文环境 Windows 默认具备)
YAHEI = "微软雅黑"
KAI = "楷体"
SONG = "宋体"
MONO = "Consolas"
GEO = "Georgia"

PAGE_W, PAGE_H = 13.333, 7.5


def C(hex_str):
    """'#RRGGBB' 或 'RRGGBB' -> RGBColor"""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ---------- 底层 XML 注入 ----------

def _set_font(run, name):
    """同时设置 latin 与东亚字体,保证中文按指定字体渲染。"""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", name)


def _tracking(run, pts):
    """字距(磅)。python-pptx 不支持,直接写 rPr@spc(1/100 磅)。"""
    run._r.get_or_add_rPr().set("spc", str(int(pts * 100)))


def set_alpha(shape, opacity_pct):
    """设置纯色填充的不透明度(0-100)。用于色场/氛围层。"""
    try:
        sF = shape.fill._xPr.find(qn("a:solidFill"))
        clr = sF.find(qn("a:srgbClr"))
        clr.append(clr.makeelement(qn("a:alpha"), {"val": str(int(opacity_pct * 1000))}))
    except Exception as exc:
        raise RuntimeError(f"set_alpha 失败: {exc}") from exc


def set_dash(shape, val="dash"):
    """边框虚线。val: dash / sysDot / lgDash 等 DrawingML 预设。"""
    ln = shape.line._get_or_add_ln()
    if ln.find(qn("a:prstDash")) is None:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": val}))


def set_vert(tf):
    """文本框竖排(东亚竖写)。"""
    tf._txBody.find(qn("a:bodyPr")).set("vert", "eaVert")


def add_transition(slide, kind="fade", speed="med"):
    """页面切换动效。kind: fade / push。"""
    sld = slide._element
    trans = sld.makeelement(qn("p:transition"), {"spd": speed})
    if kind == "fade":
        trans.append(trans.makeelement(qn("p:fade"), {}))
    elif kind == "push":
        trans.append(trans.makeelement(qn("p:push"), {"dir": "l"}))
    else:
        raise ValueError(f"不支持的切换动效: {kind}")
    sld.append(trans)


def link(shape, target_slide):
    """点击形状跳转到目标页(放映态生效)。"""
    shape.click_action.target_slide = target_slide


# ---------- 声明式图元 ----------

def rect(slide, x, y, w, h, fill=None, line=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE,
         radius=None, rot=None, alpha=None, dash=None):
    """矩形/圆角矩形/圆/三角/星形等自绘形状。坐标单位:英寸。"""
    sp = slide.shapes.add_shape(shape, IN(x), IN(y), IN(w), IN(h))
    sp.shadow.inherit = False  # 禁用默认投影,硬投影由调用方显式绘制
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(lw)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    if rot is not None:
        sp.rotation = rot
    if alpha is not None:
        set_alpha(sp, alpha)
    if dash is not None:
        set_dash(sp, dash)
    return sp


def _fill_tf(tb, lines, size, color, bold, align, font, anchor, spacing, tracking):
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        if color is not None:
            r.font.color.rgb = color
        _set_font(r, font)
        if tracking:
            _tracking(r, tracking)
    return tf


def text(slide, x, y, w, h, s, size=14, color=None, bold=False, align=PP_ALIGN.LEFT,
         font=YAHEI, anchor=MSO_ANCHOR.TOP, spacing=1.0, tracking=None, rot=None, vert=False):
    """文本框。s 支持 \\n 分段。"""
    tb = slide.shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
    tf = _fill_tf(tb, s.split("\n"), size, color, bold, align, font, anchor, spacing, tracking)
    if rot is not None:
        tb.rotation = rot
    if vert:
        set_vert(tf)
    return tb


def runs_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """单段混排文本。runs: [(文本, 字号, 颜色, 加粗, 字体), ...]"""
    tb = slide.shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    for txt, size, color, bold, font in runs:
        r = p.add_run()
        r.text = txt
        r.font.size = Pt(size)
        r.font.bold = bold
        if color is not None:
            r.font.color.rgb = color
        _set_font(r, font)
    return tb


def shape_text(sp, s, size, color, bold=False, align=PP_ALIGN.CENTER, font=YAHEI, spacing=1.0):
    """向已存在的形状内写文字(如色块按钮、印章、爆炸星)。"""
    tf = sp.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, ln in enumerate(s.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        _set_font(r, font)


# ---------- 版式 lint(文字溢出估算) ----------

def _cjk_units(s):
    """估算排版宽度单位:CJK 记 1,ASCII 记 0.55。"""
    return sum(1.0 if ord(ch) > 0x2E7F else 0.55 for ch in s)


def lint_slide(slide, page_no):
    """对每个文本框做溢出估算,返回警告列表(启发式,仅供参考)。"""
    warnings = []
    for sh in slide.shapes:
        if not sh.has_text_frame or not sh.text_frame.text.strip():
            continue
        box_w = sh.width / 914400  # EMU -> inch
        box_h = sh.height / 914400
        for p in sh.text_frame.paragraphs:
            txt = "".join(r.text for r in p.runs)
            if not txt.strip():
                continue
            size = max((r.font.size.pt for r in p.runs if r.font.size), default=14)
            spacing = p.line_spacing if isinstance(p.line_spacing, (int, float)) else 1.0
            spacing = max(spacing, 1.0)
            units_per_line = max(box_w * 72 / size, 1)
            est_lines = max(1, -(-_cjk_units(txt) // units_per_line))  # 向上取整
            est_h = est_lines * size * spacing * 1.28 / 72
            if est_h > box_h * 1.12:
                warnings.append(
                    f"page{page_no} 疑似溢出: {txt[:14]!r} 需约 {est_h:.2f}in, 框高 {box_h:.2f}in"
                )
    return warnings


# ============================================================
# 实战增强(引擎层):照片 / 图表 / logo墙 / 组织架构 / 备注 / 字阶 / 容量
# 依据汇报实战经验沉淀
# ============================================================
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from PIL import Image as _PILImage


# ---------- 照片版式 ----------

def photo_contain(slide, path, x, y, w, h, bg=None, pad=0.06):
    """完整显示不裁切:底衬 + 按比例 contain 居中。产品/证书/场景图用。"""
    if bg is not None:
        rect(slide, x, y, w, h, fill=bg)
    pic = slide.shapes.add_picture(path, IN(x), IN(y))
    iw, ih = pic.image.size
    bw, bh = w - 2 * pad, h - 2 * pad
    ratio = min(bw / (iw / 96), bh / (ih / 96))
    pw, ph = iw / 96 * ratio, ih / 96 * ratio
    pic.width, pic.height = IN(pw), IN(ph)
    pic.left = IN(x + (w - pw) / 2)
    pic.top = IN(y + (h - ph) / 2)
    return pic


def photo_cover(slide, path, x, y, w, h):
    """满版出血(裁切填满)。仅用于全屏背景大图,勿用于产品图。"""
    pic = slide.shapes.add_picture(path, IN(x), IN(y), IN(w), IN(h))
    try:
        iw, ih = pic.image.size
        br, ir = w / h, iw / ih
        if ir > br:
            c = (1 - br / ir) / 2
            pic.crop_left = pic.crop_right = c
        else:
            c = (1 - ir / br) / 2
            pic.crop_top = pic.crop_bottom = c
    except Exception:
        pass
    return pic


def logo_fit(slide, path, x, y, w, h):
    """logo 按比例 contain 居中(无底衬)。"""
    pic = slide.shapes.add_picture(path, IN(x), IN(y))
    iw, ih = pic.image.size
    ratio = min(w / (iw / 96), h / (ih / 96))
    pw, ph = iw / 96 * ratio, ih / 96 * ratio
    pic.width, pic.height = IN(pw), IN(ph)
    pic.left = IN(x + (w - pw) / 2)
    pic.top = IN(y + (h - ph) / 2)
    return pic


def logo_wall(slide, paths, x, y, w, h, per_row=6, box_fill=None, box_line=None, pad=0.12):
    """logo 墙:自动换行防溢出。在 x,y,w,h 区域内按 per_row 列网格排列。"""
    n = len(paths)
    rows = -(-n // per_row)
    cols = min(per_row, n)
    gap = 0.14
    bw = (w - (cols - 1) * gap) / cols
    bh = (h - (rows - 1) * gap) / rows
    for i, p in enumerate(paths):
        col, row = i % per_row, i // per_row
        cx = x + col * (bw + gap)
        cy = y + row * (bh + gap)
        if box_fill is not None:
            rect(slide, cx, cy, bw, bh, fill=box_fill, line=box_line, lw=0.75)
        logo_fit(slide, p, cx + pad, cy + pad, bw - 2 * pad, bh - 2 * pad)


# ---------- 原生图表 ----------

def add_chart(slide, kind, categories, series_name, values, x, y, w, h,
              highlight_idx=None, highlight_color=None, base_color=None,
              legend=False, number_format=None):
    """原生图表(柱/饼/环)。highlight_idx 用强调色突出某一项。返回 chart。"""
    type_map = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "pie": XL_CHART_TYPE.PIE,
                "doughnut": XL_CHART_TYPE.DOUGHNUT}
    if kind not in type_map:
        raise ValueError(f"不支持的图表类型: {kind}, 支持 bar/pie/doughnut")
    cd = CategoryChartData()
    cd.categories = categories
    cd.add_series(series_name, values)
    gf = slide.shapes.add_chart(type_map[kind], IN(x), IN(y), IN(w), IN(h), cd)
    ch = gf.chart
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.RIGHT
        ch.legend.font.size = Pt(9)
    plot = ch.plots[0]
    plot.has_data_labels = True
    if number_format:
        plot.data_labels.number_format = number_format
        plot.data_labels.number_format_is_linked = False
    plot.data_labels.font.size = Pt(8)
    if kind == "bar":
        ch.category_axis.tick_labels.font.size = Pt(9)
        ch.value_axis.visible = False
        if highlight_idx is not None and base_color and highlight_color:
            ser = plot.series[0]
            for idx in range(len(categories)):
                pt = ser.points[idx]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = highlight_color if idx == highlight_idx else base_color
    return ch


# ---------- 组织架构图 ----------

def connector(slide, x, y, w, h, color, lw=1.5):
    """正交连接线(用细矩形模拟,横或竖)。"""
    return rect(slide, x, y, max(w, 0.014), max(h, 0.014), fill=color)


def org_node(slide, x, y, w, h, label, fill, tcolor, sub=None, size=13):
    """组织架构节点。返回形状。"""
    b = rect(slide, x, y, w, h, fill=fill)
    tf = b.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = IN(0.04)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = tcolor
    _set_font(r, YAHEI)
    if sub:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = sub
        r2.font.size = Pt(8)
        r2.font.color.rgb = tcolor
        _set_font(r2, YAHEI)
    return b


# ---------- 演讲者备注(口播稿) ----------

def add_notes(slide, note_text):
    """向页面写入演讲者备注(口播稿)。"""
    slide.notes_slide.notes_text_frame.text = note_text


# ---------- 字阶吸附 ----------

def snap_to_scale(size, scale):
    """把字号向下吸附到最近的一级字阶,保证全篇同类文字大小一致。"""
    scale = sorted(scale)
    for s in reversed(scale):
        if size >= s:
            return s
    return scale[0]


# ---------- 容量估算 ----------

def text_capacity(w_in, h_in, size_pt, spacing=1.0):
    """估算一个框在指定字号下能装多少"排版单位"(CJK=1, ASCII=0.55)。"""
    units_per_line = max(w_in * 72 / size_pt, 1)
    lines = max(int(h_in / (size_pt * max(spacing, 1.0) * 1.28 / 72)), 1)
    return units_per_line * lines


def check_capacity(text_str, w_in, h_in, size_pt, spacing=1.0):
    """返回 (是否装得下, 占用比例)。"""
    need = _cjk_units(text_str)
    # 估算所需行数
    units_per_line = max(w_in * 72 / size_pt, 1)
    est_lines = max(1, -(-need // units_per_line))
    est_h = est_lines * size_pt * max(spacing, 1.0) * 1.28 / 72
    return est_h <= h_in, est_h / h_in if h_in else 1.0


# ============================================================
# 字体系统与艺术字变换(v3,中西文字体配对规范 + 文艺风实践)
# ============================================================

# 字体配对:西文(latin)与中文(ea)分离,按字符自动选用,兼容性最好
FONT_PAIRS = {
    "brutal": {
        "heading": ("Arial Black", "微软雅黑"),
        "body": ("Arial", "微软雅黑"),
        "display": ("Arial Black", "微软雅黑"),
        "quote": ("Georgia", "楷体"),
        "numeral": ("Arial Black", "微软雅黑"),
    },
    "editorial": {
        "heading": ("Georgia", "宋体"),
        "body": ("Arial", "微软雅黑"),
        "display": ("Georgia", "宋体"),
        "quote": ("Georgia", "楷体"),
        "numeral": ("Arial", "微软雅黑"),
    },
    "dark_premium": {
        "heading": ("Georgia", "微软雅黑"),
        "body": ("Arial", "微软雅黑"),
        "display": ("Georgia", "微软雅黑"),
        "quote": ("Georgia", "楷体"),
        "numeral": ("Arial", "微软雅黑"),
    },
}


def set_font_pair(run, latin, ea):
    """西文走 latin 字体,中文走 ea 字体(PowerPoint/WPS 按字符自动选用)。"""
    run.font.name = latin
    rPr = run._r.get_or_add_rPr()
    ea_el = rPr.find(qn("a:ea"))
    if ea_el is None:
        ea_el = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea_el)
    ea_el.set("typeface", ea)
    cs = rPr.find(qn("a:cs"))
    if cs is None:
        cs = rPr.makeelement(qn("a:cs"), {})
        rPr.append(cs)
    cs.set("typeface", ea)


def _rpr_prop_solid(run, hex_color, tag="a:solidFill"):
    rPr = run._r.get_or_add_rPr()
    el = rPr.find(qn(tag))
    if el is not None:
        rPr.remove(el)
    fill = rPr.makeelement(qn(tag), {})
    clr = fill.makeelement(qn("a:srgbClr"), {"val": hex_color.lstrip("#")})
    fill.append(clr)
    rPr.insert(0, fill)


def text_stroke(run, hex_color, width_pt=1.0):
    """描边字(艺术字):给文字加同色系或反差色轮廓。"""
    rPr = run._r.get_or_add_rPr()
    ln = rPr.find(qn("a:ln"))
    if ln is None:
        ln = rPr.makeelement(qn("a:ln"), {})
        rPr.append(ln)
    ln.set("w", str(int(width_pt * 12700)))
    fill = ln.makeelement(qn("a:solidFill"), {})
    clr = fill.makeelement(qn("a:srgbClr"), {"val": hex_color.lstrip("#")})
    fill.append(clr)
    ln.append(fill)


def text_gradient(run, c1, c2, angle=5400000):
    """渐变字(艺术字):线性渐变填充文字。"""
    rPr = run._r.get_or_add_rPr()
    old = rPr.find(qn("a:solidFill"))
    if old is not None:
        rPr.remove(old)
    grad = rPr.makeelement(qn("a:gradFill"), {"flip": "none", "rotWithShape": "0"})
    lst = grad.makeelement(qn("a:gsLst"), {})
    for pos, col in ((0, c1), (100000, c2)):
        gs = lst.makeelement(qn("a:gs"), {"pos": str(pos)})
        clr = gs.makeelement(qn("a:srgbClr"), {"val": col.lstrip("#")})
        gs.append(clr)
        lst.append(gs)
    grad.append(lst)
    lin = grad.makeelement(qn("a:lin"), {"ang": str(angle), "scaled": "0"})
    grad.append(lin)
    rPr.insert(0, grad)


def text_art(slide, x, y, w, h, s, size, color, latin=None, ea=None, bold=True,
             italic=False, tracking=None, stroke=None, stroke_w=1.0,
             gradient=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0,
             rot=None, vert=False):
    """艺术字:字体配对 / 斜体 / 字距 / 描边 / 渐变,一站式。"""
    tb = slide.shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, ln in enumerate(s.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        use_latin = latin or "Arial"
        if use_latin in RISKY_NUMERAL_FONTS and any(c.isdigit() for c in ln):
            use_latin = "Arial"  # 旧式数字字体遇数字自动切等高线字体,防数字错位
        if latin or ea:
            set_font_pair(r, use_latin, ea or "微软雅黑")
        else:
            _set_font(r, YAHEI)
        if color is not None:
            r.font.color.rgb = color
        if tracking:
            _tracking(r, tracking)
        if gradient:
            text_gradient(r, gradient[0], gradient[1])
        if stroke:
            text_stroke(r, stroke, stroke_w)
    if rot is not None:
        tb.rotation = rot
    if vert:
        set_vert(tf)
    return tb


# ============================================================
# 文字错位监测(v3.1)—— 用户实战反馈:数字/字母与文字混排会出现错位
# ============================================================
# 旧式数字(old-style figures)字体:数字高低不齐,禁用于含数字的文本
RISKY_NUMERAL_FONTS = {"Georgia", "Times New Roman", "Times New Roman PS"}
# 等高线数字(lining figures)安全字体
SAFE_NUMERAL_FONTS = {"Arial", "Arial Black", "Consolas", "微软雅黑", "Helvetica",
                      "Helvetica Neue", "宋体", "黑体", "SimSun", "SimHei"}


def _run_fonts(run):
    names = set()
    if run.font.name:
        names.add(run.font.name)
    rPr = run._r.find(qn("a:rPr"))
    if rPr is not None:
        for tag in ("a:ea", "a:cs"):
            el = rPr.find(qn(tag))
            if el is not None and el.get("typeface"):
                names.add(el.get("typeface"))
    return names


def check_text_alignment(slide, page_no):
    """文字错位监测,返回警告列表。

    巡检三类问题:
    1. 旧式数字字体(Georgia 等)用于含数字文本 → 数字高低不齐;
    2. 同一行内字号不一致 → 基线漂移;
    3. 含数字文本使用了安全清单外的西文字体 → 对方电脑缺字体时回退,度量漂移。
    """
    warnings = []
    for sh in slide.shapes:
        if not sh.has_text_frame or not sh.text_frame.text.strip():
            continue
        for p in sh.text_frame.paragraphs:
            runs = [r for r in p.runs if r.text.strip()]
            if not runs:
                continue
            txt = p.text
            has_digit = any(c.isdigit() for c in txt)
            has_cjk = any(ord(ch) > 0x2E7F for ch in txt)
            # 注:同段不同字号不做警告——DrawingML 中混排自动基线对齐,
            # 且"大数字+小单位"(如 38 项)是刻意设计,报了就是误报。
            if has_digit:
                for r in runs:
                    fonts = _run_fonts(r)
                    risky = fonts & RISKY_NUMERAL_FONTS
                    if risky:
                        warnings.append(
                            f"page{page_no} [旧式数字] {sorted(risky)[0]} 用于含数字文本 "
                            f"{txt[:16]!r} → 数字高低不齐,应改用 Arial/Consolas/微软雅黑")
                    unknown = fonts - SAFE_NUMERAL_FONTS - RISKY_NUMERAL_FONTS
                    if unknown and has_cjk:
                        warnings.append(
                            f"page{page_no} [字体未备案] {sorted(unknown)[0]} 用于中英混排 "
                            f"{txt[:16]!r} → 缺字体时回退会错位")
    return warnings


# ============================================================
# v3.3 自动护栏:透明 logo 合成白底(防黑底)/ 安全图片读取
# ============================================================
import os
import tempfile

_WHITE_BG_CACHE = {}


def ensure_rgba_on_white(path):
    """RGBA 透明底图片 -> 合成到白底(临时文件),防止直接插入变黑底。

    有透明通道且四角存在透明像素时才处理;已缓存,重复调用不重复生成。
    """
    if path in _WHITE_BG_CACHE:
        return _WHITE_BG_CACHE[path]
    result = path
    try:
        im = _PILImage.open(path)
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGBA")
            w, h = im.size
            corners = [im.getpixel(p) for p in ((1, 1), (w - 2, 1), (1, h - 2), (w - 2, h - 2))]
            has_transparent_corner = any(len(c) == 4 and c[3] < 250 for c in corners)
            if has_transparent_corner:
                bg = _PILImage.new("RGBA", im.size, (255, 255, 255, 255))
                out = _PILImage.alpha_composite(bg, im).convert("RGB")
                fd, tmp = tempfile.mkstemp(suffix=".png", prefix="wb_")
                os.close(fd)
                out.save(tmp)
                result = tmp
    except Exception as exc:
        # 仅图片读取/解析失败时按原样放行;其余错误向上暴露,禁止静默吞错
        if "cannot identify image file" in str(exc) or "DecompressedBombError" in str(exc):
            result = path
        else:
            raise
    _WHITE_BG_CACHE[path] = result
    return result


# logo_fit / logo_wall 自动走白底防护
def _logo_fit_orig(slide, path, x, y, w, h):
    pic = slide.shapes.add_picture(path, IN(x), IN(y))
    iw, ih = pic.image.size
    ratio = min(w / (iw / 96), h / (ih / 96))
    pw, ph = iw / 96 * ratio, ih / 96 * ratio
    pic.width, pic.height = IN(pw), IN(ph)
    pic.left = IN(x + (w - pw) / 2)
    pic.top = IN(y + (h - ph) / 2)
    return pic


def logo_fit(slide, path, x, y, w, h):  # noqa: F811 — 覆盖上面的旧定义,加白底防护
    return _logo_fit_orig(slide, ensure_rgba_on_white(path), x, y, w, h)


def logo_wall(slide, paths, x, y, w, h, per_row=6, box_fill=None, box_line=None, pad=0.12):  # noqa: F811
    n = len(paths)
    rows = -(-n // per_row)
    cols = min(per_row, n)
    gap = 0.14
    bw = (w - (cols - 1) * gap) / cols
    bh = (h - (rows - 1) * gap) / rows
    for i, p in enumerate(paths):
        col, row = i % per_row, i // per_row
        cx = x + col * (bw + gap)
        cy = y + row * (bh + gap)
        if box_fill is not None:
            rect(slide, cx, cy, bw, bh, fill=box_fill, line=box_line, lw=0.75)
        _logo_fit_orig(slide, ensure_rgba_on_white(p), cx + pad, cy + pad, bw - 2 * pad, bh - 2 * pad)
