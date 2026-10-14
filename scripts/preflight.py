# -*- coding: utf-8 -*-
"""preflight — 渲染前一键全检:零问题才放行(客户零二次调教的第一道闸)。

用法:
    python preflight.py <deck-spec.json>

检查项:
  E 错误(必须修):JSON 非法 / 缺 slides / 未知页面类型 / 未知风格 / 图片路径不存在 /
                  必填字段缺失 / 占位文本(lorem/TODO/XXX)
  W 警告(建议修):容量超限 / 相邻同版式 / 缺口播备注 / 单页元素过少 / 数字字段未配 numeral 安全字体说明
退出码:有 E 为 1,只有 W 为 0(打印提醒)。
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import capacity as _cap

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

PLACEHOLDER = re.compile(r"lorem|TODO|TBD|XXX|占位|待补|placeholder", re.I)

LAYOUTS_REQUIRING = {
    "cover": ["title"], "cover_type": ["title"], "toc": ["items"],
    "stats": ["title", "cards"], "compare": ["title", "before", "after"],
    "items": ["title", "items"], "thanks": [],
    "chart": ["title", "chart"], "photo_gallery": ["title", "photos"],
    "logo_wall": ["title", "logos"], "org_chart": ["title", "top", "middle"],
    "scene": ["title", "image"], "timeline": ["title", "items"],
    "big_quote": ["quote"], "section": ["title"],
    "shot_stack": ["title", "items"], "overlap": ["title", "cards"],
    "ed_cover": ["title"], "ed_split": ["title", "numeral"],
    "ed_report": ["title"], "ed_figures": ["title", "figures"], "table": ["title", "columns", "rows"], "free": ["elements"],
}
NO_NOTE_OK = {"section", "big_quote", "ed_cover"}  # 这些页型允许无口播


def _collect_images(sd):
    """收集一页里引用的所有图片路径。"""
    paths = []
    for key in ("image", "bg_image"):
        if isinstance(sd.get(key), str):
            paths.append(sd[key])
    for key in ("items", "cards", "photos"):
        for it in sd.get(key) or []:
            if isinstance(it, dict):
                for k in ("image", "path"):
                    if isinstance(it.get(k), str):
                        paths.append(it[k])
    for el in sd.get("elements") or []:
        if isinstance(el, dict) and el.get("t") == "image" and el.get("path"):
            paths.append(el["path"])
    return paths


def check(spec_path):
    errors, warns = [], []
    try:
        with open(spec_path, encoding="utf-8") as f:
            spec = json.load(f)
    except Exception as e:
        return [f"JSON 解析失败: {e}"], []

    slides = spec.get("slides")
    if not slides:
        errors.append("spec 缺少 slides 数组")
        return errors, warns
    if not isinstance(slides, list):
        errors.append("slides 必须是数组")
        return errors, warns

    style = spec.get("style", "brutal")
    if style not in ("brutal", "dark_premium", "editorial"):
        errors.append(f"未知风格 {style!r},可选 brutal / dark_premium / editorial")

    prev_type = None
    for i, sd in enumerate(slides, 1):
        stype = sd.get("type")
        tag = f"page{i}"
        if stype not in LAYOUTS_REQUIRING:
            errors.append(f"{tag} 未知页面类型 {stype!r}")
            continue
        # 必填字段
        for field in LAYOUTS_REQUIRING[stype]:
            if field not in sd or sd[field] in (None, "", []):
                errors.append(f"{tag} [{stype}] 缺必填字段 {field!r}")
        # 占位文本
        blob = json.dumps(sd, ensure_ascii=False)
        m = PLACEHOLDER.search(blob)
        if m:
            errors.append(f"{tag} [{stype}] 含占位文本 {m.group(0)!r} —— 内容没备厚就出图,必翻车")
        # 图片路径
        for p in _collect_images(sd):
            if not os.path.isabs(p):
                warns.append(f"{tag} 图片路径非绝对路径: {p}")
            elif not os.path.exists(p):
                errors.append(f"{tag} [{stype}] 图片不存在: {p}")
        # 口播备注
        if stype not in NO_NOTE_OK and not sd.get("note"):
            warns.append(f"{tag} [{stype}] 缺口播备注(note) —— 汇报人打开备注栏会空白")
        # 节奏:相邻同版式
        if prev_type == stype and stype not in ("ed_split", "ed_report"):
            warns.append(f"{tag} [{stype}] 与上一页同版式 —— 建议换版式或插 section/big_quote 换气")
        prev_type = stype
        # 容量
        warns.extend(w.replace("CAPACITY ", "") for w in _cap.check_slide(i, sd))

    # 全册建议
    if len(slides) >= 6 and not any(s.get("type") == "toc" for s in slides):
        warns.append("6 页以上建议加目录页(toc),并开启 interactive")
    if not any(s.get("type") in ("thanks", "ed_cover") for s in slides) and len(slides) > 3:
        warns.append("缺结尾页(thanks) —— 汇报没有收束")
    return errors, warns


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    errors, warns = check(sys.argv[1])
    for e in errors:
        print("E:", e)
    for w in warns:
        print("W:", w)
    if errors:
        print(f"\n预检未通过:{len(errors)} 个错误必须修复后才能渲染。")
        sys.exit(1)
    print(f"\n预检通过({len(warns)} 个提醒)。可渲染: python scripts/build_deck.py {sys.argv[1]} <输出.pptx>")
    sys.exit(0)


if __name__ == "__main__":
    main()
