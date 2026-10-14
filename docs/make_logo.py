# -*- coding: utf-8 -*-
"""汇报大师 PPT 专属图标(学 Grokbot Icon 极简纪律, PIL 精确绘制)。

设计语言(参考 grokbot-icon-studio):
- 近黑圆角方底 #18191B,暖白几何字 #F7F5F0
- 全画布 4 个形状,20% 安全边距,零渐变零阴影
概念「画框中的幻灯片」:
- 左右两根暖白竖胶囊 = 画框/ brackets
- 中央黄铜金圆角矩形 = 幻灯片
- 金块内近黑三角 = 播放/汇报
"""
from PIL import Image, ImageDraw

S = 1024  # 母版尺寸
BG = (24, 25, 27)        # #18191B 近黑
OFFW = (247, 245, 240)   # #F7F5F0 暖白
GOLD = (201, 162, 91)    # #C9A25B 黄铜金

im = Image.new("RGB", (S, S), BG)
d = ImageDraw.Draw(im)

# 圆角方底(半径 22.5%)
radius = int(S * 0.225)
d.rounded_rectangle([0, 0, S - 1, S - 1], radius=radius, fill=BG)

# 安全边距 20%
m = int(S * 0.20)

# 左右竖胶囊(画框)
pill_w = int(S * 0.085)
pill_h = int(S * 0.44)
pill_r = pill_w // 2
pill_y = (S - pill_h) // 2
d.rounded_rectangle([m, pill_y, m + pill_w, pill_y + pill_h], radius=pill_r, fill=OFFW)
d.rounded_rectangle([S - m - pill_w, pill_y, S - m, pill_y + pill_h], radius=pill_r, fill=OFFW)

# 中央幻灯片(黄铜金圆角矩形)
sl_w = int(S * 0.24)
sl_h = int(S * 0.24)
sl_x = (S - sl_w) // 2
sl_y = (S - sl_h) // 2
d.rounded_rectangle([sl_x, sl_y, sl_x + sl_w, sl_y + sl_h], radius=int(S * 0.045), fill=GOLD)

# 金块内近黑三角(播放/汇报)
tri_h = int(sl_h * 0.42)
tri_w = int(tri_h * 0.9)
cx = sl_x + sl_w // 2 + int(tri_w * 0.06)  # 视觉居中补偿
cy = sl_y + sl_h // 2
d.polygon([(cx - tri_w // 2, cy - tri_h // 2),
           (cx - tri_w // 2, cy + tri_h // 2),
           (cx + tri_w // 2, cy)], fill=BG)

im.save(r"D:\资料M2\mywork\WorkBuddy\PPT\.workbuddy\skills\huibao-ppt\docs\logo.png", optimize=True)
# 512 市场头像(同一构图,无裁切)
im.resize((512, 512), Image.LANCZOS).save(
    r"D:\资料M2\mywork\WorkBuddy\PPT\.workbuddy\skills\huibao-ppt\docs\avatar.png", optimize=True)
print("logo.png 1024 + avatar.png 512 生成完毕(无 watermark,零裁切)")
