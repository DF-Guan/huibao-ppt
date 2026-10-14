# 页面原型手册(Spec 字段契约)

`deck-spec.json` 顶层结构:

```json
{
  "style": "brutal",
  "interactive": true,
  "transition": "push",
  "slides": [ { "type": "cover", "note": "口播备注", ... } ]
}
```

- `style`(默认 `"brutal"`):整体风格,可选 `brutal` / `dark_premium`,对应 `styles/<name>.json`。
- `interactive`(默认 `true`):开启目录跳页 + 内容页"目录"返回钮。
- `transition`(默认 `"push"`):全册切换动效,可选 `"fade"`。
- `slides`:页面数组,每个对象的 `type` 决定原型,字段见下文。**加粗字段为必填**。
- 每页可带 **`note`** 字段(口播备注),渲染后自动写入该页演讲者备注。

渲染命令:`python scripts/build_deck.py <spec.json> <输出.pptx> [--style 风格]`
容量自检:`python scripts/capacity.py <spec.json>`(超限提前报警)

---

## cover · 封面

构图:黑顶条(kicker)+ 倾斜黄底大标题 + 黑色副标题条 + 倾斜导语硬投影卡 + 红星爆 badge + 几何纸屑 + 底部 ticker 黑条。

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 主标题,≤16 字最佳(超出自动缩底衬宽度,注意 lint) |
| kicker | | 顶条英文/编码标签,如 `Q3 · 2026 · SECURITY REPORT` |
| subtitle | | 黑色副标题条文字,如部门+汇报人 |
| lead | | 导语卡文字,1-2 句,≤60 字 |
| badge | | 红星爆内文字,短!如 `97%` |
| ticker | | 底部黑条文字,MONO 字体,如 `RECTIFICATION // INFOSEC // Q3-2026` |

## toc · 目录(可点击)

构图:黑色页眉牌 + 红方块 + 2×2 硬投影大卡(序号/名称/描述/accent 条)。**1-4 个条目**。
交互开启时,第 i 个条目自动链接到目录页之后的第 i 页。

| 字段 | 必填 | 说明 |
|---|---|---|
| **items** | ✅ | 数组,1-4 项:`{"num": "01", "name": "章节名", "desc": "一句话摘要"}` |
| title | | 页眉文字,默认 `目录 CONTENTS` |

## stats · 数据亮点页

构图:黄色页眉牌 + N 张硬投影数据卡(accent 高亮条 + 巨型等宽数字 + 黑牌标签)+ 可选 ticker 黑条。**2-4 张卡**。

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 页眉牌文字,≤8 字最佳 |
| **cards** | ✅ | 数组,2-4 项:`{"num": "38", "unit": "项", "label": "高危漏洞清零", "sub": "中危修复率 91%"}` |
| ticker | | 底部黑条文字 |

数字排版建议:`num` 为纯数字或短码(如 `Q4`),`unit` 为单位汉字。

## compare · 今昔对比页

构图:黑卡(旧值+红叉)→ 大箭头 → 黄卡(新值)+ 粉圆章 + 底部注释。

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 页眉牌文字 |
| **before** | ✅ | `{"num": "62%", "label": "上季度整改完成率", "tag": "BEFORE"}`(tag 可省) |
| **after** | ✅ | `{"num": "97%", "label": "本季度整改完成率", "tag": "NOW"}`(tag 可省) |
| stamp | | 粉圆章文字,≤4 字,如 `复测\n通过` |
| note | | 底部居中注释,1 句 |

## items · 条目横带页

构图:黄色通栏页眉 + N 条硬投影横带(黑序号块 + 标题 + 描述 + 右端 accent 色柱)。**1-4 条**,自动等高分布。

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 通栏页眉文字,如 `NEXT · 下一步计划` |
| **items** | ✅ | 数组,1-4 项:`{"num": "01", "title": "条目名", "desc": "一句说明"}` |
| ticker | | 底部黑条文字 |

## thanks · 结尾页

构图:纯黑页 + 84pt 酸性黄 MONO 大字 + 白色致谢行 + 黄条 + 红星爆 badge + 青色"← 目录"返回钮。

| 字段 | 必填 | 说明 |
|---|---|---|
| big | | 大字,默认 `THANKS` |
| line | | 致谢行,默认 `谢谢观看 · 敬请指正` |
| badge | | 星爆内短文字,如 `Q3` |
| footnote | | 左下 MONO 小字,如 `INFOSEC // Q3-2026 // DF` |

---

## 排页节奏建议

- 标准 6 页:`cover → toc → stats → compare → items → thanks`
- 8-12 页:中间按内容插入多个 stats / items / compare,**同一原型不连排两页**(节奏:重→轻交替)。
- 目录条目数 = 目录之后内容页数(≤4);超过 4 章时合并章节,不加第二页目录。

## chart · 图表页

构图:页眉 + 原生图表(柱/饼/环)+ 大数字 + 结论。**图表是原生对象,双击可改数据。**

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 结论式标题 |
| **chart** | ✅ | `{"kind":"bar|pie|doughnut", "categories":[...], "series":"名", "values":[...], "highlight_idx":1, "number_format":"$0,,\"M\"", "legend":true}` |
| big_num | | 右侧大数字,如 `USD 27.8M` |
| big_label | | 大数字说明 |
| insights | | 右侧结论(≤120 排版单位) |
| source | | 数据来源(左下角) |

## photo_gallery · 照片墙

构图:2-6 张照片**完整显示不裁切**(contain + 底衬),每张带说明。产品/证书/场景图用。

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 标题 |
| **photos** | ✅ | 数组,1-6 项:`{"image":"/abs/path.png", "caption":"说明"}` |

## logo_wall · logo 墙

构图:logo 网格,**自动换行防溢出**。客户/伙伴/认证 logo 用。

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 标题 |
| **logos** | ✅ | logo 图片绝对路径数组 |
| per_row | | 每行几个,默认 5 |
| lead | | 一句加粗导语 |
| body | | 补充说明 |

> 透明底(RGBA)logo 需先 `alpha_composite` 合成白底,否则变黑。

## org_chart · 组织架构图

构图:顶层节点 + 中层节点(带连接线)+ 下层子节点。

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 标题 |
| **top** | ✅ | `{"label":"顶层"}` |
| **middle** | ✅ | 数组:`{"label":"中层", "sub":"副标"}` |
| sub | | 下层标签数组(连到 `sub_parent` 指定的中层,默认 1) |
| body | | 底部说明 |

## scene · 场景大图页

构图:左/右全高场景大图 + 文字区(结论标题+正文+强调句)。颜色方案/产品理念用。

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 大标题(可 `\n` 换行) |
| **image** | ✅ | 场景图绝对路径 |
| side | | `right`(默认)/ `left` 图的位置 |
| body | | 正文 |
| lead | | 底部强调句(accent2 色) |

## timeline · 时间线

构图:横轴 + 2-6 个里程碑节点(日期/标题/说明)。

| 字段 | 必填 | 说明 |
|---|---|---|
| **title** | ✅ | 标题 |
| **items** | ✅ | 数组,2-6 项:`{"date":"2011", "title":"里程碑", "text":"说明"}` |
| body | | 底部说明 |


## v3 艺术版式(字体配对 / 艺术字 / 艺术排版)

### 字体系统(所有版式共用)
- **中西文分离配对**:`FONT_PAIRS` 定义 heading / body / display / quote 四个角色的 (西文, 中文) 字体对,PowerPoint/WPS 按字符自动选用。
- 默认配对:标题 `Arial Black / 微软雅黑`、正文 `Arial / 微软雅黑`、引文 `Georgia / 楷体`。
- 正式文气场景:heading 换 `思源宋体 / Georgia`(对方没装会退回,交付时提醒)。

### 艺术字变换(spec 支持字段)
| 字段 | 效果 |
|---|---|
| `stroke` / `stroke_w` | 描边字(需配合 cover_type / free) |
| `gradient: [c1, c2]` | 渐变字 |
| `italic` | 斜体(引文默认) |
| `tracking` | 字距(千分比 pt) |
| `vert: true` | 竖排(free 版) |
| `font: heading/body/display/quote` | 字体角色(free 版) |

### cover_type · 大字封面
标题即画面。字段:**title**(必填,`\n` 换行)、kicker、size(默认60)、stroke、stroke_w、gradient、spacing、subtitle、body、ticker、note。

### big_quote · 金句页
巨型引号 + 衬线斜体引文。字段:**quote**、by、size(默认30)、note_text、note。

### section · 章节页
超大幽灵章节号 + 章节题。字段:**title**、number(默认01)、kicker、subtitle、note。

### shot_stack · 编号图组
2-4 张图等宽卡片 + 编号 + 说明。产品系列/步骤/对比用。字段:**title**、**items**:`[{image,title,text}]`(2-4)、body、note。

### overlap · 错位叠卡
2-4 张卡片层叠错位微旋转。字段:**title**、**cards**:`[{title,big,text}]`(2-4)、rot、note。

### free · 自由版
手工艺术排版,elements 元素清单(text/rect/image/line)。text 支持全部艺术字段;rect 支持 fill/line/rot/radius/shape(rect|round|ellipse);image 支持 fit(cover|contain)。字段:**elements**、note。

## v3.2 杂志编辑风(editorial 风格专用版式)

> 审美纲领(从业界公开风格资料与多轮实战评审中提炼):
> **高级感 = 字体对撞(衬线×无衬线×mono) + 结构即装饰(发丝线/分栏/报头) + 反色焦点格 + 克制用色(颜色只做结构不做装饰)**。
> 中式豪华风(大红大蓝+Pill 指标+箭头轰炸)与 AI 渐变风(柔和圆角卡片)均已过时,不用。

### ed_cover · 杂志封面
报头信息行(双发丝线)+ 超大衬线标题 + 斜体导语 + 底部元信息三栏(竖线分隔)。
字段:**title**、kicker、issue、size(默认54)、standfirst、meta:`[{label,value}]`(≤3)、note。

### ed_split · 不对称分栏章页
超大首字数字(左)+ 竖发丝线 + 衬线标题 + 红条引文 + 正文(右,宽栏)。
字段:**title**、**numeral**、numeral_label、quote、body、kicker、note。

### ed_report · 报纸报道页
报头黑条(衬线标题+日期)+ 反色焦点格(大数字)+ 多栏正文(栏头+发丝线)+ 脚注。
字段:**title**、kicker、date、focus_big、focus_label、focus_text、columns:`[{head,body}]`(≤3)、footnote、note。

## v3.4 杂志风精修(用户认证后的挑剔式升级)

### ed_figures · 报纸行情式数据行(替代卡片式 stats)
双线报界 + 发丝线分隔的大数字 + 标签 + 副文 + 底部署名。**杂志风不要用卡片装数字。**
字段:**title**、**figures**:`[{num,label,sub}]`(2-4,数字宽度自适应缩号防撞列)、kicker、body、source、note。

### ed_cover v2(封面精修)
- 左缘**书脊竖排**(spine 字段):`"spine": "MEGA CASA · STAINLESS STEEL GROUP · JIANGMEN"`
- 大标题**横穿发丝线**(cross_at 控制线位)
- 导语**首字下沉**(drop cap)

### ed_split v2(章节页精修)
- 超大数字横穿规则线
- 正文**首字下沉**(body 首字符自动放大成 accent 色装饰字)
- 引文放大为 pull-quote 式

> 首字下沉是 editorial 的标志性武器:正文首字符放大、accent 色、右侧接排。

## v3.5 数据页与报道页再精修

### ed_figures v2:头条数字制
figures[0] 为**头条数字**(60pt 左栏),其余为配角(34pt 右列,发丝线分隔)。
报头加红色小方块点题。数字按 em 宽自适应缩号(大写0.72/数字0.56/小写0.5/空格0.28/CJK1.0)。

### ed_report v2:报名式报头 + 不等宽分栏
- 报名:黑条 + **双规则线**(粗+细)压在条下
- 分栏**不等宽**:首栏 1.5 倍宽且带**首字下沉**,其余等宽
- 焦点数字放大至 64pt(自适应)

## v4.1 table · 原生表格

原生可编辑表格(双击改数据)。字段:**title**、**columns**、**rows**、highlight_col(高亮列序号)、note_row、note。
三风格表格语言:brutal=粗黑框+黄表头 / dark_premium=细线+金表头 / editorial=横发丝线+墨表头(无竖线,报纸式)。
单元格文字≥10.5pt;列≤6、行≤8 装得从容。
