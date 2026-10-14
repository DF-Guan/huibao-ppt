---
name: huibao-ppt
display_name: 汇报大师 PPT
display_name_en: Report Master PPT
description: >-
  汇报大师 PPT —— 专为工作汇报/整改验收/述职总结/公司介绍场景打造的演示文稿生成技能。
  三风格:「新野兽派」(粗黑描边+硬投影+酸性黄)、「深色高级风」(藏青底+黄铜金)、
  「杂志编辑风」(衬线大标题+不对称分栏+反色焦点格,最高审美档)。原生可编辑 PPTX,
  23 种版式覆盖实战(图表/照片墙/logo墙/组织架构/时间线)与艺术(大字封面/金句页/
  章节页/编号图组/错位叠卡/自由版/杂志三件套);中西文字体分离配对、描边/渐变/竖排/
  斜体衬线等艺术字变换;数字等高线字体自动防护;渲染前 preflight 预检 + 渲染后
  LINT/错位监测 + PowerPoint 导出逐页验收三道闸;自带可点击目录跳页、翻页动效与
  每页口播备注。当用户要"做汇报PPT/整改汇报/述职/季度总结/公司介绍/生成可编辑PPTX",
  或对已有 PPT 不满意要求优化时使用。Keywords: PPT / presentation / slides /
  deck / report / 汇报 / 述职 / 整改 / 总结 / business review / company profile。
category: productivity-tools
version: 1.0.0
author: DF
agent_created: true
when_to_use: 用户要制作或优化工作汇报/整改验收/述职总结/公司介绍等演示文稿时;
  三风格(brutal/dark_premium/editorial)×23版式,心法引导与三道质量闸
description_zh: 三风格×23版式的原生可编辑PPTX生成器:审美心法引导+三道质量闸,用户一句话就能得到好看、耐看、能讲的汇报。
description_en: Native editable PPTX generator with 3 styles and 23 layouts, aesthetic guidance and three quality gates — one sentence in, a presentable deck out.
examples_zh:
  - 帮我做一份Q3网络安全整改汇报PPT,给管理层看
  - 把这份公司介绍优化一下,要面向澳洲客户,高端大气
  - 我想做述职PPT,但不知道选什么风格,先给我看看效果
examples_en:
  - Make me a Q3 security rectification report PPT for management
  - Optimise this company profile for Australian clients, premium feel
  - I want a performance review PPT but can't pick a style, show me samples first
---

# 汇报大师 PPT

把汇报内容变成一份**零翻车、有性格、可交互**的原生可编辑 PPTX。

> **先读心法**:`references/thinking.md` —— 审美五维诊断、引导术(样片选美)、诊断表(症状→根因→处方)、自检八问、迭代法。规则防灾难,心法出好作品。

本 skill 已内置全部实战护栏与思考过程——**按流程走即可,不需要客户二次调教**。

## 快速上手(三步)

```
1. python scripts/preflight.py <你的spec.json>     # 预检:零 E 才放行
2. python scripts/build_deck.py <spec.json> <输出.pptx> [--style brutal|dark_premium|editorial]
3. 用 PowerPoint 导出 PNG 逐页肉眼验收(见下方命令)
```

spec 怎么写?复制 `examples/deck-spec.example.json` 改内容即可;版式字段契约查
`references/page-archetypes.md`;杂志风标杆看 `examples/style-showcase.json`。

**用户说不上来要什么风格?** 一条命令出多风格样片册,让用户翻页指认:

    python scripts/sample.py --title "主题" --lead "导语" --stats "38项,高危清零;214个,权限回收"

## 运行环境

- 首选托管 venv:`C:/Users/ONDA/.workbuddy/binaries/python/envs/default/Scripts/python.exe`(已装 python-pptx、Pillow)。
- 若不可用:任意 Python 3.10+ 建隔离 venv 并 `pip install python-pptx pillow`。**禁止**全局 pip。
- 脚本一律绝对路径调用,不依赖 CWD。

## 工作流(四阶段,心法驱动)

### 阶段一 · 引导(用户只给简单提示时,按 references/thinking.md 执行)
```
- [ ] 1. 摸底(合并一轮问完):受众 / 场合 / 有没有现成材料
- [ ] 2. 定风格:风格三选一附证据;面向客户先看客户品牌审美;
        用户说不上来 → 样片选美法(同一内容 2-3 页小样让用户指认)
- [ ] 3. 复述确认:"X 风格 + Y 版式序列 + Z 交互,做 N 页,对吧?"——确认再跑
```

### 阶段二 · 设计
```
- [ ] 4. 备内容:逐页写"结论式标题" + 真实内容 + 口播备注(note);
        素材逐张核实(铁律 2);改造旧稿先逐页读原稿(铁律 3)
- [ ] 5. 写 spec(查 references/page-archetypes.md)
- [ ] 6. preflight 预检:python scripts/preflight.py <spec> —— E 全清再渲染
```

### 阶段三 · 渲染
```
- [ ] 7. python scripts/build_deck.py <spec> <输出.pptx> [--style 风格]
        (LINT / 错位监测 / 容量 三道自动闸)
```

### 阶段四 · 自检与交付
```
- [ ] 8. 自检八问(thinking.md §3):主角/结论标题/构图差异/呼吸/整齐/照片完整/
        受众适不适/备注能讲——任一为否就修
- [ ] 9. python scripts/export_preview.py <成品.pptx> <预览目录> 一键导出 PNG,逐页肉眼复核
- [ ] 10. 交付:present_files 呈现,说明交互玩法与备注位置
```

### 用户否决时(禁问"哪里不满意")
按 thinking.md §4 把负面词映射到层级(风格/构图/内容/执行/策略),
用选择题确认(A 还是 B),不让用户写小作文。

## 选风格

| 风格 | 气质 | 适合 |
|---|---|---|
| `brutal` 新野兽派 | 粗黑描边、硬投影、酸性黄、张扬 | 内部汇报、年轻团队、活泼场合 |
| `dark_premium` 深色高级风 | 藏青底、黄铜金、克制高级 | 面向客户/海外、高端产品、正式提案 |
| `editorial` 杂志编辑风 | 衬线大标题、不对称分栏、反色焦点格、报头黑条 | 品牌提案、年度回顾、研究报告——最高审美档 |

风格令牌在 `styles/<name>.json`,同一份 spec 加 `--style` 整体换肤。

## 全局铁律(每条都是实战翻车换来的,禁止违反)

0. **心法先行**:动手前过一遍 thinking.md 的五维诊断与引导术;规则解决"不出错",心法解决"好看"。
1. **数据零编造**:只用用户提供的真实数据/事实。内容薄就深挖结构做厚,绝不虚构。
2. **素材必核实**:引用图片/照片/logo 前必须逐张确认内容,**禁止凭文件名猜测**;
   批量素材先做带文件名标签的编目图逐张过目;拿不准的素材不用。
3. **改造旧稿先读原稿**:用户给旧 PPT 要优化时,先逐页导出通读,再动手;
   禁止只看素材目录就重构。
4. **照片不裁切**:产品/证书/人物/场景图一律 contain 完整显示(引擎已收紧 pad);
   cover 出血只允许用于全屏背景大图。
5. **数字必用等高线字体**:Georgia/Times 是旧式数字(高低不齐),含数字文本一律
   Arial/Consolas/微软雅黑;text_art 已自动防护,不要绕过。
6. **结论式标题**:每页标题是一句能站得住的结论,不是"概述/介绍"这类空词。
7. **口播备注**:每页写 note(演讲者提示),渲染后自动进备注栏;section/big_quote/ed_cover 可免。
8. **风格要有证据**:面向特定客户群选风格,要先研究对方的品牌审美(官网/画册/门店),不凭自我想象拍颜色。
9. **本地化**:面向海外的稿子用目标市场拼写(澳式 colour/customise/standardised),
   中文稿才用中式表达。
10. **三道闸不过不交付**:preflight(渲染前)→ LINT+错位监测(渲染中)→ COM 导出逐页目检(渲染后)。
11. **相邻不同版式**:同版式不连排(ed_split/ed_report 除外);长稿每 4-6 页插
    section/big_quote 换气;6 页以上加 toc 并开 interactive。
12. **文字进框**:lint/capacity 报警就改文案/拆页/换版式,不得硬塞。
13. **错误要暴露**:引擎异常一律向上抛,禁止静默吞错后交付残缺品。

## 高保真验收(交付前必做)

    python scripts/export_preview.py <成品.pptx> <预览目录>

自动优先 PowerPoint COM,无 Office 兜底 LibreOffice。导出后逐页核对:
无溢出/无错位/素材正确/数字整齐/链接可点。

## 长稿(12 页以上)章节导航成法

    封面 → toc(≤4 章) → section(章节页) → 内容页×N → section → ... → thanks

- 目录条目链接到各章节 section 页;内容页右上角自动有"目录"返回钮
- 每章内部按"重(数据/图表)→ 轻(条目/表格)→ 重(对比)"排节奏,章末插 big_quote 换气
- section 页:brutal/dark_premium 用 `section` 版式,editorial 用 `ed_split`

## 自检必查清单

- [ ] preflight 零 E;LINT 与错位监测零警告
- [ ] 每页标题是结论句,每页有口播备注
- [ ] 每个数字/事实/图片都能在用户材料中找到出处(已逐张核实)
- [ ] 照片 contain 完整、无裁切;logo 无黑底(引擎已自动合成白底)
- [ ] 海外稿用目标市场拼写;相邻两页版式不同;交互链接与动效已写入

## 扩展指引(新增风格/版式时)

- 新风格 = `styles/<name>.json`(令牌)+ `FONT_PAIRS` 加一条,不改渲染器。
- 新版式 = `build_deck.py` 加 `r_<名>` 注册进 RENDERERS + `capacity.py` 的 CAPACITY 标容量 +
  `preflight.py` 的 LAYOUTS_REQUIRING 标必填 + page-archetypes.md 补契约。
- 一切改动后必须重跑:示例 spec 双风格渲染 + editorial 演示稿渲染,零警告才算完。

## 参考文档

- **审美心法与引导术:`references/thinking.md`(先读)**
- 版式契约与容量:`references/page-archetypes.md`
- 可改写示例:`examples/deck-spec.example.json`
- 审美纲领(杂志风)与艺术字变换:见 page-archetypes.md 的 v3/v3.2 章节
