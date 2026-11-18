---
name: sdu-ppt-hybrid
description: 制作可编辑的山东大学 PowerPoint。默认以官方第 1 套原始演示文稿和原生母版为底稿，沿用其校园封面、正文红色页眉、居中校徽和山形页脚，结合 GitHub 版字体、网格和内容排版。也可直接用 GitHub 原版。适用于组会、开题、答辩和研究汇报。
---

# 山东大学 PPT 融合版

按用户指定选择视觉模板。两种选择共用 GitHub 版页型和可编辑字段，同一内容计划可以直接切换。

| 选项 | 视觉素材 | 排版 |
| --- | --- | --- |
| `official-first`（默认） | 从原始 `山东大学PPT模版1.pptx` 构建，继承其原生母版。封面、目录、章节页、结束页沿用对应官方视觉，常规正文沿用模板 1 红色页眉和居中校徽 | GitHub 字体、文字层级与内容编排，正文避开官方页眉 |
| `github` | 用户提供的 GitHub 仓库原模板 | 全套 GitHub 版式 |

用户明确提供专用院系母版时，按要求沿用。使用内置模板时，标题、数字、标签用思源宋体（Noto Serif SC），正文用思源黑体（Noto Sans SC）。复用模板页原有的字号、对齐、留白与强调色。需要新增页时按 [GitHub 排版规范](references/design-system.md) 套用现有母版与网格。两种视觉来源及切换方式见 [融合模板说明](references/hybrid-template.md)。

## 选页与填充

- 封面用 `cover`、`mcover` 或 `cover2`。目录用 `toc`，按实际章节填写 `目录1` 至 `目录4`。未使用的目录条目和尾部分隔线一并清除。
- 章节页用 `section`，填章节序号、标题和副标题。结束页用 `close` 或 `close2`。
- 正文按 [场景页序](references/scenarios.md) 选页。用 [页型目录](references/catalog.md) 查询形状名，或运行 `catalog`。短汇报可省略目录或章节页。
- 官方第 1 套的图表和密集正文使用原版红色页眉。较疏的正文或收束页可选 `仅标题（模板1山形页脚）`；为右下角山形留空。自行放置图文时，从标题下方开始，给居中校徽和底部页码留出距离。
- 内容按听众、时长和素材取舍。保留源报告的单位、口径和证据强度。图表和表格用原生对象。只有报告图像时保留原图及出处，按原比例放置。解释和引用可放入演讲者备注。

## 汇报文案与信息密度

- 沿用材料中的技术叫法和常用表达。方法页标题直接写操作或对象。结果页写图表能支持的具体发现，带上相关样本、量值或条件。
- 先定每页要讲的内容，再选版式。用图、表或短段落承担主要信息。保留能帮助导航的转场，合并只有章名或重复前页结论的页。对称栏、三点卡片和底部结论框按内容选用，模板留出的格子可以空出或删掉。
- 通读一遍正文，删去口号、泛化的价值判断和重复图注。图旁写听众需要观察的峰、差异或口径，补充背景放在备注。证据限制和影响解读的条件留在对应结果页。

## 生成与核对

在技能目录运行以下示例。组装需要 Python 3.9+ 和 `python-pptx`。生成预览总览另需 `Pillow`。脚本和字段详见 [API](references/api.md)。

```sh
python scripts/sdu_ppt.py catalog
python scripts/sdu_ppt.py build examples/hybrid-demo.json demo.pptx --style official-first
python scripts/sdu_ppt.py build examples/hybrid-demo.json demo-github.pptx --style github
python scripts/render.py demo.pptx render --sheet
```

实际任务另建内容计划。需要时以同一计划试导两个版本。优先在 PowerPoint 中逐页预览，检查标题字形、正文换行、校徽、页码、图片裁切、图表标签和未填字段。交付可编辑的 .pptx，按需检查原生图表的内嵌数据与原生表格结构。
