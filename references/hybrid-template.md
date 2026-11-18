# 模板 1 与 GitHub 排版的融合

`official-first` 从用户提供的《山东大学PPT模版1》原始 PPTX 构建，并在其原生母版内扩展汇报页型。模板包内保留 [用户上传的原件](../assets/template1-original.pptx)，也可对照 [山大文化网发布页](https://culture.sdu.edu.cn/info/1041/14743.htm)。[实际生成底稿](../assets/template.pptx) 包含 44 张可供装配的 GitHub 内容原型，已接入模板 1 的版式。原模板的示例正文在装配前已清除。

- 封面、目录、章节和结束页继承对应的模板 1 校园视觉。标题、副标题、目录与章节编号采用 GitHub 版层级和网格。
- 常规正文与图表页采用原模板 1 正文版式的红色页眉、居中校徽与两侧办学宗旨。为官方页眉预留垂直空间，图表页尽量扩大证据区。
- 文字较少的页可用 `仅标题（模板1山形页脚）`。山形来自原模板第 5 页，适合正文引入或收束页。图、表和图注密集时用 `仅标题` 等常规正文页。
- `github` 读取 [GitHub 原模板](../assets/template-github.pptx)。版式、生成工具和示例源自 SonnetoEnd/sdu-ppt-skill 提交 `aa40bcbf3890700d0e625f2bc963e7a48b03b8ae`。[GitHub 仓库](https://github.com/SonnetoEnd/sdu-ppt-skill)

两种模式的原型页 key 和可编辑字段相同，均采用思源宋体、思源黑体。常规内容页可以用同一 JSON 计划切换。使用 `add_blank` 添加自定义页时，模板 1 的标题位置更低，图文从约 4.1 cm 开始。换用 GitHub 原版后，可按其网格重新放置自定义对象。

JSON 顶层可填 `"style": "official-first"` 或 `"style": "github"`。命令行 `--style` 优先于 JSON。Python 可用 `Deck(style="official-first")`。明确指定 `--template` 或 JSON 中的 `template` 时采用指定 PPTX。

第一版融合技能里的 `official-cover`、`official-toc`、`official-section`、`official-close` 对应现在的 `cover`、`toc`、`section`、`close`。目录填 `目录1` 至 `目录4`。更多字段见 [页型目录](catalog.md)及[同计划示例](../examples/hybrid-demo.json)。
