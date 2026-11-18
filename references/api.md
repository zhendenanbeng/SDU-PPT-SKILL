# sdu_ppt 接口说明

`scripts/sdu_ppt.py` 只依赖 `python-pptx` 和 `lxml`（随 python-pptx 安装）。
页面用 key 指定（见 `catalog.md`），形状用名字指定。

## 目录
1. Deck
2. SlideRef：文字
3. SlideRef：结构、图片、图表、表格
4. 进度类页面专用方法
5. 自定义页面组件
6. JSON 计划（build 命令）
7. 命令行与渲染

## 1. Deck

```python
import sys; sys.path.insert(0, '<skill>/scripts')
from sdu_ppt import Deck
deck = Deck(template=None)     # 默认 <skill>/assets/template.pptx
s = deck.add('weekly')         # 复制模板页（key 或模板中的 1 起页码），追加到末尾，返回 SlideRef
s = deck.add_blank('仅标题', title='本课题拟解决的问题')  # 从母版版式新建空白页（自动带 Logo、页码）
deck.section('第一章')          # 之后添加的页面归入此节（PowerPoint「节」）
deck.save('out/deck.pptx')     # 删除全部模板原页，只保留添加的页面；每个 Deck 只能保存一次
```

可用版式：封面、目录、章节页、标题和内容、仅标题、两栏内容、比较、图片与文字、大图、引用、空白、结束页、图表解读。

## 2. SlideRef：文字

| 方法 | 说明 |
|---|---|
| `set(name, value)` | 替换整个形状的文字并保留格式。`value` 为字符串（`\n` 分段）或段落列表；段落可以是字符串、run 文本列表，或 `{'text': ..., 'like': k}`（套用模板第 k 段的样式，-1 为最后一段）。默认第 i 段套用模板第 i 段样式，超出时套用最后一段。标题占位符会自动 `fit()` |
| `runs(name, texts, para=0)` | 按顺序替换某一段里每个 run 的文字，适合「标签 + 值」这类混排；多出的值会复制最后一个 run 的格式 |
| `fit(name, max_lines=1, min_size=12)` | 按估算宽度缩小字号，使最长一段不超过 `max_lines` 行（PowerPoint 不会自动缩小代码写入的文字） |
| `style(name, size=None, bold=None, spacing=None, color=None)` | 统一调整该形状所有 run 的字号（pt）、粗细、字距（pt）、颜色 |
| `color(name, color, run=None)` | 改文字颜色：`'accent1'`、`'tx1'`、`'bg1'`、`'RRGGBB'`，或 `('tx1', lumMod, lumOff)` |
| `notes(text)` | 写演讲者备注 |
| `names()` | 返回本页全部顶层形状的 `(名字, 类型, 文字)` |

示例：

```python
s = deck.add('result')
s.set('内容占位符', [
    {'text': '实验设置', 'like': 0},              # 小标题样式
    {'text': '数据集：……；指标：……', 'like': 1},     # 正文样式
    {'text': '结果观察', 'like': 0},
    {'text': '收敛速度提升约 30%', 'like': 3},       # 项目符号样式
])
s = deck.add('figread')
s.set('解读要点', ['解读要点', ['整体趋势：', '……'], ['关键差异：', '……']])  # 粗体前缀 + 正文
```

## 3. SlideRef：结构、图片、图表、表格

| 方法 | 说明 |
|---|---|
| `remove(*names, missing_ok=False)` | 删除形状；删除 `X` 时连带删除所有 `X-…` 附属形状（序号、说明、标记、色块等） |
| `respace(prefix, left=2.0, right=31.87)` | 删掉部分条目后，把剩下的 `prefix1, prefix2…` 横向整体居中并保持原间距（如 `remove('成员4')` 后 `respace('成员')`，三栏卡片删一栏后 `respace('要点')`） |
| `geometry(name, x=None, y=None, w=None, h=None)` | 移动或缩放形状（cm，相对页面左上角） |
| `image(name, path, focus=(0.5, 0.5))` | 放入图片，按框的比例裁切填满。支持普通图片（保留圆形等形状）、图片占位符和灰色「图片N」框；`focus` 控制保留图片的哪一部分 |
| `chart(name, categories, series, title=None, number_format=None)` | 替换柱状图、折线图、环形图的数据（`series` 为 `{系列名: [数值]}`），并按主题色重新上色 |
| `table(name, rows, header=None, highlight_last=None)` | 替换表格数据行（列数必须一致）。保留斑马纹和最后一行高亮（自动识别）；下方形状随表格高度移动 |

## 4. 进度类页面专用方法

```python
s = deck.add('progress')
s.progress([('文献调研', 100, 'done'), ('模型构建', 70, 'doing'), ('论文撰写', 0, 'todo')], overall=None)
# 最多 6 项；状态取 done / doing / todo；overall 省略时取平均值；圆环、统计数字、状态胶囊会一起更新

s = deck.add('gantt')
s.gantt([('文献调研', (0, 2), (0, 2.5)),       # (名称, (计划开始, 计划结束), (实际开始, 实际结束) 或 None)
         ('模型构建', (3.5, 7), (4, 6.8)),
         ('论文撰写', (8, 11), None)],
        start_month=9, year_label='2025—26', now=6.8, now_label='3 月下旬')
# 时间单位为「月」，0 表示第一列的起点；共 12 列，最多 7 项，少于 7 项时网格和图例自动上移

s = deck.add('timeline')
s.set('阶段1', '文献调研'); s.set('阶段1-时间', '2025.10'); s.set('阶段1-说明', ['第一行', '第二行'])
s.timeline_done(2)          # 前 2 个节点显示为已完成，红色进度线延伸到第 2 个节点
```

## 5. 自定义页面组件（仅在模板没有合适页型时使用）

```python
s = deck.add_blank('仅标题', title='标题')
s.add_text(text, x, y, w, h, size=18, bold=False, color='tx1', font='sans', align='l', anchor='t',
           line_spacing=1.5, name='自定义文本', bullets=False)   # font: 'sans' 正文 / 'serif' 标题字体
s.add_box(x, y, w, h, fill='bg2', line=None, shape='rect', name='自定义色块')  # shape: rect/roundRect/ellipse
s.add_takeaway(label, text, y, h=2.4, x=2.0, w=29.87)             # 模板样式的结论框
```

坐标单位为 cm，请按 `design-system.md` 的栅格取值（内容区 x 2.0 – 31.87，y 3.6 – 17.2）。

## 6. JSON 计划（build 命令）

```json
{
  "slides": [
    {"section": "第一部分"},
    {"page": "cover",
     "set": {"标题": "第 12 周组会汇报", "副标题": "XX 课题组"},
     "runs": {"汇报信息": ["汇报人：", "张三", "　　|　　", "时间：", "2026年3月10日"]}},
    {"page": "bar",
     "chart": {"图表": {"categories": ["A", "B"], "series": {"基线": [81, 76], "本文": [88, 83]}, "title": "准确率（%）"}},
     "remove": ["要点3"], "notes": "讲稿要点……"},
    {"page": "table", "table": {"表格": {"rows": [["方法A", "86.4", "82.1", "0.842", "1.35"]]}}},
    {"page": "pictext", "image": {"图片占位符": "media/photo.jpg"}},
    {"page": "close", "set": {"结束语副标题": "欢迎批评指正"}, "style": {"结束语副标题": {"spacing": 2}}},
    {"layout": "仅标题", "title": "自定义页面"},
    {"page": "progress", "progress": {"tasks": [["文献调研", 100, "done"], ["论文撰写", 0, "todo"]]}},
    {"page": "gantt", "gantt": {"tasks": [["调研", [0, 2], [0, 2.5]], ["写作", [8, 11], null]], "start_month": 9, "now": 6.5}},
    {"page": "timeline", "timeline_done": 2},
    {"page": "team", "remove": ["成员4"], "respace": ["成员"], "image": {"成员1-照片": "media/zhang.jpg"}}
  ]
}
```

每一项依次执行：`set` → `runs` → `style` → `chart` → `table` → `image` → `progress` / `gantt` / `timeline_done` → `remove` → `respace` → `notes`。
图片路径相对于 JSON 文件所在目录。

## 7. 命令行与渲染

```bash
python scripts/sdu_ppt.py catalog            # 列出模板页面、key、可编辑形状（--md 输出 Markdown）
python scripts/sdu_ppt.py inspect out.pptx   # 查看任意 pptx 的形状名和文字（--slides 2,5）
python scripts/sdu_ppt.py build plan.json out.pptx
python scripts/render.py out.pptx out/render --sheet          # 逐页 PNG + sheet.jpg 总览
python scripts/render.py out.pptx out/render --slides 3,5     # 只渲染改动的页
```

`render.py` 默认先用 PowerPoint（Windows，通过 PowerShell COM，效果与实际放映一致），不可用时改用 LibreOffice。

## 8. 切换视觉模板

两种选项共用相同的页型 key、形状名和字体。JSON 顶层可设置 `"style": "official-first"`（默认）或 `"style": "github"`。CLI 的 `--style` 可以覆盖 JSON。同一计划可以使用两种选项分别导出。指定用户模板时，`--template` 或 JSON `template` 优先。

```sh
python scripts/sdu_ppt.py build plan.json report.pptx --style github
python scripts/sdu_ppt.py catalog --style official-first
```

Python 可用 `Deck(style="official-first")` 或 `Deck(style="github")`。
