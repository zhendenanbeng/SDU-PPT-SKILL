#!/usr/bin/env python3
"""sdu_ppt — build presentations from the Shandong University template (assets/template.pptx).

Every page in the template has a key (cover, weekly, gantt, ...) and every editable element has a
unique shape name (标题, 要点1, 要点1-说明, ...). You copy pages by key, then fill shapes by name.
The template carries the art style, so decks built this way look right without restyling.

Library
-------
    from sdu_ppt import Deck
    deck = Deck()                                    # opens ../assets/template.pptx
    s = deck.add('cover')                            # copy a template page (key or 1-based number)
    s.set('标题', '第 12 周组会汇报')                  # replace text, keep formatting
    s.runs('汇报信息', ['汇报人：', '张三', '　　|　　', '时间：', '2026年3月10日'])
    s = deck.add('bar')
    s.chart('图表', ['数据集1', '数据集2'], {'基线': [81.2, 76.5], '本文': [88.4, 83.0]})
    s.remove('要点3')                                 # removes 要点3 and every 要点3-* companion
    deck.save('out.pptx')

CLI
---
    python sdu_ppt.py catalog [--md]        # template pages, keys and editable shape names
    python sdu_ppt.py inspect deck.pptx     # shape names / text of any deck (QA, debugging)
    python sdu_ppt.py build plan.json out.pptx
"""
from __future__ import annotations

import argparse
import copy
import difflib
import json
import os
import re
import sys
import uuid

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor  # noqa: F401  (kept for callers that want explicit colours)
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.parts.chart import ChartPart
from pptx.util import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.normpath(os.path.join(HERE, '..', 'assets', 'template.pptx'))
DEFAULT_STYLE = 'official-first'
STYLE_TEMPLATES = {
    DEFAULT_STYLE: DEFAULT_TEMPLATE,
    'github': os.path.join(os.path.dirname(DEFAULT_TEMPLATE), 'template-github.pptx'),
}
CM = 360000
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
SECTION_EXT = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}'
R_ATTRS = [qn('r:embed'), qn('r:id'), qn('r:link'), qn('r:pict')]
SHAPE_TAGS = {'sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'}
ACCENTS = [MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2, MSO_THEME_COLOR.ACCENT_3,
           MSO_THEME_COLOR.ACCENT_4, MSO_THEME_COLOR.ACCENT_5, MSO_THEME_COLOR.ACCENT_6]

# Chinese name and filling notes for every template page (shown by `catalog`).
PAGES = {
    'cover': ('封面', '标题≤12字（过长自动缩小）；副标题一行；汇报信息用 runs：[标签, 值, 分隔, 标签, 值]。'),
    'toc': ('目录', '4 项，标题≤14字；少于 4 项用 remove("目录4")。'),
    'section': ('章节页', '章节序号如 "02"；标题≤12字；副标题可写英文或一句话。'),
    'motto': ('要点列表', '3 行「标签 + 内容」，用 runs("要点1", [标签, "　　", 内容])；标签 4 字，2 字标签用全角空格写成「校　　训」。'),
    'text': ('正文段落', '内容占位符，段落自动首行缩进；≤3 段、≤200 字。'),
    'twocol': ('两栏正文', '左栏内容/右栏内容，各 1–2 段。'),
    'three': ('三栏要点', '要点N=标题(≤6字)、要点N-说明(≤40字)、要点N-序号；中间卡片为强调色；只要两栏时 remove("要点3") 再 respace("要点")。'),
    'grid4': ('四宫格', '要点N=标题(≤8字)、要点N-说明(≤40字)；适合研究内容、四个方面。'),
    'compare': ('方案对比', '左/右侧标题 + 左/右侧内容（每侧 3–5 条，每条≤16字）+ 结论内容。'),
    'quote': ('引用金句', '引文≤20字，出处一行。'),
    'card': ('卡片+单图', '卡片正文 1–2 段≤150字；图片1 用 image() 替换。'),
    'card2': ('双栏卡片+竖图', '卡片左栏/卡片右栏各≤150字；图片1 为竖图。'),
    'img2': ('图文·双图', '正文≤150字；图片1、图片2 为竖图。'),
    'img3': ('图文·三图', '正文≤150字；图片1、图片2 小图，图片3 竖长图。'),
    'img4': ('图文·四图', '正文≤150字；图片1–4 四宫格。'),
    'pictext': ('图片与文字', '内容占位符≤150字；图片占位符 image() 替换（自动裁切填满）。'),
    'bigpic': ('大图展示', '左侧大图 + 右侧红色说明栏；标题≤8字，图片说明≤80字。'),
    'team': ('团队成员', '成员N=姓名，-职称，-简介，-照片(圆形，image() 替换)；人数少用 remove("成员4") 再 respace("成员") 居中。'),
    'stats': ('关键数据', '数据N 用 runs([数值, 单位])；数据N-名称≤8字；数据N-说明两行；底部结论+数据来源。'),
    'bar': ('柱状图', 'chart("图表", 类别, {系列: 数值})；要点N 与系列一一对应，系列少于 3 个时删掉多余要点。'),
    'line': ('折线图', '同柱状图；适合趋势、训练曲线。'),
    'donut': ('环形图', 'chart("图表", 类别, {"占比": 数值})；要点1–4 对应类别颜色。'),
    'table': ('数据表格', 'table("表格", rows, header=…)；最后一行默认按「本文方法」高亮；表注、结论随表格高度自动下移。'),
    'mcover': ('组会封面', '标题如「第 12 周组会汇报」；副标题写课题组与本周主题。'),
    'weekly': ('本周工作概览', '看板1–3 = 已完成/进行中/下周计划；事项i-k 每项≤12字（可带进度百分比）；删项 remove("事项2-2") 后同步改「看板2-数量」；底部本周小结。'),
    'result': ('实验结果与分析', '左侧内容占位符用 paras：小标题 like=0、正文 like=1、要点 like=3；右侧图片占位符放结果图。'),
    'issues': ('问题与讨论', '问题N(≤10字) + 问题N-说明(runs: [现象标签, 现象, ...] 两段)；右侧讨论N 为希望讨论的问题。'),
    'paper': ('论文信息', '论文题目(英文≤2行)、作者、信息1–6(期刊/时间/影响因子/分区/DOI/被引)、关键词1–4、推荐理由(结论内容)。'),
    'figread': ('图表解读', '图片占位符放论文原图；图题、图源；解读要点：第 1 段小标题，后续段落为带粗体前缀的要点(runs)。'),
    'review': ('文献评述', '评述1–3 = 主要贡献/不足与局限/对本课题的启发，每栏 3 条≤18字；底部总体评价。'),
    'litcmp': ('文献对比', '5 列表格：文献/研究对象/方法/结论/局限；最后一行「本研究」高亮；底部研究空白。'),
    'background': ('研究背景与意义', '背景正文 2 段；要点1–3 关键词卡片；意义1–2 = 理论/实践意义。'),
    'framework': ('研究框架', '层标签1–3（应用/方法/数据层）；模块r-k 第一段名称≤5字、第二段说明≤8字。'),
    'process': ('技术路线', '步骤1–4 runs([编号, "  名称"])；步骤N-说明 3 条；底部预期成果。'),
    'conclusion': ('结论与展望', '左侧内容自动编号 3 条，右侧内容 3 条（每条≤18字）；底部一句话总结。'),
    'progress': ('进度总览', '用 progress(tasks) 一次性设置任务、进度条、状态与圆环（≤6 项）。'),
    'gantt': ('甘特图', '用 gantt(tasks, start_month=…, now=…) 设置计划/实际条与当前时间线（≤7 项）。'),
    'timeline': ('研究进度安排', '阶段N(名称)、阶段N-时间、阶段N-说明(两段)；timeline_done(k) 设置已完成到第 k 个节点。'),
    'nextplan': ('下阶段计划', '计划1–3 = 近期/中期/远期要点（计划N-阶段 runs([阶段, "　时间范围"])）；底部需要支持。'),
    'cover2': ('答辩封面', '标题=论文题目；副标题=答辩类型；信息1–4 = 答辩人/导师/专业/日期（信息N-标签可改）。'),
    'close2': ('答辩结束页', '标题如「恳请各位老师批评指正」；副标题写答辩人与导师。'),
    'refs': ('参考文献', '内容占位符每段一条，格式 "[n]\\t条目"（制表符对齐悬挂缩进）；底部说明可删。'),
    'close': ('结束页', '标题「谢谢大家」；副标题是英文大字距样式，改成中文时加 style("结束语副标题", spacing=2)。'),
    'guide': ('使用说明（隐藏页，勿复制到成品）', '模板自身的说明页。'),
}
STATUS = {'done': '已完成', 'doing': '进行中', 'todo': '未开始'}
ANCHORS = {'t': 't', 'm': 'ctr', 'ctr': 'ctr', 'b': 'b'}


def _local(el):
    return etree.QName(el).localname


def _name(el):
    c = el.find('.//' + qn('p:cNvPr'))
    return c.get('name') if c is not None else None


def _iter_shapes(tree, deep=False):
    for el in tree.iterchildren():
        if _local(el) in SHAPE_TAGS:
            yield el
            if deep and _local(el) == 'grpSp':
                yield from _iter_shapes(el, True)


def _xfrm(el):
    """(off, ext) elements of a shape (sp/pic/cxnSp use spPr/xfrm, grpSp grpSpPr/xfrm, graphicFrame p:xfrm)."""
    for path in ('p:spPr/a:xfrm', 'p:grpSpPr/a:xfrm', 'p:xfrm'):
        x = el.find(path, {'p': NS_P, 'a': NS_A})
        if x is not None:
            return x.find(qn('a:off')), x.find(qn('a:ext'))
    return None, None


def _solid_xml(color, ns=True):
    """color: 'accent1' / 'bg1' / 'tx1' / ('tx1', lumMod, lumOff) / 'RRGGBB'."""
    decl = f' xmlns:a="{NS_A}"' if ns else ''
    if isinstance(color, tuple):
        c, lm, lo = color
        return (f'<a:solidFill{decl}><a:schemeClr val="{c}"><a:lumMod val="{lm}"/>'
                f'<a:lumOff val="{lo}"/></a:schemeClr></a:solidFill>')
    if re.fullmatch(r'[0-9A-Fa-f]{6}', color):
        return f'<a:solidFill{decl}><a:srgbClr val="{color.upper()}"/></a:solidFill>'
    return f'<a:solidFill{decl}><a:schemeClr val="{color}"/></a:solidFill>'


GREY = ('tx1', 75000, 25000)


def _set_fill(parent, color):
    """Replace the fill child (solidFill/noFill/gradFill) of spPr or rPr with a solid fill, keeping schema order."""
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill'):
        old = parent.find(qn(tag))
        if old is not None:
            old.addprevious(etree.fromstring(_solid_xml(color)))
            parent.remove(old)
            return
    new = etree.fromstring(_solid_xml(color))
    anchor = None
    for tag in ('a:prstGeom', 'a:custGeom', 'a:xfrm', 'a:ln'):
        anchor = parent.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None and _local(anchor) != 'ln':
        anchor.addnext(new)
    elif anchor is not None:
        anchor.addprevious(new)
    else:
        parent.insert(0, new)


def _set_line(spPr, color=None, width_pt=1.0):
    if color is None:
        xml = f'<a:ln xmlns:a="{NS_A}"><a:noFill/></a:ln>'
    else:
        xml = f'<a:ln xmlns:a="{NS_A}" w="{int(width_pt * 12700)}">{_solid_xml(color, ns=False)}</a:ln>'
    new = etree.fromstring(xml)
    old = spPr.find(qn('a:ln'))
    if old is not None:
        old.addprevious(new)
        spPr.remove(old)
        return
    for tag in ('a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'):
        nxt = spPr.find(qn(tag))
        if nxt is not None:
            nxt.addprevious(new)
            return
    spPr.append(new)


class SlideRef:
    """A slide being built. Every method returns self, so calls can be chained."""

    def __init__(self, deck, slide, key):
        self.deck, self.slide, self.key = deck, slide, key

    # ------------------------------------------------------------------ lookup
    @property
    def _tree(self):
        return self.slide.shapes._spTree

    def _find(self, name, required=True):
        top = [el for el in _iter_shapes(self._tree) if _name(el) == name]
        hits = top or [el for el in _iter_shapes(self._tree, deep=True) if _name(el) == name]
        if len(hits) == 1:
            return hits[0]
        if not hits:
            if not required:
                return None
            names = [_name(el) for el in _iter_shapes(self._tree, deep=True)]
            close = difflib.get_close_matches(name, names, n=6, cutoff=0.3)
            raise KeyError(f'[{self.key}] 没有名为「{name}」的形状。相近的名字：{close}。用 inspect 查看全部名字。')
        raise KeyError(f'[{self.key}] 「{name}」不唯一（{len(hits)} 个）。')

    def _obj(self, el):
        def walk(shapes):
            for s in shapes:
                if s._element is el:
                    return s
                if hasattr(s, 'shapes'):
                    r = walk(s.shapes)
                    if r is not None:
                        return r
            return None
        return walk(self.slide.shapes)

    def names(self):
        """[(name, kind, text)] for every top-level shape."""
        out = []
        for el in _iter_shapes(self._tree):
            kind = _local(el)
            ph = el.find('.//' + qn('p:ph'))
            if ph is not None:
                kind = f'placeholder:{ph.get("type", "obj")}'
            if kind == 'graphicFrame':
                kind = 'chart' if el.find('.//' + qn('c:chart')) is not None else 'table'
            text = ''.join(t.text or '' for t in el.iter(qn('a:t')))
            out.append((_name(el), kind, text))
        return out

    # ------------------------------------------------------------------ text
    @staticmethod
    def _txbody(el):
        tb = el.find(qn('p:txBody'))
        if tb is None:
            raise ValueError(f'「{_name(el)}」没有文本框（可能是图片、图表或组合）。')
        return tb

    @staticmethod
    def _fill_para(p, item):
        """item: str (one run) or list[str] (run texts, keeping each run's own formatting)."""
        for el in list(p):
            if _local(el) in ('fld', 'br'):
                p.remove(el)
        runs = p.findall(qn('a:r'))
        texts = [item] if isinstance(item, str) else list(item)
        if not runs:
            end = p.find(qn('a:endParaRPr'))
            r = etree.SubElement(p, qn('a:r'))
            rpr = copy.deepcopy(end) if end is not None else etree.Element(qn('a:rPr'))
            rpr.tag = qn('a:rPr')
            r.append(rpr)
            etree.SubElement(r, qn('a:t'))
            if end is not None:
                end.addprevious(r)
            runs = [r]
        while len(runs) < len(texts):
            clone = copy.deepcopy(runs[-1])
            runs[-1].addnext(clone)
            runs.append(clone)
        for r in runs[len(texts):]:
            p.remove(r)
        for r, t in zip(runs, texts):
            node = r.find(qn('a:t'))
            node.text = t
            if t != t.strip():
                node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    def set(self, name, value):
        """Replace the whole text of a shape.

        value: str ('\\n' starts a new paragraph) or a list of paragraphs, where each paragraph is
        a str, a list of run texts, or {'text': str | list, 'like': k} to copy the style of the
        template's k-th paragraph (default: the paragraph at the same position, or the last one).
        """
        tb = self._txbody(self._find(name))
        items = value.split('\n') if isinstance(value, str) else list(value)
        tpl = tb.findall(qn('a:p'))
        new = []
        for i, item in enumerate(items):
            like = min(i, len(tpl) - 1)
            if isinstance(item, dict):
                like = item.get('like', like)
                like = like if like >= 0 else len(tpl) + like
                item = item['text']
            p = copy.deepcopy(tpl[max(0, min(like, len(tpl) - 1))])
            self._fill_para(p, item)
            new.append(p)
        for p in tpl:
            tb.remove(p)
        for p in new:
            tb.append(p)
        ph = self._find(name).find('.//' + qn('p:ph'))
        if ph is not None and ph.get('type') in ('title', 'ctrTitle'):
            self.fit(name, max_lines=1)
        return self

    def _inherited(self, el, attr):
        """Run attribute (sz/spc/b) inherited from the layout placeholder or the master text styles."""
        ph = el.find('.//' + qn('p:ph'))
        if ph is None:
            return None
        typ, idx = ph.get('type', 'obj'), ph.get('idx')
        lay = self.slide.slide_layout
        path = './/' + qn('a:lstStyle') + '/' + qn('a:lvl1pPr') + '/' + qn('a:defRPr')
        for lph in lay.placeholders:
            lp = lph._element.find('.//' + qn('p:ph'))
            if (idx is not None and lp.get('idx') == idx) or (idx is None and lp.get('type', 'obj') == typ):
                d = lph._element.find(path)
                if d is not None and d.get(attr):
                    return d.get(attr)
        style = 'p:titleStyle' if typ in ('title', 'ctrTitle') else 'p:bodyStyle'
        d = lay.slide_master._element.find('.//' + qn(style) + '/' + qn('a:lvl1pPr') + '/' + qn('a:defRPr'))
        return d.get(attr) if d is not None else None

    def fit(self, name, max_lines=1, min_size=12):
        """Shrink the font so the longest paragraph fits in max_lines (estimate: CJK 1 em, Latin 0.55 em).
        Applied automatically to title placeholders by set(); PowerPoint does not autofit text written by code."""
        el = self._find(name)
        shape = self._obj(el)
        runs = list(el.iter(qn('a:r')))
        if not runs or shape is None:
            return self
        rpr0 = runs[0].find(qn('a:rPr'))
        size = int((rpr0 is not None and rpr0.get('sz')) or self._inherited(el, 'sz') or 1800) / 100
        spc = int((rpr0 is not None and rpr0.get('spc')) or self._inherited(el, 'spc') or 0) / 100
        avail = shape.width / 12700 * max_lines * 0.94
        worst_em, worst_n = 0.0, 0
        for p in el.iter(qn('a:p')):
            text = ''.join(t.text or '' for t in p.iter(qn('a:t')))
            em = sum(1.0 if ord(ch) >= 0x2E80 else 0.55 for ch in text)
            if em * size + spc * len(text) > worst_em * size + spc * worst_n:
                worst_em, worst_n = em, len(text)
        if worst_em and worst_em * size + spc * worst_n > avail:
            new = max(min_size, int((avail - spc * worst_n) / worst_em * 2) / 2)
            for r in runs:
                rpr = r.find(qn('a:rPr'))
                if rpr is None:
                    rpr = etree.SubElement(r, qn('a:rPr'))
                    r.insert(0, rpr)
                rpr.set('sz', str(int(new * 100)))
        return self

    def style(self, name, size=None, bold=None, spacing=None, color=None):
        """Adjust every run of a shape: size (pt), bold, spacing (letter spacing in pt), color."""
        for r in self._find(name).iter(qn('a:r')):
            rpr = r.find(qn('a:rPr'))
            if rpr is None:
                rpr = etree.SubElement(r, qn('a:rPr'))
                r.insert(0, rpr)
            if size is not None:
                rpr.set('sz', str(int(size * 100)))
            if bold is not None:
                rpr.set('b', '1' if bold else '0')
            if spacing is not None:
                rpr.set('spc', str(int(spacing * 100)))
            if color is not None:
                _set_fill(rpr, color)
        return self

    def runs(self, name, texts, para=0):
        """Set the run texts of one paragraph, keeping each run's formatting (label + value etc.)."""
        tb = self._txbody(self._find(name))
        self._fill_para(tb.findall(qn('a:p'))[para], list(texts))
        return self

    def color(self, name, color, run=None):
        """Recolour text runs of a shape: 'accent1', 'tx1', 'bg1', a hex string, or ('tx1', lumMod, lumOff)."""
        el = self._find(name)
        for i, r in enumerate(el.iter(qn('a:r'))):
            if run is None or i == run:
                rpr = r.find(qn('a:rPr'))
                if rpr is None:
                    rpr = etree.SubElement(r, qn('a:rPr'))
                    r.insert(0, rpr)
                _set_fill(rpr, color)
        return self

    # ------------------------------------------------------------------ structure
    def remove(self, *names, missing_ok=False):
        """Delete shapes. Removing 'X' also removes its companions named 'X-…' (说明, 序号, 标记 …)."""
        for name in names:
            hits = [el for el in _iter_shapes(self._tree) if _name(el) == name or (_name(el) or '').startswith(name + '-')]
            if not hits:
                if missing_ok:
                    continue
                self._find(name)  # raises with suggestions
            for el in hits:
                rids = {el2.get(a) for el2 in el.iter() for a in R_ATTRS if el2.get(a)}
                el.getparent().remove(el)
                for rid in rids:
                    self._drop_if_unused(rid)
        return self

    def _drop_if_unused(self, rid):
        """Remove a relationship only when nothing on the slide still points at it (r:embed, r:id, r:link)."""
        if any(el.get(a) == rid for el in self.slide._element.iter() for a in R_ATTRS):
            return
        try:
            self.slide.part.rels.pop(rid)
        except KeyError:
            pass

    def respace(self, prefix, left=2.0, right=31.87):
        """Re-centre the remaining items 「prefix1, prefix2, …」 horizontally after some were removed,
        keeping the original gap between items (e.g. respace('成员') after remove('成员4'))."""
        items = []
        i = 1
        while i <= 20:
            els = [el for el in _iter_shapes(self._tree)
                   if _name(el) == f'{prefix}{i}' or (_name(el) or '').startswith(f'{prefix}{i}-')]
            if els:
                boxes = [(int(o.get('x')), int(o.get('x')) + int(e.get('cx'))) for o, e in (_xfrm(el) for el in els) if o is not None]
                items.append((min(b[0] for b in boxes), max(b[1] for b in boxes), els))
            i += 1
        if not items:
            return self
        items.sort(key=lambda t: t[0])
        gap = items[1][0] - items[0][1] if len(items) > 1 else 0
        total = sum(r - l for l, r, _ in items) + gap * (len(items) - 1)
        x = int(left * CM + ((right - left) * CM - total) / 2)
        for l, r, els in items:
            for el in els:
                off, _ = _xfrm(el)
                off.set('x', str(int(off.get('x')) + x - l))
            x += (r - l) + gap
        return self

    def geometry(self, name, x=None, y=None, w=None, h=None):
        """Move / resize a shape (centimetres from the slide's top-left corner)."""
        off, ext = _xfrm(self._find(name))
        if x is not None:
            off.set('x', str(int(x * CM)))
        if y is not None:
            off.set('y', str(int(y * CM)))
        if w is not None:
            ext.set('cx', str(int(w * CM)))
        if h is not None:
            ext.set('cy', str(int(h * CM)))
        return self

    def _shift_below(self, y_emu, delta, skip):
        for el in _iter_shapes(self._tree):
            if el is skip:
                continue
            off, _ = _xfrm(el)
            if off is not None and int(off.get('y')) >= y_emu - CM * 0.05:
                if el.find('.//' + qn('p:ph')) is not None and _name(el) == '灯片编号':
                    continue
                off.set('y', str(int(off.get('y')) + delta))

    # ------------------------------------------------------------------ media
    @staticmethod
    def _cover(pic, img_size, focus=(0.5, 0.5)):
        iw, ih = img_size
        fa, ia = pic.width / pic.height, iw / ih
        pic.crop_left = pic.crop_right = pic.crop_top = pic.crop_bottom = 0
        if ia > fa:
            keep = fa / ia
            pic.crop_left = (1 - keep) * focus[0]
            pic.crop_right = 1 - keep - pic.crop_left
        else:
            keep = ia / fa
            pic.crop_top = (1 - keep) * focus[1]
            pic.crop_bottom = 1 - keep - pic.crop_top

    def image(self, name, path, focus=(0.5, 0.5)):
        """Put an image into a picture, picture placeholder or grey 图片N frame; crops to fill.
        focus=(fx, fy) chooses which part of the image survives the crop (0..1)."""
        el = self._find(name)
        tag = _local(el)
        shape = self._obj(el)
        if tag == 'pic':
            part, rid = self.slide.part.get_or_add_image_part(path)
            blip = el.find('.//' + qn('a:blip'))
            old = blip.get(qn('r:embed'))
            blip.set(qn('r:embed'), rid)
            if old and old != rid:
                self._drop_if_unused(old)
            self._cover(shape, part.image.size, focus)
            return self
        ph = el.find('.//' + qn('p:ph'))
        if tag == 'sp' and ph is not None and ph.get('type') == 'pic':
            pic = shape.insert_picture(path)
            pic._element.find('.//' + qn('p:cNvPr')).set('name', name)
            return self
        x, y, w, h = shape.left, shape.top, shape.width, shape.height
        pic = self.slide.shapes.add_picture(path, x, y, w, h)
        self._cover(pic, pic.image.size, focus)
        pic._element.find('.//' + qn('p:cNvPr')).set('name', name)
        el.addprevious(pic._element)
        el.getparent().remove(el)
        return self

    def chart(self, name, categories, series, title=None, number_format=None):
        """Replace a chart's data (bar / line / doughnut). series: {series name: [values]}.
        Colours are re-applied in theme order so added series stay on-brand."""
        gf = self._obj(self._find(name))
        cd = CategoryChartData(number_format=number_format) if number_format else CategoryChartData()
        cd.categories = list(categories)
        for sname, vals in series.items():
            cd.add_series(sname, list(vals))
        chart = gf.chart
        chart.replace_data(cd)
        for plot in chart.plots:
            kind = _local(plot._element)
            for i, ser in enumerate(plot.series):
                color = ACCENTS[i % len(ACCENTS)]
                if kind == 'barChart':
                    ser.format.fill.solid()
                    ser.format.fill.fore_color.theme_color = color
                elif kind == 'lineChart':
                    ser.format.line.color.theme_color = color
                    ser.marker.format.line.color.theme_color = color
                    ser.marker.format.fill.solid()
                    ser.marker.format.fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
                elif kind in ('doughnutChart', 'pieChart'):
                    for k in range(len(cd.categories)):
                        pt = ser.points[k]
                        pt.format.fill.solid()
                        pt.format.fill.fore_color.theme_color = ACCENTS[k % len(ACCENTS)]
                        pt.format.line.color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
                        pt.format.line.width = Pt(2.25)
        if title is not None and chart.has_title:
            p = chart.chart_title.text_frame.paragraphs[0]
            if p.runs:
                p.runs[0].text = title
                for r in p.runs[1:]:
                    r._r.getparent().remove(r._r)
        return self

    def table(self, name, rows, header=None, highlight_last=None):
        """Replace table rows (list of lists of str). The table keeps its zebra striping; the last
        row keeps the template's highlight style (e.g. 本文方法 / 本研究) unless highlight_last=False.
        Shapes below the table move with its new height."""
        gf = self._find(name)
        tbl = gf.find('.//' + qn('a:tbl'))
        trs = tbl.findall(qn('a:tr'))
        head, data = trs[0], trs[1:]
        ncol = len(head.findall(qn('a:tc')))
        if header is not None:
            for tc, t in zip(head.findall(qn('a:tc')), header):
                self._fill_para(tc.find(qn('a:txBody')).find(qn('a:p')), str(t))
        if highlight_last is None:
            def bold(tr):
                r = tr.find('.//' + qn('a:rPr'))
                return r is not None and r.get('b') == '1'
            highlight_last = len(data) >= 3 and bold(data[-1]) and not bold(data[0])
        plain = data[:-1] if highlight_last and len(data) > 1 else data
        old_h = sum(int(tr.get('h')) for tr in trs)
        for tr in data:
            tbl.remove(tr)
        for i, row in enumerate(rows):
            if len(row) != ncol:
                raise ValueError(f'表格「{name}」有 {ncol} 列，第 {i + 1} 行给了 {len(row)} 个值。')
            src = data[-1] if (highlight_last and i == len(rows) - 1) else plain[i % len(plain)]
            tr = copy.deepcopy(src)
            for tc, t in zip(tr.findall(qn('a:tc')), row):
                self._fill_para(tc.find(qn('a:txBody')).find(qn('a:p')), str(t))
            tbl.append(tr)
        new_h = sum(int(tr.get('h')) for tr in tbl.findall(qn('a:tr')))
        off, ext = _xfrm(gf)
        bottom = int(off.get('y')) + old_h
        ext.set('cy', str(new_h))
        if new_h != old_h:
            self._shift_below(bottom, new_h - old_h, gf)
        return self

    def notes(self, text):
        """Speaker notes (演讲者备注)."""
        self.slide.notes_slide.notes_text_frame.text = text
        return self

    # ------------------------------------------------------------------ page-specific helpers
    def progress(self, tasks, overall=None):
        """进度总览 page. tasks: [(name, percent 0-100, 'done' | 'doing' | 'todo'), ...] (≤ 6)."""
        if len(tasks) > 6:
            raise ValueError('进度总览最多 6 项；更多任务请拆页或改用甘特图。')
        template_bar = next((self._find(f'任务{k}-进度条', False) for k in range(1, 7) if self._find(f'任务{k}-进度条', False) is not None), None)
        for i, (name, pct, status) in enumerate(tasks, 1):
            self.set(f'任务{i}', name)
            track = self._find(f'任务{i}-进度槽')
            t_off, t_ext = _xfrm(track)
            bar = self._find(f'任务{i}-进度条', False)
            if pct > 0:
                if bar is None:
                    bar = copy.deepcopy(template_bar)
                    bar.find('.//' + qn('p:cNvPr')).set('name', f'任务{i}-进度条')
                    bar.find('.//' + qn('p:cNvPr')).set('id', str(self._next_id()))
                    track.addnext(bar)
                b_off, b_ext = _xfrm(bar)
                b_off.set('x', t_off.get('x'))
                b_off.set('y', t_off.get('y'))
                b_ext.set('cy', t_ext.get('cy'))
                b_ext.set('cx', str(max(int(int(t_ext.get('cx')) * min(pct, 100) / 100), int(0.3 * CM))))
            elif bar is not None:
                bar.getparent().remove(bar)
            self.runs(f'任务{i}-百分比', [f'{pct:g}%'])
            self.color(f'任务{i}-百分比', 'accent1' if pct > 0 else GREY)
            tag = self._find(f'任务{i}-状态')
            spPr = tag.find(qn('p:spPr'))
            if status == 'done':
                _set_fill(spPr, 'accent1')
                _set_line(spPr, None)
                txt = 'bg1'
            elif status == 'doing':
                _set_fill(spPr, 'bg1')
                _set_line(spPr, 'accent1', 1.0)
                txt = 'accent1'
            else:
                _set_fill(spPr, 'bg1')
                _set_line(spPr, ('tx1', 50000, 50000), 1.0)
                txt = GREY
            self.runs(f'任务{i}-状态', [STATUS[status]])
            self.color(f'任务{i}-状态', txt)
        for k in range(len(tasks) + 1, 7):
            self.remove(f'任务{k}', missing_ok=True)
        self.remove(f'任务{len(tasks)}-分隔', missing_ok=True)
        if overall is None:
            overall = round(sum(t[1] for t in tasks) / max(len(tasks), 1))
        ring = self._find('进度环')
        end = (270 + overall / 100 * 360) % 360 if overall < 100 else 269.99
        for gd in ring.iter(qn('a:gd')):
            if gd.get('name') == 'adj2':
                gd.set('fmla', f'val {int(end * 60000)}')
        if overall <= 0:
            self.remove('进度环')
        self.runs('总进度', [f'{overall:g}', '%'])
        for k, st in enumerate(('done', 'doing', 'todo'), 1):
            self.runs(f'统计{k}', [STATUS[st] + ' ', str(sum(1 for t in tasks if t[2] == st))])
        return self

    def gantt(self, tasks, start_month=1, year_label=None, now=None, now_label=None):
        """甘特图 page. tasks: [(name, (plan_start, plan_end), (actual_start, actual_end) | None), ...] (≤ 7).
        Times are in months from the left edge of the first column (0 = start of the first month).
        now: same unit, moves the dashed 当前 line."""
        if len(tasks) > 7:
            raise ValueError('甘特图最多 7 项。')
        g0 = _xfrm(self._find('网格0'))[0]
        g12 = _xfrm(self._find('网格12'))[0]
        x0 = int(g0.get('x')) + 4762
        mw = (int(g12.get('x')) - int(g0.get('x'))) / 12
        for m in range(1, 13):
            self.runs(f'月份{m}', [f'{(start_month - 1 + m - 1) % 12 + 1}月'])
        if year_label:
            self.runs('年份', [year_label])
        tpl_act = self._find('任务1-实际', False)
        n1 = _xfrm(self._find('任务1'))
        rh = int(_xfrm(self._find('任务2'))[0].get('y')) - int(n1[0].get('y'))
        for i, (name, plan, act) in enumerate(tasks, 1):
            self.set(f'任务{i}', name)
            p_el = self._find(f'任务{i}-计划')
            p_off, p_ext = _xfrm(p_el)
            p_off.set('x', str(int(x0 + plan[0] * mw)))
            p_ext.set('cx', str(max(int((plan[1] - plan[0]) * mw), int(0.3 * CM))))
            a_el = self._find(f'任务{i}-实际', False)
            if act:
                if a_el is None:
                    a_el = copy.deepcopy(tpl_act)
                    a_el.find('.//' + qn('p:cNvPr')).set('name', f'任务{i}-实际')
                    a_el.find('.//' + qn('p:cNvPr')).set('id', str(self._next_id()))
                    p_el.addnext(a_el)
                a_off, a_ext = _xfrm(a_el)
                a_off.set('x', str(int(x0 + act[0] * mw)))
                a_off.set('y', str(int(p_off.get('y')) + (int(p_ext.get('cy')) - int(a_ext.get('cy'))) // 2))
                a_ext.set('cx', str(max(int((act[1] - act[0]) * mw), int(0.3 * CM))))
            elif a_el is not None:
                a_el.getparent().remove(a_el)
        for k in range(len(tasks) + 1, 8):
            self.remove(f'任务{k}', missing_ok=True)
        shrink = (7 - len(tasks)) * rh
        if shrink:
            for m in range(13):
                _, ext = _xfrm(self._find(f'网格{m}'))
                ext.set('cy', str(int(ext.get('cy')) - shrink))
            _, ext = _xfrm(self._find('当前时间线'))
            ext.set('cy', str(int(ext.get('cy')) - shrink))
            for nm in ('当前时间标签', '图例-计划', '图例文字-计划', '图例-实际', '图例文字-实际', '图例-当前', '图例文字-当前'):
                off, _ = _xfrm(self._find(nm))
                off.set('y', str(int(off.get('y')) - shrink))
        if now is not None:
            x = int(x0 + now * mw)
            off, _ = _xfrm(self._find('当前时间线'))
            off.set('x', str(x))
            off, ext = _xfrm(self._find('当前时间标签'))
            off.set('x', str(x - int(ext.get('cx')) // 2))
        if now_label:
            self.runs('图例文字-当前', [f'当前时间（{now_label}）'])
        return self

    def timeline_done(self, k):
        """研究进度安排 page: stages 1..k shown as completed (filled nodes + red progress line)."""
        n = 0
        while self._find(f'阶段{n + 1}-节点', False) is not None:
            n += 1
        for i in range(1, n + 1):
            spPr = self._find(f'阶段{i}-节点').find(qn('p:spPr'))
            if i <= k:
                _set_fill(spPr, 'accent1')
                _set_line(spPr, 'bg1', 3.0)
            else:
                _set_fill(spPr, 'bg1')
                _set_line(spPr, 'accent1', 2.25)
        bar = self._find('已完成进度', False)
        if bar is not None:
            if k <= 0:
                bar.getparent().remove(bar)
            else:
                off, ext = _xfrm(bar)
                n_off, n_ext = _xfrm(self._find(f'阶段{min(k, n)}-节点'))
                cx = int(n_off.get('x')) + int(n_ext.get('cx')) // 2
                ext.set('cx', str(max(cx - int(off.get('x')), 0)))
        return self

    # ------------------------------------------------------------------ custom shapes (style-consistent)
    def _next_id(self):
        ids = [int(c.get('id')) for c in self._tree.iter(qn('p:cNvPr')) if c.get('id', '').isdigit()]
        return max(ids + [1]) + 1

    def _append(self, xml):
        el = etree.fromstring(f'<root xmlns:a="{NS_A}" xmlns:p="{NS_P}" xmlns:r="{NS_R}">{xml}</root>')[0]
        el.find('.//' + qn('p:cNvPr')).set('id', str(self._next_id()))
        ext = self._tree.find(qn('p:extLst'))
        if ext is not None:
            ext.addprevious(el)
        else:
            self._tree.append(el)
        return el

    @staticmethod
    def _run_xml(text, size, bold, color, font):
        fonts = {'serif': '<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/>',
                 'sans': '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/>'}[font]
        t = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        b = ' b="1"' if bold else ''
        fill = _solid_xml(color, ns=False)
        return (f'<a:r><a:rPr lang="zh-CN" altLang="en-US" sz="{int(size * 100)}"{b} dirty="0">'
                f'{fill}{fonts}</a:rPr><a:t>{t}</a:t></a:r>')

    def add_text(self, text, x, y, w, h, size=18, bold=False, color='tx1', font='sans', align='l', anchor='t',
                 line_spacing=1.5, name='自定义文本', bullets=False):
        """Text box in the template's style (cm). text: str or list of paragraphs."""
        paras = text.split('\n') if isinstance(text, str) else text
        algn = {'l': 'l', 'c': 'ctr', 'ctr': 'ctr', 'r': 'r', 'j': 'just'}[align]
        bu = (f'<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/>' if bullets else '<a:buNone/>')
        mar = f' marL="{int(0.5 * CM)}" indent="{-int(0.5 * CM)}"' if bullets else ''
        ps = ''.join(f'<a:p><a:pPr algn="{algn}"{mar}><a:lnSpc><a:spcPct val="{int(line_spacing * 100000)}"/></a:lnSpc>{bu}</a:pPr>'
                     f'{self._run_xml(p, size, bold, color, font)}</a:p>' for p in paras)
        xml = (f'<p:sp><p:nvSpPr><p:cNvPr id="0" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
               f'<a:xfrm><a:off x="{int(x * CM)}" y="{int(y * CM)}"/><a:ext cx="{int(w * CM)}" cy="{int(h * CM)}"/></a:xfrm>'
               f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" '
               f'tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="{ANCHORS[anchor]}"><a:noAutofit/>'
               f'</a:bodyPr><a:lstStyle/>{ps}</p:txBody></p:sp>')
        self._append(xml)
        return self

    def add_box(self, x, y, w, h, fill='bg2', line=None, shape='rect', name='自定义色块'):
        """Flat panel / disc in theme colours (cm). shape: rect | roundRect | ellipse."""
        ln = f'<a:ln w="12700">{_solid_xml(line, ns=False)}</a:ln>' if line else '<a:ln><a:noFill/></a:ln>'
        fill_xml = _solid_xml(fill, ns=False)
        xml = (f'<p:sp><p:nvSpPr><p:cNvPr id="0" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
               f'<a:xfrm><a:off x="{int(x * CM)}" y="{int(y * CM)}"/><a:ext cx="{int(w * CM)}" cy="{int(h * CM)}"/></a:xfrm>'
               f'<a:prstGeom prst="{shape}"><a:avLst/></a:prstGeom>{fill_xml}{ln}'
               f'</p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>')
        self._append(xml)
        return self

    def add_takeaway(self, label, text, y, h=2.4, x=2.0, w=29.87):
        """The template's 结论框: grey bar, red serif label, thin divider, one-sentence takeaway."""
        lw = len(label) * 18 / 72 * 2.54 + 0.15
        self.add_box(x, y, w, h, 'bg2', name='结论底框')
        self.add_text(label, x + 0.8, y, lw, h, size=18, bold=True, color='accent1', font='serif', anchor='m', line_spacing=1.0, name='结论标签')
        self.add_box(x + 0.8 + lw + 0.6, y + 0.55, 0.035, h - 1.1, ('accent1', 60000, 40000), name='结论分隔')
        tx = x + 0.8 + lw + 1.2
        self.add_text(text, tx, y, x + w - 0.8 - tx, h, size=17, anchor='m', line_spacing=1.3, name='结论内容')
        return self


def _template_path(template=None, style=None):
    if template is not None:
        return template
    style = DEFAULT_STYLE if style is None else style
    if style not in STYLE_TEMPLATES:
        raise ValueError(f'未知视觉模板 {style!r}。可选值：{", ".join(STYLE_TEMPLATES)}')
    return STYLE_TEMPLATES[style]


class Deck:
    """Assemble a new presentation from template pages. Template pages are dropped on save."""

    def __init__(self, template=None, *, style=None):
        self.prs = Presentation(_template_path(template, style))
        self._tpl = list(self.prs.slides)
        self._tpl_rids = [s.rId for s in self.prs.slides._sldIdLst]
        self.keys = {}
        for i, s in enumerate(self._tpl, 1):
            k = s._element.cSld.get('name') or f'page{i}'
            self.keys[k] = s
        self._sections = []
        self._unsectioned = []
        self._saved = False

    def _resolve(self, page):
        if isinstance(page, int):
            return self._tpl[page - 1], self._tpl[page - 1]._element.cSld.get('name') or f'page{page}'
        if page not in self.keys:
            close = difflib.get_close_matches(page, list(self.keys), n=5, cutoff=0.3)
            raise KeyError(f'模板里没有页面「{page}」。相近的 key：{close}。运行 `python sdu_ppt.py catalog` 查看全部。')
        return self.keys[page], page

    def section(self, name):
        """Start a PowerPoint section (节) for the slides added after this call."""
        self._sections.append((name, []))
        return self

    def _register(self):
        sld_id = self.prs.slides._sldIdLst[-1].id
        (self._sections[-1][1] if self._sections else self._unsectioned).append(sld_id)

    def _clone_chart(self, chart_part):
        package = self.prs.part.package
        partname = package.next_partname('/ppt/charts/chart%d.xml')
        el = copy.deepcopy(chart_part._element)
        ext = el.find(qn('c:externalData'))
        if ext is not None:
            el.remove(ext)
        new_part = ChartPart(partname, chart_part.content_type, package, el)
        xlsx = chart_part.chart_workbook.xlsx_part
        if xlsx is not None:
            new_part.chart_workbook.update_from_xlsx_blob(xlsx.blob)
        return new_part

    def add(self, page):
        """Copy a template page (key like 'weekly' or 1-based template number) to the end of the deck."""
        src, key = self._resolve(page)
        slide = self.prs.slides.add_slide(src.slide_layout)
        tree = slide.shapes._spTree
        for el in list(tree)[2:]:
            tree.remove(el)
        for el in list(src.shapes._spTree)[2:]:
            if _local(el) != 'extLst':
                tree.append(copy.deepcopy(el))
        bg = src._element.cSld.find(qn('p:bg'))
        if bg is not None:
            slide._element.cSld.insert(0, copy.deepcopy(bg))
        slide._element.cSld.set('name', key)
        rid_map = {}
        for rid, rel in src.part.rels.items():
            if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
                continue
            if rel.is_external:
                rid_map[rid] = slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype == RT.CHART:
                rid_map[rid] = slide.part.relate_to(self._clone_chart(rel.target_part), RT.CHART)
            else:
                rid_map[rid] = slide.part.relate_to(rel.target_part, rel.reltype)
        for el in tree.iter():
            for attr in R_ATTRS:
                v = el.get(attr)
                if v is not None and v in rid_map:
                    el.set(attr, rid_map[v])
        self._register()
        return SlideRef(self, slide, key)

    def add_blank(self, layout, title=None):
        """New slide from a master layout (e.g. '仅标题', '标题和内容', '图表解读', '空白') for custom pages."""
        lay = next((l for l in self.prs.slide_layouts if l.name == layout), None)
        if lay is None:
            raise KeyError(f'没有版式「{layout}」。可用：{[l.name for l in self.prs.slide_layouts]}')
        slide = self.prs.slides.add_slide(lay)
        for ph in slide.placeholders:
            lph = lay.placeholders.get(idx=ph.placeholder_format.idx)
            if lph is not None:
                ph._element.find('.//' + qn('p:cNvPr')).set('name', lph.name)
        sld_num = next((p for p in lay.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER), None)
        ref = SlideRef(self, slide, f'blank:{layout}')
        if sld_num is not None:
            ref._append(f'<p:sp><p:nvSpPr><p:cNvPr id="0" name="灯片编号"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
                        f'<p:nvPr><p:ph type="sldNum" sz="quarter" idx="{sld_num.placeholder_format.idx}"/></p:nvPr></p:nvSpPr>'
                        f'<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{{8A1D2C3E-5B4F-4C6D-9E7F-1A2B3C4D5E6F}}" '
                        f'type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld></a:p></p:txBody></p:sp>')
        if title is not None and slide.shapes.title is not None:
            ref.set(slide.shapes.title.name, title)
        self._register()
        return ref

    def save(self, path):
        if self._saved:
            raise RuntimeError('Deck 只能保存一次；请重新创建 Deck。')
        sld_lst = self.prs.slides._sldIdLst
        for sld in list(sld_lst):
            if sld.rId in self._tpl_rids:
                rid = sld.rId
                sld_lst.remove(sld)
                self.prs.part.drop_rel(rid)
        pres = self.prs.part._element
        ext_lst = pres.find(qn('p:extLst'))
        if ext_lst is not None:
            for ext in list(ext_lst):
                if ext.get('uri') == SECTION_EXT:
                    ext_lst.remove(ext)
        if self._sections:
            sections = list(self._sections)
            if self._unsectioned:
                sections.insert(0, ('默认节', self._unsectioned))
            if ext_lst is None:
                ext_lst = etree.SubElement(pres, qn('p:extLst'))
            ext = etree.Element(qn('p:ext'))
            ext.set('uri', SECTION_EXT)
            sl = etree.SubElement(ext, f'{{{NS_P14}}}sectionLst', nsmap={'p14': NS_P14})
            for name, ids in sections:
                sec = etree.SubElement(sl, f'{{{NS_P14}}}section')
                sec.set('name', name)
                sec.set('id', '{' + str(uuid.uuid4()).upper() + '}')
                lst = etree.SubElement(sec, f'{{{NS_P14}}}sldIdLst')
                for sid in ids:
                    etree.SubElement(lst, f'{{{NS_P14}}}sldId').set('id', str(sid))
            ext_lst.insert(0, ext)
        self.prs.part.rename_slide_parts([s.rId for s in sld_lst])
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self.prs.save(path)
        self._saved = True
        return path


# ====================================================================== CLI
def _describe(slide):
    ref = SlideRef(None, slide, slide._element.cSld.get('name'))
    rows = []
    for name, kind, text in ref.names():
        if (kind == 'sp' and not text.strip()) or (name or '').startswith('装饰'):
            continue
        if kind in ('cxnSp',):
            continue
        label = {'pic': '图片', 'grpSp': '组合/图片框', 'chart': '图表', 'table': '表格'}.get(kind, '')
        if kind.startswith('placeholder:'):
            label = '占位符'
        rows.append((name, label, text))
    return rows


def catalog(template=None, md=False, *, style=None):
    prs = Presentation(_template_path(template, style))
    sec_of = {}
    pres = prs.part._element
    for sec in pres.iter(f'{{{NS_P14}}}section'):
        for sid in sec.iter(f'{{{NS_P14}}}sldId'):
            sec_of[int(sid.get('id'))] = sec.get('name')
    out = []
    ids = [s.id for s in prs.slides._sldIdLst]
    for i, s in enumerate(prs.slides, 1):
        key = s._element.cSld.get('name') or f'page{i}'
        cname, note = PAGES.get(key, (key, ''))
        out.append(f'\n### {i:02d} · `{key}` · {cname}' if md else f'\n[{i:02d}] {key}  {cname}  （{sec_of.get(ids[i - 1], "")} / 版式：{s.slide_layout.name}）')
        if md:
            out.append(f'- 分节：{sec_of.get(ids[i - 1], "")}；版式：{s.slide_layout.name}')
            out.append(f'- 填写：{note}')
            out.append('- 可编辑形状：')
        else:
            out.append(f'    提示：{note}')
        for name, label, text in _describe(s):
            t = text.replace('\n', ' ').replace('\t', ' ')
            t = t if len(t) <= 34 else t[:33] + '…'
            out.append(f'  - `{name}`{(" ［" + label + "］") if label else ""}{"：" + t if t else ""}' if md
                       else f'    {name:<16} {label:<6} {t}')
    return '\n'.join(out)


def inspect_deck(path, only=None):
    prs = Presentation(path)
    lines = []
    for i, s in enumerate(prs.slides, 1):
        if only and i not in only:
            continue
        lines.append(f'\n[{i:02d}] {s._element.cSld.get("name") or ""}  （版式：{s.slide_layout.name}）')
        for name, kind, text in SlideRef(None, s, '').names():
            t = text.replace('\n', ' ')
            lines.append(f'    {name:<18} {kind:<22} {t[:40]}')
    return '\n'.join(lines)


def build(plan_path, out_path, template=None, *, style=None):
    """Build a deck from a JSON plan (see references/api.md)."""
    with open(plan_path, encoding='utf-8') as f:
        plan = json.load(f)
    base = os.path.dirname(os.path.abspath(plan_path))
    if template is None:
        template = plan.get('template')
    if template and not os.path.isabs(template):
        template = os.path.join(base, template)
    deck = Deck(template, style=style if style is not None else plan.get('style'))
    for item in plan['slides']:
        if 'section' in item and 'page' not in item and 'layout' not in item:
            deck.section(item['section'])
            continue
        s = deck.add(item['page']) if 'page' in item else deck.add_blank(item['layout'], title=item.get('title'))
        for name, val in item.get('set', {}).items():
            s.set(name, val)
        for name, val in item.get('runs', {}).items():
            s.runs(name, val)
        for name, st in item.get('style', {}).items():
            s.style(name, **st)
        for name, spec in item.get('chart', {}).items():
            s.chart(name, spec['categories'], spec['series'], title=spec.get('title'), number_format=spec.get('number_format'))
        for name, spec in item.get('table', {}).items():
            s.table(name, spec['rows'], header=spec.get('header'), highlight_last=spec.get('highlight_last'))
        for name, img in item.get('image', {}).items():
            s.image(name, img if os.path.isabs(img) else os.path.join(base, img))
        if 'progress' in item:
            s.progress([tuple(t) for t in item['progress']['tasks']], item['progress'].get('overall'))
        if 'gantt' in item:
            g = item['gantt']
            s.gantt([(t[0], tuple(t[1]), tuple(t[2]) if len(t) > 2 and t[2] else None) for t in g['tasks']],
                    start_month=g.get('start_month', 1), year_label=g.get('year_label'), now=g.get('now'), now_label=g.get('now_label'))
        if 'timeline_done' in item:
            s.timeline_done(item['timeline_done'])
        if item.get('remove'):
            s.remove(*item['remove'])
        for prefix in item.get('respace', []):
            s.respace(prefix)
        if item.get('notes'):
            s.notes(item['notes'])
    return deck.save(out_path)


def main(argv=None):
    ap = argparse.ArgumentParser(description='SDU template deck builder')
    sub = ap.add_subparsers(dest='cmd', required=True)
    c = sub.add_parser('catalog', help='list template pages and editable shapes')
    c.add_argument('--template')
    c.add_argument('--style', choices=tuple(STYLE_TEMPLATES))
    c.add_argument('--md', action='store_true', help='markdown output')
    i = sub.add_parser('inspect', help='list shapes of a deck')
    i.add_argument('deck')
    i.add_argument('--slides', help='comma separated slide numbers')
    b = sub.add_parser('build', help='build a deck from a JSON plan')
    b.add_argument('plan')
    b.add_argument('out')
    b.add_argument('--template')
    b.add_argument('--style', choices=tuple(STYLE_TEMPLATES))
    args = ap.parse_args(argv)
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if args.cmd == 'catalog':
        print(catalog(args.template, args.md, style=args.style))
    elif args.cmd == 'inspect':
        only = {int(x) for x in args.slides.split(',')} if args.slides else None
        print(inspect_deck(args.deck, only))
    else:
        print(build(args.plan, args.out, args.template, style=args.style))


if __name__ == '__main__':
    main()
