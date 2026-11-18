"""进度汇报示例：python examples/progress_report.py out/progress.pptx"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'scripts'))
from sdu_ppt import Deck  # noqa: E402

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'out', 'progress.pptx')
deck = Deck()

s = deck.add('cover')
s.set('标题', '课题中期进度汇报')
s.set('副标题', '基于多源数据的城市交通状态预测研究')
s.runs('汇报信息', ['汇报人：', '赵六', '　　|　　', '时间：', '2026年3月24日'])

s = deck.add('progress')
s.progress([('文献调研', 100, 'done'), ('数据采集', 100, 'done'), ('模型构建', 75, 'doing'),
            ('对比实验', 30, 'doing'), ('论文撰写', 0, 'todo')])

s = deck.add('gantt')
s.set('标题', '项目进度（2025.09 — 2026.08）')
s.gantt([('文献调研', (0, 2), (0, 2.5)), ('数据采集', (1.5, 4), (2, 4.5)), ('模型构建', (3.5, 7), (4, 6.8)),
         ('对比实验', (6, 9), (6.2, 6.8)), ('论文撰写', (8, 11), None), ('答辩准备', (11, 12), None)],
        start_month=9, year_label='2025—26', now=6.8, now_label='3 月下旬')

s = deck.add('timeline')
s.set('标题', '阶段进展')
stages = [('2025.10', '文献调研', ('完成 60 篇文献精读', '确定研究问题')),
          ('2025.12', '数据采集', ('接入 3 类交通数据', '完成清洗与对齐')),
          ('2026.03', '模型构建', ('完成时空图模型', '调参中')),
          ('2026.05', '对比实验', ('与 4 种基线比较', '完成消融实验')),
          ('2026.07', '论文撰写', ('完成学位论文初稿', '准备投稿'))]
for i, (date, name, desc) in enumerate(stages, 1):
    s.set(f'阶段{i}-时间', date)
    s.set(f'阶段{i}', name)
    s.set(f'阶段{i}-说明', list(desc))
s.timeline_done(3)

s = deck.add('nextplan')
s.set('计划1', ['完成模型调参与稳定性测试', '补充两组对比实验'])
s.set('计划2', ['完成全部对比与消融实验', '撰写小论文并投稿'])
s.set('计划3', ['完成学位论文初稿', '准备预答辩'])
s.set('结论内容', '需要协调 GPU 资源约 200 卡时，并请导师审阅实验设计。')

deck.add('close')
print(deck.save(out))
