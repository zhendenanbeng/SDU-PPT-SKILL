"""文献汇报示例：python examples/literature_report.py out/literature.pptx"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'scripts'))
from sdu_ppt import Deck  # noqa: E402

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'out', 'literature.pptx')
deck = Deck()

deck.section('论文概况')
s = deck.add('cover')
s.set('标题', '小样本学习中的特征重加权')          # ≤12 字；更长的标题会被 fit() 自动缩小
s.set('副标题', '文献汇报 · 课题组文献研读第 3 期')
s.runs('汇报信息', ['汇报人：', '王五', '　　|　　', '时间：', '2026年3月17日'])

s = deck.add('paper')
s.set('论文题目', 'Feature Re-weighting for Few-Shot Image Classification')
s.set('作者', 'Author A, Author B, Author C, et al.')
for k, val in enumerate(['Journal Name', '2024 年 5 月', 'IF xx.x', '中科院 1 区', '10.xxxx/xxxxx', 'xx 次'], 1):
    s.set(f'信息{k}', val)
for k, kw in enumerate(['小样本学习', '特征重加权', '元学习', '图像分类'], 1):
    s.set(f'关键词{k}', kw)
s.set('结论内容', '方法简洁、可复现，其特征重加权思路可直接用于本课题的分类模块。')

deck.section('方法与结果')
s = deck.add('text')
s.set('标题', '研究背景与问题')
s.set('内容占位符', ['小样本学习希望模型只用少量标注样本就能识别新类别，但特征提取器在新类别上往往泛化不足。',
                    '本文提出按类别动态调整特征通道权重的方法，使模型更关注具有判别性的特征。'])

s = deck.add('figread')
s.image('图片占位符', os.path.join(HERE, 'media', 'accuracy_bars.png'))
s.set('图题', '图 3　不同方法在三个数据集上的准确率对比')
s.set('图源', '图片来源：Author A, et al. Journal Name, 2024.（示例图）')
s.set('解读要点', ['解读要点',
                   ['整体趋势：', '三个数据集上均取得最高准确率，平均领先约 3 个百分点。'],
                   ['关键差异：', '数据集 2 难度最大，改进主要来自困难样本。'],
                   ['值得借鉴：', '每组结果都报告了误差范围，结论更可信。']])

deck.section('评述')
s = deck.add('review')
s.set('评述1', ['提出通道重加权模块，结构简单', '在三个基准上取得最优结果', '代码与数据公开，便于复现'])
s.set('评述2', ['只验证了图像分类任务', '对超参数较为敏感', '缺少计算开销分析'])
s.set('评述3', ['可替换本课题的特征融合模块', '沿用其评价协议便于比较', '可在小样本检测任务中验证'])
s.set('结论内容', '思路清晰、实验扎实，建议作为本课题下一阶段的主要参考方法。')

s = deck.add('litcmp')
s.table('表格', [['[1] 作者A 等（2021）', '公开图像数据集', '原型网络', '度量学习简单有效', '特征泛化不足'],
                 ['[2] 作者B 等（2023）', '公开图像数据集', '元学习 + 注意力', '提升困难样本识别', '训练成本较高'],
                 ['本研究', '自建小样本数据集', '特征重加权改进', '兼顾精度与效率', '扩展到检测任务']])
s.set('结论内容', '现有方法多关注度量方式，较少讨论特征通道的选择，这正是本课题的切入点。')

# 自定义页面：用母版版式 + 风格组件
s = deck.add_blank('仅标题', title='本课题拟解决的问题')
s.add_text(['如何在样本极少时选出最有判别力的特征通道？', '如何在不显著增加计算量的前提下提升泛化能力？'],
           2.0, 4.2, 29.87, 4.0, size=22, bullets=True)
s.add_takeaway('研究思路', '以特征重加权为基础，结合轻量化注意力模块，并在自建数据集上验证。', y=12.2)

s = deck.add('refs')
s.set('内容占位符', ['[1]\tAUTHOR A, AUTHOR B. Title of the first reference[J]. Journal Name, 2021, 10(2): 1-12.',
                    '[2]\tAUTHOR C, AUTHOR D. Title of the second reference[C]//Proceedings of the Conference. 2023: 100-110.'])
s.remove('参考文献说明')
deck.add('close')
print(deck.save(out))
