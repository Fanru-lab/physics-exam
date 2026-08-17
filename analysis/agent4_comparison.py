# -*- coding: utf-8 -*-
"""
Agent 4：第一大题与第三大题（前后测）对比分析报告
功能：对比前测、后测答题数据，分析迷思概念的纠正情况，生成 Word 对比报告
"""
import os
from datetime import datetime
from report_utils import grade_four_tier, collect_misconceptions, _build_report

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def generate_comparison_report(pretest_answers, posttest_answers, pretest_questions, posttest_questions,
                               student_name, student_id, output_dir=None):
    """生成前后测对比报告"""
    pre_results = grade_four_tier(pretest_questions, pretest_answers, 'q')
    post_results = grade_four_tier(posttest_questions, posttest_answers, 'p')

    pre_miscon = collect_misconceptions(pre_results)
    post_miscon = collect_misconceptions(post_results)
    corrected = [m for m in pre_miscon if m not in post_miscon]
    remained = [m for m in pre_miscon if m in post_miscon]
    new_miscon = [m for m in post_miscon if m not in pre_miscon]

    pre_wrong = sum(1 for r in pre_results if r['is_wrong'])
    post_wrong = sum(1 for r in post_results if r['is_wrong'])
    pre_correct = len(pre_results) - pre_wrong
    post_correct = len(post_results) - post_wrong

    # 逐迷思对比表
    # 前测题目和后测题目都覆盖11个迷思（部分题目覆盖多个迷思），这里按迷思汇总
    miscon_keys = ['迷思1', '迷思2', '迷思3', '迷思4', '迷思5', '迷思6',
                   '迷思7', '迷思8', '迷思9', '迷思10', '迷思11']

    def miscon_wrong(results):
        """统计每个迷思在前/后测中的错题数"""
        d = {}
        for r in results:
            if r['is_wrong']:
                for m in r['miscon'].split('+'):
                    d[m] = d.get(m, 0) + 1
        return d

    pre_wrong_map = miscon_wrong(pre_results)
    post_wrong_map = miscon_wrong(post_results)

    headers = ['迷思概念', '前测错题数', '后测错题数', '变化']
    rows = []
    for mk in miscon_keys:
        pre_c = pre_wrong_map.get(mk, 0)
        post_c = post_wrong_map.get(mk, 0)
        if pre_c == 0 and post_c == 0:
            change = '—'
        elif post_c < pre_c:
            change = '↓ 改善'
        elif post_c > pre_c:
            change = '↑ 退步'
        else:
            change = '→ 未变'
        rows.append([mk, pre_c, post_c, change])

    sections = [
        {
            'heading': '一、总体对比',
            'table': {
                'headers': ['项目', '前测（第一大题）', '后测（第三大题）'],
                'rows': [
                    ['答对题数', pre_correct, post_correct],
                    ['答错题数', pre_wrong, post_wrong],
                    ['存在迷思数', len(pre_miscon), len(post_miscon)],
                ],
            },
        },
        {
            'heading': '二、逐迷思概念对比',
            'table': {'headers': headers, 'rows': rows},
        },
        {
            'heading': '三、迷思概念纠正情况',
            'texts': [
                f'已纠正的迷思概念：{"、".join(corrected) if corrected else "无"}',
                f'仍存在的迷思概念：{"、".join(remained) if remained else "无"}',
                f'新增的迷思概念：{"、".join(new_miscon) if new_miscon else "无"}',
            ],
        },
        {
            'heading': '四、学习效果评价',
            'texts': [
                f'本次学习共纠正了 {len(corrected)} 个迷思概念。',
                '若后测错题数明显减少，说明纠错学习有效；若仍有迷思概念未纠正，建议针对性复习。',
            ],
        },
    ]

    if output_dir is None:
        output_dir = os.path.join(BASE_DIR, 'student_data', f'{student_name}_{student_id}', 'reports')
    os.makedirs(output_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, '前后测对比报告.docx')
    _build_report('前后测对比分析报告', student_name, student_id, sections, output_path)
    return output_path


if __name__ == '__main__':
    print('Agent4 前后测对比报告生成器（由 app.py 调用）')
