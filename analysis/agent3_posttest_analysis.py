# -*- coding: utf-8 -*-
"""
Agent 3：第三大题（后测）诊断报告
功能：读取后测答题数据，按四阶批改，生成 Word 诊断报告
"""
import os
from datetime import datetime
from report_utils import grade_four_tier, collect_misconceptions, _build_report

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def generate_posttest_report(posttest_answers, posttest_questions, student_name, student_id, output_dir=None):
    results = grade_four_tier(posttest_questions, posttest_answers, 'p')
    miscon = collect_misconceptions(results)

    total = len(results)
    wrong = sum(1 for r in results if r['is_wrong'])
    correct = total - wrong
    stubborn = [r for r in results if r['is_wrong'] and r['conf1'] == '很有信心']

    headers = ['题号', '迷思概念', '答案阶', '理由阶', '答案信心', '理由信心', '结果']
    rows = []
    for r in results:
        status = '✓ 正确' if not r['is_wrong'] else ('✗答案' if r['answer_wrong'] else '✗理由')
        rows.append([
            r['num'], r['miscon'], f'{r["answer"]}/{r["answer_key"]}',
            f'{r["reason"]}/{r["reason_key"]}', r['conf1'], r['conf2'], status
        ])

    sections = [
        {
            'heading': '一、逐题批改结果',
            'table': {'headers': headers, 'rows': rows},
        },
        {
            'heading': '二、迷思概念诊断',
            'texts': [
                f'答对 {correct} 题，答错 {wrong} 题。',
                f'仍存在的迷思概念：{"、".join(miscon) if miscon else "无（全部掌握）"}',
            ],
        },
        {
            'heading': '三、信心指数分析',
            'texts': [
                f'顽固迷思（答错且信心十足）：{len(stubborn)} 个',
            ],
        },
    ]

    if output_dir is None:
        output_dir = os.path.join(BASE_DIR, 'student_data', f'{student_name}_{student_id}', 'reports')
    os.makedirs(output_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, f'后测诊断报告_{ts}.docx')
    _build_report('第三大题（后测）诊断报告', student_name, student_id, sections, output_path)
    return output_path


if __name__ == '__main__':
    print('Agent3 后测诊断报告生成器（由 app.py 调用）')
