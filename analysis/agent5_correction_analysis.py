# -*- coding: utf-8 -*-
"""
Agent 5：第二大题（纠错环节）诊断报告
功能：读取纠错答题数据，批改并生成 Word 诊断报告
"""
import os
from datetime import datetime
from report_utils import grade_correction, _build_report

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def generate_correction_report(correction_answers, correction_questions, student_name, student_id, output_dir=None):
    results = grade_correction(correction_questions, correction_answers)

    total = len(results)
    wrong = sum(1 for r in results if r['is_wrong'])
    correct = total - wrong

    headers = ['题号', '迷思概念', '答案阶', '结果']
    rows = []
    for r in results:
        status = '✓ 正确' if not r['is_wrong'] else '✗ 错误'
        rows.append([
            r['num'], f'{r["miscon"]}-{r["name"]}', f'{r["answer"]}/{r["answer_key"]}', status
        ])

    miscon = sorted(set(r['miscon'] for r in results if r['is_wrong']))

    sections = [
        {
            'heading': '一、纠错题批改结果',
            'table': {'headers': headers, 'rows': rows},
        },
        {
            'heading': '二、纠错效果',
            'texts': [
                f'共 {total} 道纠错题，答对 {correct} 题，答错 {wrong} 题。',
                f'仍存在问题的迷思概念：{"、".join(miscon) if miscon else "无"}',
                '说明：纠错题答对说明对应迷思概念已初步纠正，答错则需进一步巩固。',
            ],
        },
    ]

    if output_dir is None:
        output_dir = os.path.join(BASE_DIR, 'student_data', f'{student_name}_{student_id}', 'reports')
    os.makedirs(output_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, '纠错诊断报告.docx')
    _build_report('第二大题（纠错环节）诊断报告', student_name, student_id, sections, output_path)
    return output_path


if __name__ == '__main__':
    print('Agent5 纠错诊断报告生成器（由 app.py 调用）')
