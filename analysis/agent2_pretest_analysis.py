# -*- coding: utf-8 -*-
"""
Agent 2：第一大题（前测）诊断报告
功能：读取前测答题数据，按四阶（答案阶+理由阶+信心指数）批改，生成 Word 诊断报告
"""
import os, json
from datetime import datetime
from report_utils import (
    grade_four_tier, collect_misconceptions, _build_report
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def generate_pretest_report(pretest_answers, pretest_questions, student_name, student_id, output_dir=None):
    """
    生成前测诊断报告（Word 文档）

    参数：
        pretest_answers: dict，前测答题数据 {q0~q9: {answer, reason, conf1, conf2}}
        pretest_questions: list，前测题目列表
        student_name, student_id: 学生信息
        output_dir: 输出目录（默认 student_data/姓名_学号/reports）
    返回：报告文件路径
    """
    results = grade_four_tier(pretest_questions, pretest_answers, 'q')
    miscon = collect_misconceptions(results)

    # 统计
    total = len(results)
    wrong = sum(1 for r in results if r['is_wrong'])
    correct = total - wrong

    # 顽固迷思（答错且信心高）
    stubborn = [r for r in results if r['is_wrong'] and r['conf1'] == '很有信心']

    # 表格
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
                f'存在的迷思概念：{"、".join(miscon) if miscon else "无"}',
            ],
        },
        {
            'heading': '三、信心指数分析',
            'texts': [
                f'顽固迷思（答错且对答案很有信心）：{len(stubborn)} 个',
                '说明：答错但信心十足，说明是根深蒂固的错误观念，需重点纠正。',
            ],
        },
    ]

    if output_dir is None:
        output_dir = os.path.join(BASE_DIR, 'student_data', f'{student_name}_{student_id}', 'reports')
    os.makedirs(output_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, '前测诊断报告.docx')
    _build_report('第一大题（前测）诊断报告', student_name, student_id, sections, output_path)
    return output_path


if __name__ == '__main__':
    print('Agent2 前测诊断报告生成器（由 app.py 调用）')
