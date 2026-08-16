# -*- coding: utf-8 -*-
"""
Agent6 —— 第二大题（纠错环节）动态生成器

功能：读取用户第一大题（前测）的答题情况，判断错题（答案阶或理由阶任一出错即算错），
从习题纠正库中对应的板块调取一道题，组成第二大题的内容。

判定规则：
- 答案阶（answer）与标准答案 answer_key 不一致 → 错
- 理由阶（reason）与标准理由 reason_key 不一致 → 错
- 二者任一出错，该题即为"错题"
"""
import os, json, random

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
QUESTIONS_PATH = os.path.join(BASE_DIR, 'test_data', 'questions.json')
CORRECTION_PATH = os.path.join(BASE_DIR, 'test_data', 'correction_bank.json')


def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def split_misconceptions(miscon_str):
    """拆分迷思标签，如 '迷思10+11' -> ['迷思10', '迷思11']"""
    return miscon_str.split('+')


def judge_wrong(answer_choice, reason_choice, answer_key, reason_key):
    """判断一道题是否答错（答案阶或理由阶任一出错即错）"""
    answer_wrong = (answer_choice != answer_key)
    reason_wrong = (reason_choice != reason_key)
    return {
        'answer_wrong': answer_wrong,
        'reason_wrong': reason_wrong,
        'is_wrong': answer_wrong or reason_wrong,
    }


def generate_correction(pretest_answers, questions=None, correction_bank=None):
    """
    根据第一大题答题情况，生成第二大题内容。

    参数：
        pretest_answers: dict，第一大题答题数据，格式：
            {'q0': {'answer': 'A', 'reason': 'B', 'conf1': '很有信心', 'conf2': '一般'}, ...}
            （key 为 'q'+题号-1，或 'q'+题号，两种都兼容）
        questions: 前测题目列表（可选，默认从 questions.json 加载）
        correction_bank: 纠正库（可选，默认从 correction_bank.json 加载）

    返回：
        {
            'wrong_list': [{'num': 1, 'miscon': '迷思1', 'answer_wrong': True, 'reason_wrong': False, ...}, ...],
            'correction_questions': [{'miscon': '迷思1', 'name': '力与功混淆', 'stem': ..., 'options': [...], 'answer': 'C'}, ...],
            'total_wrong': int,
        }
    """
    if questions is None:
        questions = load_json(QUESTIONS_PATH)['pretest']
    if correction_bank is None:
        correction_bank = load_json(CORRECTION_PATH)

    wrong_list = []
    correction_questions = []
    used_bank = {}  # 记录每个迷思已使用的纠正题，避免重复

    for q in questions:
        qnum = q['num']
        # 兼容两种 key 格式：'q0'（0-based）和 'q1'（1-based）
        student = pretest_answers.get(f'q{qnum - 1}') or pretest_answers.get(f'q{qnum}') or {}
        answer_choice = student.get('answer', '')
        reason_choice = student.get('reason', '')

        result = judge_wrong(answer_choice, reason_choice, q['answer_key'], q['reason_key'])

        if result['is_wrong']:
            wrong_list.append({
                'num': qnum,
                'miscon': q['miscon'],
                'src': q['src'],
                'answer_wrong': result['answer_wrong'],
                'reason_wrong': result['reason_wrong'],
                'student_answer': answer_choice,
                'student_reason': reason_choice,
                'correct_answer': q['answer_key'],
                'correct_reason': q['reason_key'],
            })
            # 从纠正库对应板块调取题目
            for mk in split_misconceptions(q['miscon']):
                bank = correction_bank.get(mk)
                if not bank or not bank.get('questions'):
                    continue
                qs = bank['questions']
                # 取该板块还未使用的题目
                used_idx = used_bank.get(mk, set())
                available = [i for i in range(len(qs)) if i not in used_idx]
                if not available:
                    # 题目用完，随机复用
                    idx = random.randint(0, len(qs) - 1)
                else:
                    idx = available[0]
                    used_idx.add(idx)
                    used_bank[mk] = used_idx
                selected = qs[idx]
                correction_questions.append({
                    'miscon': mk,
                    'name': bank.get('name', mk),
                    'stem': selected['stem'],
                    'options': selected['options'],
                    'answer': selected['answer'],
                })

    return {
        'wrong_list': wrong_list,
        'correction_questions': correction_questions,
        'total_wrong': len(wrong_list),
    }


if __name__ == '__main__':
    # 测试
    test_answers = {
        'q0': {'answer': 'A', 'reason': 'B'},  # 第1题：答案A对，理由B对 -> 正确
        'q1': {'answer': 'A', 'reason': 'B'},  # 第2题：答案C对，学生答A -> 错
        'q2': {'answer': 'B', 'reason': 'B'},  # 第3题：正确
        'q3': {'answer': 'C', 'reason': 'C'},  # 第4题：答案B对，学生答C -> 错
        'q4': {'answer': 'C', 'reason': 'C'},  # 第5题：正确
        'q5': {'answer': 'B', 'reason': 'B'},  # 第6题：正确
        'q6': {'answer': 'C', 'reason': 'C'},  # 第7题：正确
        'q7': {'answer': 'C', 'reason': 'C'},  # 第8题：正确
        'q8': {'answer': 'B', 'reason': 'B'},  # 第9题：正确
        'q9': {'answer': 'B', 'reason': 'B'},  # 第10题：正确
    }
    result = generate_correction(test_answers)
    print(f"错题数: {result['total_wrong']}")
    for w in result['wrong_list']:
        print(f"  第{w['num']}题（{w['miscon']}）答案错:{w['answer_wrong']} 理由错:{w['reason_wrong']}")
    print(f"生成的第二大题题目数: {len(result['correction_questions'])}")
    for c in result['correction_questions']:
        print(f"  [{c['miscon']}-{c['name']}] {c['stem'][:30]}... 答案:{c['answer']}")
