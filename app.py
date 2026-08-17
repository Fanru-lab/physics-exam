# -*- coding: utf-8 -*-
"""Flask 云后端 —— 机械能守恒章节·四阶测试问卷答题系统
功能：登录建档 + 前测（第一大题）+ 动态纠错（第二大题）+ 后测（第三大题）+ 诊断报告
"""
import os
import sys
import json
import traceback
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STUDENT_DATA_DIR = os.path.join(BASE_DIR, 'student_data')
os.makedirs(STUDENT_DATA_DIR, exist_ok=True)

sys.path.insert(0, os.path.join(BASE_DIR, 'analysis'))

# 导入题目数据 + 纠正生成器
def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

QUESTIONS = load_json(os.path.join(BASE_DIR, 'test_data', 'questions.json'))
CORRECTION_BANK = load_json(os.path.join(BASE_DIR, 'test_data', 'correction_bank.json'))

try:
    from agent6_correction_generator import generate_correction
    CORRECTION_READY = True
except ImportError:
    CORRECTION_READY = False
    print('[WARN] 纠正生成器未加载')

# 导入四个诊断报告生成 agent
try:
    from agent2_pretest_analysis import generate_pretest_report
    from agent3_posttest_analysis import generate_posttest_report
    from agent4_comparison import generate_comparison_report
    from agent5_correction_analysis import generate_correction_report
    REPORTS_READY = True
except ImportError as e:
    REPORTS_READY = False
    print(f'[WARN] 报告生成模块未加载: {e}')

app = Flask(__name__, static_folder=BASE_DIR, static_url_path='')


@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'login.html')


@app.route('/index.html')
def exam():
    return send_from_directory(BASE_DIR, 'index.html')


@app.route('/api/health')
def health():
    return jsonify({'status': 'ok'})


@app.route('/api/login', methods=['POST'])
def login():
    """登录：创建学生文件夹"""
    try:
        data = request.get_json(force=True)
        name = data.get('name', '').strip()
        sid = data.get('studentId', '').strip()
        folder_name = f'{name}_{sid}' if name and sid else '_unknown'
        student_dir = os.path.join(STUDENT_DATA_DIR, folder_name)
        os.makedirs(student_dir, exist_ok=True)
        for sub in ['answers', 'reports']:
            os.makedirs(os.path.join(student_dir, sub), exist_ok=True)
        print(f'[LOGIN] {folder_name}')
        return jsonify({'success': True, 'folder': folder_name})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/questions/<section>')
def get_questions(section):
    """返回前测/后测题目数据"""
    if section not in ('pretest', 'posttest'):
        return jsonify({'error': 'Invalid section'}), 400
    return jsonify({'questions': QUESTIONS[section]})


@app.route('/api/submit/pretest', methods=['POST'])
def submit_pretest():
    """提交第一大题（前测）：锁定并保存数据，生成第二大题 + 前测诊断报告"""
    try:
        data = request.get_json(force=True)
        data['server_timestamp'] = datetime.now().isoformat()
        folder_name = _get_folder(data)
        _save_answer(data, folder_name, 'pretest')

        # 生成第二大题（纠错题）
        correction_result = None
        if CORRECTION_READY:
            correction_result = generate_correction(data.get('answers', {}), QUESTIONS['pretest'], CORRECTION_BANK)

        # 生成第一大题诊断报告
        report = None
        if REPORTS_READY:
            try:
                report_dir = os.path.join(STUDENT_DATA_DIR, folder_name, 'reports')
                p = generate_pretest_report(data.get('answers', {}), QUESTIONS['pretest'],
                                            data.get('studentName', ''), data.get('studentId', ''), report_dir)
                report = os.path.basename(p)
            except Exception as e:
                print(f'[Agent2 前测报告失败] {traceback.format_exc()}')

        return jsonify({
            'success': True,
            'section': 'pretest',
            'folder': folder_name,
            'correction': correction_result,
            'report': report,
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/submit/correction', methods=['POST'])
def submit_correction():
    """提交第二大题（纠错）：锁定并保存数据，生成纠错诊断报告"""
    try:
        data = request.get_json(force=True)
        data['server_timestamp'] = datetime.now().isoformat()
        folder_name = _get_folder(data)
        _save_answer(data, folder_name, 'correction')

        # 生成第二大题诊断报告
        report = None
        if REPORTS_READY and data.get('correction_questions'):
            try:
                report_dir = os.path.join(STUDENT_DATA_DIR, folder_name, 'reports')
                p = generate_correction_report(data.get('answers', {}), data.get('correction_questions', []),
                                               data.get('studentName', ''), data.get('studentId', ''), report_dir)
                report = os.path.basename(p)
            except Exception as e:
                print(f'[Agent5 纠错报告失败] {traceback.format_exc()}')

        return jsonify({'success': True, 'section': 'correction', 'folder': folder_name, 'report': report})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/submit/posttest', methods=['POST'])
def submit_posttest():
    """提交第三大题（后测）：锁定并保存数据，生成后测诊断报告 + 前后测对比报告"""
    try:
        data = request.get_json(force=True)
        data['server_timestamp'] = datetime.now().isoformat()
        folder_name = _get_folder(data)
        _save_answer(data, folder_name, 'posttest')
        _save_complete_data(data, folder_name)

        # 生成后测诊断报告 + 前后测对比报告
        reports = None
        if REPORTS_READY:
            try:
                report_dir = os.path.join(STUDENT_DATA_DIR, folder_name, 'reports')
                os.makedirs(report_dir, exist_ok=True)
                student_name = data.get('studentName', '')
                student_id = data.get('studentId', '')
                pretest_answers = data.get('pretest_answers', {})
                posttest_answers = data.get('answers', {})

                p3 = generate_posttest_report(posttest_answers, QUESTIONS['posttest'],
                                              student_name, student_id, report_dir)
                p4 = generate_comparison_report(pretest_answers, posttest_answers,
                                                QUESTIONS['pretest'], QUESTIONS['posttest'],
                                                student_name, student_id, report_dir)
                reports = {
                    'posttest': os.path.basename(p3),
                    'comparison': os.path.basename(p4),
                }
            except Exception as e:
                print(f'[Agent3/4 报告失败] {traceback.format_exc()}')

        return jsonify({
            'success': True,
            'section': 'posttest',
            'folder': folder_name,
            'report': reports,
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate_reports', methods=['POST'])
def generate_reports():
    """手动触发生成四份 Word 诊断报告"""
    try:
        data = request.get_json(force=True)
        folder_name = _get_folder(data)

        # 读取该学生最新保存的完整答题数据
        complete = _load_complete_data(folder_name)
        if complete is None:
            return jsonify({'error': '未找到该学生的答题数据'}), 404

        # 生成四份 Word 诊断报告
        reports = None
        try:
            reports = _generate_all_reports(complete, folder_name)
        except Exception as e:
            print(f'[REPORT ERROR] {traceback.format_exc()}')

        return jsonify({
            'success': True,
            'folder': folder_name,
            'report': reports,
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ===== 辅助函数 =====

def _get_folder(data):
    name = data.get('studentName', '')
    sid = data.get('studentId', '')
    return f'{name}_{sid}' if name and sid else '_anonymous'


def _save_answer(data, folder_name, section):
    base = os.path.join(STUDENT_DATA_DIR, folder_name)
    answer_dir = os.path.join(base, 'answers')
    os.makedirs(answer_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'answer_{section}_{ts}.json'
    filepath = os.path.join(answer_dir, filename)
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'[SAVED] {folder_name}/{section}: {filename}')
    return filepath


def _save_complete_data(data, folder_name):
    """保存完整答题数据（前测+纠错+后测），供手动生成报告使用"""
    base = os.path.join(STUDENT_DATA_DIR, folder_name)
    answer_dir = os.path.join(base, 'answers')
    os.makedirs(answer_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'complete_{ts}.json'
    filepath = os.path.join(answer_dir, filename)
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'[COMPLETE] {folder_name}: {filename}')
    return filepath


def _load_complete_data(folder_name):
    """读取该学生最新保存的完整答题数据"""
    base = os.path.join(STUDENT_DATA_DIR, folder_name)
    answer_dir = os.path.join(base, 'answers')
    if not os.path.exists(answer_dir):
        return None
    files = [f for f in os.listdir(answer_dir) if f.startswith('complete_') and f.endswith('.json')]
    if not files:
        return None
    files.sort(reverse=True)  # 最新在前
    latest = os.path.join(answer_dir, files[0])
    with open(latest, 'r', encoding='utf-8') as f:
        return json.load(f)


def _generate_all_reports(data, folder_name):
    """调用四个 agent 生成四份 Word 诊断报告"""
    if not REPORTS_READY:
        return None

    student_name = data.get('studentName', '')
    student_id = data.get('studentId', '')
    pretest = QUESTIONS['pretest']
    posttest = QUESTIONS['posttest']

    pretest_answers = data.get('pretest_answers', {})
    posttest_answers = data.get('answers', {})
    correction_answers = data.get('correction_answers', {})
    correction_questions = data.get('correction_questions', [])

    report_dir = os.path.join(STUDENT_DATA_DIR, folder_name, 'reports')
    os.makedirs(report_dir, exist_ok=True)

    reports = {}

    # 1. 第一大题（前测）诊断报告 —— Agent2
    try:
        p1 = generate_pretest_report(pretest_answers, pretest, student_name, student_id, report_dir)
        reports['pretest'] = os.path.basename(p1)
    except Exception as e:
        print(f'[Agent2 前测报告失败] {traceback.format_exc()}')

    # 2. 第二大题（纠错）诊断报告 —— Agent5
    if correction_questions:
        try:
            p2 = generate_correction_report(correction_answers, correction_questions, student_name, student_id, report_dir)
            reports['correction'] = os.path.basename(p2)
        except Exception as e:
            print(f'[Agent5 纠错报告失败] {traceback.format_exc()}')

    # 3. 第三大题（后测）诊断报告 —— Agent3
    try:
        p3 = generate_posttest_report(posttest_answers, posttest, student_name, student_id, report_dir)
        reports['posttest'] = os.path.basename(p3)
    except Exception as e:
        print(f'[Agent3 后测报告失败] {traceback.format_exc()}')

    # 4. 前后测对比报告 —— Agent4
    try:
        p4 = generate_comparison_report(pretest_answers, posttest_answers, pretest, posttest,
                                        student_name, student_id, report_dir)
        reports['comparison'] = os.path.basename(p4)
    except Exception as e:
        print(f'[Agent4 对比报告失败] {traceback.format_exc()}')

    return reports


# ===== 数据同步接口 =====

@app.route('/api/sync/list_students')
def list_students():
    students = []
    if os.path.exists(STUDENT_DATA_DIR):
        for folder in os.listdir(STUDENT_DATA_DIR):
            folder_path = os.path.join(STUDENT_DATA_DIR, folder)
            if os.path.isdir(folder_path):
                sections = []
                answers_dir = os.path.join(folder_path, 'answers')
                if os.path.exists(answers_dir):
                    for f in os.listdir(answers_dir):
                        if f.endswith('.json'):
                            sections.append(f'answers/{f}')
                reports_dir = os.path.join(folder_path, 'reports')
                if os.path.exists(reports_dir):
                    for f in os.listdir(reports_dir):
                        sections.append(f'reports/{f}')
                parts = folder.rsplit('_', 1)
                name = parts[0] if len(parts) >= 2 else folder
                students.append({'name': name, 'folder': folder, 'sections': sections})
    return jsonify({'students': students})


@app.route('/api/sync/download')
def download_file():
    folder = request.args.get('folder', '')
    section = request.args.get('section', '')
    if not folder or not section:
        return jsonify({'error': 'Missing parameters'}), 400
    filepath = os.path.join(STUDENT_DATA_DIR, folder, section)
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404
    from flask import send_file
    return send_file(filepath, as_attachment=True)


if __name__ == '__main__':
    print('=' * 50)
    print('  机械能守恒章节·四阶测试问卷系统')
    print(f'  数据目录: {STUDENT_DATA_DIR}')
    print(f'  纠正生成器: {"已启用" if CORRECTION_READY else "未启用"}')
    print('=' * 50)
    app.run(host='0.0.0.0', port=5000, debug=False)
