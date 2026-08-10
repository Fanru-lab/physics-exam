# -*- coding: utf-8 -*-
"""Flask 云后端 —— 完整线上答题系统，数据实时保存并生成Word诊断报告"""
import os
import sys
import json
import traceback
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory

# 工作目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STUDENT_DATA_DIR = os.path.join(BASE_DIR, 'student_data')
os.makedirs(STUDENT_DATA_DIR, exist_ok=True)

# 导入分析模块
sys.path.insert(0, os.path.join(BASE_DIR, 'analysis'))
try:
    from analyzer import analyze_section, generate_docx_report
    ANALYZER_READY = True
except ImportError:
    ANALYZER_READY = False
    print('[WARN] analyzer not available')

app = Flask(__name__, static_folder=BASE_DIR, static_url_path='')


@app.route('/')
def index():
    """首页 -> 登录页"""
    return send_from_directory(BASE_DIR, 'login.html')


@app.route('/index.html')
def exam():
    """考试页"""
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
        for sub in ['answers/pretest', 'answers/correction', 'answers/posttest', 'reports']:
            os.makedirs(os.path.join(student_dir, sub), exist_ok=True)
        
        print(f'[LOGIN] {folder_name}')
        return jsonify({'success': True, 'folder': folder_name})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/submit/<section>', methods=['POST'])
def submit(section):
    """提交答题数据 + 实时分析"""
    try:
        data = request.get_json(force=True)
        data['server_timestamp'] = datetime.now().isoformat()
        
        # 确定学生文件夹
        student_name = data.get('studentName', '')
        student_id = data.get('studentId', '')
        folder_name = f'{student_name}_{student_id}' if student_name and student_id else '_anonymous'
        base = os.path.join(STUDENT_DATA_DIR, folder_name)
        answer_dir = os.path.join(base, 'answers', section)
        report_dir = os.path.join(base, 'reports')
        os.makedirs(answer_dir, exist_ok=True)
        os.makedirs(report_dir, exist_ok=True)
        
        # 保存答题数据
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'answer_{section}_{ts}.json'
        filepath = os.path.join(answer_dir, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print(f'[SAVED] {folder_name}/{section}: {filename}')
        
        # 实时分析
        report_info = None
        if ANALYZER_READY:
            try:
                student_answers = data.get('answers', {})
                analysis = analyze_section(section, student_answers)
                if analysis:
                    analysis['student_name'] = student_name
                    analysis['student_id'] = student_id
                    report_name = f'{section}_diagnosis_{ts}'
                    report_path = generate_docx_report(analysis, report_name, output_dir=report_dir)
                    if report_path:
                        report_info = {
                            'file': os.path.basename(report_path),
                            'score': analysis['total_score'],
                            'max_score': analysis['max_score'],
                            'ratio': f"{analysis['ratio']:.1f}%",
                            'level': analysis['level']
                        }
                        print(f'[REPORT] {folder_name}: {os.path.basename(report_path)} ({analysis["total_score"]}/{analysis["max_score"]})')
            except Exception as e:
                print(f'[ANALYZE ERROR] {traceback.format_exc()}')
        
        return jsonify({
            'success': True, 'section': section,
            'folder': folder_name, 'file': filename,
            'report': report_info
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ===== 数据同步接口（供本地 sync.py 调用） =====

@app.route('/api/sync/list_students')
def list_students():
    """列出所有学生及其答题数据"""
    students = []
    if os.path.exists(STUDENT_DATA_DIR):
        for folder in os.listdir(STUDENT_DATA_DIR):
            folder_path = os.path.join(STUDENT_DATA_DIR, folder)
            if os.path.isdir(folder_path):
                sections = []
                # 扫描 answers 目录
                answers_dir = os.path.join(folder_path, 'answers')
                if os.path.exists(answers_dir):
                    for sub in os.listdir(answers_dir):
                        sub_path = os.path.join(answers_dir, sub)
                        if os.path.isdir(sub_path):
                            for f in os.listdir(sub_path):
                                if f.endswith('.json'):
                                    sections.append(f'answers/{sub}/{f}')
                # 扫描 reports 目录
                reports_dir = os.path.join(folder_path, 'reports')
                if os.path.exists(reports_dir):
                    for f in os.listdir(reports_dir):
                        if f.endswith('.docx'):
                            sections.append(f'reports/{f}')
                
                # 尝试从文件夹名提取姓名学号
                parts = folder.rsplit('_', 1)
                name = parts[0] if len(parts) >= 2 else folder
                
                students.append({
                    'name': name,
                    'folder': folder,
                    'sections': sections
                })
    return jsonify({'students': students})


@app.route('/api/sync/download')
def download_file():
    """下载单个学生文件"""
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
    print('  电磁感应习题课 - 云后端服务')
    print(f'  数据目录: {STUDENT_DATA_DIR}')
    print(f'  分析引擎: {"已启用" if ANALYZER_READY else "未启用"}')
    print('=' * 50)
    app.run(host='0.0.0.0', port=5000, debug=False)
