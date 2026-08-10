# -*- coding: utf-8 -*-
"""云端数据同步脚本 —— 从 PythonAnywhere 下载学生数据到本地开发空间"""
import os, sys, json, urllib.request, urllib.error

# 配置
CLOUD_URL = 'https://Fanru.pythonanywhere.com'
LOCAL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'student_data')
SECRET = 'physics-sync-2026'

def sync():
    """从云端下载所有学生数据"""
    os.makedirs(LOCAL_DIR, exist_ok=True)
    
    try:
        # 1. 获取学生列表
        req = urllib.request.Request(f'{CLOUD_URL}/api/sync/list_students')
        resp = urllib.request.urlopen(req, timeout=10)
        data = json.loads(resp.read().decode('utf-8'))
        students = data.get('students', [])
        
        print(f'找到 {len(students)} 个学生:')
        for s in students:
            print(f'  {s["name"]} ({s["folder"]}) - {s["sections"]} 份答题')
        
        if not students:
            print('云端暂无学生数据。')
            return
        
        # 2. 下载每个学生的数据
        for s in students:
            print(f'\n正在同步: {s["folder"]}...')
            
            # 创建本地文件夹
            local_student = os.path.join(LOCAL_DIR, s['folder'])
            os.makedirs(local_student, exist_ok=True)
            
            for section_name in s.get('sections', []):
                parts = section_name.split('/')
                if len(parts) >= 2:
                    sec_type = parts[0]  # answers 或 reports
                    sec_sub = parts[1]   # pretest/correction/posttest 或 报告文件名
                    
                    if sec_type == 'answers':
                        dest_dir = os.path.join(local_student, 'answers', sec_sub)
                        os.makedirs(dest_dir, exist_ok=True)
                        # 下载答题数据
                        try:
                            url = f'{CLOUD_URL}/api/sync/download?folder={s["folder"]}&section={section_name}'
                            req2 = urllib.request.Request(url)
                            resp2 = urllib.request.urlopen(req2, timeout=15)
                            
                            # 提取文件名
                            content_disposition = resp2.headers.get('Content-Disposition', '')
                            filename = section_name.split('/')[-1] + '.json'
                            if 'filename=' in content_disposition:
                                filename = content_disposition.split('filename=')[-1].strip('"')
                            
                            filepath = os.path.join(dest_dir, filename)
                            with open(filepath, 'wb') as f:
                                f.write(resp2.read())
                            print(f'  ✓ 下载: {section_name} → {filename}')
                        except urllib.error.HTTPError as e:
                            print(f'  ✗ 下载失败: {section_name} ({e.code})')
                    
                    elif sec_type == 'reports':
                        dest_dir = os.path.join(local_student, 'reports')
                        os.makedirs(dest_dir, exist_ok=True)
                        try:
                            url = f'{CLOUD_URL}/api/sync/download?folder={s["folder"]}&section={section_name}'
                            req2 = urllib.request.Request(url)
                            resp2 = urllib.request.urlopen(req2, timeout=15)
                            
                            filename = section_name.split('/')[-1]
                            filepath = os.path.join(dest_dir, filename)
                            with open(filepath, 'wb') as f:
                                f.write(resp2.read())
                            print(f'  ✓ 下载: {section_name} → {filename}')
                        except urllib.error.HTTPError as e:
                            print(f'  ✗ 下载失败: {section_name} ({e.code})')
        
        print(f'\n同步完成！数据已保存至: {LOCAL_DIR}')
        
    except urllib.error.URLError as e:
        print(f'无法连接云端服务器: {e}')
        print('请确认 PythonAnywhere 服务正在运行。')
    except Exception as e:
        print(f'同步出错: {e}')
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    sync()
