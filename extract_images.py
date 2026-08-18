# -*- coding: utf-8 -*-
"""从题目库原卷版docx中提取题目配图，保存到 static/images/，并更新 questions.json 的 image 字段

定位方式：文件 + 题型号 + 题号（题型号与题号联合唯一确定题目，兼容第八章题号在题型间重复的情况）
"""
import os, re, json, shutil, docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

BASE = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE, 'static', 'images')

FILES = {
    '专题04': os.path.join(BASE, '题目库', '专题04 功、功率和机车的启动（期末专项训练） (原卷版).docx'),
    '专题06': os.path.join(BASE, '题目库', '专题06 机械能守恒及其应用（期末专项训练） (原卷版).docx'),
    '第八章': os.path.join(BASE, '题目库', '第八章 机械能守恒定律（举一反三重难点训练）（原卷版）.docx'),
}

# 题目 -> (文件tag, 题型号, 题号)
MAPPING = {
    ('pretest', 1): ('专题04', 1, 1),
    ('pretest', 2): ('专题04', 1, 3),
    ('pretest', 3): ('专题04', 5, 17),
    ('pretest', 4): ('专题06', 2, 5),
    ('pretest', 5): ('专题06', 2, 7),
    ('pretest', 6): ('专题06', 2, 6),
    ('pretest', 7): ('专题06', 1, 1),
    ('pretest', 8): ('第八章', 10, 37),
    ('posttest', 1): ('专题04', 1, 2),
    ('posttest', 2): ('第八章', 1, 6),
    ('posttest', 3): ('专题04', 5, 19),
    ('posttest', 4): ('第八章', 4, 15),
    ('posttest', 5): ('专题06', 3, 10),
    ('posttest', 6): ('第八章', 4, 16),
    ('posttest', 8): ('第八章', 10, 38),
}

TYPE_RE = re.compile(r'题型\s*(\d+)')


def current_type_no(txt):
    """从题型标题中提取题型号，如 '题型01：xxx' 或 '【题型4 标题】'"""
    m = TYPE_RE.search(txt)
    return int(m.group(1)) if m else None


def extract_question_images(doc, type_no, qnum):
    """提取 docx 中「题型 type_no 的第 qnum 题」的所有图片"""
    body = doc.element.body
    images = []
    cur_type = None
    collecting = False
    for child in body.iterchildren():
        if child.tag != qn('w:p'):
            continue
        txt = Paragraph(child, doc).text.strip()
        # 更新题型
        tno = current_type_no(txt)
        if tno is not None:
            cur_type = tno
            continue
        # 题号
        m = re.match(r'^(\d+)[．.](.*)', txt)
        if m:
            n = int(m.group(1))
            if cur_type == type_no and n == qnum:
                collecting = True
            elif collecting:
                # 已经越过了目标题
                break
        if not collecting:
            continue
        for blip in child.findall('.//' + qn('a:blip')):
            rId = blip.get(qn('r:embed'))
            if rId:
                part = doc.part.related_parts[rId]
                images.append((part.content_type, part.blob))
    return images


def main():
    # 清空旧的图片目录
    if os.path.exists(IMG_DIR):
        shutil.rmtree(IMG_DIR)
    os.makedirs(IMG_DIR, exist_ok=True)

    docs = {}
    result = {}

    for (section, num), (tag, tno, qnum) in MAPPING.items():
        if tag not in docs:
            docs[tag] = docx.Document(FILES[tag])
        imgs = extract_question_images(docs[tag], tno, qnum)
        key = f'{section}_{num}'
        result[(section, num)] = []
        if not imgs:
            print(f'⚠ {key}: 未提取到图片（{tag} 题型{tno} 第{qnum}题）')
            continue
        for i, (ctype, blob) in enumerate(imgs):
            ext = 'png' if 'png' in ctype else ('jpg' if 'jpeg' in ctype else 'bin')
            fname = f'{key}.{ext}' if len(imgs) == 1 else f'{key}_{i + 1}.{ext}'
            with open(os.path.join(IMG_DIR, fname), 'wb') as f:
                f.write(blob)
            result[(section, num)].append(fname)
        print(f'✓ {key}: {len(imgs)} 张 -> {", ".join(result[(section, num)])}')

    # 更新 questions.json
    qpath = os.path.join(BASE, 'test_data', 'questions.json')
    with open(qpath, 'r', encoding='utf-8') as f:
        qdata = json.load(f)

    updated = 0
    for section in ['pretest', 'posttest']:
        for q in qdata[section]:
            key = (section, q['num'])
            if key in result and result[key]:
                names = result[key]
                q['image'] = names[0] if len(names) == 1 else names
                updated += 1

    with open(qpath, 'w', encoding='utf-8') as f:
        json.dump(qdata, f, ensure_ascii=False, indent=2)

    print(f'\n已更新 {updated} 道题的 image 字段')
    print('图片总数:', sum(len(v) for v in result.values()))


if __name__ == '__main__':
    main()
