# -*- coding: utf-8 -*-
"""提取 docx 文本，输出 UTF-8 文本文件"""
import zipfile, re, sys, html

def extract(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8', errors='replace')
    xml = re.sub(r'</w:p>', '\n', xml)          # 段落在换行
    xml = re.sub(r'</w:tr>', '\n', xml)          # 表格行换行
    xml = re.sub(r'</w:tc>', ' | ', xml)         # 单元格分隔
    text = re.sub(r'<[^>]+>', '', xml)
    text = html.unescape(text)
    # 压缩多余空行
    lines = [l.rstrip() for l in text.split('\n')]
    lines = [l for l in lines if l.strip()]
    return '\n'.join(lines)

if __name__ == '__main__':
    out_path = sys.argv[1]
    in_paths = sys.argv[2:]
    parts = []
    for p in in_paths:
        parts.append('=' * 30)
        parts.append('FILE: ' + p)
        parts.append('=' * 30)
        parts.append(extract(p))
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(parts))
    print('OK ->', out_path)
