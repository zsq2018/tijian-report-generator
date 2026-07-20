#!/usr/bin/env python3
import docx, re, os

WORK_DIR = "/Users/zhaoshuqing/Desktop/体检报告"
files = [f for f in os.listdir(WORK_DIR) if f.endswith('.docx') and f != '郑万军.docx']

for fname in sorted(files):
    path = os.path.join(WORK_DIR, fname)
    d = docx.Document(path)
    text = "\n".join([p.text for p in d.paragraphs[:80] if p.text.strip()])
    name = re.search(r'姓\s*名\s*[：:\t]?\s*(\S+)', text)
    idcard = re.search(r'(\d{17}[\dXx])', text)
    age = re.search(r'年\s*龄\s*[：:\t]?\s*(\d+)', text)
    phone = re.search(r'(1[3-9]\d{9})', text)
    number = re.search(r'体检[编代]码\s*[：:\t]?\s*(\d+)', text)
    print(f"--- {fname} ---")
    print(f"  姓名: {name.group(1) if name else 'N/A'}")
    print(f"  身份证: {idcard.group(1) if idcard else 'N/A'}")
    print(f"  年龄: {age.group(1) if age else 'N/A'}")
    print(f"  电话: {phone.group(1) if phone else 'N/A'}")
    print(f"  体检编码: {number.group(1) if number else 'N/A'}")
    print()
