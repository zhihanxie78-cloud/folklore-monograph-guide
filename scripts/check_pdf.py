# -*- coding: utf-8 -*-
"""PDF 文本层检测——判断是文本版还是扫描版，决定提取路线还是知识化路线。

用法:
    PYTHONUTF8=1 python check_pdf.py "D:/路径/书名.pdf"
    PYTHONUTF8=1 python check_pdf.py "D:/路径/*.pdf"   # 支持 glob 通配（文件名带特殊字符时用）
"""
import sys, io, os, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import fitz  # PyMuPDF


def check_one(pdf_path):
    print("=" * 70)
    print(f"PDF: {pdf_path}")
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    print(f"页数: {total_pages}")

    # 统计每页文本层字符数
    chars_per_page = []
    for i in range(total_pages):
        chars_per_page.append(len(doc[i].get_text().strip()))
    doc.close()

    text_pages = sum(1 for c in chars_per_page if c > 20)
    total_chars = sum(chars_per_page)
    print(f"含文本层的页数(>20字符/页): {text_pages}/{total_pages}")
    print(f"全书文本层总字符数: {total_chars}")

    if text_pages >= total_pages * 0.8 and total_chars > 5000:
        verdict = "TEXT 有文本层 → 可用 PyMuPDF 逐章提取，证据索引可定位到页"
    elif text_pages == 0:
        verdict = "SCAN 纯扫描版（无文本层）→ 转知识化处理，README 声明"
    else:
        verdict = "MIXED 部分文本层 → 混合处理：有文本的章节提取，其余知识化"

    print(f"判定: {verdict}")

    # 前 3 个有文本的页面样本
    shown = 0
    doc = fitz.open(pdf_path)
    for i in range(total_pages):
        if shown >= 3:
            break
        t = doc[i].get_text().strip()
        if t:
            print(f"\n--- 第 {i + 1} 页样本 ---")
            print(t[:300])
            shown += 1
    doc.close()
    print("=" * 70)
    print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    pattern = sys.argv[1]
    if any(ch in pattern for ch in "*?["):
        files = sorted(glob.glob(pattern))
    else:
        files = [pattern]
    if not files:
        print(f"未找到匹配文件: {pattern}")
        sys.exit(1)
    for f in files:
        check_one(f)
