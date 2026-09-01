#!/usr/bin/env python3
"""抽取 pdf 文本层；扫描件页面无文本层，会标记出来。

用法: python3 extract_pdf.py <输入.pdf> <输出.txt>
扫描件对策：找同名数字版 Excel（审计 PDF 的正表页常是扫描图，附注页有文本）；
正文数字一律以数字版 Excel 为准，PDF 附注（实收资本变动、收入构成、减值明细）是重要证据源。
"""
import sys
import pdfplumber


def main(path, out):
    lines = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            lines.append(f"\n##### PAGE {i} #####")
            t = page.extract_text()
            lines.append(t if t else "!!! 本页无文本层（扫描件）")
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"pages={len(pdf.pages)} -> {out}")


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
