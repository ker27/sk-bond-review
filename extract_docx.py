#!/usr/bin/env python3
"""抽取 docx 全文：段落标记 [P#]，表格标记 ===TABLE#N=== 与 T#N R#M 行。

用法: python3 extract_docx.py <输入.docx> <输出.txt>
合并单元格：遍历 row._tr 下的 w:tc，避免 row.cells 的重复单元格问题。
"""
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


def main(path, out):
    doc = Document(path)
    lines = []
    pi = ti = 0
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            p = Paragraph(child, doc)
            text = p.text.strip()
            if text:
                pi += 1
                lines.append(f"[P{pi}] {text}")
        elif child.tag == qn('w:tbl'):
            t = Table(child, doc)
            ti += 1
            lines.append(f"===TABLE#{ti}===")
            for ri, row in enumerate(t.rows, 1):
                cells = []
                for tc in row._tr.findall(qn('w:tc')):
                    # 不能用 tc.itertext()：部分 docx 经 python-docx 加载后
                    # 元素 .text 会异常重复（实测文本被 3 倍复制），
                    # 只取 w:t 的原始文本最可靠。
                    text = ''.join(t.text or '' for t in tc.findall('.//' + qn('w:t')))
                    cells.append(text.strip().replace('\n', ' '))
                lines.append(f"T#{ti} R#{ri} | " + " | ".join(cells))
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"paragraphs={pi} tables={ti} -> {out}")


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
