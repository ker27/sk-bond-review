# ABS 募集说明书数据复核（abs-data-review）

对照源文件（审计报告、财报、现金流预测等）复核 ABS 募集说明书全文的工具包与技能：
**每一段文字、每一张表都要核，核心目标是不留勾稽错误、不留文字描述错误。**

## 内容

| 文件 | 说明 |
|---|---|
| `ABS数据复核技能手册.html` | 网页版复核工作台（单文件，双击即可打开）：工作流手册 + 可打勾核对清单 + 自动扫描（占位符 / 错别字 / 资产负债表·现金流量表勾稽 / 术语一致性），支持上传 .txt / .docx / .xlsx / .pdf（扫描件可选 OCR），可导出 Markdown 报告 |
| `skill/SKILL.md` | Claude Code 技能说明（完整核对方法论与关键坑） |
| `skill/scripts/` | 文档抽取脚本（docx / xlsx / pdf → 可检索文本） |

## 网页工作台怎么用

1. 双击打开 `ABS数据复核技能手册.html`（无需安装任何东西，数据不出本机）。
2. 粘贴说明书文本（建议先用 `skill/scripts/extract_docx.py` 抽取，输出格式可自动识别章节），或直接上传 .txt / .docx / .xlsx / .pdf。
3. 勾选扫描项目 → 运行扫描 → 得到按【高/中/低】分级、按章节定位的问题清单。
4. 复制 Markdown 报告 / 下载 .md / 一键生成 Claude 深核提示词，交给 Claude 做最终判定。

> 网页做机械检查（规则、勾稽计算）；"数字谁对谁错、是否审计重述未更新、跨文件对比"等深度判定请交给 Claude。

## 与 Claude Code 配合（推荐）

1. 安装技能：把 `skill/` 目录复制到 `~/.claude/skills/abs-data-review/`。
2. 安装依赖：`pip3 install python-docx openpyxl pdfplumber`。
3. 对话示例：
   - `/abs-data-review`
   - `根据文件夹里的各种审计报告，复核一下其他文件的财务数据`
   - `根据文件夹里的标准条款，核对一下说明书里有没有错误的地方`

## 报告格式约定

- 定位用「第几章第几节」，不用表号；每条引用原文短语（便于 Word 里 Ctrl+F）。
- 每条写清：说明书写的是什么 / 源文件是什么 / 哪个对、应改为多少 / 证据（源文件名）。
- 分级：**【高】**勾稽不平、数据矛盾、审计重述未更新、被引用的空表；**【中】**占位符、文字错误、术语/口径不一致；**【低】**标点、格式。
- 核对无误的章节明确写「未发现错误」；源文件自身问题单独报告。

## 本地抽取脚本

```bash
# 说明书：段落 [P#] + 表格 ===TABLE#N=== / T#N R#M 行
python3 skill/scripts/extract_docx.py "说明书.docx" /tmp/fuch/prospectus.txt
# Excel：每 sheet 输出 坐标=值（保留类型，防 0/空/字符串混淆）
python3 skill/scripts/extract_xlsx.py "财报.xlsx" /tmp/fuch/caiwu.txt
# PDF：pdfplumber 取文本层；扫描件页面输出"无文本层"
python3 skill/scripts/extract_pdf.py "审计报告.pdf" /tmp/fuch/pdf_audit.txt
```
