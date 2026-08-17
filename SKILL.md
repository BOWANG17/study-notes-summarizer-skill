---
name: study-notes-summarizer
description: "This skill turns a folder of mixed-format study notes (Word docx or doc, PDF text or scanned, images, PowerPoint pptx) into organized, per-month summary Word documents. Notes are classified into six smart sections: vocabulary, grammar, listening, speaking, reading, writing. A section is omitted when the source notes contain nothing for it. Use this skill when a user wants to consolidate many scattered notes into clean revision summaries, especially before an exam. Subject-agnostic (language, exam prep, coursework). All parsing uses free local tools; no paid API required."
agent_created: true
---

# Study Notes Summarizer

## Overview

把一堆零散的学习笔记（Word / PDF / 图片 / PPT，可能有几十份）整理成**按月归档、按板块分类**
的考前总结 Word 文档。整条流水线只用免费本地工具，不依赖任何付费 API 或外部连接器。

- 输入：一个本地文件夹，里面是用户的笔记原文件（`.docx` / `.doc` / `.pdf` / 图片 / `.pptx`）。
- 处理：① 统一解析成 Markdown → ② 按月份（`x.xx` 文件名）聚合 → ③ 按六板块智能分类 →
  ④ 用 tencent-docx 渲染成排版好的 `.docx`。
- 输出：`final/{月份}笔记总结.docx`，以及无日期素材的 `final/{科目}基础素材提炼.docx`。

## When To Use

- 用户说"把我的笔记整理成总结 / 考前复习文档 / 月度总结"。
- 用户有一堆 `.docx`/`.pdf`/图片/PPT 笔记，想合并、归类、去重。
- 用户备考、复习，需要按词汇/语法/听说读写组织材料。
- 触发关键词：笔记总结、复习资料、考前整理、月度归档、把文档归类。

## Pipeline

### Step 0 — 确认参数
向用户确认（或读取记忆）：① 笔记源文件夹路径；② 科目名（如"德语 B1"）；
③ 是否已有无日期素材的归宿偏好。源文件夹**只读取/解析，绝不移动、重命名、删除原文件**。

### Step 1 — 解析（scripts/parse_notes.py）
运行统一解析脚本，把源文件夹所有支持的格式转成 Markdown：
```bash
python3 scripts/parse_notes.py --source "源文件夹" --out "工作区/parsed" --log "工作区/parsed/processed.log"
```
- 脚本按扩展名分流：`.docx`→python-docx，`.doc`→macOS textutil，`.pdf`文本→pdfplumber，
  `.pdf`扫描/图片→tesseract OCR，`.pptx`→python-pptx，图片→OCR（可选另存 PDF）。
- 用 `processed.log` 记录已处理文件，**重跑不会重复解析**；缺失某引擎时跳过并提示，装好后可重试。
- 超大文件自动分块为 `_part1.md / _part2.md`。

### Step 2 — 阅读与分类
读取 `parsed/*.md`，按 `references/section_guide.md` 的六板块规则判断每份笔记命中的板块，
并标记原文中的错误用于 ⚠️ 易错标注。

### Step 3 — 生成总结（references/summary_prompt.md）
按 `references/summary_prompt.md` 的模板：
- 把同月份（`x.xx` → 对应月）的文件合并为一份 Markdown 总结。
- 按"单词词汇 → 语法 → 听 → 说 → 读 → 写"组织，**无内容的板块整段省略**。
- 抓原文错误做 ⚠️ 易错标注。
- 无日期的通用素材单独成《基础素材提炼》文档（除非用户指定编入某月）。

### Step 4 — 渲染 Word（tencent-docx）
把 Step 3 的 Markdown 交给内置 **tencent-docx** skill（tdoc-orchestrator），
走 S1 创作 → S2 美化 → S3 转 docx，输出 `final/{月份}笔记总结.docx`。
词汇表用表格、⚠️易错用底纹卡片。

### Step 5 — 记录与交付
- 更新项目的 `processed.log`（脚本自动完成）与记忆（哪些月已生成、哪些待补）。
- 把生成的 `.docx` 用 present_files 展示给用户。

## Parameters

| 参数 | 说明 | 默认 |
|---|---|---|
| 源文件夹 | 用户笔记原文件所在目录（绝对路径） | 必填 |
| 科目名 | 如"德语 B1"，用于标题与提炼文档命名 | 必填 |
| 工作区 | parsed/ summaries/ final/ 存放处 | 源文件夹同级或指定 |
| 无日期素材归宿 | 单独提炼 / 编入某月 / 两者都要 | 默认单独提炼 |

## Prerequisites（全部免费）

脚本所需的 Python 库（装到运行 Python 的环境即可）：
```bash
pip install python-docx pdfplumber python-pptx img2pdf pytesseract
```
- `.doc` / 老 `.ppt` 解析依赖 **macOS 自带 textutil**（仅 macOS；其他平台请先转成 .docx/.pptx）。
- **图片 / 扫描 PDF 的 OCR** 需要系统安装 `tesseract` + `poppler`：
  `brew install tesseract poppler`（无 brew 时见 tesseract 官网下载安装包）。
  缺此二者时，图片/扫描 PDF 会被跳过并提示，其余格式不受影响。
- 最终 `.docx` 渲染依赖内置 **tencent-docx** skill（始终可用）。

若某库缺失，脚本运行时会打印明确的 `pip install` 提示，不会静默失败。

## Notes

- **只读取，不改动原文件**：解析、总结全程在副本/派生文件上进行，桌面/源文件夹原文件保持原样。
- **增量友好**：`processed.log` 保证"扔一篇理解一篇、断点续跑"，适合长期累积的笔记。
- **科目无关**：德语 B1 只是示例；换成英语、考研、任何科目，只需改科目名与板块权重。
- **扫描/手写图片质量**决定 OCR 效果；清晰打印体最佳，潦草手写可能需人工校对。
