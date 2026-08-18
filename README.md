# Study Notes Summarizer

> 一个把多种格式学习笔记（Word / PDF / 图片 / PPT）整理成**按月归档、按板块分类**的考前总结 Word 文档的 WorkBuddy Skill。全程使用免费本地工具，**零付费 API、零外部连接器**。

*A WorkBuddy skill that turns a folder of mixed-format study notes (Word / PDF / images / PPT) into organized, per-month, section-classified revision-summary Word documents — using only free local tools, no paid API.*

---

## 特性 (Features)

- **多格式解析**：`.docx` / `.doc` / `.pdf`（文本型与扫描型）/ `.pptx` / 图片，统一转成 Markdown。
- **全免费本地引擎**：`python-docx`、`macOS textutil`、`pdfplumber`、`python-pptx`、`tesseract OCR`、`img2pdf`——不联网、不按量计费。
- **自动出 Word，不依赖 WorkBuddy**：`scripts/build_docx.py` 用 python-docx 把总结自动渲染成 `.docx`，AI 跑完分类直接调用，**用户零命令**；任何能跑 Python 的 AI 工具（WorkBuddy / Cursor / Claude Code 等）都能用。
- **按月聚合**：文件名带 `x.xx`（如 `7.14`、`8.16`）的笔记自动归入对应月份总结。
- **六板块智能分类**：`单词词汇 / 语法 / 听 / 说 / 读 / 写`，某板块源材料没有内容就**整段省略**。
- **⚠️ 易错标注**：自动抓取原文中的错误，集中标出供考前避坑。
- **增量友好**：`processed.log` 记录已处理文件，重跑不会重复解析，断点续跑。
- **跨平台容错**：缺某个引擎时优雅跳过并提示安装命令，不会静默失败。
- **科目无关**：德语 B1 只是示例，换成英语、考研、任何科目只需改科目名。

## 目录结构

```
study-notes-summarizer/
├── SKILL.md              # Skill 定义：触发条件、完整流水线、参数说明
├── scripts/
│   ├── parse_notes.py    # 统一免费解析脚本（多格式 → Markdown）
│   └── build_docx.py     # Markdown 总结 → Word 渲染器（python-docx，自动出 .docx）
├── references/
│   ├── section_guide.md  # 六板块分类规则与智能解析约定
│   └── summary_prompt.md # 总结生成提示词模板
├── requirements.txt      # Python 依赖一览
├── LICENSE               # MIT 许可证
└── README.md             # 本文件
```

## 环境要求 (Prerequisites)

解析脚本所需的 Python 库（装到运行 Python 的环境即可）：

```bash
pip install python-docx pdfplumber python-pptx img2pdf pytesseract
# 或一键装本仓库列出的依赖：
pip install -r requirements.txt
```

- **`.doc` / 老 `.ppt` 解析**依赖 macOS 自带 `textutil`（仅 macOS；其他平台请先转成 `.docx` / `.pptx`）。
- **图片 / 扫描 PDF 的 OCR** 需要系统安装 `tesseract` + `poppler`：
  - macOS：`brew install tesseract poppler`
  - 无 brew 时见 [tesseract 官网](https://github.com/tesseract-ocr/tesseract) 下载安装包
  - 缺这两者时，图片/扫描 PDF 会被跳过并提示，其余格式不受影响。
- 最终 `.docx` 渲染用 `scripts/build_docx.py`（python-docx），**不依赖 WorkBuddy 内置 skill**，任何能跑 Python 的 AI 工具都能自动出 Word。

若某库缺失，脚本运行时会打印明确的 `pip install` 提示，不会静默失败。

## 安装 (Installation)

作为 WorkBuddy 自定义 skill 安装：

```bash
# 方式一：直接把仓库放到 skills 目录
git clone <本仓库地址> ~/.workbuddy/skills/study-notes-summarizer

# 方式二：手动复制
cp -r study-notes-summarizer ~/.workbuddy/skills/
```

安装后在对话中说"用 study-notes-summarizer 整理我的笔记"即可触发。

## 使用 (Usage)

**最简流程**（让 AI 接管）：

1. 把笔记原文件放进任意一个本地文件夹（例如 `~/Desktop/我的笔记`）。
2. 对 WorkBuddy 说：*"用 study-notes-summarizer 整理 `~/Desktop/我的笔记`，科目是德语 B1"*。
3. Skill 会自动：解析 → 按月聚合 → 六板块分类 → 用 `build_docx.py` 自动生成 `final/{月份}笔记总结.docx`（用户零命令）。

**只想跑解析脚本**（命令行）：

```bash
python3 scripts/parse_notes.py \
  --source "/path/to/笔记" \
  --out   "/path/to/笔记/../parsed" \
  --log   "/path/to/笔记/../parsed/processed.log"
```

参数说明：

| 参数 | 说明 | 默认 |
|---|---|---|
| `--source` | 笔记源文件夹（必填） | — |
| `--out` | 解析输出目录 | `<source>/../parsed` |
| `--log` | `processed.log` 路径 | `<out>/processed.log` |
| `--force` | 忽略 `processed.log`，强制重解析全部 | 关 |

## 在其他 AI 工具里用（无需 WorkBuddy）

本 skill 的 Word 生成已改用 `scripts/build_docx.py`（python-docx），**不依赖 WorkBuddy 内置的 tencent-docx**。
因此在任何能读文件 + 跑命令的 AI 工具里都能跑通整条流水线：

1. 把本仓库放到该工具能访问的目录（WorkBuddy 是 `~/.workbuddy/skills/`，其他工具随意）。
2. 让 AI 读 `SKILL.md` + `references/*.md` 了解流程与提示词。
3. AI 执行：`parse_notes.py` 解析 → 按 `summary_prompt.md` 做六板块分类（产出 Markdown）→ `build_docx.py` 自动出 `.docx`。

支持的工具示例：**WorkBuddy、Cursor、Claude Code、Cline、aider** 等任何带终端执行能力的 AI 助手。
（纯网页版 ChatGPT/Claude 无法跑本地脚本，需配合上述带终端的工具。）

## 参数 / 配置

| 参数 | 说明 | 默认 |
|---|---|---|
| 源文件夹 | 用户笔记原文件所在目录 | 必填 |
| 科目名 | 如"德语 B1"，用于标题与提炼文档命名 | 必填 |
| 工作区 | `parsed/` `summaries/` `final/` 存放处 | 源文件夹同级或指定 |
| 无日期素材归宿 | 单独提炼 / 编入某月 / 两者都要 | 默认单独提炼 |

## 工作原理 (How it works)

```
投喂文件夹 → parse_notes.py 解析为 Markdown
           → AI 按月份(x.xx)聚合 + 六板块分类
           → build_docx.py 自动渲染为 .docx（python-docx，不依赖 WorkBuddy）
           → 每天可由自动化定时重跑（增量）
```

详见 `SKILL.md` 的 Pipeline 章节。

## 注意事项 / 局限 (Notes)

- **只读取、不改动原文件**：解析与总结全程在派生文件上进行，源文件夹原文件保持原样。
- **`.doc` / `.ppt` 老格式**依赖 macOS `textutil`，非 macOS 需先转换为新格式。
- **OCR 默认语言为 `de+eng`**（德语+英语，契合德语学习场景）。若要其他语言，改 `scripts/parse_notes.py` 中 `_ocr_pdf` / `_ocr_image` 的 `lang` 参数即可（如 `chi_sim+eng`）。
- **扫描 / 手写图片质量**决定 OCR 效果；清晰打印体最佳，潦草手写可能需人工校对。

## 许可证 (License)

[MIT](./LICENSE) —— 可自由使用、修改、分发，包括商用。
