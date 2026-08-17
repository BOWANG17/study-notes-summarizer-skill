#!/usr/bin/env python3
"""
study-notes-summarizer — 统一免费解析脚本
==========================================
把学习笔记文件夹里的多种格式，全部解析成 Markdown 文本，落到 parsed/ 目录，
并写 processed.log 防止重复处理。全程使用免费本地工具，不依赖任何付费 API。

支持的格式与对应免费引擎：
  .docx            -> python-docx
  .doc (老格式)    -> macOS 自带 textutil  (仅 macOS；其他平台给出提示)
  .pdf (文本型)    -> pdfplumber
  .pdf (扫描型)    -> pdftoppm(poppler) + tesseract OCR
  .png/.jpg/...    -> tesseract OCR  (可选同时用 img2pdf 另存一份 PDF)
  .pptx            -> python-pptx
  .ppt (老格式)    -> 尝试 macOS textutil；失败则提示

缺失某个引擎时，脚本会打印清晰的安装提示并跳过该文件（不写进 processed.log，
下次工具装好后会自动重试），不会静默失败。

用法：
  python3 parse_notes.py --source "/path/to/笔记" --out "/path/to/parsed" [--log processed.log] [--force]
"""
import argparse
import os
import sys
import subprocess
import shutil
from pathlib import Path

# ---------- 工具可用性探测 ----------
def has(cmd):
    return shutil.which(cmd) is not None

def py_has(mod):
    try:
        __import__(mod)
        return True
    except Exception:
        return False

TEXTUTIL = has("textutil")
PDFPLUMBER = py_has("pdfplumber")
PYTHON_PPTX = py_has("pptx")
IMG2PDF = py_has("img2pdf")
PYTESSERACT = py_has("pytesseract")
TESSERACT_BIN = has("tesseract")
PDFTOPPM_BIN = has("pdftoppm")   # poppler 的一部分，扫描PDF转图用

OCR_OK = TESSERACT_BIN and PYTESSERACT  # 图片/扫描PDF OCR 可用

# ---------- 分发表 ----------
EXT_DISPATCH = {
    ".docx": "parse_docx",
    ".doc": "parse_doc",
    ".pdf": "parse_pdf",
    ".pptx": "parse_pptx",
    ".ppt": "parse_ppt",
}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp"}

SUPPORTED = set(EXT_DISPATCH) | IMAGE_EXTS


# ---------- 各格式解析器（懒加载，缺失即提示） ----------
def parse_docx(path: Path):
    import docx
    d = docx.Document(str(path))
    lines = []
    for p in d.paragraphs:
        if p.text.strip():
            lines.append(p.text)
    for ti, table in enumerate(d.tables, 1):
        lines.append(f"\n[表格 {ti}]")
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            lines.append(" | ".join(cells))
    return "\n".join(lines), "docx"


def parse_doc(path: Path, out_dir: Path):
    if not TEXTUTIL:
        raise RuntimeError(
            "解析 .doc 需要 macOS 自带 textutil。非 macOS 环境请先把 .doc 另存为 .docx 再投喂。"
        )
    tmp = out_dir / f"._{path.stem}.txt"
    subprocess.run(["textutil", "-convert", "txt", "-output", str(tmp), str(path)],
                   check=True, capture_output=True)
    text = tmp.read_text(encoding="utf-8", errors="ignore")
    tmp.unlink(missing_ok=True)
    return text, "doc"


def parse_pdf(path: Path, out_dir: Path):
    import pdfplumber
    texts = []
    with pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            t = page.extract_text() or ""
            if t.strip():
                texts.append(t)
    joined = "\n\n".join(texts)
    # 文本极少 -> 疑似扫描件，尝试 OCR
    if len(joined.strip()) < 60:
        if OCR_OK and PDFTOPPM_BIN:
            ocr = _ocr_pdf(path, out_dir)
            if ocr.strip():
                return ocr, "pdf-ocr"
            return joined, "pdf-scan-empty"
        else:
            missing = []
            if not PDFTOPPM_BIN:
                missing.append("poppler (pdftoppm)")
            if not TESSERACT_BIN:
                missing.append("tesseract")
            if not PYTESSERACT:
                missing.append("pytesseract(python 库)")
            raise RuntimeError(
                f"PDF「{path.name}」疑似扫描件，需要 OCR 但缺少：{', '.join(missing)}。"
                f" 请安装：brew install tesseract poppler && pip install pytesseract"
            )
    return joined, "pdf-text"


def _ocr_pdf(path: Path, out_dir: Path):
    import pytesseract
    from pdf2image import convert_from_path
    pages = convert_from_path(str(path), dpi=200)
    out = []
    for i, img in enumerate(pages, 1):
        out.append(f"[第 {i} 页]\n" + pytesseract.image_to_string(img, lang="de+eng"))
    return "\n\n".join(out)


def _ocr_image(path: Path):
    import pytesseract
    from PIL import Image
    img = Image.open(str(path))
    return pytesseract.image_to_string(img, lang="de+eng")


def parse_pptx(path: Path):
    import pptx
    prs = pptx.Presentation(str(path))
    blocks = []
    for si, slide in enumerate(prs.slides, 1):
        blocks.append(f"\n[幻灯片 {si}]")
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    blocks.append(txt)
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                blocks.append(f"(备注) {notes}")
    return "\n".join(blocks), "pptx"


def parse_ppt(path: Path, out_dir: Path):
    if not TEXTUTIL:
        raise RuntimeError(
            "解析老格式 .ppt 需要 macOS textutil。其他平台请先转成 .pptx 再投喂。"
        )
    tmp = out_dir / f"._{path.stem}.txt"
    subprocess.run(["textutil", "-convert", "txt", "-output", str(tmp), str(path)],
                   check=True, capture_output=True)
    text = tmp.read_text(encoding="utf-8", errors="ignore")
    tmp.unlink(missing_ok=True)
    return text, "ppt"


def parse_image(path: Path, out_dir: Path):
    if not OCR_OK:
        missing = []
        if not TESSERACT_BIN:
            missing.append("tesseract")
        if not PYTESSERACT:
            missing.append("pytesseract(python 库)")
        raise RuntimeError(
            f"图片「{path.name}」需要 OCR 但缺少：{', '.join(missing)}。"
            f" 请安装：brew install tesseract && pip install pytesseract"
        )
    text = _ocr_image(path)
    # 可选：另存一份 PDF（满足“图片转 PDF”需求）
    pdf_path = None
    if IMG2PDF:
        try:
            import img2pdf
            pdf_path = out_dir / f"{path.stem}.pdf"
            with open(pdf_path, "wb") as f:
                f.write(img2pdf.convert(str(path)))
        except Exception:
            pdf_path = None
    return text, ("image-ocr", pdf_path)


# ---------- 主流程 ----------
def write_markdown(out_dir: Path, stem: str, raw_text: str, meta: dict, chunk=40000):
    out_dir.mkdir(parents=True, exist_ok=True)
    header = (
        f"# {meta['title']}\n\n"
        f"> 来源: `{meta['source']}`\n"
        f"> 格式: {meta['format']}\n"
        f"> 解析引擎: {meta['engine']}\n\n---\n\n"
    )
    text = raw_text.strip()
    if not text:
        text = "（未能从该文件提取到文本内容）"
    # 超大文件分块
    if len(text) <= chunk:
        (out_dir / f"{stem}.md").write_text(header + text, encoding="utf-8")
        return [f"{stem}.md"]
    parts = []
    for i in range(0, len(text), chunk):
        part = text[i:i+chunk]
        fn = f"{stem}_part{i//chunk+1}.md"
        (out_dir / fn).write_text(header + part, encoding="utf-8")
        parts.append(fn)
    return parts


def main():
    ap = argparse.ArgumentParser(description="统一免费解析学习笔记为 Markdown")
    ap.add_argument("--source", required=True, help="笔记源文件夹")
    ap.add_argument("--out", default=None, help="解析输出目录（默认 <source>/../parsed）")
    ap.add_argument("--log", default=None, help="processed.log 路径（默认 <out>/processed.log）")
    ap.add_argument("--force", action="store_true", help="忽略 processed.log，强制重新解析全部")
    args = ap.parse_args()

    src = Path(args.source).expanduser().resolve()
    if not src.is_dir():
        print(f"❌ 源文件夹不存在: {src}", file=sys.stderr)
        sys.exit(1)

    out = Path(args.out).expanduser().resolve() if args.out else (src.parent / "parsed")
    out.mkdir(parents=True, exist_ok=True)
    log = Path(args.log).expanduser().resolve() if args.log else (out / "processed.log")

    done = set()
    if log.exists() and not args.force:
        for line in log.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                done.add(line)

    files = sorted([p for p in src.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED])
    print(f"源文件夹: {src}")
    print(f"输出目录: {out}")
    print(f"待处理文件数: {len(files)}\n")

    parsed, skipped = [], []
    for p in files:
        ext = p.suffix.lower()
        if p.name in done and not args.force:
            continue
        try:
            if ext in IMAGE_EXTS:
                text, (fmt, pdf_path) = parse_image(p, out)
                engine = "tesseract OCR"
                extra = f" (另存PDF: {pdf_path.name})" if pdf_path else ""
            elif ext == ".docx":
                text, fmt = parse_docx(p)
                engine = "python-docx"
            elif ext == ".doc":
                text, fmt = parse_doc(p, out)
                engine = "macOS textutil"
            elif ext == ".pdf":
                text, fmt = parse_pdf(p, out)
                engine = "pdfplumber" if fmt == "pdf-text" else "tesseract OCR"
            elif ext == ".pptx":
                text, fmt = parse_pptx(p)
                engine = "python-pptx"
            elif ext == ".ppt":
                text, fmt = parse_ppt(p, out)
                engine = "macOS textutil"
            else:
                skipped.append((p.name, "不支持的扩展名"))
                continue

            parts = write_markdown(
                out, p.stem, text,
                {"title": p.name, "source": str(p), "format": ext, "engine": engine}
            )
            parsed.append((p.name, parts, fmt))
            with log.open("a", encoding="utf-8") as f:
                f.write(p.name + "\n")
            print(f"✅ {p.name}  ->  {', '.join(parts)}  [{fmt}]")
        except Exception as e:
            skipped.append((p.name, str(e)))
            print(f"⚠️  {p.name}  跳过: {e}", file=sys.stderr)

    print(f"\n=== 完成 ===")
    print(f"本次解析: {len(parsed)} 个；跳过: {len(skipped)} 个")
    if skipped:
        print("跳过的文件（修好环境后重跑即可自动重试）：")
        for n, r in skipped:
            print(f"  - {n}: {r}")


if __name__ == "__main__":
    main()
