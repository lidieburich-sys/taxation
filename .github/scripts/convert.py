"""Convert user-supplied TX decks on an ephemeral GitHub Actions runner."""
from collections import Counter
from pathlib import Path
from urllib.parse import quote
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

import fitz
from PIL import Image, ImageDraw

ROOT = Path.cwd()
WORK = Path(tempfile.mkdtemp(prefix="tx-conversion-"))
QA = ROOT / "qa"
QA.mkdir(exist_ok=True)
MANIFEST = json.loads((ROOT / "metadata/source-manifest.json").read_text())
PREVIOUS_REPORT = ROOT / "metadata/conversion-report.json"
PREVIOUS_HASHES = ({r["pdf"]: r["sha256"] for r in json.loads(PREVIOUS_REPORT.read_text())["files"]}
                   if PREVIOUS_REPORT.exists() else {})
ASSETS = {hashlib.sha256(p.read_bytes()).hexdigest(): p
          for p in (ROOT / "source-assets").iterdir() if p.is_file()}


def normalized(text):
    return Counter(c for c in unicodedata.normalize("NFKC", text)
                   if c.isalnum())


def slide_texts(path):
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with zipfile.ZipFile(path) as z:
        rels = {e.attrib["Id"]: e.attrib["Target"] for e in ET.fromstring(
            z.read("ppt/_rels/presentation.xml.rels"))}
        texts = []
        for e in ET.fromstring(z.read("ppt/presentation.xml")).findall("p:sldIdLst/p:sldId", ns):
            target = rels[e.attrib[f"{{{ns['r']}}}id"]]
            part = target.lstrip("/") if target.startswith("/") else "ppt/" + target
            root = ET.fromstring(z.read(part))
            texts.append("".join(x.text or "" for x in root.findall(".//a:t", ns)))
        return texts


def convert(item):
    source, info, category = item
    out = WORK / ("out-" + info["sha256"][:12])
    out.mkdir()
    profile = WORK / ("profile-" + info["sha256"][:12])
    options = {"ExportHiddenSlides": {"type": "boolean", "value": "true"},
               "ExportNotesPages": {"type": "boolean", "value": "false"},
               "UseLosslessCompression": {"type": "boolean", "value": "true"},
               "ReduceImageResolution": {"type": "boolean", "value": "false"}}
    result = subprocess.run([
        "libreoffice", "-env:UserInstallation=" + profile.as_uri(),
        "--headless", "--convert-to", "pdf:impress_pdf_Export:" + json.dumps(options),
        "--outdir", str(out), str(source)], capture_output=True, text=True, timeout=240)
    pdf = out / (source.stem + ".pdf")
    if result.returncode or not pdf.exists():
        raise RuntimeError(f"Conversion failed: {source.name}\n{result.stdout}\n{result.stderr}")
    texts = slide_texts(source)
    doc = fitz.open(pdf)
    if len(doc) != info["slide_count"] or len(doc) != len(texts):
        raise RuntimeError(f"Page count mismatch: {source.name}: {len(doc)} / {info['slide_count']}")
    checks = []
    for index, page in enumerate(doc):
        expected = normalized(texts[index])
        actual = normalized(page.get_text())
        retained = sum((expected & actual).values()) / max(1, sum(expected.values()))
        checks.append({"page": index + 1, "text_retention": round(retained, 4),
                       "source_text_chars": sum(expected.values()),
                       "pdf_text_chars": sum(actual.values())})
    if any(p["source_text_chars"] > 30 and p["pdf_text_chars"] == 0 for p in checks):
        raise RuntimeError(f"Missing searchable text: {source.name}")
    selected = sorted({0, len(doc) - 1, max(range(len(texts)), key=lambda i: len(texts[i]))})
    index = re.search(r"\d+", source.stem).group()
    key = ("lecture" if category == "主课" else "answer") + f"-{int(index):02}"
    for page_index in selected:
        page = doc[page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(480 / page.rect.width, 480 / page.rect.width), alpha=False)
        pix.save(QA / f"{key}-p{page_index + 1:03}.png")
    for check in checks:
        if check["text_retention"] < 0.85 and check["source_text_chars"] > 30:
            page = doc[check["page"] - 1]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            pix.save(QA / f"review-{key}-p{check['page']:03}.png")
    doc.close()
    target = ROOT / pdf.name
    if target.exists() and hashlib.sha256(target.read_bytes()).hexdigest() != PREVIOUS_HASHES.get(target.name):
        raise RuntimeError(f"Refusing to overwrite different existing PDF: {target.name}")
    shutil.copy2(pdf, target)
    record = {"source": info["filename"], "pdf": target.name, "category": category,
              "pages": len(texts), "hidden_slides_included": info["hidden_slide_count"],
              "bytes": target.stat().st_size, "sha256": hashlib.sha256(target.read_bytes()).hexdigest(),
              "minimum_text_retention": min(p["text_retention"] for p in checks),
              "page_checks": checks}
    print(json.dumps({k: record[k] for k in ["pdf", "pages", "minimum_text_retention"]}, ensure_ascii=False), flush=True)
    shutil.rmtree(profile, ignore_errors=True)
    return record


items = []
for archive in MANIFEST["archives"]:
    asset = ASSETS[archive["sha256"]]
    category = "作业答案" if "答案" in archive["filename"] else "主课"
    with zipfile.ZipFile(asset) as z:
        for info in archive["pptx_files"]:
            data = z.read(info["filename"])
            assert hashlib.sha256(data).hexdigest() == info["sha256"]
            assert Path(info["filename"]).name == info["filename"]
            source = WORK / info["filename"]
            source.write_bytes(data)
            items.append((source, info, category))
assert len(items) == MANIFEST["expected_pptx_count"]
# PyMuPDF must not run from concurrent threads. Sequential conversions also keep
# the runner's font cache and memory usage predictable.
records = [convert(item) for item in items]
assert sum(r["pages"] for r in records) == MANIFEST["expected_slide_count"]
tax = MANIFEST["tax_table"]
shutil.copy2(ASSETS[tax["sha256"]], ROOT / tax["filename"])
tax_doc = fitz.open(ROOT / tax["filename"])
tax_pages = len(tax_doc)
tax_doc[0].get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False).save(QA / "tax-table-p001.png")
tax_doc.close()
report = {"converted_pdfs": len(records), "converted_pages": sum(r["pages"] for r in records),
          "tax_table_pages": tax_pages, "hidden_slides_included": sum(r["hidden_slides_included"] for r in records),
          "engine": subprocess.check_output(["libreoffice", "--version"], text=True).strip(),
          "files": records}
(ROOT / "metadata/conversion-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")

thumbs = sorted(QA.glob("lecture-*.png")) + sorted(QA.glob("answer-*.png"))
for start in range(0, len(thumbs), 12):
    sheet = Image.new("RGB", (1440, 4 * 300), "#dddddd")
    draw = ImageDraw.Draw(sheet)
    for offset, path in enumerate(thumbs[start:start + 12]):
        image = Image.open(path)
        image.thumbnail((480, 270))
        x, y = offset % 3 * 480, offset // 3 * 300
        sheet.paste(image, (x, y + 24))
        draw.text((x + 8, y + 6), path.stem, fill="black")
    sheet.save(QA / f"contact-{start // 12 + 1:02}.jpg", quality=90)

rows = ["# Taxation", "", "ACCA TX UK 课程资料。税率表适用考试期间：2026 年 6 月至 2027 年 3 月。", "",
        f"共 26 份主课 PDF、24 份作业答案 PDF（共 {report['converted_pages']:,} 页），另附原版考试税率表。", "",
        "## 主课", "", "| 课次 | 文件 | 页数 |", "| --- | --- | ---: |"]
for category in ["主课", "作业答案"]:
    if category != "主课":
        rows += ["", "## 作业答案", "", "| 课次 | 文件 | 页数 |", "| --- | --- | ---: |"]
    for record in sorted((r for r in records if r["category"] == category), key=lambda r: int(re.search(r"\d+", r["source"]).group())):
        number = int(re.search(r"\d+", record["source"]).group())
        rows.append(f"| {number} | [{record['pdf']}]({quote(record['pdf'])}) | {record['pages']} |")
rows += ["", "## 税率表", "", f"[{tax['filename']}]({quote(tax['filename'])})", "",
         "## 原始资料与转换", "", "[原始 ZIP 与税率表](https://github.com/lidieburich-sys/taxation/releases/tag/tx-source-2026)", "",
         "PDF 由 GitHub Actions 在远端使用 LibreOffice 转换，保留原文件名及幻灯片顺序。原稿中的 1 张隐藏幻灯片也已导出；动画在 PDF 中显示为静态页面。", "",
         "远端字体替代可能造成少量字形和间距差异，原始 PPTX 保留在压缩包中。", "",
         "[来源清单](metadata/source-manifest.json) · [转换页数与文件校验记录](metadata/conversion-report.json)", ""]
(ROOT / "README.md").write_text("\n".join(rows))
shutil.rmtree(WORK)
print(f"Complete: {len(records)} converted PDFs, {report['converted_pages']} slides; tax table {tax_pages} pages.")
