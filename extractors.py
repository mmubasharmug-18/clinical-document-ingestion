import hashlib
import io
import json
import os
import re
import shutil
import time
import zipfile
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path

import pytesseract
from bs4 import BeautifulSoup
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import ExifTags, Image, ImageOps, ImageSequence

try:
    import fitz
except ImportError:
    fitz = None

WINDOWS_TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
LIST_START = re.compile(r"^(?:[\u2022\-*]|\d{1,2}[.)])\s+")
IMAGE_MAGIC = (b"\x89PNG", b"\xff\xd8\xff", b"II*\x00", b"MM\x00*", b"GIF8")


def find_tesseract():
    cmd = os.environ.get("TESSERACT_CMD")
    if not cmd and not shutil.which("tesseract") and os.path.exists(WINDOWS_TESSERACT):
        cmd = WINDOWS_TESSERACT
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd
        return os.path.exists(cmd)
    return shutil.which("tesseract") is not None


TESSERACT_OK = find_tesseract()


@dataclass
class Options:
    ocr: bool = True
    dpi: int = 200
    lang: str = "eng"


@dataclass
class Block:
    kind: str
    text: str
    page: int = None
    level: int = None
    source: str = "text"
    confidence: float = None
    rows: list = None


@dataclass
class ExtractedDocument:
    source: dict
    doc_type: str
    method: str = ""
    metadata: dict = field(default_factory=dict)
    blocks: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    pages: int = None
    seconds: float = 0.0

    def text(self):
        return "\n\n".join(b.text for b in self.blocks)

    def page_text(self, number):
        return "\n\n".join(b.text for b in self.blocks if b.page == number)

    def stats(self):
        text = self.text()
        ocr = [b for b in self.blocks if b.source == "ocr"]
        confs = [b.confidence for b in ocr if b.confidence is not None]
        return {
            "pages": self.pages,
            "blocks": len(self.blocks),
            "words": len(text.split()),
            "characters": len(text),
            "ocr_pages": len({b.page for b in ocr}),
            "avg_ocr_confidence": round(sum(confs) / len(confs), 2) if confs else None,
            "seconds": round(self.seconds, 2),
        }

    def to_dict(self):
        blocks = [{k: v for k, v in asdict(b).items() if v is not None} for b in self.blocks]
        return {
            "source": self.source,
            "type": self.doc_type,
            "method": self.method,
            "metadata": self.metadata,
            "stats": self.stats(),
            "warnings": self.warnings,
            "blocks": blocks,
            "text": self.text(),
        }

    def to_json(self):
        return json.dumps(self.to_dict(), indent=2, ensure_ascii=False, default=str)


def tidy(text, keep_lines=False):
    text = text.replace("\u00a0", " ").replace("\u200b", "").replace("\r", "")
    text = re.sub(r"(\w)-\n([a-z])", r"\1\2", text)
    text = re.sub(r"[ \t]+", " ", text)
    if keep_lines:
        text = re.sub(r" ?\n ?", "\n", text)
    else:
        text = re.sub(r"\s*\n\s*", " ", text)
    return text.strip()


def table_text(rows):
    return "\n".join(" | ".join(r) for r in rows)


def detect_type(file_name, data):
    head = data[:8]
    ext = Path(file_name).suffix.lower()
    if head.startswith(b"%PDF"):
        return "pdf"
    if head.startswith(b"PK"):
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                if "word/document.xml" in z.namelist():
                    return "docx"
        except zipfile.BadZipFile:
            pass
    if head.startswith(IMAGE_MAGIC) or (data[:4] == b"RIFF" and data[8:12] == b"WEBP"):
        return "image"
    if ext == ".bmp" and head.startswith(b"BM"):
        return "image"
    sample = data[:4000].lower()
    if ext in {".html", ".htm"} or b"<html" in sample or b"<!doctype html" in sample:
        return "html"
    if b"\x00" not in data[:4000]:
        return "txt"
    return "unknown"


def prepare_for_ocr(img):
    img = ImageOps.autocontrast(img.convert("L"))
    if max(img.size) < 1500:
        img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    return img


def ocr_blocks(img, lang, page):
    data = pytesseract.image_to_data(
        prepare_for_ocr(img), lang=lang, output_type=pytesseract.Output.DICT
    )
    groups = {}
    for i, word in enumerate(data["text"]):
        word = word.strip()
        if not word:
            continue
        key = (data["block_num"][i], data["par_num"][i])
        group = groups.setdefault(key, {"words": [], "conf": []})
        group["words"].append(word)
        conf = float(data["conf"][i])
        if conf >= 0:
            group["conf"].append(conf)
    blocks = []
    for key in sorted(groups):
        group = groups[key]
        conf = None
        if group["conf"]:
            conf = round(sum(group["conf"]) / len(group["conf"]) / 100, 2)
        text = tidy(" ".join(group["words"]))
        kind = "list_item" if LIST_START.match(text) else "paragraph"
        blocks.append(Block(kind, text, page=page, source="ocr", confidence=conf))
    return blocks


def low_confidence_warning(blocks, label):
    confs = [b.confidence for b in blocks if b.confidence is not None]
    if confs and sum(confs) / len(confs) < 0.6:
        return f"{label}: OCR confidence is low, check the text by eye."
    return None


def pdf_date(value):
    m = re.match(r"D:(\d{4})(\d{2})(\d{2})(\d{2})?(\d{2})?(\d{2})?", value or "")
    if not m:
        return value
    y, mo, d, h, mi, s = (int(x) if x else 0 for x in m.groups())
    try:
        return datetime(y, mo, d, h, mi, s).isoformat()
    except ValueError:
        return value


def extract_pdf(data, opts):
    if fitz is None:
        raise RuntimeError("PyMuPDF is not installed")
    doc = fitz.open(stream=data, filetype="pdf")
    if doc.needs_pass:
        raise ValueError("PDF is password protected")
    meta = {k: v for k, v in (doc.metadata or {}).items() if v}
    for key in ("creationDate", "modDate"):
        if key in meta:
            meta[key] = pdf_date(meta[key])
    outline = doc.get_toc()
    if outline:
        meta["outline_entries"] = len(outline)

    counts = Counter()
    raw = []
    for page in doc:
        found = [b for b in page.get_text("dict")["blocks"] if b.get("type") == 0]
        raw.append((page, found))
        for b in found:
            for line in b["lines"]:
                for span in line["spans"]:
                    counts[round(span["size"], 1)] += len(span["text"].strip())
    body = counts.most_common(1)[0][0] if counts else 11

    blocks = []
    warnings = []
    for number, (page, found) in enumerate(raw, 1):
        page_blocks = []
        for b in found:
            lines, sizes, bold = [], [], []
            for line in b["lines"]:
                spans = [s for s in line["spans"] if s["text"].strip()]
                if not spans:
                    continue
                lines.append("".join(s["text"] for s in spans))
                sizes += [s["size"] for s in spans]
                bold += [bool(s["flags"] & 16) for s in spans]
            text = tidy("\n".join(lines))
            if not text:
                continue
            top = max(sizes)
            is_heading = len(lines) <= 2 and len(text) <= 120 and (
                top >= body * 1.15 or (all(bold) and len(text) <= 80)
            )
            if is_heading:
                level = 1 if top >= body * 1.6 else 2 if top >= body * 1.3 else 3
                page_blocks.append(Block("heading", text, page=number, level=level))
            else:
                kind = "list_item" if LIST_START.match(text) else "paragraph"
                page_blocks.append(Block(kind, text, page=number))

        if sum(len(b.text) for b in page_blocks) < 25:
            if not opts.ocr:
                warnings.append(f"Page {number} has almost no text. It may be scanned, turn on OCR.")
            elif not TESSERACT_OK:
                warnings.append(f"Page {number} looks scanned but Tesseract was not found.")
            else:
                try:
                    pix = page.get_pixmap(dpi=opts.dpi)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    page_blocks = ocr_blocks(img, opts.lang, number)
                    if not page_blocks:
                        warnings.append(f"Page {number}: OCR found no text.")
                    note = low_confidence_warning(page_blocks, f"Page {number}")
                    if note:
                        warnings.append(note)
                except Exception as exc:
                    warnings.append(f"Page {number}: OCR failed ({exc})")
        blocks += page_blocks
    pages = len(raw)
    doc.close()
    return meta, blocks, pages, warnings


def extract_docx(data, opts):
    doc = Document(io.BytesIO(data))
    blocks = []
    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            para = Paragraph(child, doc)
            text = tidy(para.text)
            if not text:
                continue
            style = (para.style.name if para.style is not None else "") or ""
            if style.startswith("Heading") or style == "Title":
                digits = re.findall(r"\d+", style)
                level = int(digits[0]) if digits else 1
                blocks.append(Block("heading", text, level=level))
            elif "List" in style or LIST_START.match(text):
                blocks.append(Block("list_item", text))
            else:
                blocks.append(Block("paragraph", text))
        elif tag == "tbl":
            table = Table(child, doc)
            rows = [[tidy(c.text) for c in row.cells] for row in table.rows]
            rows = [r for r in rows if any(r)]
            if rows:
                blocks.append(Block("table", table_text(rows), rows=rows))

    props = doc.core_properties
    meta = {}
    for name in ("title", "author", "subject", "keywords", "category", "last_modified_by", "revision", "created", "modified"):
        value = getattr(props, name, None)
        if value:
            meta[name] = value.isoformat() if isinstance(value, datetime) else value
    header_lines = []
    for section in doc.sections:
        header_lines += [tidy(p.text) for p in section.header.paragraphs if p.text.strip()]
    if header_lines:
        meta["header_text"] = " / ".join(dict.fromkeys(header_lines))

    warnings = []
    if len(doc.inline_shapes):
        warnings.append(f"{len(doc.inline_shapes)} embedded image(s) were not read.")
    return meta, blocks, None, warnings


def extract_html(data, opts):
    soup = BeautifulSoup(data, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    meta = {}
    if soup.title and soup.title.get_text(strip=True):
        meta["title"] = tidy(soup.title.get_text())
    if soup.html and soup.html.get("lang"):
        meta["language"] = soup.html.get("lang")
    for tag in soup.find_all("meta"):
        name = tag.get("name") or tag.get("property")
        if name and tag.get("content"):
            meta[name.lower()] = tag.get("content")
    link = soup.find("link", rel="canonical")
    if link and link.get("href"):
        meta["canonical_url"] = link.get("href")

    root = soup.body or soup
    blocks = []
    wanted = ["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "pre", "table"]
    for el in root.find_all(wanted):
        if el.name != "table" and el.find_parent("table"):
            continue
        if el.name == "p" and el.find_parent("li"):
            continue
        if el.name == "table":
            rows = []
            for tr in el.find_all("tr"):
                cells = [tidy(c.get_text(" ")) for c in tr.find_all(["th", "td"])]
                if any(cells):
                    rows.append(cells)
            if rows:
                blocks.append(Block("table", table_text(rows), rows=rows))
            continue
        if el.name == "li":
            for nested in el.find_all(["ul", "ol"]):
                nested.extract()
        text = tidy(el.get_text(" "), keep_lines=el.name == "pre")
        if not text:
            continue
        if el.name.startswith("h"):
            blocks.append(Block("heading", text, level=int(el.name[1])))
        elif el.name == "li":
            blocks.append(Block("list_item", text))
        else:
            blocks.append(Block("paragraph", text))

    warnings = []
    if not blocks:
        text = tidy(root.get_text("\n"), keep_lines=True)
        blocks = [Block("paragraph", p) for p in re.split(r"\n{2,}", text) if p.strip()]
        warnings.append("No standard tags found, plain text was used.")
    return meta, blocks, None, warnings


def decode_text(data):
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1"), "latin-1"


def looks_like_heading(line):
    return len(line) <= 80 and (line.isupper() or line.endswith(":"))


def extract_txt(data, opts):
    text, encoding = decode_text(data)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    blocks = []
    for chunk in re.split(r"\n\s*\n", text):
        lines = [ln.strip() for ln in chunk.split("\n") if ln.strip()]
        if lines and looks_like_heading(lines[0]):
            blocks.append(Block("heading", lines[0].rstrip(":"), level=2))
            lines = lines[1:]
        if not lines:
            continue
        if all(LIST_START.match(ln) for ln in lines):
            blocks += [Block("list_item", ln) for ln in lines]
        else:
            blocks.append(Block("paragraph", tidy("\n".join(lines), keep_lines=True)))
    meta = {"encoding": encoding, "lines": text.count("\n") + 1}
    return meta, blocks, None, []


def extract_image(data, opts):
    img = Image.open(io.BytesIO(data))
    meta = {
        "format": img.format,
        "width": img.width,
        "height": img.height,
        "mode": img.mode,
        "frames": getattr(img, "n_frames", 1),
    }
    if img.info.get("dpi"):
        meta["dpi"] = [round(float(x)) for x in img.info["dpi"]]
    try:
        exif = img.getexif()
        for tag_id, value in exif.items():
            name = ExifTags.TAGS.get(tag_id)
            if name in {"Make", "Model", "Software", "DateTime"}:
                meta[name.lower()] = str(value)
    except Exception:
        pass

    if not TESSERACT_OK:
        raise RuntimeError("Tesseract was not found, images can't be read")
    blocks, warnings = [], []
    for number, frame in enumerate(ImageSequence.Iterator(img), 1):
        found = ocr_blocks(frame.convert("RGB"), opts.lang, number)
        if not found:
            warnings.append(f"Frame {number}: OCR found no text.")
        note = low_confidence_warning(found, f"Frame {number}")
        if note:
            warnings.append(note)
        blocks += found
    return meta, blocks, meta["frames"], warnings


EXTRACTORS = {
    "pdf": (extract_pdf, "PyMuPDF + Tesseract for scanned pages"),
    "docx": (extract_docx, "python-docx"),
    "html": (extract_html, "BeautifulSoup"),
    "txt": (extract_txt, "Plain text reader"),
    "image": (extract_image, "Tesseract OCR"),
}


def extract_document(file_name, data, ocr=True, dpi=200, lang="eng"):
    start = time.perf_counter()
    kind = detect_type(file_name, data)
    source = {
        "file_name": file_name,
        "extension": Path(file_name).suffix.lower(),
        "detected_type": kind,
        "size_bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
        "extracted_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    doc = ExtractedDocument(source=source, doc_type=kind)
    if kind not in EXTRACTORS:
        doc.warnings.append("File type is not supported. Use PDF, DOCX, TXT, HTML or an image.")
        doc.seconds = time.perf_counter() - start
        return doc
    func, doc.method = EXTRACTORS[kind]
    try:
        doc.metadata, doc.blocks, doc.pages, warnings = func(data, Options(ocr, dpi, lang))
        doc.warnings += warnings
        if not doc.blocks and not doc.warnings:
            doc.warnings.append("No text was found in this file.")
    except Exception as exc:
        doc.warnings.append(f"Could not read this file: {exc}")
    doc.seconds = time.perf_counter() - start
    return doc


def extract_many(files, **options):
    return [extract_document(name, data, **options) for name, data in files]
