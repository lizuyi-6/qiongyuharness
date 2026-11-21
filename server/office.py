import hashlib
import math
import re
import socket
import struct
from io import BytesIO
from pathlib import PurePosixPath
from zipfile import BadZipFile, ZipFile

from typing import Any, Literal

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches as DocInches, Pt as DocPt, RGBColor as DocRGB
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import BaseModel, ConfigDict, Field

try:
    from PIL import Image as PilImage, ImageDraw as PilImageDraw
except ImportError:  # pragma: no cover
    PilImage = None
    PilImageDraw = None

from .identity import fail
from .store import audit, instant, uid

MAX_FILE = 12 * 1024 * 1024
MIMES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pdf": "application/pdf",
    "txt": "text/plain; charset=utf-8",
    "md": "text/markdown; charset=utf-8",
    "png": "image/png",
}
UPLOAD_EXTENSIONS = frozenset({"docx", "pptx", "xlsx", "pdf", "txt", "md"})
INVALID_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f\x7f]')
WP = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DRAW = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WORD_PART = re.compile(r"word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$")
SLIDE_PART = re.compile(r"ppt/slides/slide(\d+)\.xml$")


class Table(BaseModel):
    columns: list[str] = Field(min_length=2, max_length=5)
    rows: list[list[str]] = Field(min_length=1, max_length=8)

    def clean(self):
        columns = [clean_text(" ".join(cell.split()), 40) for cell in self.columns]
        rows = []
        for row in self.rows:
            if len(row) > len(columns):
                fail(422, "表格某一行的格数多于表头")
            # 短行补空格子，避免整份文稿因一行少写一格而失败；多写的格子不截断，直接拒绝。
            rows.append([clean_text(" ".join(cell.split()), 80, False) for cell in row] + [""] * (len(columns) - len(row)))
        return Table(columns=columns, rows=rows)


class Section(BaseModel):
    heading: str = Field(min_length=1, max_length=80)
    points: list[str] = Field(default_factory=list, max_length=8)
    conclusion: str = Field(default="", max_length=200)
    notes: str = Field(default="", max_length=2000)
    table: Table | None = None

    def clean(self):
        if not self.points and not self.table:
            fail(422, "每一节至少需要一条要点或一张表格")
        return Section(heading=clean_text(self.heading, 80), points=[clean_text(point, 600) for point in self.points],
                       conclusion=clean_text(" ".join(self.conclusion.split()), 200, False),
                       notes=clean_text(self.notes, 2000, False),
                       table=self.table.clean() if self.table else None)


class Plan(BaseModel):
    title: str = Field(min_length=1, max_length=110)
    subtitle: str = Field(default="", max_length=140)
    sections: list[Section] = Field(min_length=1, max_length=19)

    def clean(self):
        return Plan(title=clean_text(self.title, 110), subtitle=clean_text(self.subtitle, 140, False),
                    sections=[part.clean() for part in self.sections])


def clean_text(text, limit, required=True):
    value = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]", "", text).strip()
    if len(value) > limit or required and not value:
        fail(422, "请检查标题和段落长度")
    return value


def safe_name(name):
    value = INVALID_NAME.sub("_", name.strip()).strip(". ")[:130]
    if not value:
        fail(422, "文件名无效")
    return value


def parse_xml(data):
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False))


def check_zip(data, ext):
    required = {"docx": "word/document.xml", "pptx": "ppt/presentation.xml", "xlsx": "xl/workbook.xml"}
    expected_type = {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    }
    try:
        with ZipFile(BytesIO(data)) as zf:
            infos = zf.infolist()
            names = zf.namelist()
            if (len(infos) > 6000 or len(set(names)) != len(names)
                or sum(i.file_size for i in infos) > 96 * 1024 * 1024
                or required[ext] not in names or "[Content_Types].xml" not in names):
                raise ValueError
            for info in infos:
                path = PurePosixPath(info.filename)
                low = info.filename.lower()
                if (path.is_absolute() or ".." in path.parts or "\\" in info.filename
                    or info.flag_bits & 1 or "vbaproject" in low or "/activex/" in low or "/embeddings/" in low
                    or low.startswith("_xmlsignatures/") or "/externallinks/" in low):
                    raise ValueError
                if low.endswith(".rels"):
                    for rel in parse_xml(zf.read(info)):
                        if rel.get("TargetMode", "").lower() == "external":
                            target = rel.get("Target", "").lower()
                            if not rel.get("Type", "").endswith("/hyperlink") or not target.startswith(("https://", "http://", "mailto:")):
                                raise ValueError
            expected = next((node.get("ContentType") for node in parse_xml(zf.read("[Content_Types].xml"))
                             if node.get("PartName") == f"/{required[ext]}"), None)
            if expected != expected_type[ext]:
                raise ValueError
    except (BadZipFile, ValueError, etree.XMLSyntaxError, KeyError, RuntimeError, NotImplementedError):
        fail(422, "Office 文件格式无效，或包含宏、外部对象、嵌入对象或签名")


def scan(data, target):
    if not target:
        return
    host, port = target.rsplit(":", 1)
    try:
        with socket.create_connection((host, int(port)), timeout=12) as sock:
            sock.sendall(b"zINSTREAM\0")
            for start in range(0, len(data), 65536):
                part = data[start:start+65536]
                sock.sendall(struct.pack("!I", len(part)) + part)
            sock.sendall(b"\0\0\0\0")
            with sock.makefile("rb") as source:
                result = source.readline(4096)
        if b"FOUND" in result:
            fail(422, "文件未通过恶意内容扫描")
        if b"stream: OK\0" not in result:
            fail(503, "文件扫描服务暂不可用")
    except (OSError, ValueError):
        fail(503, "文件扫描服务暂不可用")


def verify(name, data, scanner=""):
    name = safe_name(name)
    ext = name.rsplit(".", 1)[-1].lower()
    if ext not in UPLOAD_EXTENSIONS or not data or len(data) > MAX_FILE:
        fail(422, "请选择支持的文档格式，单个文件最多 12 MB")
    if ext in {"docx", "pptx", "xlsx"}:
        check_zip(data, ext)
    if ext == "pdf" and not data.startswith(b"%PDF-"):
        fail(422, "PDF 文件格式无效")
    if ext in {"txt", "md"}:
        try:
            data.decode("utf-8-sig")
        except UnicodeDecodeError:
            fail(422, "文本附件需使用 UTF-8 编码")
    scan(data, scanner)
    return name, ext


def add_file(db, actor, name, data, ext, label, assignment=None, template=None, parent=None, run=None):
    key = uid()
    digest = hashlib.sha256(data).hexdigest()
    db.execute(
        """INSERT INTO artifacts(id,owner_id,assignment_id,template_id,parent_id,run_id,name,extension,
           digest,size,data,label,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (key, actor, assignment, template, parent, run, name, ext, digest, len(data), data, label, instant()),
    )
    audit(db, "file." + label, actor, assignment, template, artifact_id=key, sha256=digest)
    return key


def nodes_for(root, namespace):
    for paragraph in root.iter(f"{namespace}p"):
        nodes = [n for n in paragraph.iter(f"{namespace}t")
                 if n.text and next(n.iterancestors(f"{namespace}p"), None) is paragraph]
        if nodes:
            yield nodes


def extract(data, ext, limit=16000):
    if ext in {"txt", "md"}:
        return data.decode("utf-8-sig")[:limit]
    if ext == "pdf":
        try:
            from pypdf import PdfReader
            pdf = PdfReader(BytesIO(data))
            if pdf.is_encrypted:
                return "该 PDF 已加密"
            return "\n".join((page.extract_text() or "")[:2000] for page in list(pdf.pages)[:12])[:limit]
        except Exception:
            return ""
    if ext not in {"docx", "pptx"}:
        return ""
    parts = []
    size = 0
    with ZipFile(BytesIO(data)) as zf:
        names = zf.namelist()
        if ext == "pptx":
            names = sorted((name for name in names if SLIDE_PART.fullmatch(name)),
                           key=lambda name: int(SLIDE_PART.fullmatch(name)[1]))
        else:
            names = sorted((name for name in names if WORD_PART.fullmatch(name)),
                           key=lambda name: name != "word/document.xml")
        for i, name in enumerate(names, 1):
            if ext == "pptx":
                parts.append(f"【第 {i} 页】")
            root = parse_xml(zf.read(name))
            for nodes in nodes_for(root, WP if ext == "docx" else DRAW):
                line = "".join(n.text for n in nodes)
                parts.append(line)
                size += len(line) + 1
                if size >= limit:
                    return "\n".join(parts)[:limit]
    return "\n".join(parts)[:limit]


def replace_in_runs(nodes, old, new):
    text = "".join(n.text for n in nodes)
    matches = list(re.finditer(re.escape(old), text))
    if not matches:
        return 0
    bounds = [0]
    for node in nodes:
        bounds.append(bounds[-1] + len(node.text))
    for match in reversed(matches):
        start, end = match.span()
        for i, node in enumerate(nodes):
            if bounds[i+1] <= start or bounds[i] >= end:
                continue
            left = max(0, start - bounds[i])
            right = min(bounds[i+1], end) - bounds[i]
            node.text = (node.text or "")[:left] + (new if bounds[i] <= start else "") + (node.text or "")[right:]
            if node.text != node.text.strip():
                node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return len(matches)


def edit_copy(data, ext, old, new):
    if ext not in {"docx", "pptx"}:
        fail(422, "局部替换支持 Word 和 PowerPoint")
    old, new = clean_text(old, 300), clean_text(new, 500, False)
    if old == new:
        fail(422, "请填写不同的替换文字")
    out = BytesIO()
    count = 0
    matcher = WORD_PART if ext == "docx" else SLIDE_PART
    with ZipFile(BytesIO(data)) as original, ZipFile(out, "w") as result:
        for member in original.infolist():
            payload = original.read(member)
            if matcher.fullmatch(member.filename):
                root = parse_xml(payload)
                changed = sum(replace_in_runs(nodes, old, new) for nodes in nodes_for(root, WP if ext == "docx" else DRAW))
                if changed:
                    payload = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
                    count += changed
            result.writestr(member, payload)
    if not count:
        fail(422, "没有找到对应文字。请复制同一段落中的连续文字")
    return out.getvalue(), count


def word(plan):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = DocInches(0.8)
    section.bottom_margin = DocInches(0.75)
    section.left_margin = section.right_margin = DocInches(0.86)
    for name in ("Normal", "Title", "Subtitle", "Heading 1", "List Bullet"):
        style = doc.styles[name]
        style.font.name = "Microsoft YaHei"
        style.element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal = doc.styles["Normal"]
    normal.font.size = DocPt(11)
    normal.font.color.rgb = DocRGB.from_string("273651")
    normal.paragraph_format.space_after = DocPt(9)
    normal.paragraph_format.line_spacing = 1.35
    doc.styles["Heading 1"].font.color.rgb = DocRGB.from_string("294CC0")
    doc.styles["Title"].font.size = DocPt(28)
    doc.add_heading(plan.title, 0)
    if plan.subtitle:
        doc.add_paragraph(plan.subtitle, "Subtitle")
    for section in plan.sections:
        doc.add_heading(section.heading, 1)
        for point in section.points:
            doc.add_paragraph(point)
        if section.table:
            table = doc.add_table(rows=1 + len(section.table.rows), cols=len(section.table.columns))
            table.style = "Table Grid"
            for row_index, row in enumerate([section.table.columns, *section.table.rows]):
                for col_index, value in enumerate(row):
                    cell = table.cell(row_index, col_index)
                    cell.text = value
                    if row_index == 0:
                        for run in cell.paragraphs[0].runs:
                            run.bold = True
            doc.add_paragraph()
        if section.conclusion:
            doc.add_paragraph().add_run("结论：" + conclusion_text(section.conclusion)).bold = True
    doc.core_properties.author = "穹驭办公"
    doc.core_properties.title = plan.title
    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def deck_box(slide, x, y, width, height, fill):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(width), Inches(height))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor.from_string(fill)
    box.line.fill.background()
    return box


def deck_text(slide, text, x, y, width, height, size, color, bold=False):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = Pt(0)
    frame.margin_top = frame.margin_bottom = Pt(0)
    paragraph = frame.paragraphs[0]
    paragraph.space_after = Pt(0)
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    run._r.get_or_add_rPr().append(etree.Element(f"{DRAW}ea", typeface="Microsoft YaHei"))
    return shape


PAGE_POINTS = 6
TABLE_PAGE_POINTS = 3
CJK = re.compile(r"[⺀-鿿豈-﫿＀-￯　-〿]")


def conclusion_text(text):
    return re.sub(r"^\s*结论\s*[:：]\s*", "", text)


def text_em(text):
    return sum(1.0 if CJK.match(char) else 0.56 for char in text)


def fit_text(text, width, height, sizes):
    """按估算的换行行数选能放下的最大字号，返回 (字号, 实际占用高度英寸)。都放不下时用最小字号。"""
    for size in sizes:
        per_line = max(1.0, width * 72 / size * 0.94)
        lines = max(1, math.ceil(text_em(text) / per_line))
        used = lines * size * 1.32 / 72
        if used <= height or size == sizes[-1]:
            return size, used
    return sizes[-1], height


def split_point(point):
    """"小标题：说明" 拆成 (小标题, 说明)；没有合适的小标题时返回 ("", 原文)。"""
    for mark in ("：", ":"):
        label, sep, body = point.partition(mark)
        if sep and 1 <= len(label.strip()) <= 16 and body.strip():
            return label.strip(), body.strip()
    return "", point


def paginate_deck(plan):
    pages = []
    for part in plan.sections:
        points = [" ".join(point.split()) for point in part.points]
        first = TABLE_PAGE_POINTS if part.table else PAGE_POINTS
        chunks = [points[:first]] + [points[i:i+PAGE_POINTS] for i in range(first, len(points), PAGE_POINTS)]
        for index, chunk in enumerate(chunks):
            suffix = "（续）" if index else ""
            pages.append(Section(heading=part.heading + suffix if len(part.heading) < 77 else part.heading,
                                 points=chunk, conclusion=part.conclusion if index == len(chunks) - 1 else "",
                                 notes=part.notes if index == 0 else "", table=part.table if index == 0 else None))
    if len(pages) > 30:
        fail(422, "演示内容超过 30 页，请精简段落后再生成")
    return pages


def deck_cards(slide, points, x, y, width, height):
    count = len(points)
    columns = 2 if count == 4 else count
    rows = 2 if count == 4 else 1
    gap = .28
    card_w = (width - gap * (columns - 1)) / columns
    card_h = (height - gap * (rows - 1)) / rows
    for index, point in enumerate(points):
        cx = x + (index % columns) * (card_w + gap)
        cy = y + (index // columns) * (card_h + gap)
        deck_box(slide, cx, cy, card_w, card_h, "FFFFFF")
        deck_box(slide, cx, cy, card_w, .06, "345BDD")
        label, body = split_point(point)
        inner_w = card_w - .5
        top = cy + .28
        if label:
            size, used = fit_text(label, inner_w, .85, (20, 18, 16, 15, 14))
            deck_text(slide, label, cx + .25, top, inner_w, used + .05, size, "19284F", True)
            top += used + .18
        body_h = cy + card_h - .22 - top
        size, _ = fit_text(body, inner_w, body_h, (20, 18, 17, 16, 15, 14, 13, 12, 11))
        deck_text(slide, body, cx + .25, top, inner_w, body_h, size, "30405C")


def deck_rows(slide, points, x, y, width, height):
    count = len(points)
    row_h = height / count
    labelled = any(split_point(point)[0] for point in points)
    label_w = 2.55 if labelled else 0
    for index, point in enumerate(points):
        ry = y + index * row_h
        label, body = split_point(point)
        if labelled:
            text = label or "要点"
            size, used = fit_text(text, label_w - .25, row_h - .12, (17, 16, 15, 14, 13))
            deck_text(slide, text, x + .25, ry + max(.06, (row_h - used) / 2), label_w - .25, used + .05, size, "2D4FC4", True)
        else:
            deck_box(slide, x + .05, ry + row_h / 2 - .035, .07, .07, "345BDD")
        body_x = x + (label_w + .2 if labelled else .3)
        body_w = width - (body_x - x)
        size, used = fit_text(body, body_w, row_h - .1, (17, 16, 15, 14, 13, 12, 11))
        deck_text(slide, body, body_x, ry + max(.05, (row_h - used) / 2), body_w, min(used + .05, row_h), size, "30405C")
        if index < count - 1:
            deck_box(slide, x, ry + row_h - .006, width, .012, "D6DFF2")


def deck_table(slide, table, x, y, width, height):
    columns = len(table.columns)
    first_w = width * (.24 if columns >= 3 else .34)
    other_w = (width - first_w) / (columns - 1)
    widths = [first_w] + [other_w] * (columns - 1)
    lines = [table.columns, *table.rows]
    row_h = min(.72, height / len(lines))
    for row_index, row in enumerate(lines):
        cx = x
        ry = y + row_index * row_h
        for col_index, value in enumerate(row):
            header = row_index == 0
            fill = "345BDD" if header else ("FFFFFF" if row_index % 2 else "EEF2FB")
            deck_box(slide, cx, ry, widths[col_index] - .02, row_h - .02, fill)
            if value:
                size, used = fit_text(value, widths[col_index] - .24, row_h - .1, (16, 15, 14, 13, 12, 11, 10))
                color = "FFFFFF" if header else ("19284F" if col_index == 0 else "30405C")
                deck_text(slide, value, cx + .12, ry + max(.04, (row_h - used) / 2), widths[col_index] - .24,
                          min(used + .04, row_h - .04), size, color, header or col_index == 0)
            cx += widths[col_index]
    return row_h * len(lines)


def deck_page(slide, part, number):
    deck_box(slide, 0, 0, 13.333, 7.5, "F6F8FD")
    deck_box(slide, .83, .66, .07, .46, "345BDD")
    size, _ = fit_text(part.heading, 11.3, .8, (28, 25, 22, 19))
    deck_text(slide, part.heading, 1.0, .6, 11.3, .8, size, "19284F", True)
    x, width, top, bottom = 1.0, 11.33, 1.62, 6.6
    if part.conclusion:
        text = "结论：" + conclusion_text(part.conclusion)
        bar_h = .66
        bar_y = 6.6 - bar_h
        deck_box(slide, x, bar_y, width, bar_h, "E6EDFF")
        deck_box(slide, x, bar_y, .08, bar_h, "345BDD")
        size, used = fit_text(text, width - .5, bar_h - .14, (17, 16, 15, 14, 13, 12))
        deck_text(slide, text, x + .3, bar_y + (bar_h - used) / 2, width - .5, used + .04, size, "1F3FAF", True)
        bottom = bar_y - .22
    area = bottom - top
    if part.table:
        points_h = min(len(part.points) * .62, area * .38) if part.points else 0
        used = deck_table(slide, part.table, x, top, width, area - points_h - (.2 if points_h else 0))
        if part.points:
            deck_rows(slide, part.points, x, top + used + .2, width, bottom - (top + used + .2))
    elif part.points and len(part.points) <= 4:
        deck_cards(slide, part.points, x, top, width, area)
    elif part.points:
        deck_rows(slide, part.points, x, top, width, area)
    if part.notes:
        slide.notes_slide.notes_text_frame.text = part.notes
    deck_box(slide, 1.0, 6.82, 11.25, .012, "D6DFF2")
    deck_text(slide, "穹驭办公", 1.0, 7.0, 4, .31, 10, "50618A")
    deck_text(slide, f"{number:02}", 11.7, 7.0, .5, .31, 10, "50618A")


def powerpoint(plan):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    cover = prs.slides.add_slide(blank)
    deck_box(cover, 0, 0, 13.333, 7.5, "17264A")
    deck_box(cover, 10.25, 0, 3.083, 7.5, "2544AB")
    deck_box(cover, 0.85, 2.13, 0.62, 0.065, "8EA7FF")
    deck_text(cover, "穹驭办公", 0.85, 0.75, 6, .5, 16, "A4B7FA", True)
    deck_text(cover, plan.title, 0.85, 2.53, 9.1, 2.35, 34 if len(plan.title) < 50 else 25, "FFFFFF", True)
    deck_text(cover, plan.subtitle or "项目资料", 0.85, 5.14, 9.1, 1.25, 17, "CBD6FA")
    for i, part in enumerate(paginate_deck(plan), 2):
        deck_page(prs.slides.add_slide(blank), part, i)
    prs.core_properties.title = plan.title
    prs.core_properties.author = "穹驭办公"
    out = BytesIO()
    prs.save(out)
    return out.getvalue()


HEX_COLOR_RE = re.compile(r"^#?([0-9a-fA-F]{6}|[0-9a-fA-F]{3})$")
SLIDE_W = 13.333
SLIDE_H = 7.5
_ELEMENT_TYPE_ALIASES = {
    "rect": "rect",
    "rectangle": "rect",
    "box": "rect",
    "bar": "rect",
    "block": "rect",
    "shape": "rect",
    "round_rect": "round_rect",
    "rounded_rect": "round_rect",
    "rounded_rectangle": "round_rect",
    "roundrect": "round_rect",
    "card": "round_rect",
    "pill": "round_rect",
    "badge": "round_rect",
    "tag": "round_rect",
    "ellipse": "ellipse",
    "oval": "ellipse",
    "circle": "ellipse",
    "dot": "ellipse",
    "line": "line",
    "divider": "line",
    "rule": "line",
    "separator": "line",
    "connector": "line",
    "arrow": "line",
    "text": "text",
    "textbox": "text",
    "label": "text",
    "heading": "text",
    "title": "text",
    "paragraph": "text",
    "image": "image",
    "img": "image",
    "picture": "image",
    "photo": "image",
    "table": "table",
    "chart": "chart",
    "bar_chart": "chart",
    "hbar": "chart",
    "vbar": "chart",
    "kpi": "chart",
    "kpi_row": "chart",
    "metrics": "chart",
    "progress": "chart",
}


def _to_float(value: Any, default: float) -> float:
    if isinstance(value, bool) or value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value) if math.isfinite(value) else default
    if isinstance(value, str):
        match = re.search(r"-?\d+(?:\.\d+)?", value)
        if match:
            num = float(match.group(0))
            return num if math.isfinite(num) else default
    return default


def _to_int(value: Any, default: int | None = None) -> int | None:
    if isinstance(value, bool) or value is None:
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, float) and math.isfinite(value):
        return int(round(value))
    if isinstance(value, str):
        match = re.search(r"\d+", value)
        if match:
            return int(match.group(0))
    return default


def _clamp(value: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, value))


def _coerce_opacity(value: Any, default: float | None = None) -> float | None:
    if value is None or isinstance(value, bool):
        return default
    num = _to_float(value, -1.0)
    if num < 0:
        return default
    if num > 1.0:
        num = num / 100.0
    return _clamp(num, 0.0, 1.0)


def _coerce_str(value: Any, limit: int = 2000) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        value = "\n".join(_coerce_str(item, limit) for item in value if item is not None)
    text = str(value)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]", "", text).strip()
    return text[:limit]


class OverlaySpec(BaseModel):
    model_config = ConfigDict(extra="ignore")
    color: str = "#000000"
    opacity: float = 0.65
    mode: Literal["solid", "scrim_left", "scrim_bottom"] = "scrim_left"


class DeckDesign(BaseModel):
    model_config = ConfigDict(extra="ignore")
    concept: str = ""
    palette: dict[str, str] = Field(default_factory=dict)
    font: str = "Microsoft YaHei"


class SlideElement(BaseModel):
    model_config = ConfigDict(extra="ignore")
    type: Literal["rect", "round_rect", "ellipse", "line", "text", "image", "table", "chart"]
    x: float = 0.0
    y: float = 0.0
    w: float = 1.0
    h: float = 0.5
    fill: str | None = None
    gradient: dict[str, Any] | None = None
    opacity: float | None = None
    border: str | None = None
    border_width: float = 1.0
    radius: float | None = None
    shadow: bool = False
    text: str | None = None
    size: float = 16.0
    bold: bool = False
    italic: bool = False
    color: str | None = None
    font: str | None = None
    align: Literal["left", "center", "right"] = "left"
    valign: Literal["top", "middle", "bottom"] = "top"
    line_spacing: float = 1.28
    image_index: int | None = None
    fit: Literal["cover", "contain"] = "cover"
    columns: list[str] | None = None
    rows: list[list[str]] | None = None
    header_fill: str | None = None
    header_color: str | None = None
    row_fill: str | None = None
    alt_fill: str | None = None
    text_color: str | None = None
    border_color: str | None = None
    chart_kind: Literal["hbar", "vbar", "kpi_row"] = "hbar"
    chart_title: str | None = None
    unit: str | None = None
    items: list[dict[str, Any]] | None = None


class FreeSlide(BaseModel):
    model_config = ConfigDict(extra="ignore")
    title: str = ""
    bg: str = "#FFFFFF"
    bg_image: int | None = None
    bg_overlay: OverlaySpec | None = None
    notes: str = ""
    elements: list[SlideElement] = Field(default_factory=list)


class DeckSpec(BaseModel):
    model_config = ConfigDict(extra="ignore")
    title: str = "演示文稿"
    subtitle: str = ""
    design: DeckDesign = Field(default_factory=DeckDesign)
    slides: list[FreeSlide] = Field(min_length=1, max_length=30)

    def clean(self):
        return DeckSpec(
            title=_coerce_str(self.title, 110) or "演示文稿",
            subtitle=_coerce_str(self.subtitle, 140),
            design=self.design,
            slides=self.slides,
        )


def _parse_coord(value: Any, default: float, axis_size: float) -> tuple[float, bool]:
    """解析坐标数值，同时识别显式百分号（如 '88%'）。返回 (数值, 是否显式百分比)。"""
    if isinstance(value, str) and "%" in value:
        num = _to_float(value, -9999.0)
        if num > -9990.0:
            return (num / 100.0) * axis_size, True
    return _to_float(value, default), False


def _detect_coord_scale(raw_slides: list[Any]) -> tuple[float, float]:
    """自动识别 AI 输出的坐标系单位（英寸 / 0-100 百分比 / 1920x1080 / 1280x720 / 960x540 / 1000x562.5）。"""
    xs: list[float] = []
    ys: list[float] = []
    rights: list[float] = []
    bottoms: list[float] = []
    out_of_inch = 0
    total = 0

    for slide_item in raw_slides[:30]:
        if not isinstance(slide_item, dict):
            continue
        raw_els = slide_item.get("elements")
        if not isinstance(raw_els, list):
            continue
        for el in raw_els[:120]:
            if not isinstance(el, dict):
                continue
            raw_type = str(el.get("type") or "").strip().lower().replace("-", "_").replace(" ", "_")
            el_type = _ELEMENT_TYPE_ALIASES.get(raw_type)
            if not el_type:
                continue
            rx, x_pct = _parse_coord(el.get("x", el.get("left")), 0.0, SLIDE_W)
            ry, y_pct = _parse_coord(el.get("y", el.get("top")), 0.0, SLIDE_H)
            rw, w_pct = _parse_coord(el.get("w", el.get("width")), 1.0, SLIDE_W)
            rh, h_pct = _parse_coord(el.get("h", el.get("height")), 0.02 if el_type == "line" else 0.5, SLIDE_H)
            if x_pct or y_pct or w_pct or h_pct:
                continue
            total += 1
            right = rx + max(0.0, rw)
            bottom = ry + max(0.0, rh)
            xs.append(rx)
            ys.append(ry)
            rights.append(right)
            bottoms.append(bottom)
            if right > 14.5 or bottom > 8.5 or rx > 13.8 or ry > 7.9:
                out_of_inch += 1

    if total == 0 or out_of_inch == 0:
        return 1.0, 1.0
    # 偶发一两个微小越界元素仍属于英寸坐标系；只有大面积超界才判定为非英寸坐标系
    if out_of_inch < max(2, int(total * 0.2)):
        return 1.0, 1.0

    max_right = max(rights)
    max_bottom = max(bottoms)
    # 1) 0-100 百分比坐标系（最常见：x/y/w/h 落在 0..100）
    if max_right <= 108.0 and max_bottom <= 108.0 and (max_right > 15.0 or max_bottom > 9.0):
        return SLIDE_W / 100.0, SLIDE_H / 100.0
    # 2) 1920x1080 像素坐标系
    if 1350.0 < max_right <= 2100.0 and 760.0 < max_bottom <= 1200.0:
        return SLIDE_W / 1920.0, SLIDE_H / 1080.0
    # 3) 1600x900 像素坐标系
    if 1350.0 < max_right <= 1700.0 and 760.0 < max_bottom <= 980.0:
        return SLIDE_W / 1600.0, SLIDE_H / 900.0
    # 4) 1280x720 像素坐标系
    if 1050.0 < max_right <= 1350.0 and 580.0 < max_bottom <= 780.0:
        return SLIDE_W / 1280.0, SLIDE_H / 720.0
    # 5) 1000x562.5 或 960x540 坐标系
    if 600.0 < max_right <= 1050.0 and 350.0 < max_bottom <= 620.0:
        ref_w = 1000.0 if max_right > 975.0 else 960.0
        ref_h = ref_w * (SLIDE_H / SLIDE_W)
        return SLIDE_W / ref_w, SLIDE_H / ref_h
    # 6) 通用超界放大坐标系兜底
    scale_x = SLIDE_W / max_right if max_right > 14.5 else 1.0
    scale_y = SLIDE_H / max_bottom if max_bottom > 8.5 else 1.0
    return scale_x, scale_y


def coerce_deck_spec(raw: Any, fallback_title: str = "演示文稿") -> DeckSpec:
    """对 AI 输出的自由画布 JSON 做强容错归一化，自动换算误用的百分比/像素坐标系，避免因别名或越界导致空白页。"""
    if not isinstance(raw, dict):
        fail(502, "办公助理返回了无法读取的结果，请重试当前请求")
    raw_slides = raw.get("slides")
    if not isinstance(raw_slides, list) or not raw_slides:
        fail(502, "办公助理返回了无法读取的结果，请重试当前请求")

    coord_sx, coord_sy = _detect_coord_scale(raw_slides)

    raw_design = raw.get("design") if isinstance(raw.get("design"), dict) else {}
    raw_palette = raw_design.get("palette") if isinstance(raw_design.get("palette"), dict) else {}
    palette = {
        str(k): v.strip()
        for k, v in raw_palette.items()
        if isinstance(v, str) and parse_hex_color(v)
    }
    design = DeckDesign(
        concept=_coerce_str(raw_design.get("concept"), 300),
        palette=palette,
        font=_coerce_str(raw_design.get("font"), 60) or "Microsoft YaHei",
    )

    slides: list[FreeSlide] = []
    for s_idx, item in enumerate(raw_slides[:30]):
        if not isinstance(item, dict):
            continue
        raw_bg = item.get("bg")
        if isinstance(raw_bg, dict):
            raw_bg = raw_bg.get("color") or raw_bg.get("fill")
        bg = raw_bg.strip() if isinstance(raw_bg, str) and parse_hex_color(raw_bg) else (palette.get("bg") or "#FFFFFF")

        bg_image = _to_int(item.get("bg_image"))
        if bg_image is not None and not (1 <= bg_image <= 12):
            bg_image = None

        raw_ov = item.get("bg_overlay")
        bg_overlay = None
        if isinstance(raw_ov, dict):
            ov_color = raw_ov.get("color") if isinstance(raw_ov.get("color"), str) and parse_hex_color(raw_ov.get("color")) else bg
            ov_opacity = _coerce_opacity(raw_ov.get("opacity"), 0.68)
            ov_mode_raw = str(raw_ov.get("mode") or ("scrim_left" if s_idx == 0 else "solid")).strip().lower()
            ov_mode = ov_mode_raw if ov_mode_raw in {"solid", "scrim_left", "scrim_bottom"} else "scrim_left"
            bg_overlay = OverlaySpec(color=ov_color, opacity=ov_opacity if ov_opacity is not None else 0.68, mode=ov_mode)
        elif isinstance(raw_ov, str) and parse_hex_color(raw_ov):
            bg_overlay = OverlaySpec(color=raw_ov.strip(), opacity=0.68, mode="scrim_left" if s_idx == 0 else "solid")
        elif bg_image is not None and s_idx == 0:
            bg_overlay = OverlaySpec(color=bg, opacity=0.72, mode="scrim_left")

        elements: list[SlideElement] = []
        raw_elements = item.get("elements") if isinstance(item.get("elements"), list) else []
        for el in raw_elements[:120]:
            if not isinstance(el, dict):
                continue
            raw_type = str(el.get("type") or "").strip().lower().replace("-", "_").replace(" ", "_")
            el_type = _ELEMENT_TYPE_ALIASES.get(raw_type)
            if not el_type:
                continue
            rx, x_pct = _parse_coord(el.get("x", el.get("left")), 0.0, SLIDE_W)
            ry, y_pct = _parse_coord(el.get("y", el.get("top")), 0.0, SLIDE_H)
            rw, w_pct = _parse_coord(el.get("w", el.get("width")), 1.0, SLIDE_W)
            rh, h_pct = _parse_coord(el.get("h", el.get("height")), 0.02 if el_type == "line" else 0.5, SLIDE_H)
            sx = 1.0 if x_pct else coord_sx
            sy = 1.0 if y_pct else coord_sy
            sw = 1.0 if w_pct else coord_sx
            sh = 1.0 if h_pct else coord_sy
            x = _clamp(rx * sx, -2.0, 15.0)
            y = _clamp(ry * sy, -2.0, 9.5)
            w = _clamp(rw * sw, 0.015, 16.0)
            h = _clamp(rh * sh, 0.015, 10.0)
            if el_type == "line" and el.get("x2") is not None and el.get("y2") is not None:
                rx2, x2_pct = _parse_coord(el.get("x2"), rx + rw, SLIDE_W)
                ry2, y2_pct = _parse_coord(el.get("y2"), ry, SLIDE_H)
                x2 = rx2 * (1.0 if x2_pct else coord_sx)
                y2 = ry2 * (1.0 if y2_pct else coord_sy)
                x, y = min(x, x2), min(y, y2)
                w = _clamp(abs(x2 - x), 0.015, 16.0)
                h = _clamp(abs(y2 - y), 0.015, 10.0)

            radius = None
            if el.get("radius") is not None:
                r_val = _to_float(el.get("radius"), 0.1)
                if r_val > 0.5:
                    r_val = min(0.5, r_val / max(1.0, min(w, h) * 72.0))
                radius = _clamp(r_val, 0.0, 0.5)
                if el_type == "rect" and radius > 0.001:
                    el_type = "round_rect"

            align_raw = str(el.get("align") or "left").strip().lower()
            align = align_raw if align_raw in {"left", "center", "right"} else "left"
            valign_raw = str(el.get("valign") or "top").strip().lower()
            if valign_raw == "center":
                valign_raw = "middle"
            valign = valign_raw if valign_raw in {"top", "middle", "bottom"} else "top"
            fit_raw = str(el.get("fit") or "cover").strip().lower()
            fit = fit_raw if fit_raw in {"cover", "contain"} else "cover"

            img_idx = _to_int(el.get("image_index", el.get("index")))
            if img_idx is not None and not (1 <= img_idx <= 12):
                img_idx = None

            cols = None
            rows = None
            if el_type == "table":
                raw_cols = el.get("columns")
                raw_rows = el.get("rows")
                if isinstance(raw_cols, list) and raw_cols:
                    cols = [_coerce_str(c, 60) for c in raw_cols[:8]]
                if isinstance(raw_rows, list):
                    rows = [
                        [_coerce_str(cell, 160) for cell in row[:8]]
                        for row in raw_rows[:14]
                        if isinstance(row, list)
                    ]

            raw_text = el.get("text")
            text_val = _coerce_str(raw_text, 2000) if raw_text is not None else None
            if el_type == "text" and not text_val:
                continue

            gradient = None
            raw_grad = el.get("gradient")
            if isinstance(raw_grad, dict):
                g_from = raw_grad.get("from") or raw_grad.get("start")
                g_to = raw_grad.get("to") or raw_grad.get("end")
                if isinstance(g_from, str) and parse_hex_color(g_from) and isinstance(g_to, str) and parse_hex_color(g_to):
                    gradient = {
                        "from": g_from.strip(),
                        "to": g_to.strip(),
                        "angle": _clamp(_to_float(raw_grad.get("angle"), 0.0), 0.0, 360.0),
                    }

            chart_kind = "hbar"
            chart_items = None
            chart_title = None
            chart_unit = None
            if el_type == "chart":
                ck_raw = str(el.get("chart_kind") or el.get("kind") or raw_type or "hbar").strip().lower()
                if ck_raw in {"vbar", "column", "vertical"}:
                    chart_kind = "vbar"
                elif ck_raw in {"kpi", "kpi_row", "metrics", "stats"}:
                    chart_kind = "kpi_row"
                else:
                    chart_kind = "hbar"
                chart_title = _coerce_str(el.get("title") or el.get("chart_title"), 100) or None
                chart_unit = _coerce_str(el.get("unit"), 24) or None
                raw_items = el.get("items") or el.get("data")
                if isinstance(raw_items, list):
                    chart_items = []
                    for it in raw_items[:10]:
                        if not isinstance(it, dict):
                            continue
                        lbl = _coerce_str(it.get("label") or it.get("name"), 60)
                        disp = _coerce_str(it.get("display") or it.get("value_text") or it.get("value"), 40)
                        val = _to_float(it.get("value"), 0.0)
                        sub = _coerce_str(it.get("sub") or it.get("note") or it.get("desc"), 100)
                        if lbl or disp:
                            chart_items.append({
                                "label": lbl,
                                "value": max(0.0, val),
                                "display": disp,
                                "sub": sub,
                                "highlight": bool(it.get("highlight")),
                                "color": it.get("color") if isinstance(it.get("color"), str) and parse_hex_color(it.get("color")) else None,
                            })
                if not chart_items:
                    continue

            raw_fill = el.get("fill") or el.get("background") or el.get("bg") or el.get("fill_color") or el.get("background_color")
            raw_border = el.get("border") or el.get("stroke") or el.get("border_color")
            raw_color = el.get("color") or (el.get("text_color") if el_type == "text" else None) or el.get("font_color")
            elements.append(SlideElement(
                type=el_type,
                x=x,
                y=y,
                w=w,
                h=h,
                fill=raw_fill if isinstance(raw_fill, str) else None,
                gradient=gradient,
                opacity=_coerce_opacity(el.get("opacity")),
                border=raw_border if isinstance(raw_border, str) else None,
                border_width=_clamp(_to_float(el.get("border_width", el.get("stroke_width")), 1.0), 0.0, 12.0),
                radius=radius,
                shadow=bool(el.get("shadow")),
                text=text_val,
                size=_clamp(_to_float(el.get("size", el.get("font_size")), 16.0), 8.0, 120.0),
                bold=bool(el.get("bold") or str(el.get("font_weight") or "").lower() in {"bold", "600", "700", "800"}),
                italic=bool(el.get("italic")),
                color=raw_color if isinstance(raw_color, str) else None,
                font=_coerce_str(el.get("font"), 60) or None,
                align=align,
                valign=valign,
                line_spacing=_clamp(_to_float(el.get("line_spacing"), 1.28), 0.9, 2.4),
                image_index=img_idx,
                fit=fit,
                columns=cols,
                rows=rows,
                header_fill=el.get("header_fill") if isinstance(el.get("header_fill"), str) else None,
                header_color=el.get("header_color") if isinstance(el.get("header_color"), str) else None,
                row_fill=el.get("row_fill") if isinstance(el.get("row_fill"), str) else None,
                alt_fill=el.get("alt_fill") if isinstance(el.get("alt_fill"), str) else None,
                text_color=el.get("text_color") if isinstance(el.get("text_color"), str) else None,
                border_color=el.get("border_color") if isinstance(el.get("border_color"), str) else None,
                chart_kind=chart_kind,
                chart_title=chart_title,
                unit=chart_unit,
                items=chart_items,
            ))

        slides.append(FreeSlide(
            title=_coerce_str(item.get("title"), 100) or f"第 {s_idx + 1} 页",
            bg=bg,
            bg_image=bg_image,
            bg_overlay=bg_overlay,
            notes=_coerce_str(item.get("notes"), 2000),
            elements=elements,
        ))

    if not slides:
        fail(502, "办公助理返回了无法读取的结果，请重试当前请求")
    return DeckSpec(
        title=_coerce_str(raw.get("title"), 110) or _coerce_str(fallback_title, 110) or "演示文稿",
        subtitle=_coerce_str(raw.get("subtitle"), 140),
        design=design,
        slides=slides,
    )


def parse_hex_color(value: Any, default: str | None = None) -> str | None:
    if not isinstance(value, str):
        return default
    raw = value.strip()
    match = HEX_COLOR_RE.fullmatch(raw)
    if not match:
        return default
    digits = match.group(1).upper()
    if len(digits) == 3:
        digits = "".join(ch * 2 for ch in digits)
    return digits


def _apply_alpha(color_elem, opacity: float | None):
    if opacity is None or opacity >= 0.999:
        return
    alpha_val = max(0, min(100000, int(round(opacity * 100000))))
    for child in list(color_elem):
        if child.tag == f"{DRAW}alpha":
            color_elem.remove(child)
    alpha = etree.Element(f"{DRAW}alpha")
    alpha.set("val", str(alpha_val))
    color_elem.append(alpha)


def _set_shape_fill(shape, color_hex: str | None, opacity: float | None = None):
    if not color_hex:
        shape.fill.background()
        return
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color_hex)
    if opacity is not None and opacity < 0.999:
        srgb = shape._element.find(f".//{DRAW}solidFill/{DRAW}srgbClr")
        if srgb is not None:
            _apply_alpha(srgb, opacity)


def _is_dark_hex(hex_str: str | None) -> bool:
    code = parse_hex_color(hex_str)
    if not code:
        return False
    r, g, b = int(code[0:2], 16), int(code[2:4], 16), int(code[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 142.0


def _set_shape_gradient(shape, stops: list[tuple[int, str, float]], angle_deg: float = 0.0):
    """用 DrawingML <a:gradFill> 渲染多段线性渐变（支持起止透明度，用于电影感封面蒙层与高端卡片）。"""
    sp_pr = shape._element.spPr
    for tag in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"):
        existing = sp_pr.find(f"{DRAW}{tag}")
        if existing is not None:
            sp_pr.remove(existing)
    grad_fill = etree.Element(f"{DRAW}gradFill")
    gs_lst = etree.SubElement(grad_fill, f"{DRAW}gsLst")
    for pos_pct, hex_code, alpha in stops:
        gs = etree.SubElement(gs_lst, f"{DRAW}gs", pos=str(max(0, min(100000, int(pos_pct)))))
        srgb = etree.SubElement(gs, f"{DRAW}srgbClr", val=hex_code)
        if alpha < 0.999:
            etree.SubElement(srgb, f"{DRAW}alpha", val=str(max(0, min(100000, int(round(alpha * 100000))))))
    ang_val = int(round((angle_deg % 360.0) * 60000))
    etree.SubElement(grad_fill, f"{DRAW}lin", ang=str(ang_val), scaled="1")
    ln = sp_pr.find(f"{DRAW}ln")
    if ln is not None:
        sp_pr.insert(sp_pr.index(ln), grad_fill)
    else:
        sp_pr.append(grad_fill)


def _apply_soft_shadow(shape, is_dark_bg: bool = False):
    """为卡片或高亮模块施加现代 Keynote 风格柔和外投影。"""
    sp_pr = shape._element.spPr
    existing = sp_pr.find(f"{DRAW}effectLst")
    if existing is not None:
        sp_pr.remove(existing)
    eff_lst = etree.Element(f"{DRAW}effectLst")
    shdw = etree.SubElement(
        eff_lst,
        f"{DRAW}outerShdw",
        blurRad="152400",
        dist="38100",
        dir="5400000",
        algn="ctr",
        rotWithShape="0",
    )
    srgb = etree.SubElement(shdw, f"{DRAW}srgbClr", val="000000")
    etree.SubElement(srgb, f"{DRAW}alpha", val="22000" if is_dark_bg else "12000")
    sp_pr.append(eff_lst)


def _set_shape_border(shape, border_hex: str | None, border_width: float = 1.0):
    if not border_hex or border_width <= 0:
        shape.line.fill.background()
        return
    shape.line.color.rgb = RGBColor.from_string(border_hex)
    shape.line.width = Pt(border_width)


def _png_dimensions(blob: bytes) -> tuple[int, int] | None:
    if len(blob) >= 24 and blob[:8] == b"\x89PNG\r\n\x1a\n" and blob[12:16] == b"IHDR":
        w, h = struct.unpack(">II", blob[16:24])
        if w > 0 and h > 0:
            return w, h
    return None


def _prepare_picture_blob(blob: bytes, w: float, h: float, fit: str = "cover", corner_radius_in: float = 0.0) -> bytes:
    """用 Pillow 做真正的高清等比裁切与 4x 超采样圆角遮罩，避免拉伸变形和生硬直角。"""
    if PilImage is None or PilImageDraw is None:
        return blob
    if fit != "cover" and corner_radius_in <= 0:
        return blob
    try:
        img = PilImage.open(BytesIO(blob)).convert("RGBA")
        iw, ih = img.size
        if iw <= 0 or ih <= 0:
            return blob
        box_ratio = max(0.05, w / max(0.05, h))
        img_ratio = iw / ih
        if fit == "cover":
            if img_ratio > box_ratio:
                new_w = max(1, int(round(ih * box_ratio)))
                left = (iw - new_w) // 2
                img = img.crop((left, 0, left + new_w, ih))
            elif img_ratio < box_ratio:
                new_h = max(1, int(round(iw / box_ratio)))
                top = (ih - new_h) // 2
                img = img.crop((0, top, iw, top + new_h))
        cw, ch = img.size
        if corner_radius_in > 0:
            px_per_in = cw / max(0.1, w)
            radius_px = max(4, min(int(round(corner_radius_in * px_per_in)), min(cw, ch) // 4))
            scale = 2
            mask = PilImage.new("L", (cw * scale, ch * scale), 0)
            draw = PilImageDraw.Draw(mask)
            draw.rounded_rectangle((0, 0, cw * scale - 1, ch * scale - 1), radius=radius_px * scale, fill=255)
            mask = mask.resize((cw, ch), PilImage.Resampling.LANCZOS)
            img.putalpha(mask)
        out = BytesIO()
        img.save(out, format="PNG", optimize=True)
        return out.getvalue()
    except Exception:
        return blob


def _add_fitted_picture(slide, blob: bytes, x: float, y: float, w: float, h: float, fit: str = "cover", corner_radius_in: float = 0.0):
    processed = _prepare_picture_blob(blob, w, h, fit=fit, corner_radius_in=corner_radius_in)
    dims = _png_dimensions(processed)
    if not dims or fit == "cover":
        return slide.shapes.add_picture(BytesIO(processed), Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    iw, ih = dims
    img_ratio = iw / ih
    box_ratio = w / h
    if img_ratio > box_ratio:
        draw_w = w
        draw_h = w / img_ratio
        draw_x = x
        draw_y = y + (h - draw_h) / 2
    else:
        draw_h = h
        draw_w = h * img_ratio
        draw_x = x + (w - draw_w) / 2
        draw_y = y
    return slide.shapes.add_picture(BytesIO(processed), Inches(draw_x), Inches(draw_y),
                                    width=Inches(draw_w), height=Inches(draw_h))


def _add_free_chart(slide, el: SlideElement, palette: dict[str, str], default_font: str, is_dark_slide: bool):
    """将结构化 chart 原语渲染为 100% 原生可编辑的咨询级数据可视化（hbar / vbar / kpi_row）。"""
    items = el.items or []
    if not items:
        return
    primary = parse_hex_color(palette.get("primary")) or "2563EB"
    accent = parse_hex_color(palette.get("accent")) or primary
    text_hex = parse_hex_color(palette.get("text")) or ("F8FAFC" if is_dark_slide else "111827")
    muted_hex = parse_hex_color(palette.get("muted")) or ("94A3B8" if is_dark_slide else "64748B")
    surface_hex = parse_hex_color(palette.get("surface")) or ("1E293B" if is_dark_slide else "FFFFFF")
    border_hex = parse_hex_color(palette.get("border")) or ("334155" if is_dark_slide else "E2E8F0")
    track_hex = "1E293B" if is_dark_slide else "E8EEF6"

    top = el.y
    if el.chart_title:
        title_el = SlideElement(
            type="text", x=el.x, y=top, w=el.w, h=0.38,
            text=el.chart_title, size=14.5, bold=True, color="#" + text_hex, font=default_font,
        )
        _add_free_text(slide, title_el, default_font, text_hex)
        top += 0.46

    avail_h = max(0.8, el.y + el.h - top)

    if el.chart_kind == "kpi_row":
        n = len(items)
        gap = 0.24
        card_w = (el.w - gap * (n - 1)) / max(1, n)
        for idx, it in enumerate(items):
            cx = el.x + idx * (card_w + gap)
            hl = bool(it.get("highlight"))
            card_fill = primary if hl else surface_hex
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(top), Inches(card_w), Inches(avail_h))
            try:
                card.adjustments[0] = 0.08
            except Exception:
                pass
            _set_shape_fill(card, card_fill)
            _set_shape_border(card, None if hl else border_hex, 0.0 if hl else 0.75)
            _apply_soft_shadow(card, is_dark_slide)
            bar_color = "FFFFFF" if hl else (parse_hex_color(it.get("color")) or accent)
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx + 0.28), Inches(top + 0.24), Inches(0.42), Inches(0.06))
            _set_shape_fill(bar, bar_color, 0.9)
            bar.line.fill.background()

            val_color = "FFFFFF" if hl else (parse_hex_color(it.get("color")) or primary)
            lbl_color = "E2E8F0" if hl else text_hex
            sub_color = "CBD5E1" if hl else muted_hex
            display_val = str(it.get("display") or it.get("value") or "")
            if el.unit and el.unit not in display_val:
                display_val = f"{display_val}{el.unit}"
            val_h = min(1.05, avail_h * 0.42)
            _add_free_text(slide, SlideElement(
                type="text", x=cx + 0.26, y=top + 0.44, w=card_w - 0.52, h=val_h,
                text=display_val, size=36.0 if n <= 3 else 30.0, bold=True, color="#" + val_color, font=default_font, valign="middle",
            ), default_font, val_color)
            _add_free_text(slide, SlideElement(
                type="text", x=cx + 0.26, y=top + 0.44 + val_h + 0.06, w=card_w - 0.52, h=0.42,
                text=str(it.get("label") or ""), size=14.5, bold=True, color="#" + lbl_color, font=default_font,
            ), default_font, lbl_color)
            if it.get("sub"):
                sub_y = top + 0.44 + val_h + 0.48
                sub_h = max(0.35, top + avail_h - sub_y - 0.18)
                _add_free_text(slide, SlideElement(
                    type="text", x=cx + 0.26, y=sub_y, w=card_w - 0.52, h=sub_h,
                    text=str(it.get("sub")), size=12.5, color="#" + sub_color, font=default_font, line_spacing=1.25,
                ), default_font, sub_color)
        return

    max_val = max((float(it.get("value") or 0.0) for it in items), default=1.0) or 1.0

    if el.chart_kind == "vbar":
        n = len(items)
        gap = min(0.35, el.w * 0.04)
        col_w = (el.w - gap * (n - 1)) / max(1, n)
        bar_w = min(0.85, col_w * 0.56)
        label_h = 0.52
        val_h = 0.34
        plot_h = max(0.8, avail_h - label_h - val_h - 0.14)
        base_y = top + val_h + plot_h
        axis = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(el.x), Inches(base_y), Inches(el.w), Inches(0.015))
        _set_shape_fill(axis, border_hex)
        axis.line.fill.background()
        for idx, it in enumerate(items):
            cx = el.x + idx * (col_w + gap)
            ratio = _clamp(float(it.get("value") or 0.0) / max_val, 0.06, 1.0)
            bh = max(0.18, plot_h * ratio)
            bx = cx + (col_w - bar_w) / 2
            by = base_y - bh
            bar_color = parse_hex_color(it.get("color")) or (accent if it.get("highlight") else primary)
            bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(bx), Inches(by), Inches(bar_w), Inches(bh))
            try:
                bar.adjustments[0] = 0.08
            except Exception:
                pass
            _set_shape_fill(bar, bar_color, 1.0 if it.get("highlight") else 0.78)
            bar.line.fill.background()
            disp = str(it.get("display") or it.get("value") or "")
            _add_free_text(slide, SlideElement(
                type="text", x=cx, y=by - val_h, w=col_w, h=val_h,
                text=disp, size=13.5, bold=bool(it.get("highlight")), color="#" + (bar_color if it.get("highlight") else text_hex),
                font=default_font, align="center", valign="bottom",
            ), default_font, text_hex)
            _add_free_text(slide, SlideElement(
                type="text", x=cx, y=base_y + 0.08, w=col_w, h=label_h,
                text=str(it.get("label") or ""), size=12.5, bold=bool(it.get("highlight")), color="#" + muted_hex,
                font=default_font, align="center", valign="top",
            ), default_font, muted_hex)
        return

    # 默认 hbar：咨询级横向对比条形图
    n = len(items)
    row_h = min(0.76, avail_h / max(1, n))
    label_w = min(2.6, el.w * 0.27)
    val_w = min(1.35, el.w * 0.16)
    track_x = el.x + label_w + 0.16
    track_w = max(1.2, el.w - label_w - val_w - 0.34)
    bar_h = min(0.32, row_h * 0.48)
    for idx, it in enumerate(items):
        ry = top + idx * row_h
        by = ry + (row_h - bar_h) / 2
        hl = bool(it.get("highlight"))
        bar_color = parse_hex_color(it.get("color")) or (accent if hl else primary)
        _add_free_text(slide, SlideElement(
            type="text", x=el.x, y=ry + 0.02, w=label_w, h=row_h - 0.04,
            text=str(it.get("label") or ""), size=13.5, bold=hl, color="#" + (text_hex if hl else muted_hex),
            font=default_font, align="left", valign="middle",
        ), default_font, text_hex)
        track = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(track_x), Inches(by), Inches(track_w), Inches(bar_h))
        try:
            track.adjustments[0] = 0.35
        except Exception:
            pass
        _set_shape_fill(track, track_hex)
        track.line.fill.background()

        ratio = _clamp(float(it.get("value") or 0.0) / max_val, 0.05, 1.0)
        fill_w = max(0.18, track_w * ratio)
        fill_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(track_x), Inches(by), Inches(fill_w), Inches(bar_h))
        try:
            fill_bar.adjustments[0] = 0.35
        except Exception:
            pass
        _set_shape_fill(fill_bar, bar_color, 1.0 if hl else 0.72)
        fill_bar.line.fill.background()

        disp = str(it.get("display") or it.get("value") or "")
        _add_free_text(slide, SlideElement(
            type="text", x=track_x + track_w + 0.16, y=ry + 0.02, w=val_w, h=row_h - 0.04,
            text=disp, size=14.0, bold=True, color="#" + (bar_color if hl else text_hex),
            font=default_font, align="right", valign="middle",
        ), default_font, text_hex)


def _fit_free_text_size(text: str, width: float, height: float, target_size: float, line_spacing: float) -> float:
    lines_raw = text.replace("\r", "").split("\n")
    size = float(target_size)
    min_size = max(9.0, min(11.0, size * 0.65))
    while size > min_size:
        per_line = max(1.0, width * 72.0 / size * 0.92)
        total_lines = sum(max(1, math.ceil(text_em(part) / per_line)) for part in lines_raw)
        used_h = total_lines * size * max(1.15, line_spacing) / 72.0
        if used_h <= height * 1.06:
            return round(size, 1)
        size -= 1.0 if size <= 24 else 2.0
    return round(min_size, 1)


def _add_free_text(slide, el: SlideElement, default_font: str, default_color: str):
    raw = clean_text(el.text or "", 2000, False)
    if not raw:
        return
    shape = slide.shapes.add_textbox(Inches(el.x), Inches(el.y), Inches(el.w), Inches(el.h))
    if el.fill:
        fill_hex = parse_hex_color(el.fill)
        if fill_hex:
            _set_shape_fill(shape, fill_hex, el.opacity)
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = Pt(2)
    frame.margin_top = frame.margin_bottom = Pt(1)
    valign_map = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    frame.vertical_anchor = valign_map.get(el.valign, MSO_ANCHOR.TOP)

    font_name = (el.font or default_font or "Microsoft YaHei").strip() or "Microsoft YaHei"
    color_hex = parse_hex_color(el.color, default_color) or "182231"
    fitted_size = _fit_free_text_size(raw, el.w, el.h, el.size, el.line_spacing)

    paragraphs = raw.replace("\r", "").split("\n")
    for idx, line in enumerate(paragraphs):
        p = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        p.alignment = align_map.get(el.align, PP_ALIGN.LEFT)
        p.space_after = Pt(2 if idx < len(paragraphs) - 1 else 0)
        p.line_spacing = el.line_spacing
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(fitted_size)
        run.font.bold = el.bold
        run.font.italic = el.italic
        run.font.color.rgb = RGBColor.from_string(color_hex)
        run._r.get_or_add_rPr().append(etree.Element(f"{DRAW}ea", typeface=font_name))


def _add_free_table(slide, el: SlideElement, palette: dict[str, str], default_font: str):
    cols = [clean_text(" ".join(str(c).split()), 60, False) for c in (el.columns or [])]
    if len(cols) < 1:
        return
    rows_raw = el.rows or []
    rows = []
    for row in rows_raw[:14]:
        cleaned = [clean_text(" ".join(str(cell).split()), 120, False) for cell in row[:len(cols)]]
        cleaned += [""] * (len(cols) - len(cleaned))
        rows.append(cleaned)
    all_rows = [cols, *rows]
    col_count = len(cols)
    first_w = el.w * (0.26 if col_count >= 3 else 0.36) if col_count > 1 else el.w
    other_w = (el.w - first_w) / max(1, col_count - 1) if col_count > 1 else el.w
    widths = [first_w] + [other_w] * (col_count - 1)
    row_h = min(0.75, el.h / max(1, len(all_rows)))

    header_fill = parse_hex_color(el.header_fill) or parse_hex_color(palette.get("primary")) or "1E293B"
    header_color = parse_hex_color(el.header_color) or "FFFFFF"
    row_fill = parse_hex_color(el.row_fill) or parse_hex_color(palette.get("surface")) or "FFFFFF"
    alt_fill = parse_hex_color(el.alt_fill) or parse_hex_color(palette.get("bg")) or "F8FAFC"
    text_color = parse_hex_color(el.text_color) or parse_hex_color(palette.get("text")) or "1E293B"
    border_hex = parse_hex_color(el.border_color) or parse_hex_color(palette.get("border"))

    for r_idx, row in enumerate(all_rows):
        cx = el.x
        ry = el.y + r_idx * row_h
        is_header = r_idx == 0
        for c_idx, value in enumerate(row):
            cw = widths[c_idx]
            cell_fill = header_fill if is_header else (row_fill if r_idx % 2 else alt_fill)
            box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(ry),
                                         Inches(max(0.05, cw - 0.02)), Inches(max(0.05, row_h - 0.02)))
            _set_shape_fill(box, cell_fill)
            _set_shape_border(box, border_hex, 0.5 if border_hex else 0.0)
            if value:
                cell_el = SlideElement(
                    type="text",
                    x=cx + 0.08,
                    y=ry + 0.04,
                    w=max(0.1, cw - 0.18),
                    h=max(0.1, row_h - 0.1),
                    text=value,
                    size=13.5 if is_header else 12.5,
                    bold=is_header or c_idx == 0,
                    color="#" + (header_color if is_header else text_color),
                    font=default_font,
                    align="left",
                    valign="middle",
                    line_spacing=1.18,
                )
                _add_free_text(slide, cell_el, default_font, text_color)
            cx += cw


def render_freeform_pptx(deck: DeckSpec, image_blobs: list[bytes] | tuple[bytes, ...] = ()) -> bytes:
    """无模板自由画布渲染器：完全按 AI 输出的视觉设计系统、坐标原语与生图资产生成可编辑 PPTX。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
    blank = prs.slide_layouts[6]

    palette = deck.design.palette if deck.design else {}
    default_font = (deck.design.font if deck.design else "") or "Microsoft YaHei"
    default_bg = parse_hex_color(palette.get("bg"), "FFFFFF") or "FFFFFF"
    default_text = parse_hex_color(palette.get("text"), "182231") or "182231"

    shape_map = {
        "rect": MSO_SHAPE.RECTANGLE,
        "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
        "ellipse": MSO_SHAPE.OVAL,
        "line": MSO_SHAPE.RECTANGLE,
    }

    total_slides = len(deck.slides)
    for s_idx, slide_spec in enumerate(deck.slides):
        slide = prs.slides.add_slide(blank)
        bg_hex = parse_hex_color(slide_spec.bg, default_bg) or default_bg
        is_dark_slide = _is_dark_hex(bg_hex)
        bg_fill = slide.background.fill
        bg_fill.solid()
        bg_fill.fore_color.rgb = RGBColor.from_string(bg_hex)

        # 可选全屏背景图 + 电影感渐变蒙层或纯色蒙层
        if slide_spec.bg_image and 1 <= slide_spec.bg_image <= len(image_blobs):
            _add_fitted_picture(slide, image_blobs[slide_spec.bg_image - 1], 0.0, 0.0, SLIDE_W, SLIDE_H, fit="cover", corner_radius_in=0.0)
            if slide_spec.bg_overlay:
                ov_hex = parse_hex_color(slide_spec.bg_overlay.color, bg_hex)
                if ov_hex and slide_spec.bg_overlay.opacity > 0:
                    overlay = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                                     Inches(SLIDE_W), Inches(SLIDE_H))
                    op = _clamp(slide_spec.bg_overlay.opacity, 0.15, 0.95)
                    if slide_spec.bg_overlay.mode == "scrim_left":
                        _set_shape_gradient(
                            overlay,
                            [(0, ov_hex, min(0.92, op + 0.14)), (52000, ov_hex, op), (100000, ov_hex, max(0.12, op * 0.25))],
                            angle_deg=0.0,
                        )
                    elif slide_spec.bg_overlay.mode == "scrim_bottom":
                        _set_shape_gradient(
                            overlay,
                            [(0, ov_hex, max(0.12, op * 0.22)), (55000, ov_hex, op * 0.72), (100000, ov_hex, min(0.92, op + 0.16))],
                            angle_deg=90.0,
                        )
                    else:
                        _set_shape_fill(overlay, ov_hex, op)
                    overlay.line.fill.background()

        # 非封面页防“无头页”兜底：若顶部区域没有任何文字，自动补齐页面主标题
        if s_idx > 0 and slide_spec.title and not slide_spec.bg_image:
            has_header_text = any(el.type == "text" and el.y < 1.35 for el in slide_spec.elements)
            if not has_header_text:
                header_color = "F8FAFC" if is_dark_slide else (parse_hex_color(palette.get("text")) or "111827")
                _add_free_text(slide, SlideElement(
                    type="text", x=0.68, y=0.52, w=11.9, h=0.68,
                    text=slide_spec.title, size=25.0, bold=True, color="#" + header_color, font=default_font,
                ), default_font, header_color)

        for el in slide_spec.elements:
            if el.type in shape_map:
                w = max(0.01, el.w)
                h = max(0.01, el.h)
                shape = slide.shapes.add_shape(shape_map[el.type], Inches(el.x), Inches(el.y), Inches(w), Inches(h))
                if el.type == "round_rect":
                    try:
                        shape.adjustments[0] = max(0.0, min(0.5, float(el.radius if el.radius is not None else 0.07)))
                    except Exception:
                        pass
                fill_hex = parse_hex_color(el.fill)
                border_hex = parse_hex_color(el.border)
                if el.type == "line" and not fill_hex and border_hex:
                    fill_hex = border_hex
                    border_hex = None
                if el.gradient and parse_hex_color(el.gradient.get("from")) and parse_hex_color(el.gradient.get("to")):
                    g_alpha = el.opacity if el.opacity is not None else 1.0
                    _set_shape_gradient(
                        shape,
                        [(0, parse_hex_color(el.gradient["from"]), g_alpha), (100000, parse_hex_color(el.gradient["to"]), g_alpha)],
                        angle_deg=float(el.gradient.get("angle") or 0.0),
                    )
                else:
                    _set_shape_fill(shape, fill_hex, el.opacity)
                _set_shape_border(shape, border_hex, el.border_width)
                # 浅色底上的大卡片自动给柔和投影，告别平面死板感
                auto_card_shadow = (
                    el.shadow or (
                        el.type == "round_rect"
                        and not is_dark_slide
                        and fill_hex is not None
                        and fill_hex != bg_hex
                        and w >= 2.2
                        and h >= 1.2
                        and (el.opacity is None or el.opacity >= 0.9)
                    )
                )
                if auto_card_shadow:
                    _apply_soft_shadow(shape, is_dark_slide)
            elif el.type == "image":
                if el.image_index and 1 <= el.image_index <= len(image_blobs):
                    img_w = max(0.1, el.w)
                    img_h = max(0.1, el.h)
                    full_bleed = img_w >= SLIDE_W - 0.5 and img_h >= SLIDE_H - 0.5
                    corner_r = 0.0 if full_bleed else (el.radius * min(img_w, img_h) if el.radius is not None else 0.14)
                    _add_fitted_picture(slide, image_blobs[el.image_index - 1], el.x, el.y,
                                        img_w, img_h, fit=el.fit, corner_radius_in=corner_r)
            elif el.type == "text":
                _add_free_text(slide, el, default_font, default_text)
            elif el.type == "table":
                _add_free_table(slide, el, palette, default_font)
            elif el.type == "chart":
                _add_free_chart(slide, el, palette, default_font, is_dark_slide)

        # 统一杂志级细线页脚与页码（封面页与全出血背景过渡页不加，且不与底部自绘内容冲突）
        if s_idx > 0 and not slide_spec.bg_image:
            has_footer = any(el.type == "text" and el.y >= 6.95 for el in slide_spec.elements)
            if not has_footer:
                foot_line = "334155" if is_dark_slide else (parse_hex_color(palette.get("border")) or "DCE3EE")
                foot_text = "94A3B8" if is_dark_slide else (parse_hex_color(palette.get("muted")) or "64748B")
                rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.68), Inches(7.04), Inches(11.97), Inches(0.01))
                _set_shape_fill(rule, foot_line, 0.55)
                rule.line.fill.background()
                _add_free_text(slide, SlideElement(
                    type="text", x=0.68, y=7.10, w=8.5, h=0.28,
                    text=deck.title[:48], size=9.5, color="#" + foot_text, font=default_font,
                ), default_font, foot_text)
                _add_free_text(slide, SlideElement(
                    type="text", x=10.8, y=7.10, w=1.85, h=0.28,
                    text=f"{s_idx + 1:02d} / {total_slides:02d}", size=9.5, bold=True, color="#" + foot_text, font=default_font, align="right",
                ), default_font, foot_text)

        if slide_spec.notes:
            slide.notes_slide.notes_text_frame.text = clean_text(slide_spec.notes, 2000, False)

    prs.core_properties.title = deck.title
    prs.core_properties.author = "穹驭办公"
    out = BytesIO()
    prs.save(out)
    return out.getvalue()
