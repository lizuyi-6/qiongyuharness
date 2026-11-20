import hashlib
import math
import re
import socket
import struct
from io import BytesIO
from pathlib import PurePosixPath
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches as DocInches, Pt as DocPt, RGBColor as DocRGB
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pydantic import BaseModel, Field

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
