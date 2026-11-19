import hashlib
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
}
INVALID_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f\x7f]')
WP = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DRAW = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WORD_PART = re.compile(r"word/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$")
SLIDE_PART = re.compile(r"ppt/slides/slide(\d+)\.xml$")


class Section(BaseModel):
    heading: str = Field(min_length=1, max_length=80)
    points: list[str] = Field(min_length=1, max_length=8)

    def clean(self):
        return Section(heading=clean_text(self.heading, 80), points=[clean_text(point, 600) for point in self.points])


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
    if ext not in MIMES or not data or len(data) > MAX_FILE:
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


def paginate_deck(plan):
    pages = []
    for part in plan.sections:
        lines = []
        for point in part.points:
            lines.extend(chunk[i:i+82] for chunk in point.splitlines() for i in range(0, len(chunk), 82) if chunk)
        for offset in range(0, len(lines), 4):
            suffix = "（续）" if offset else ""
            pages.append(Section(heading=part.heading + suffix if len(part.heading) < 77 else part.heading,
                                 points=lines[offset:offset+4]))
    if len(pages) > 30:
        fail(422, "演示内容超过 30 页，请精简段落后再生成")
    return pages


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
        slide = prs.slides.add_slide(blank)
        deck_box(slide, 0, 0, 13.333, 7.5, "F6F8FD")
        deck_box(slide, .83, .78, .07, .43, "345BDD")
        deck_text(slide, part.heading, 1.0, .73, 11.45, 1.05, 28 if len(part.heading) < 34 else 19, "19284F", True)
        for j, point in enumerate(part.points):
            y = 2.0 + j * 1.05
            deck_box(slide, 1.02, y + .14, .065, .065, "345BDD")
            deck_text(slide, point, 1.23, y, 10.98, .91, 17 if len(point) <= 60 else 15, "30405C")
        deck_box(slide, 1.0, 6.82, 11.25, .012, "D6DFF2")
        deck_text(slide, "穹驭办公", 1.0, 7.0, 4, .31, 10, "50618A")
        deck_text(slide, f"{i:02}", 11.7, 7.0, .5, .31, 10, "50618A")
    prs.core_properties.title = plan.title
    prs.core_properties.author = "穹驭办公"
    out = BytesIO()
    prs.save(out)
    return out.getvalue()
