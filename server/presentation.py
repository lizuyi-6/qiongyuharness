import base64
from io import BytesIO
from zipfile import ZipFile

from fastapi import APIRouter
from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pydantic import BaseModel, Field

from .files import file_for
from .identity import fail
from .office import DRAW, clean_text, parse_xml
from .store import database
from .tasks import User

router = APIRouter(prefix="/api/v1/files")
PRESENTATION = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


class Selection(BaseModel):
    file_id: str = Field(min_length=1,max_length=100)
    slide: int = Field(ge=1, le=160)
    shape_id: int | None = Field(default=None, ge=1)


class Replacement(BaseModel):
    shape_id: int = Field(ge=1)
    text: str = Field(min_length=1,max_length=500)


def pptx_file(db, key, user):
    row=file_for(db,user,key)
    if row["extension"]!="pptx":
        fail(422,"此文件不是 PPTX 演示文稿")
    return row


def load_presentation(data):
    try:
        result=Presentation(BytesIO(data))
        if not 1 <= len(result.slides) <= 160:
            fail(422,"预览支持 160 页以内的 PPTX")
        return result
    except (KeyError, ValueError, TypeError, OSError, etree.XMLSyntaxError):
        fail(422,"这个 PPTX 无法预览")


def get_page(prs, page):
    if page<1 or page>len(prs.slides):
        fail(404,"幻灯片不存在")
    return prs.slides[page-1]


def editable_text(shape):
    if not shape.has_text_frame:
        return None
    paragraphs=shape.text_frame.paragraphs
    if len(paragraphs)!=1 or not paragraphs[0].runs or not shape.text.strip() or "\n" in shape.text or "\v" in shape.text:
        return None
    return shape.text


def color_value(color, default):
    return f"#{color.rgb}" if color.type == MSO_COLOR_TYPE.RGB else default


def shape_fill(shape):
    if hasattr(shape,"fill") and shape.fill.type==MSO_FILL_TYPE.SOLID:
        return color_value(shape.fill.fore_color,"#e9ecef")
    return None


def shape_opacity(shape):
    try:
        alpha = shape._element.find(f".//{DRAW}solidFill/{DRAW}srgbClr/{DRAW}alpha")
        if alpha is not None and alpha.get("val"):
            return round(max(0.0, min(1.0, int(alpha.get("val")) / 100000.0)), 3)
    except Exception:
        pass
    return 1.0


def shape_border(shape):
    try:
        if hasattr(shape, "line") and shape.line.fill.type == MSO_FILL_TYPE.SOLID:
            return color_value(shape.line.color, None)
    except Exception:
        pass
    return None


def shape_kind(shape):
    if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
        try:
            if shape.auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
                return "round_rect"
            if shape.auto_shape_type == MSO_SHAPE.OVAL:
                return "ellipse"
        except Exception:
            pass
    return "rect"


def selected_layout(prs,page):
    slide=get_page(prs,page)
    scale=1000/(prs.slide_width or 1)
    objects=[]
    image_budget=6*1024*1024
    for shape in list(slide.shapes)[:180]:
        item={key:round(int(getattr(shape,key) or 0)*scale,2) for key in ("left","top","width","height")}
        fill = shape_fill(shape)
        border = shape_border(shape)
        opacity = shape_opacity(shape) if fill else 1.0
        item.update(id=shape.shape_id, fill=fill, shape=shape_kind(shape))
        if border:
            item["border"] = border
        if opacity < 0.999:
            item["opacity"] = opacity
        if shape.shape_type==MSO_SHAPE_TYPE.PICTURE and image_budget>0:
            picture=shape.image
            if picture.content_type in {"image/png","image/jpeg","image/gif"} and len(picture.blob)<=image_budget:
                item["image"]="data:"+picture.content_type+";base64,"+base64.b64encode(picture.blob).decode()
                image_budget-=len(picture.blob)
        if shape.has_text_frame and shape.text.strip():
            first_p=next((p for p in shape.text_frame.paragraphs if p.text.strip()), None)
            first=next((run for p in shape.text_frame.paragraphs for run in p.runs if run.text.strip()),None)
            font=first.font if first else None
            pt=font.size if font and font.size else 280000
            align = "left"
            if first_p and first_p.alignment == PP_ALIGN.CENTER:
                align = "center"
            elif first_p and first_p.alignment == PP_ALIGN.RIGHT:
                align = "right"
            valign = "top"
            if shape.text_frame.vertical_anchor == MSO_ANCHOR.MIDDLE:
                valign = "middle"
            elif shape.text_frame.vertical_anchor == MSO_ANCHOR.BOTTOM:
                valign = "bottom"
            item.update(text=shape.text[:900], size=round(int(pt)*scale,2), bold=bool(font and font.bold),
                        color=color_value(font.color,"#182231") if font else "#182231",
                        align=align, valign=valign,
                        editable=editable_text(shape) is not None)
        if item.get("text") or item.get("image") or item["fill"] or item.get("border"):
            objects.append(item)
    bg=slide.background.fill
    color=color_value(bg.fore_color,"#ffffff") if bg.type==MSO_FILL_TYPE.SOLID else "#ffffff"
    return {"page":page,"width":1000,"height":round(prs.slide_height*scale,2),"background":color,"objects":objects}


@router.get("/{key}/slides")
def slides(key: str,user:User):
    with database() as db:
        row=pptx_file(db,key,user)
        data=bytes(row["data"])
        name=row["name"]
    prs=load_presentation(data)
    pages=[]
    for index,slide in enumerate(prs.slides,1):
        titles=[shape.text.strip().splitlines()[0] for shape in slide.shapes if shape.has_text_frame and shape.text.strip()]
        meaningful=[t for t in titles if t!="穹驭办公" and not t.isdigit()]
        pages.append({"page":index,"title":(meaningful[0] if meaningful else f"第 {index} 页")[:85]})
    return {"id":key,"name":name,"pages":pages}


@router.get("/{key}/slides/{page}")
def slide_preview(key: str,page:int,user:User):
    with database() as db:
        data=bytes(pptx_file(db,key,user)["data"])
    return selected_layout(load_presentation(data),page)


def edit_scope(data, selection):
    prs=load_presentation(data)
    slide=get_page(prs,selection.slide)
    parts=[]
    for shape in slide.shapes:
        if selection.shape_id and shape.shape_id!=selection.shape_id:
            continue
        text=editable_text(shape)
        if text:
            parts.append({"shape_id":shape.shape_id,"text":text[:475]})
    if not parts:
        fail(422,"所选范围没有可修改的单段文本")
    if len(parts)>45:
        fail(422,"这页文本较多，请单独选择一个文本框")
    return {"page":selection.slide,"boxes":parts,"part":str(slide.part.partname).lstrip("/")}


def revise_slide(content, scope, replacements):
    boxes={item["shape_id"]:item["text"] for item in scope["boxes"]}
    changes={}
    for item in replacements:
        if item.shape_id not in boxes or item.shape_id in changes:
            fail(422,"修改包含未选中的文本框")
        raw=clean_text(item.text,500)
        limit=max(85, min(500,len(boxes[item.shape_id])+35))
        if "\n" in raw or "\r" in raw or len(raw)>limit:
            fail(422,"修改后的文字超出版面范围，请缩短内容")
        if raw!=boxes[item.shape_id]:
            changes[item.shape_id]=raw
    if not changes:
        fail(422,"未生成实际文本改动")
    result=BytesIO()
    with ZipFile(BytesIO(content)) as original, ZipFile(result,"w") as final:
        root=parse_xml(original.read(scope["part"]))
        count=0
        for node in root.iter(f"{PRESENTATION}sp"):
            ident=node.find(f"{PRESENTATION}nvSpPr/{PRESENTATION}cNvPr")
            if ident is None or int(ident.get("id","0")) not in changes:
                continue
            replacement=changes[int(ident.get("id"))]
            frame=node.find(f"{PRESENTATION}txBody")
            texts=list(frame.iter(f"{DRAW}t")) if frame is not None else []
            if not texts:
                fail(422,"所选文本框已不可修改")
            texts[0].text=replacement
            for text in texts[1:]:
                text.text=""
            count+=1
        if count!=len(changes):
            fail(422,"所选文本框不受当前编辑器支持")
        for member in original.infolist():
            final.writestr(member,etree.tostring(root,encoding="UTF-8",xml_declaration=True,standalone=True)
                           if member.filename==scope["part"] else original.read(member))
    return result.getvalue()
