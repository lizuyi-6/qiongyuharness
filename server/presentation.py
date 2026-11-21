import base64
from collections import OrderedDict
import hashlib
from io import BytesIO
import re
import threading
from zipfile import ZipFile

from fastapi import APIRouter
from fastapi.responses import Response
from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches
from pydantic import BaseModel, Field

try:
    from PIL import Image as PilImage
except Exception:
    PilImage = None

from .files import file_for, public_file
from .identity import fail
from .office import DRAW, add_file, clean_text, parse_xml
from .store import database, instant, pack, unpack
from .tasks import User

router = APIRouter(prefix="/api/v1/files")
PRESENTATION = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_SAFE_MEDIA_DIGEST = re.compile(r"^[0-9a-f]{16,40}$")
_PREVIEW_CACHE_LOCK = threading.Lock()
_PREVIEW_CACHE: OrderedDict[tuple[str, str], dict] = OrderedDict()
_PREVIEW_CACHE_MAX_DECKS = 12
_PREVIEW_CACHE_MAX_BYTES = 20 * 1024 * 1024
_PREVIEW_CACHE_BYTES = 0


class RegionBox(BaseModel):
    """圈选矩形区域（16:9 自由画布英寸坐标：x, y, w, h）"""
    x: float = Field(ge=-1.0, le=14.5)
    y: float = Field(ge=-1.0, le=8.5)
    w: float = Field(gt=0.05, le=15.0)
    h: float = Field(gt=0.05, le=9.0)


class Selection(BaseModel):
    file_id: str = Field(min_length=1,max_length=100)
    slide: int = Field(ge=1, le=160)
    shape_id: int | None = Field(default=None, ge=1)
    shape_ids: list[int] = Field(default_factory=list, max_length=60)
    region: RegionBox | None = None
    label: str | None = Field(default=None, max_length=120)


class Replacement(BaseModel):
    shape_id: int = Field(ge=1)
    text: str = Field(min_length=1,max_length=500)


class LayoutChange(BaseModel):
    shape_id: int = Field(ge=1)
    x: float = Field(ge=-2.0, le=15.0)
    y: float = Field(ge=-2.0, le=9.5)
    w: float = Field(gt=0.05, le=16.0)
    h: float = Field(gt=0.05, le=10.0)


class SlideLayoutUpdate(BaseModel):
    changes: list[LayoutChange] = Field(min_length=1, max_length=80)


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


def shape_gradient(shape):
    try:
        grad = shape._element.find(f".//{DRAW}gradFill")
        if grad is None:
            return None
        stops = []
        for gs in grad.findall(f"./{DRAW}gsLst/{DRAW}gs"):
            pos = int(gs.get("pos", "0")) / 100000.0
            srgb = gs.find(f"{DRAW}srgbClr")
            if srgb is None or not srgb.get("val"):
                continue
            alpha_el = srgb.find(f"{DRAW}alpha")
            alpha = int(alpha_el.get("val")) / 100000.0 if alpha_el is not None and alpha_el.get("val") else 1.0
            stops.append({"pos": round(max(0.0, min(1.0, pos)), 3), "color": f"#{srgb.get('val')}", "alpha": round(max(0.0, min(1.0, alpha)), 3)})
        if len(stops) < 2:
            return None
        lin = grad.find(f"{DRAW}lin")
        angle = int(lin.get("ang", "0")) / 60000.0 if lin is not None and lin.get("ang") else 0.0
        return {"angle": round(angle, 1), "stops": stops}
    except Exception:
        return None


def shape_has_shadow(shape):
    try:
        return shape._element.find(f".//{DRAW}effectLst/{DRAW}outerShdw") is not None
    except Exception:
        return False


def shape_radius(shape):
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and shape.auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
            if len(shape.adjustments) > 0:
                return round(max(0.0, min(0.5, float(shape.adjustments[0]))), 3)
    except Exception:
        pass
    return None


def _compress_preview_image(blob: bytes, content_type: str) -> tuple[str, bytes]:
    """将 PPT 内嵌高分辨率 PNG 极速压缩为轻量预览图（JPEG/WebP），体积缩小 20-40 倍且耗时仅数毫秒。"""
    if content_type in {"image/jpeg", "image/webp"} and len(blob) <= 180 * 1024:
        return content_type, blob
    if PilImage is None:
        return content_type, blob
    try:
        with PilImage.open(BytesIO(blob)) as src:
            has_alpha = False
            if src.mode in ("RGBA", "LA") or (src.mode == "P" and "transparency" in src.info):
                lo, _hi = src.convert("RGBA").getchannel("A").getextrema()
                has_alpha = lo < 250
            img = src.convert("RGBA" if has_alpha else "RGB")
        if img.width > 1280 or img.height > 720:
            img.thumbnail((1280, 720), PilImage.Resampling.BILINEAR)
        out = BytesIO()
        if has_alpha:
            img.save(out, format="WEBP", quality=80, method=1)
            return "image/webp", out.getvalue()
        img.save(out, format="JPEG", quality=82, optimize=False)
        return "image/jpeg", out.getvalue()
    except Exception:
        return content_type, blob


def _evict_preview_cache_locked():
    global _PREVIEW_CACHE_BYTES
    while _PREVIEW_CACHE and (len(_PREVIEW_CACHE) > _PREVIEW_CACHE_MAX_DECKS or _PREVIEW_CACHE_BYTES > _PREVIEW_CACHE_MAX_BYTES):
        _, oldest = _PREVIEW_CACHE.popitem(last=False)
        _PREVIEW_CACHE_BYTES = max(0, _PREVIEW_CACHE_BYTES - int(oldest.get("bytes", 0)))


def _build_deck_preview_bundle(key: str, data: bytes) -> dict:
    prs = load_presentation(data)
    pages = []
    layouts: dict[int, dict] = {}
    media: dict[str, tuple[str, bytes]] = {}
    total_media_bytes = 0
    for index, slide in enumerate(prs.slides, 1):
        titles = [shape.text.strip().splitlines()[0] for shape in slide.shapes if shape.has_text_frame and shape.text.strip()]
        meaningful = [t for t in titles if t != "穹驭办公" and not t.isdigit()]
        pages.append({"page": index, "title": (meaningful[0] if meaningful else f"第 {index} 页")[:85]})
        layouts[index] = selected_layout(
            prs, index, media_store=media, media_url_prefix=f"/api/v1/files/{key}/slides/media"
        )
    for _mime, raw_bytes in media.values():
        total_media_bytes += len(raw_bytes)
    return {"pages": pages, "layouts": layouts, "media": media, "bytes": total_media_bytes}


def _get_deck_preview_bundle(key: str, digest: str, data: bytes) -> dict:
    global _PREVIEW_CACHE_BYTES
    cache_key = (key, digest or "")
    with _PREVIEW_CACHE_LOCK:
        cached = _PREVIEW_CACHE.get(cache_key)
        if cached is not None:
            _PREVIEW_CACHE.move_to_end(cache_key)
            return cached
    bundle = _build_deck_preview_bundle(key, data)
    with _PREVIEW_CACHE_LOCK:
        existing = _PREVIEW_CACHE.get(cache_key)
        if existing is not None:
            _PREVIEW_CACHE.move_to_end(cache_key)
            return existing
        _PREVIEW_CACHE[cache_key] = bundle
        _PREVIEW_CACHE_BYTES += int(bundle.get("bytes", 0))
        _evict_preview_cache_locked()
    return bundle


def selected_layout(prs, page, *, media_store: dict | None = None, media_url_prefix: str = ""):
    slide=get_page(prs,page)
    scale=1000/(prs.slide_width or 1)
    objects=[]
    image_budget=6*1024*1024
    for shape in list(slide.shapes)[:180]:
        item={key:round(int(getattr(shape,key) or 0)*scale,2) for key in ("left","top","width","height")}
        fill = shape_fill(shape)
        gradient = shape_gradient(shape) if not fill else None
        border = shape_border(shape)
        opacity = shape_opacity(shape) if fill else 1.0
        radius = shape_radius(shape)
        item.update(id=shape.shape_id, fill=fill, shape=shape_kind(shape))
        if gradient:
            item["gradient"] = gradient
        if border:
            item["border"] = border
        if opacity < 0.999:
            item["opacity"] = opacity
        if radius is not None:
            item["radius"] = radius
        if shape_has_shadow(shape):
            item["shadow"] = True
        if shape.shape_type==MSO_SHAPE_TYPE.PICTURE and image_budget>0:
            picture=shape.image
            if picture.content_type in {"image/png","image/jpeg","image/gif","image/webp"} and len(picture.blob)<=image_budget:
                if media_store is not None and media_url_prefix:
                    img_hash = hashlib.sha1(picture.blob).hexdigest()[:20]
                    if img_hash not in media_store:
                        media_store[img_hash] = _compress_preview_image(picture.blob, picture.content_type)
                    item["image"] = f"{media_url_prefix}/{img_hash}"
                else:
                    mime, compressed = _compress_preview_image(picture.blob, picture.content_type)
                    item["image"] = "data:" + mime + ";base64," + base64.b64encode(compressed).decode()
                image_budget-=len(picture.blob)
            item["kind"] = "image"
            item["selectable"] = True
            if item["width"] >= 920 and item["height"] >= 500:
                item["is_background"] = True
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
                        kind="text", selectable=True,
                        editable=editable_text(shape) is not None)
        if not item.get("kind"):
            item["kind"] = "shape"
            item["selectable"] = True
        if item.get("text") or item.get("image") or item["fill"] or item.get("gradient") or item.get("border"):
            objects.append(item)
    bg=slide.background.fill
    color=color_value(bg.fore_color,"#ffffff") if bg.type==MSO_FILL_TYPE.SOLID else "#ffffff"
    return {"page":page,"width":1000,"height":round(prs.slide_height*scale,2),"background":color,"objects":objects}


def _load_authorized_bundle(key: str, user: User) -> tuple[str, dict]:
    with database() as db:
        row = pptx_file(db, key, user)
        name = row["name"]
        digest = str(row["digest"] or "")
        cache_key = (key, digest)
        with _PREVIEW_CACHE_LOCK:
            cached = _PREVIEW_CACHE.get(cache_key)
            if cached is not None:
                _PREVIEW_CACHE.move_to_end(cache_key)
                return name, cached
        data = bytes(row["data"])
    return name, _get_deck_preview_bundle(key, digest, data)


@router.get("/{key}/slides")
def slides(key: str, user: User):
    name, bundle = _load_authorized_bundle(key, user)
    return {
        "id": key,
        "name": name,
        "pages": bundle["pages"],
        "first_page": bundle["layouts"].get(1),
    }


@router.get("/{key}/slides/media/{img_hash}")
def slide_media(key: str, img_hash: str, user: User):
    if not _SAFE_MEDIA_DIGEST.fullmatch(img_hash):
        fail(404, "图片资源不存在")
    _name, bundle = _load_authorized_bundle(key, user)
    entry = bundle["media"].get(img_hash)
    if not entry:
        fail(404, "图片资源不存在")
    mime, payload = entry
    return Response(
        payload,
        media_type=mime,
        headers={
            "Cache-Control": "private, max-age=86400, immutable",
            "X-Content-Type-Options": "nosniff",
        },
    )


@router.get("/{key}/slides/{page}")
def slide_preview(key: str, page: int, user: User):
    _name, bundle = _load_authorized_bundle(key, user)
    layout = bundle["layouts"].get(page)
    if layout is None:
        fail(404, "幻灯片不存在")
    return layout


def _sync_plan_element_coords(plan, page: int, moved_boxes: list[tuple[float, float, float, float, float, float, float, float]]):
    """把手动拖动/缩放的新坐标同步回自由画布 DeckSpec，保证后续 AI 圈选修改看到最新排版。"""
    if not isinstance(plan, dict):
        return False
    slides_list = plan.get("slides")
    if not isinstance(slides_list, list) or not (1 <= page <= len(slides_list)):
        return False
    slide_dict = slides_list[page - 1]
    elements = slide_dict.get("elements") if isinstance(slide_dict, dict) else None
    if not isinstance(elements, list) or not elements:
        return False
    used_indices: set[int] = set()
    changed = False
    for old_x, old_y, old_w, old_h, new_x, new_y, new_w, new_h in moved_boxes:
        best_idx = -1
        best_dist = 0.65
        for idx, el in enumerate(elements):
            if idx in used_indices or not isinstance(el, dict):
                continue
            ex = float(el.get("x") or 0.0)
            ey = float(el.get("y") or 0.0)
            ew = float(el.get("w") or 0.0)
            eh = float(el.get("h") or 0.0)
            dist = abs(ex - old_x) + abs(ey - old_y) + abs(ew - old_w) * 0.5 + abs(eh - old_h) * 0.5
            if dist < best_dist:
                best_dist = dist
                best_idx = idx
        if best_idx >= 0:
            used_indices.add(best_idx)
            elements[best_idx]["x"] = round(new_x, 3)
            elements[best_idx]["y"] = round(new_y, 3)
            elements[best_idx]["w"] = round(new_w, 3)
            elements[best_idx]["h"] = round(new_h, 3)
            changed = True
    return changed


@router.post("/{key}/slides/{page}/layout")
def update_slide_layout(key: str, page: int, payload: SlideLayoutUpdate, user: User):
    with database() as db:
        row = pptx_file(db, key, user)
        data = bytes(row["data"])
    prs = load_presentation(data)
    slide = get_page(prs, page)
    change_map = {item.shape_id: item for item in payload.changes}
    moved_boxes: list[tuple[float, float, float, float, float, float, float, float]] = []
    emu_per_inch = 914400.0
    updated = 0
    for shape in slide.shapes:
        if shape.shape_id not in change_map:
            continue
        item = change_map[shape.shape_id]
        old_x = int(shape.left or 0) / emu_per_inch
        old_y = int(shape.top or 0) / emu_per_inch
        old_w = int(shape.width or 0) / emu_per_inch
        old_h = int(shape.height or 0) / emu_per_inch
        shape.left = Inches(item.x)
        shape.top = Inches(item.y)
        shape.width = Inches(max(0.05, item.w))
        shape.height = Inches(max(0.05, item.h))
        moved_boxes.append((old_x, old_y, old_w, old_h, item.x, item.y, item.w, item.h))
        updated += 1
    if not updated:
        fail(422, "未找到可移动的幻灯片组件")
    out = BytesIO()
    prs.save(out)
    new_bytes = out.getvalue()

    with database(write=True) as db:
        row = pptx_file(db, key, user)
        new_id = add_file(
            db, user["id"], row["name"], new_bytes, "pptx", "edited",
            assignment=row["assignment_id"], parent=row["id"], run=row["run_id"],
        )
        new_file = public_file(db, new_id)
        if row["run_id"]:
            turn_row = db.one("SELECT id, conversation_id, result FROM office_turns WHERE id=?", (row["run_id"],))
            if turn_row:
                res = unpack(turn_row["result"] or "{}")
                if isinstance(res.get("file"), dict) and res["file"].get("id") == row["id"]:
                    res["file"] = new_file
                if res.get("plan"):
                    _sync_plan_element_coords(res["plan"], page, moved_boxes)
                db.execute("UPDATE office_turns SET result=? WHERE id=?", (pack(res), turn_row["id"]))
                db.execute("UPDATE office_conversations SET updated_at=? WHERE id=?", (instant(), turn_row["conversation_id"]))
    return {"file": new_file, "page": page, "updated": updated}


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
