"""
Pure-Python OOXML (.pptx) builder.

Assembles the minimum set of Office Open XML parts required for a valid,
editable PowerPoint file (ECMA-376), then zips them into a .pptx container
using the standard library only.

No python-pptx or other PowerPoint libraries are used — this module writes
raw OOXML strings and packages them with `zipfile`.

Public API
----------
build_pptx(deck: DeckInput) -> bytes
    Returns the .pptx binary for the given deck.

DeckInput / SlideInput
    Typed dicts describing the deck content.

Units
-----
OOXML uses English Metric Units (EMU): 914,400 EMU = 1 inch.
We render a 16:9 slide (12,192,000 x 6,858,000 EMU = 13.333in x 7.5in).
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass, field
from html import escape
from typing import List, Optional


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #

@dataclass
class SlideInput:
    title: str
    bullets: List[str] = field(default_factory=list)
    image_description: Optional[str] = None


@dataclass
class DeckInput:
    title: str
    slides: List[SlideInput] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Layout constants (EMU)
# --------------------------------------------------------------------------- #

SLIDE_W = 12_192_000
SLIDE_H = 6_858_000

MARGIN_X = 685_800   # 0.75in
TITLE_Y = 457_200    # 0.5in
TITLE_H = 900_000

BODY_Y = 1_500_000
BODY_W_HALF = 5_200_000    # left column when image placeholder present
BODY_W_FULL = SLIDE_W - 2 * MARGIN_X
BODY_H = 4_600_000

IMAGE_X = MARGIN_X + BODY_W_HALF + 300_000
IMAGE_Y = BODY_Y
IMAGE_W = SLIDE_W - IMAGE_X - MARGIN_X
IMAGE_H = BODY_H


def _esc(s: str) -> str:
    """XML-escape a string safely."""
    return escape(s, quote=True)


# --------------------------------------------------------------------------- #
# OOXML part templates
# --------------------------------------------------------------------------- #

def _content_types_xml(slide_count: int) -> str:
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{i + 1}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(slide_count)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        f'{overrides}'
        '</Types>'
    )


def _root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="ppt/presentation.xml"/>'
        '</Relationships>'
    )


def _presentation_xml(slide_count: int) -> str:
    slide_ids = "".join(
        f'<p:sldId id="{256 + i}" r:id="rId{i + 2}"/>' for i in range(slide_count)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{slide_ids}</p:sldIdLst>'
        f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        '</p:presentation>'
    )


def _presentation_rels_xml(slide_count: int) -> str:
    slide_rels = "".join(
        f'<Relationship Id="rId{i + 2}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        f'Target="slides/slide{i + 1}.xml"/>'
        for i in range(slide_count)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        'Target="slideMasters/slideMaster1.xml"/>'
        f'{slide_rels}'
        f'<Relationship Id="rId{slide_count + 2}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"/>'
        '</Relationships>'
    )


def _theme_xml() -> str:
    # Minimal but valid Office theme. PowerPoint requires theme1.xml to exist.
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">'
        '<a:themeElements>'
        '<a:clrScheme name="Office">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="1F2937"/></a:dk2>'
        '<a:lt2><a:srgbClr val="E5E7EB"/></a:lt2>'
        '<a:accent1><a:srgbClr val="2563EB"/></a:accent1>'
        '<a:accent2><a:srgbClr val="10B981"/></a:accent2>'
        '<a:accent3><a:srgbClr val="F59E0B"/></a:accent3>'
        '<a:accent4><a:srgbClr val="EF4444"/></a:accent4>'
        '<a:accent5><a:srgbClr val="8B5CF6"/></a:accent5>'
        '<a:accent6><a:srgbClr val="EC4899"/></a:accent6>'
        '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
        '</a:clrScheme>'
        '<a:fontScheme name="Office">'
        '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
        '</a:fontScheme>'
        '<a:fmtScheme name="Office">'
        '<a:fillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '</a:fillStyleLst>'
        '<a:lnStyleLst>'
        '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        '</a:lnStyleLst>'
        '<a:effectStyleLst>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '</a:effectStyleLst>'
        '<a:bgFillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '</a:bgFillStyleLst>'
        '</a:fmtScheme>'
        '</a:themeElements>'
        '</a:theme>'
    )


def _slide_master_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldMaster '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld>'
        '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree>'
        '</p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
        'accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        '</p:sldMaster>'
    )


def _slide_master_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="../theme/theme1.xml"/>'
        '</Relationships>'
    )


def _slide_layout_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldLayout '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'type="blank" preserve="1">'
        '<p:cSld name="Blank">'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree>'
        '</p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sldLayout>'
    )


def _slide_layout_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        'Target="../slideMasters/slideMaster1.xml"/>'
        '</Relationships>'
    )


def _slide_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>'
    )


# --------------------------------------------------------------------------- #
# Slide shape helpers
# --------------------------------------------------------------------------- #

def _text_shape(
    sp_id: int,
    name: str,
    x: int,
    y: int,
    cx: int,
    cy: int,
    paragraphs_xml: str,
) -> str:
    return (
        f'<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="{_esc(name)}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/>'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr wrap="square" rtlCol="0" anchor="t"/>'
        f'<a:lstStyle/>'
        f'{paragraphs_xml}'
        f'</p:txBody>'
        f'</p:sp>'
    )


def _title_paragraph(text: str) -> str:
    return (
        '<a:p>'
        '<a:pPr algn="l"/>'
        f'<a:r><a:rPr lang="en-US" sz="4000" b="1"/><a:t>{_esc(text)}</a:t></a:r>'
        '</a:p>'
    )


def _bullet_paragraph(text: str) -> str:
    return (
        '<a:p>'
        '<a:pPr marL="285750" indent="-285750"><a:buChar char="\u2022"/></a:pPr>'
        f'<a:r><a:rPr lang="en-US" sz="2000"/><a:t>{_esc(text)}</a:t></a:r>'
        '</a:p>'
    )


def _image_placeholder(sp_id: int, description: str) -> str:
    """
    Dashed-border rectangle acting as an image placeholder.

    Phase 2 hookup: replace this with a <p:pic> referencing a real image part
    (ppt/media/imageN.jpg) and a matching relationship in the slide's _rels.
    The x/y/cx/cy geometry can be reused as-is.
    """
    label = f"[Image] {description}"
    body = (
        '<a:p><a:pPr algn="ctr"/>'
        '<a:r><a:rPr lang="en-US" sz="1600" i="1">'
        '<a:solidFill><a:srgbClr val="6B7280"/></a:solidFill>'
        f'</a:rPr><a:t>{_esc(label)}</a:t></a:r>'
        '</a:p>'
    )
    return (
        f'<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="Image Placeholder"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{IMAGE_X}" y="{IMAGE_Y}"/>'
        f'<a:ext cx="{IMAGE_W}" cy="{IMAGE_H}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="F3F4F6"/></a:solidFill>'
        f'<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">'
        f'<a:solidFill><a:srgbClr val="9CA3AF"/></a:solidFill>'
        f'<a:prstDash val="dash"/>'
        f'</a:ln>'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>'
        f'<a:lstStyle/>'
        f'{body}'
        f'</p:txBody>'
        f'</p:sp>'
    )


def _slide_xml(index: int, slide: SlideInput, deck_title: str) -> str:
    has_image = bool(slide.image_description)
    body_w = BODY_W_HALF if has_image else BODY_W_FULL

    title_text = slide.title if index > 0 else (slide.title or deck_title)
    paragraphs = [_title_paragraph(title_text)]
    title_shape = _text_shape(2, "Title", MARGIN_X, TITLE_Y, BODY_W_FULL, TITLE_H, "".join(paragraphs))

    bullet_paragraphs = "".join(_bullet_paragraph(b) for b in slide.bullets) or (
        '<a:p><a:r><a:rPr lang="en-US" sz="2000"/><a:t></a:t></a:r></a:p>'
    )
    body_shape = _text_shape(3, "Body", MARGIN_X, BODY_Y, body_w, BODY_H, bullet_paragraphs)

    image_shape = _image_placeholder(4, slide.image_description) if has_image else ""

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld>'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{title_shape}{body_shape}{image_shape}'
        '</p:spTree>'
        '</p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sld>'
    )


# --------------------------------------------------------------------------- #
# Package assembly
# --------------------------------------------------------------------------- #

def build_pptx(deck: DeckInput) -> bytes:
    """Assemble every OOXML part and zip them into a valid .pptx binary."""
    # Prepend a synthetic title slide as slide 1.
    title_slide = SlideInput(
        title=deck.title,
        bullets=[],
        image_description=None,
    )
    all_slides = [title_slide] + list(deck.slides)
    n = len(all_slides)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _content_types_xml(n))
        z.writestr("_rels/.rels", _root_rels_xml())
        z.writestr("ppt/presentation.xml", _presentation_xml(n))
        z.writestr("ppt/_rels/presentation.xml.rels", _presentation_rels_xml(n))
        z.writestr("ppt/theme/theme1.xml", _theme_xml())
        z.writestr("ppt/slideMasters/slideMaster1.xml", _slide_master_xml())
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", _slide_master_rels_xml())
        z.writestr("ppt/slideLayouts/slideLayout1.xml", _slide_layout_xml())
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", _slide_layout_rels_xml())
        for i, slide in enumerate(all_slides):
            z.writestr(f"ppt/slides/slide{i + 1}.xml", _slide_xml(i, slide, deck.title))
            z.writestr(f"ppt/slides/_rels/slide{i + 1}.xml.rels", _slide_rels_xml())

    return buf.getvalue()
