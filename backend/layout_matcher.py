"""Deterministic layout matching: source slide type → template layout."""

from enum import Enum
from dataclasses import dataclass, field
from typing import Optional
from pptx import Presentation
from pptx.util import Emu


class SlideType(Enum):
    TITLE_BODY = "title_body"          # A: title + text content
    TITLE_TWO_COL = "title_two_col"    # B: title + two columns
    TITLE_IMAGE = "title_image"        # C: title + image only
    TITLE_TEXT_IMAGE = "title_text_image"  # D: title + text + image
    TITLE_TABLE = "title_table"        # E: title + table
    TITLE_ONLY = "title_only"          # F: section/divider
    BLANK = "blank"                    # G: custom/blank


@dataclass
class TemplateLayout:
    index: int
    name: str
    slide_type: SlideType
    has_title_ph: bool = False
    has_body_ph: bool = False
    has_image_ph: bool = False
    placeholder_count: int = 0


def classify_source_slide(parsed_slide) -> SlideType:
    has_title = parsed_slide.has_title
    has_body = parsed_slide.has_body
    has_images = parsed_slide.has_images
    has_table = parsed_slide.has_table
    body_count = len([b for b in parsed_slide.body_boxes if b.full_text.strip()])

    if has_table:
        return SlideType.TITLE_TABLE
    if has_images and has_body:
        return SlideType.TITLE_TEXT_IMAGE
    if has_images and not has_body:
        return SlideType.TITLE_IMAGE
    if body_count >= 2:
        return SlideType.TITLE_TWO_COL
    if has_body:
        return SlideType.TITLE_BODY
    if has_title:
        return SlideType.TITLE_ONLY
    return SlideType.BLANK


def _classify_layout(layout) -> TemplateLayout:
    placeholders = list(layout.placeholders)

    def _ph(ph):
        try:
            return ph.placeholder_format if ph.is_placeholder else None
        except Exception:
            return None

    has_title = any(_ph(ph) and _ph(ph).idx == 0 for ph in placeholders)
    has_body = any(_ph(ph) and _ph(ph).idx == 1 for ph in placeholders)
    has_image = any(
        _ph(ph) and _ph(ph).type is not None and
        str(_ph(ph).type) in ("PICTURE(18)", "OBJECT(14)", "MEDIA(16)")
        for ph in placeholders
    )

    count = len(placeholders)
    name = layout.name.lower()

    # Determine slide type from layout name keywords first
    if any(k in name for k in ("title slide", "section", "divider")):
        slide_type = SlideType.TITLE_ONLY
    elif any(k in name for k in ("two content", "comparison", "two col")):
        slide_type = SlideType.TITLE_TWO_COL
    elif any(k in name for k in ("picture", "image", "content with caption")):
        slide_type = SlideType.TITLE_TEXT_IMAGE if has_body else SlideType.TITLE_IMAGE
    elif any(k in name for k in ("blank",)):
        slide_type = SlideType.BLANK
    elif has_title and has_body:
        slide_type = SlideType.TITLE_BODY
    elif has_title and not has_body:
        slide_type = SlideType.TITLE_ONLY
    else:
        slide_type = SlideType.BLANK

    return TemplateLayout(
        index=None,  # set by caller
        name=layout.name,
        slide_type=slide_type,
        has_title_ph=has_title,
        has_body_ph=has_body,
        has_image_ph=has_image,
        placeholder_count=count,
    )


def analyze_template(template_path: str) -> list[TemplateLayout]:
    prs = Presentation(template_path)
    layouts = []
    for i, layout in enumerate(prs.slide_layouts):
        tl = _classify_layout(layout)
        tl.index = i
        layouts.append(tl)
    return layouts


# Priority order when multiple layouts match the same type
_TYPE_FALLBACK = [
    SlideType.TITLE_BODY,
    SlideType.TITLE_ONLY,
    SlideType.BLANK,
]


def find_best_layout(slide_type: SlideType, layouts: list[TemplateLayout]) -> TemplateLayout:
    """Return best matching template layout for a given slide type."""
    # Exact match first
    exact = [l for l in layouts if l.slide_type == slide_type]
    if exact:
        if slide_type == SlideType.TITLE_ONLY:
            # Prefer section-style layouts over the cover "Title Slide"
            non_cover = [l for l in exact if "title slide" not in l.name.lower()]
            return non_cover[0] if non_cover else exact[0]
        return exact[0]

    # Fallback chain
    fallbacks = {
        SlideType.TITLE_TWO_COL: [SlideType.TITLE_BODY, SlideType.BLANK],
        SlideType.TITLE_IMAGE: [SlideType.TITLE_BODY, SlideType.BLANK],
        SlideType.TITLE_TEXT_IMAGE: [SlideType.TITLE_BODY, SlideType.BLANK],
        SlideType.TITLE_TABLE: [SlideType.TITLE_BODY, SlideType.BLANK],
        SlideType.TITLE_ONLY: [SlideType.TITLE_BODY, SlideType.BLANK],
        SlideType.BLANK: [SlideType.TITLE_BODY],
    }
    for fallback_type in fallbacks.get(slide_type, []):
        candidates = [l for l in layouts if l.slide_type == fallback_type]
        if candidates:
            return candidates[0]

    # Last resort: first layout
    return layouts[0]
