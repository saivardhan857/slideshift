"""V2 compatibility model — classify a source PPTX shape by how well the
transfer engine can carry it into the college template.

This is a vocabulary + a pure classifier. It does not move any content; it
tells the parser/transfer layers what each shape *is* so they can salvage,
skip, or warn honestly (rule: never claim an unsupported object is supported).

Levels
------
SUPPORTED            transfer engine handles it fully (text, image, table,
                     picture placeholder)
PARTIALLY_SUPPORTED  some content can be salvaged, the rest is lost
                     (e.g. a grouped shape whose text we lift out)
PRESERVED_AS_IS      carries no transferable content and is intentionally left
                     out — nothing is lost, so no warning
                     (decorative lines / connectors / plain auto-shapes)
UNSUPPORTED          recognised shape kind the engine cannot transfer
SKIPPED_WITH_WARNING content the engine cannot transfer AND the user should be
                     told (chart, SmartArt, OLE object, media, WordArt, an
                     empty group)

Byte-level preservation of complex objects (charts / OLE) into the rebuilt
deck is deliberately out of scope for Phase 1.
"""

from dataclasses import dataclass
from enum import Enum
from typing import Optional

from pptx.enum.shapes import MSO_SHAPE_TYPE

_SMARTART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
_DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'


class SupportLevel(str, Enum):
    SUPPORTED = "SUPPORTED"
    PARTIALLY_SUPPORTED = "PARTIALLY_SUPPORTED"
    PRESERVED_AS_IS = "PRESERVED_AS_IS"
    UNSUPPORTED = "UNSUPPORTED"
    SKIPPED_WITH_WARNING = "SKIPPED_WITH_WARNING"


@dataclass
class ShapeCompat:
    level: SupportLevel
    kind: str          # "text" | "image" | "table" | "chart" | "smartart" |
                       # "group" | "ole" | "media" | "wordart" | "decorative" | "unknown"
    label: str         # human-readable
    detail: str = ""   # optional extra context


def _is_smartart(shape) -> bool:
    try:
        el = shape.element
        if el.find(f'.//{{{_SMARTART_NS}}}relIds') is not None:
            return True
        gd = el.find(f'.//{{{_DRAWING_NS}}}graphicData')
        return gd is not None and gd.get('uri') == _SMARTART_NS
    except Exception:
        return False


def _group_has_text(shape) -> bool:
    """True if a group (recursively) contains at least one non-empty text frame."""
    try:
        children = list(shape.shapes)
    except Exception:
        return False
    for sub in children:
        try:
            if sub.shape_type == MSO_SHAPE_TYPE.GROUP:
                if _group_has_text(sub):
                    return True
            elif sub.has_text_frame and sub.text_frame.text.strip():
                return True
        except Exception:
            continue
    return False


def classify_shape(shape) -> ShapeCompat:
    """Classify a single python-pptx shape. Pure — reads, never mutates."""
    # --- fully supported content ---
    try:
        if shape.has_table:
            return ShapeCompat(SupportLevel.SUPPORTED, "table", "Table")
    except Exception:
        pass

    try:
        st = shape.shape_type
    except Exception:
        st = None

    if st == MSO_SHAPE_TYPE.PICTURE:
        return ShapeCompat(SupportLevel.SUPPORTED, "image", "Image")

    try:
        if shape.is_placeholder and shape.element.tag.endswith('}pic'):
            return ShapeCompat(SupportLevel.SUPPORTED, "image", "Image (picture placeholder)")
    except Exception:
        pass

    try:
        if shape.has_text_frame:
            return ShapeCompat(SupportLevel.SUPPORTED, "text", "Text")
    except Exception:
        pass

    # --- groups: salvage what we can ---
    if st == MSO_SHAPE_TYPE.GROUP:
        if _group_has_text(shape):
            return ShapeCompat(
                SupportLevel.PARTIALLY_SUPPORTED, "group", "Grouped Shapes",
                "text content salvaged; other elements not transferred",
            )
        return ShapeCompat(
            SupportLevel.SKIPPED_WITH_WARNING, "group", "Grouped Shapes",
            "no transferable text found",
        )

    # --- content-bearing objects the engine cannot transfer ---
    if st == MSO_SHAPE_TYPE.CHART:
        return ShapeCompat(SupportLevel.SKIPPED_WITH_WARNING, "chart", "Chart")
    if st == MSO_SHAPE_TYPE.TEXT_EFFECT:
        return ShapeCompat(SupportLevel.SKIPPED_WITH_WARNING, "wordart", "WordArt")
    if st in (MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT, MSO_SHAPE_TYPE.LINKED_OLE_OBJECT):
        return ShapeCompat(SupportLevel.SKIPPED_WITH_WARNING, "ole", "Embedded Object")
    if st == MSO_SHAPE_TYPE.MEDIA:
        return ShapeCompat(SupportLevel.SKIPPED_WITH_WARNING, "media", "Media (Video/Audio)")
    if _is_smartart(shape):
        return ShapeCompat(SupportLevel.SKIPPED_WITH_WARNING, "smartart", "SmartArt")

    # --- decorative, no transferable content: leave it out, no warning ---
    if st in (MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.AUTO_SHAPE,
              MSO_SHAPE_TYPE.FREEFORM, MSO_SHAPE_TYPE.PLACEHOLDER):
        return ShapeCompat(SupportLevel.PRESERVED_AS_IS, "decorative",
                           "Decorative shape", "no transferable content")

    return ShapeCompat(SupportLevel.UNSUPPORTED, "unknown", "Unknown shape")


def demo():
    """Self-check — runs the enum through the level vocabulary."""
    levels = {l.value for l in SupportLevel}
    assert levels == {
        "SUPPORTED", "PARTIALLY_SUPPORTED", "PRESERVED_AS_IS",
        "UNSUPPORTED", "SKIPPED_WITH_WARNING",
    }
    c = ShapeCompat(SupportLevel.SUPPORTED, "text", "Text")
    assert c.level == "SUPPORTED"          # str-Enum compares to plain strings
    assert SupportLevel.SKIPPED_WITH_WARNING == "SKIPPED_WITH_WARNING"
    print("[OK] compat vocabulary")


if __name__ == "__main__":
    demo()
