"""Phase 1 — compat.py support-level classification.

Self-contained — no pytest.  Run:  python test_compat.py
"""
import io
import os
import sys
import tempfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from PIL import Image

sys.path.insert(0, str(Path(__file__).parent))

from compat import SupportLevel, classify_shape, demo as compat_demo

_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_passed = _failed = 0


def check(name, cond, detail=""):
    global _passed, _failed
    if cond:
        _passed += 1
        print(f"  [OK]   {name}")
    else:
        _failed += 1
        print(f"  [FAIL] {name}  {detail}")


def _group_xml(with_text: bool) -> str:
    if with_text:
        sp = (
            '<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="TB{i}"/><p:cNvSpPr txBox="1"/>'
            '<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="{y}"/>'
            '<a:ext cx="4000000" cy="900000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:r><a:t>Grouped line {i}</a:t></a:r></a:p></p:txBody></p:sp>'
        )
    else:
        # auto-shapes with NO txBody -> has_text_frame is False
        sp = (
            '<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Rect{i}"/><p:cNvSpPr/>'
            '<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="{y}"/>'
            '<a:ext cx="1000000" cy="500000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>'
        )
    return (
        f'<p:grpSp xmlns:p="{_P}" xmlns:a="{_A}">'
        '<p:nvGrpSpPr><p:cNvPr id="40" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="4000000" cy="2000000"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="4000000" cy="2000000"/></a:xfrm></p:grpSpPr>'
        + sp.format(i=41, y=0) + sp.format(i=42, y=1000000) +
        '</p:grpSp>'
    )


def _smartart_xml() -> str:
    return (
        f'<p:graphicFrame xmlns:p="{_P}" xmlns:a="{_A}" xmlns:dgm="{_DGM}" xmlns:r="{_R}">'
        '<p:nvGraphicFramePr><p:cNvPr id="99" name="SmartArt 1"/><p:cNvGraphicFramePr/><p:nvPr/>'
        '</p:nvGraphicFramePr>'
        '<p:xfrm><a:off x="457200" y="457200"/><a:ext cx="4572000" cy="2743200"/></p:xfrm>'
        f'<a:graphic><a:graphicData uri="{_DGM}">'
        '<dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>'
        '</a:graphicData></a:graphic></p:graphicFrame>'
    )


def _first_shape(slide):
    # injected XML is appended last in the shape tree
    return list(slide.shapes)[-1]


def run():
    compat_demo()  # vocabulary self-check

    prs = Presentation()

    # SUPPORTED: text placeholder
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Hi"
    check("text -> SUPPORTED", classify_shape(s.shapes.title).level == SupportLevel.SUPPORTED)

    # SUPPORTED: image
    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    buf = io.BytesIO()
    Image.new("RGB", (60, 60), (10, 120, 200)).save(buf, format="PNG")
    buf.seek(0)
    pic = s2.shapes.add_picture(buf, Inches(1), Inches(1), Inches(2), Inches(2))
    check("image -> SUPPORTED", classify_shape(pic).level == SupportLevel.SUPPORTED)

    # SUPPORTED: table
    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    tbl = s3.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(2))
    check("table -> SUPPORTED", classify_shape(tbl).level == SupportLevel.SUPPORTED)

    # SKIPPED_WITH_WARNING: chart
    s4 = prs.slides.add_slide(prs.slide_layouts[5])
    cd = ChartData()
    cd.categories = ["A", "B"]
    cd.add_series("S", (1, 2))
    ch = s4.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(1), Inches(1), Inches(4), Inches(3), cd)
    c = classify_shape(ch)
    check("chart -> SKIPPED_WITH_WARNING", c.level == SupportLevel.SKIPPED_WITH_WARNING and c.kind == "chart")

    # SKIPPED_WITH_WARNING: SmartArt
    s5 = prs.slides.add_slide(prs.slide_layouts[5])
    s5.shapes._spTree.append(etree.fromstring(_smartart_xml()))
    c = classify_shape(_first_shape(s5))
    check("smartart -> SKIPPED_WITH_WARNING", c.level == SupportLevel.SKIPPED_WITH_WARNING and c.kind == "smartart")

    # PRESERVED_AS_IS: decorative auto-shape (line)
    s6 = prs.slides.add_slide(prs.slide_layouts[5])
    from pptx.enum.shapes import MSO_CONNECTOR
    conn = s6.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1), Inches(1), Inches(3), Inches(1))
    check("line/connector -> PRESERVED_AS_IS", classify_shape(conn).level == SupportLevel.PRESERVED_AS_IS)

    # PARTIALLY_SUPPORTED: group with text
    s7 = prs.slides.add_slide(prs.slide_layouts[5])
    s7.shapes._spTree.append(etree.fromstring(_group_xml(with_text=True)))
    c = classify_shape(_first_shape(s7))
    check("group w/ text -> PARTIALLY_SUPPORTED", c.level == SupportLevel.PARTIALLY_SUPPORTED and c.kind == "group")

    # SKIPPED_WITH_WARNING: empty group
    s8 = prs.slides.add_slide(prs.slide_layouts[5])
    s8.shapes._spTree.append(etree.fromstring(_group_xml(with_text=False)))
    c = classify_shape(_first_shape(s8))
    check("empty group -> SKIPPED_WITH_WARNING", c.level == SupportLevel.SKIPPED_WITH_WARNING and c.kind == "group")

    print(f"\n{'='*40}\ncompat: {_passed} passed, {_failed} failed")
    sys.exit(1 if _failed else 0)


if __name__ == "__main__":
    run()
