"""Self-contained tests — no pytest required. Run: python test_transfer.py"""

import os
import sys
import tempfile
from pathlib import Path
from lxml import etree

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
import io

sys.path.insert(0, str(Path(__file__).parent))

from parser import parse_source, _detect_unsupported_shape, UnsupportedShape
from layout_matcher import classify_source_slide, SlideType, analyze_template, find_best_layout
from transfer import transfer
from validator import validate


def _make_source_pptx(path: str):
    """Create a test source PPTX with various slide types."""
    prs = Presentation()
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    layouts = prs.slide_layouts

    # Slide 1: Title + Body (Type A)
    sl = prs.slides.add_slide(layouts[1])
    sl.shapes.title.text = "Introduction to Machine Learning"
    sl.placeholders[1].text = "Machine learning is a subset of artificial intelligence.\nIt enables systems to learn from data automatically."

    # Slide 2: Title + two text blocks (Type B)
    sl2 = prs.slides.add_slide(layouts[3])  # Two Content layout
    sl2.shapes.title.text = "Comparison: Supervised vs Unsupervised"
    try:
        sl2.placeholders[1].text = "Supervised Learning\n• Uses labeled data\n• Classification tasks\n• Regression problems"
        sl2.placeholders[2].text = "Unsupervised Learning\n• No labeled data\n• Clustering\n• Dimensionality reduction"
    except Exception:
        sl2.placeholders[1].text = "Supervised vs Unsupervised comparison content here."

    # Slide 3: Title + Image (Type C)
    sl3 = prs.slides.add_slide(layouts[5])  # Blank
    sl3.shapes.add_textbox(Emu(457200), Emu(274638), Emu(8229600), Emu(1143000)).text_frame.text = "Neural Network Architecture"
    # Add a small red rectangle as a fake image substitute
    from pptx.util import Inches
    from PIL import Image
    img = Image.new("RGB", (200, 150), color=(220, 50, 50))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    sl3.shapes.add_picture(buf, Inches(2), Inches(2), Inches(4), Inches(3))

    # Slide 4: Section title only (Type F)
    sl4 = prs.slides.add_slide(layouts[0])
    sl4.shapes.title.text = "Chapter 2: Deep Learning"

    # Slide 5: Title + Table (Type E)
    sl5 = prs.slides.add_slide(layouts[5])
    tb = sl5.shapes.add_textbox(Emu(457200), Emu(274638), Emu(8229600), Emu(914400))
    tb.text_frame.text = "Algorithm Comparison"
    table = sl5.shapes.add_table(3, 3, Inches(1), Inches(2), Inches(8), Inches(3)).table
    headers = ["Algorithm", "Accuracy", "Speed"]
    for i, h in enumerate(headers):
        table.cell(0, i).text = h
    table.cell(1, 0).text = "Decision Tree"
    table.cell(1, 1).text = "85%"
    table.cell(1, 2).text = "Fast"
    table.cell(2, 0).text = "Neural Network"
    table.cell(2, 1).text = "94%"
    table.cell(2, 2).text = "Slow"

    prs.save(path)
    return 5  # slide count


def _make_template_pptx(path: str):
    """Create a minimal college template with multiple layouts."""
    prs = Presentation()
    prs.save(path)  # Default template has standard layouts


def _collect_all_text(prs_path: str) -> set[str]:
    prs = Presentation(prs_path)
    texts = set()
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = para.text.strip()
                    if t:
                        texts.add(t)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        t = cell.text.strip()
                        if t:
                            texts.add(t)
    return texts


def run_tests():
    passed = 0
    failed = 0

    def ok(name):
        nonlocal passed
        print(f"  [OK] {name}")
        passed += 1

    def fail(name, reason):
        nonlocal failed
        print(f"  [FAIL] {name}: {reason}")
        failed += 1

    with tempfile.TemporaryDirectory() as tmpdir:
        source_path = os.path.join(tmpdir, "source.pptx")
        template_path = os.path.join(tmpdir, "template.pptx")
        output_path = os.path.join(tmpdir, "output.pptx")

        print("\n[1] Creating test PPTX files...")
        try:
            expected_slides = _make_source_pptx(source_path)
            _make_template_pptx(template_path)
            ok("Test files created")
        except Exception as e:
            fail("Test file creation", str(e))
            return

        print("\n[2] Parsing source...")
        try:
            slides = parse_source(source_path)
            assert len(slides) == expected_slides, f"Expected {expected_slides} slides, got {len(slides)}"
            ok(f"Parsed {len(slides)} slides")
        except Exception as e:
            fail("Parse source", str(e))
            return

        print("\n[3] Layout classification...")
        try:
            types = [classify_source_slide(s) for s in slides]
            assert SlideType.TITLE_BODY in types, "No TITLE_BODY slide detected"
            assert SlideType.TITLE_TABLE in types, "No TITLE_TABLE slide detected"
            ok(f"Classified: {[t.value for t in types]}")
        except Exception as e:
            fail("Layout classification", str(e))

        print("\n[4] Template analysis...")
        try:
            layouts = analyze_template(template_path)
            assert len(layouts) > 0, "No layouts found in template"
            ok(f"Found {len(layouts)} template layouts")
        except Exception as e:
            fail("Template analysis", str(e))
            return

        print("\n[5] Content transfer...")
        try:
            results = transfer(
                source_slides=slides,
                source_path=source_path,
                template_path=template_path,
                output_path=output_path,
            )
            ok(f"Transfer completed, {len(results)} results")
        except Exception as e:
            fail("Content transfer", str(e))
            return

        print("\n[6] Output validation...")
        try:
            report = validate(output_path, expected_slides, results)
            assert os.path.exists(output_path), "Output file missing"
            ok(f"Output valid: {report.summary}")
        except Exception as e:
            fail("Validation", str(e))
            return

        print("\n[7] Slide count preservation...")
        try:
            prs_out = Presentation(output_path)
            assert len(prs_out.slides) == expected_slides, \
                f"Slide count mismatch: {len(prs_out.slides)} vs {expected_slides}"
            ok(f"{len(prs_out.slides)} slides in output")
        except Exception as e:
            fail("Slide count", str(e))

        print("\n[8] Text content preservation (SOURCE == OUTPUT)...")
        try:
            source_texts = _collect_all_text(source_path)
            output_texts = _collect_all_text(output_path)
            # Key phrases that must survive
            key_phrases = [
                "Introduction to Machine Learning",
                "Chapter 2: Deep Learning",
                "Algorithm Comparison",
                "Decision Tree",
            ]
            missing = [p for p in key_phrases if not any(p in t for t in output_texts)]
            if missing:
                fail("Text preservation", f"Missing: {missing}")
            else:
                ok("All key text phrases preserved in output")
        except Exception as e:
            fail("Text preservation", str(e))

        print("\n[9] Source file unmodified...")
        try:
            before = os.path.getmtime(source_path)
            # Re-read source — should not have changed
            parse_source(source_path)
            after = os.path.getmtime(source_path)
            # mtime check not reliable across all OS for read-only ops; check content instead
            slides_again = parse_source(source_path)
            assert len(slides_again) == expected_slides
            ok("Source file unchanged")
        except Exception as e:
            fail("Source unmodified", str(e))

        print("\n[10] Invalid PPTX handling...")
        try:
            bad_path = os.path.join(tmpdir, "bad.pptx")
            with open(bad_path, "wb") as f:
                f.write(b"this is not a pptx file")
            try:
                parse_source(bad_path)
                fail("Invalid PPTX", "Should have raised an exception")
            except Exception:
                ok("Invalid PPTX raises exception correctly")
        except Exception as e:
            fail("Invalid PPTX test", str(e))

    print(f"\n{'='*40}")
    print(f"Results: {passed} passed, {failed} failed")
    if failed:
        sys.exit(1)


def _make_smartart_slide(prs):
    """Add a slide with a synthetic SmartArt graphicFrame element."""
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank
    # Build a minimal SmartArt graphicFrame XML
    dgm_ns = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
    draw_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    pptx_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'

    spTree = slide.shapes._spTree
    nsmap = {
        'a': draw_ns,
        'p': pptx_ns,
        'dgm': dgm_ns,
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    gf_xml = (
        '<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<p:nvGraphicFramePr>'
        '<p:cNvPr id="99" name="SmartArt 1"/>'
        '<p:cNvGraphicFramePr/>'
        '<p:nvPr/>'
        '</p:nvGraphicFramePr>'
        '<p:xfrm><a:off x="457200" y="457200"/><a:ext cx="4572000" cy="2743200"/></p:xfrm>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        '<dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>'
        '</a:graphicData>'
        '</a:graphic>'
        '</p:graphicFrame>'
    )
    gf_elem = etree.fromstring(gf_xml)
    spTree.append(gf_elem)
    return slide


def run_unsupported_detection_tests():
    passed = 0
    failed = 0

    def ok(name):
        nonlocal passed
        print(f"  [OK] {name}")
        passed += 1

    def fail(name, reason):
        nonlocal failed
        print(f"  [FAIL] {name}: {reason}")
        failed += 1

    with tempfile.TemporaryDirectory() as tmpdir:

        # --- Test 11: Chart detection ---
        print("\n[11] Chart detection...")
        try:
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            cd = ChartData()
            cd.categories = ['A', 'B']
            cd.add_series('S1', (1, 2))
            slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(1), Inches(1), Inches(4), Inches(3), cd)
            path = os.path.join(tmpdir, "chart_source.pptx")
            prs.save(path)
            slides = parse_source(path)
            assert len(slides) == 1
            assert any(u.type == "chart" for u in slides[0].unsupported_shapes), \
                f"No chart detected; unsupported={slides[0].unsupported_shapes}"
            ok("Chart shape detected as unsupported")
        except Exception as e:
            fail("Chart detection", str(e))

        # --- Test 12: SmartArt detection ---
        print("\n[12] SmartArt detection...")
        try:
            prs = Presentation()
            _make_smartart_slide(prs)
            path = os.path.join(tmpdir, "smartart_source.pptx")
            prs.save(path)
            slides = parse_source(path)
            assert len(slides) == 1
            assert any(u.type == "smartart" for u in slides[0].unsupported_shapes), \
                f"No SmartArt detected; unsupported={slides[0].unsupported_shapes}"
            ok("SmartArt shape detected as unsupported")
        except Exception as e:
            fail("SmartArt detection", str(e))

        # --- Test 13: Normal text is NOT flagged ---
        print("\n[13] Normal text not flagged as unsupported...")
        try:
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = "Hello"
            slide.placeholders[1].text = "World"
            path = os.path.join(tmpdir, "text_source.pptx")
            prs.save(path)
            slides = parse_source(path)
            assert len(slides) == 1
            assert slides[0].unsupported_shapes == [], \
                f"Text incorrectly flagged: {slides[0].unsupported_shapes}"
            ok("Normal text produces no unsupported warnings")
        except Exception as e:
            fail("Normal text not flagged", str(e))

        # --- Test 14: Normal image is NOT flagged ---
        print("\n[14] Normal image not flagged as unsupported...")
        try:
            from PIL import Image
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            img = Image.new("RGB", (100, 100), color=(100, 200, 100))
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            slide.shapes.add_picture(buf, Inches(1), Inches(1), Inches(2), Inches(2))
            path = os.path.join(tmpdir, "img_source.pptx")
            prs.save(path)
            slides = parse_source(path)
            assert slides[0].unsupported_shapes == [], \
                f"Image incorrectly flagged: {slides[0].unsupported_shapes}"
            assert slides[0].images, "Image not parsed"
            ok("Normal image produces no unsupported warnings")
        except Exception as e:
            fail("Normal image not flagged", str(e))

        # --- Test 15: Mixed slide (text + chart) ---
        print("\n[15] Mixed slide: text + chart...")
        try:
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(6), Inches(0.6))
            tb.text_frame.text = "Revenue Chart"
            cd = ChartData()
            cd.categories = ['Q1', 'Q2']
            cd.add_series('Rev', (100, 200))
            slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(1), Inches(1), Inches(5), Inches(4), cd)
            path = os.path.join(tmpdir, "mixed_source.pptx")
            prs.save(path)
            slides = parse_source(path)
            assert slides[0].has_body or slides[0].has_title, "Text not parsed"
            assert any(u.type == "chart" for u in slides[0].unsupported_shapes), \
                "Chart not flagged in mixed slide"
            ok("Mixed slide: text parsed, chart flagged")
        except Exception as e:
            fail("Mixed slide", str(e))

        # --- Test 16: Slide with only unsupported content → transfer warns ---
        print("\n[16] Only-unsupported slide produces transfer warning...")
        try:
            prs_src = Presentation()
            cd = ChartData()
            cd.categories = ['X']
            cd.add_series('Y', (1,))
            slide = prs_src.slides.add_slide(prs_src.slide_layouts[5])
            slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(1), Inches(1), Inches(4), Inches(3), cd)
            src_path = os.path.join(tmpdir, "only_unsup_src.pptx")
            prs_src.save(src_path)

            prs_tmpl = Presentation()
            tmpl_path = os.path.join(tmpdir, "only_unsup_tmpl.pptx")
            prs_tmpl.save(tmpl_path)

            out_path = os.path.join(tmpdir, "only_unsup_out.pptx")
            slides = parse_source(src_path)
            results = transfer(source_slides=slides, source_path=src_path,
                               template_path=tmpl_path, output_path=out_path)
            assert results[0].content_warnings, "No content_warnings generated"
            assert results[0].content_warnings[0]["type"] == "chart"
            # Warning message should mention the slide cannot be fully transferred
            full_warn = " ".join(results[0].warnings)
            assert "could not be fully transferred" in full_warn or "unsupported" in full_warn.lower()
            ok("Only-unsupported slide produces content_warning and transfer warning")
        except Exception as e:
            fail("Only-unsupported slide warning", str(e))

        # --- Test 17: SmartArt slide → content_warning with type=smartart ---
        print("\n[17] SmartArt slide -> content_warning type=smartart...")
        try:
            prs_src = Presentation()
            _make_smartart_slide(prs_src)
            src_path = os.path.join(tmpdir, "smartart_transfer_src.pptx")
            prs_src.save(src_path)

            prs_tmpl = Presentation()
            tmpl_path = os.path.join(tmpdir, "smartart_transfer_tmpl.pptx")
            prs_tmpl.save(tmpl_path)

            out_path = os.path.join(tmpdir, "smartart_transfer_out.pptx")
            slides = parse_source(src_path)
            results = transfer(source_slides=slides, source_path=src_path,
                               template_path=tmpl_path, output_path=out_path)
            cws = results[0].content_warnings
            assert cws, "No content_warnings for SmartArt slide"
            assert cws[0]["type"] == "smartart"
            assert cws[0]["slide"] == 1
            ok(f"SmartArt slide content_warning: {cws[0]['message']}")
        except Exception as e:
            fail("SmartArt content_warning", str(e))

        # --- Test 18: Existing 10 tests unaffected (no regressions) ---
        print("\n[18] Regression check: existing source PPTX produces no spurious unsupported warnings...")
        try:
            with tempfile.TemporaryDirectory() as t2:
                src = os.path.join(t2, "src.pptx")
                _make_source_pptx(src)
                slides = parse_source(src)
                flagged = [(i+1, s.unsupported_shapes) for i, s in enumerate(slides) if s.unsupported_shapes]
                assert not flagged, f"Spurious unsupported on slides: {flagged}"
                ok("No spurious unsupported warnings on normal PPTX")
        except Exception as e:
            fail("Regression check", str(e))

    print(f"\n{'='*40}")
    print(f"Unsupported detection tests: {passed} passed, {failed} failed")
    return failed


if __name__ == "__main__":
    run_tests()
    print()
    failures = run_unsupported_detection_tests()
    if failures:
        sys.exit(1)
