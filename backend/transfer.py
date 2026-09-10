"""Core content transfer engine."""

import io
import zipfile
from typing import Optional
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn as _qn


def _dedupe_pptx(path: str):
    """Remove duplicate ZIP entries from a PPTX, keeping the last occurrence of
    each name (python-pptx can serialise a stripped slide's parts twice, which
    corrupts the file).

    Streamed copy: one pass over infolist() (metadata only) picks the last index
    per filename, a second pass copies just those entries. Peak memory is the
    single largest entry rather than the whole decompressed package, which
    matters for large image-heavy decks on a small instance.
    """
    with zipfile.ZipFile(path, 'r') as src:
        infos = src.infolist()
        last_idx = {}
        for i, info in enumerate(infos):
            last_idx[info.filename] = i
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as dst:
            for i, info in enumerate(infos):
                if last_idx[info.filename] == i:
                    dst.writestr(info, src.read(info.filename))
    with open(path, 'wb') as f:
        f.write(buf.getvalue())

from parser import ParsedSlide, TextBox, Paragraph, TextRun, ImageData, TableData
from layout_matcher import TemplateLayout, SlideType, find_best_layout, analyze_template, classify_source_slide
from layout_safety import plan_fit


# ponytail: hardcoded for CUCOM template whose logo occupies ~0-1.88in in background image
_TITLE_MIN_LEFT = Inches(2.0)

# ponytail: measured from the CUCOM background image (8000x4500 full-bleed).
# The saturated red wedge's left edge, within the body placeholder's vertical
# span (y 2.0-6.76in), runs from ~9.30in at the top to ~11.1in mid-slide.
# Clamp body text to the worst-case edge minus a ~0.2in margin so lecture text
# never crosses the red. The wedge itself stays fully visible. Titles, images
# and tables are not touched.  upgrade path: per-line diagonal clamp if the
# uniform right edge ever wastes too much width on short slides.
_BODY_SAFE_LEFT = Inches(0.92)    # CUCOM title/body left edge
_BODY_SAFE_RIGHT = Inches(9.1)
_COL_GUTTER = Inches(0.3)


def _place_ph(ph, left, width):
    """Reposition a placeholder while preserving its inherited top/height.

    Setting .left/.width materialises an <a:xfrm> that would otherwise zero the
    y-offset and height (same gotcha the title-repositioning code works around).
    """
    top, height = ph.top, ph.height
    ph.left = left
    ph.top = top
    ph.height = height
    ph.width = max(width, Inches(1.5))

_NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
_NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def _extract_design_images(template_prs):
    """Extract background picture shapes from a representative template content slide."""
    ref_slide = None
    for slide in template_prs.slides:
        if 'title and content' in slide.slide_layout.name.lower():
            ref_slide = slide
            break
    if ref_slide is None and template_prs.slides:
        ref_slide = template_prs.slides[0]
    if ref_slide is None:
        return []

    src_part = ref_slide.part
    spTree = ref_slide.element.find(f'{{{_NS_P}}}cSld/{{{_NS_P}}}spTree')
    if spTree is None:
        return []

    result = []
    for child in spTree:
        if child.tag.split('}')[-1] != 'pic':
            continue
        blip = child.find(f'.//{{{_NS_A}}}blip')
        if blip is None:
            continue
        rId = blip.get(f'{{{_NS_R}}}embed')
        if not rId:
            continue
        try:
            blob = src_part.related_part(rId).blob
        except Exception:
            continue
        xfrm = child.find(f'.//{{{_NS_A}}}xfrm')
        if xfrm is None:
            continue
        off = xfrm.find(f'{{{_NS_A}}}off')
        ext = xfrm.find(f'{{{_NS_A}}}ext')
        if off is None or ext is None:
            continue
        result.append((blob, int(off.get('x', 0)), int(off.get('y', 0)),
                       int(ext.get('cx', 0)), int(ext.get('cy', 0))))
    return result


def _add_background_images(dest_slide, design_images):
    """Inject background images behind all slide content."""
    spTree = dest_slide.element.find(f'{{{_NS_P}}}cSld/{{{_NS_P}}}spTree')
    if spTree is None:
        return
    insert_idx = 2  # after nvGrpSpPr and grpSpPr
    for blob, left, top, width, height in design_images:
        try:
            pic = dest_slide.shapes.add_picture(io.BytesIO(blob), left, top, width, height)
            el = pic.element
            spTree.remove(el)
            spTree.insert(insert_idx, el)
            insert_idx += 1
        except Exception:
            pass


class TransferResult:
    def __init__(self):
        self.slide_index: int = 0
        self.warnings: list[str] = []
        self.errors: list[str] = []
        self.overflow: bool = False
        self.content_warnings: list[dict] = []
        # V2 Phase 4 diagnostics: one layout_safety action per body column filled
        # ("none" | "reduce_spacing" | "reduce_font" | "overflow").
        self.fit_actions: list[str] = []


def _apply_run_formatting(dest_run, src_run: TextRun):
    if src_run.bold is not None:
        dest_run.font.bold = src_run.bold
    if src_run.italic is not None:
        dest_run.font.italic = src_run.italic
    if src_run.font_size is not None:
        dest_run.font.size = Pt(src_run.font_size)
    if src_run.font_name:
        dest_run.font.name = src_run.font_name
    if src_run.font_color:
        try:
            r = int(src_run.font_color[0:2], 16)
            g = int(src_run.font_color[2:4], 16)
            b = int(src_run.font_color[4:6], 16)
            dest_run.font.color.rgb = RGBColor(r, g, b)
        except Exception:
            pass


def _write_paragraphs_to_tf(tf, paragraphs: list[Paragraph], clear_first=True):
    """Write paragraphs into a text frame, preserving runs and formatting."""
    if clear_first:
        # Clear existing paragraphs except the first (python-pptx requires ≥1)
        tf.clear()

    # python-pptx: tf starts with one empty paragraph after clear()
    for i, para in enumerate(paragraphs):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.level = para.level

        for j, run in enumerate(para.runs):
            if j == 0:
                # Use existing run in paragraph if available
                if p.runs:
                    r = p.runs[0]
                    r.text = run.text
                else:
                    r = p.add_run()
                    r.text = run.text
            else:
                r = p.add_run()
                r.text = run.text
            _apply_run_formatting(r, run)

        # If no runs, set plain text
        if not para.runs and para.full_text:
            if p.runs:
                p.runs[0].text = para.full_text
            else:
                r = p.add_run()
                r.text = para.full_text


def _add_image_to_slide(slide, img: ImageData, placeholder=None):
    """Add an image to a slide, fitting into placeholder bounds if given."""
    if placeholder is not None:
        left = placeholder.left
        top = placeholder.top
        width = placeholder.width
        height = placeholder.height
    else:
        left = img.left
        top = img.top
        width = img.width
        height = img.height

    img_stream = io.BytesIO(img.blob)
    # Scale to fit while preserving aspect ratio
    src_w, src_h = img.width, img.height
    if src_w > 0 and src_h > 0:
        ratio = min(width / src_w, height / src_h)
        width = int(src_w * ratio)
        height = int(src_h * ratio)

    slide.shapes.add_picture(img_stream, left, top, width, height)


def _add_table_to_slide(slide, tbl: TableData, placeholder=None):
    """Add a table to a slide."""
    if tbl.row_count == 0 or tbl.col_count == 0:
        return

    if placeholder is not None:
        left, top = placeholder.left, placeholder.top
        width, height = placeholder.width, placeholder.height
    else:
        left, top = tbl.left, tbl.top
        width, height = tbl.width, tbl.height

    rows, cols = tbl.row_count, tbl.col_count
    dest_table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    for r_idx, row in enumerate(tbl.rows):
        for c_idx, cell in enumerate(row):
            dest_cell = dest_table.cell(r_idx, c_idx)
            _write_paragraphs_to_tf(dest_cell.text_frame, cell.paragraphs)


def transfer(
    source_slides: list[ParsedSlide],
    source_path: str,
    template_path: str,
    output_path: str,
    progress_callback=None,
) -> list[TransferResult]:
    """Transfer all source slides into the template, save to output_path."""

    def _progress(msg):
        if progress_callback:
            progress_callback(msg)

    _progress("reading_template")
    template_layouts = analyze_template(template_path)
    dest_prs = Presentation(template_path)

    # Extract background design images before stripping template slides
    design_images = _extract_design_images(dest_prs)

    # Remove all existing slides — must drop parts AND relationships to avoid
    # duplicate ZIP entries (which corrupt the output file)
    sldIdLst = dest_prs.slides._sldIdLst
    prs_part = dest_prs.part
    for sldId in list(sldIdLst):
        rId = sldId.get(_qn('r:id'))
        try:
            slide_part = prs_part.related_part(rId)
            prs_part.package._parts.pop(slide_part.partname, None)
        except Exception:
            pass
        prs_part.drop_rel(rId)
        sldIdLst.remove(sldId)

    results = []
    _progress("matching")

    total_slides = len(source_slides)
    for i, parsed in enumerate(source_slides):
        if progress_callback:
            progress_callback(f"slide:{i+1}/{total_slides}")
        result = TransferResult()
        result.slide_index = parsed.index

        # Classify and match layout
        slide_type = classify_source_slide(parsed)
        best_layout = find_best_layout(slide_type, template_layouts)
        dest_layout = dest_prs.slide_layouts[best_layout.index]
        dest_slide = dest_prs.slides.add_slide(dest_layout)
        _add_background_images(dest_slide, design_images)

        # Map placeholders by idx
        ph_map = {ph.placeholder_format.idx: ph for ph in dest_slide.placeholders}

        # --- Transfer title ---
        if parsed.title and 0 in ph_map:
            title_ph = ph_map[0]
            if title_ph.left < _TITLE_MIN_LEFT:
                orig_right = title_ph.left + title_ph.width
                orig_top = title_ph.top       # read before xfrm override is created
                orig_height = title_ph.height
                title_ph.left = _TITLE_MIN_LEFT  # creates xfrm, zeros y
                title_ph.top = orig_top          # restore y
                title_ph.height = orig_height
                title_ph.width = max(orig_right - _TITLE_MIN_LEFT, Inches(4))
            try:
                _write_paragraphs_to_tf(title_ph.text_frame, parsed.title.paragraphs)
            except Exception as e:
                result.errors.append(f"Title transfer failed: {e}")
        elif parsed.title and 0 not in ph_map:
            # No title placeholder — add as text box at top
            try:
                slide_w = dest_prs.slide_width
                txBox = dest_slide.shapes.add_textbox(
                    Emu(0), Emu(0), slide_w, Pt(40).emu
                )
                _write_paragraphs_to_tf(txBox.text_frame, parsed.title.paragraphs)
            except Exception as e:
                result.errors.append(f"Title fallback failed: {e}")

        # --- Transfer body text ---
        body_text_boxes = [b for b in parsed.body_boxes if b.full_text.strip()]

        def _merge_boxes(boxes):
            from parser import Paragraph as Para, TextRun as TR
            merged = []
            for i, box in enumerate(boxes):
                if i > 0:
                    merged.append(Para(runs=[TR(text="")]))
                merged.extend(box.paragraphs)
            return merged

        def _fill_body(ph, paras):
            """Write body paragraphs, enable PowerPoint shrink-to-fit, then run
            the layout-safety fitter (layout_safety.plan_fit) on this column's
            actual dimensions. The fitter applies the least-destructive
            adjustment in order: nothing -> reduce line spacing -> reduce font
            (floored at 75% of the intended size). If content still overflows
            after that, a per-slide warning is recorded and the best-effort
            floor is baked into normAutofit so the deck still renders."""
            _write_paragraphs_to_tf(ph.text_frame, paras)
            tf = ph.text_frame
            try:
                tf.word_wrap = True
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            except Exception:
                return None

            plan = plan_fit(paras, ph.width, ph.height)
            result.fit_actions.append(plan.action)
            if plan.font_scale < 1.0 or plan.line_spacing_reduction > 0.0:
                na = tf._txBody.find(f'.//{{{_NS_A}}}bodyPr/{{{_NS_A}}}normAutofit')
                if na is not None:
                    na.set('fontScale', str(int(round(plan.font_scale * 100000))))
                    na.set('lnSpcReduction', str(int(round(plan.line_spacing_reduction * 100000))))
            if plan.warning:
                result.overflow = True
                result.warnings.append(f"⚠ Slide {parsed.index + 1} — {plan.warning}.")
            return plan

        if body_text_boxes and 1 in ph_map:
            try:
                # Two-column source + a layout that actually has a 2nd content
                # placeholder → spread the boxes across both columns instead of
                # stacking them all in the (often half-width) first one.
                if len(body_text_boxes) >= 2 and 2 in ph_map:
                    # BUG-2: if either column reaches into the red wedge, rebalance
                    # both into the text-safe band (0.92 .. 9.1in) as equal columns.
                    r1 = ph_map[1].left + ph_map[1].width
                    r2 = ph_map[2].left + ph_map[2].width
                    if r1 > _BODY_SAFE_RIGHT or r2 > _BODY_SAFE_RIGHT:
                        col_w = (_BODY_SAFE_RIGHT - _BODY_SAFE_LEFT - _COL_GUTTER) // 2
                        _place_ph(ph_map[1], _BODY_SAFE_LEFT, col_w)
                        _place_ph(ph_map[2], _BODY_SAFE_LEFT + col_w + _COL_GUTTER, col_w)

                    mid = (len(body_text_boxes) + 1) // 2
                    left, right = body_text_boxes[:mid], body_text_boxes[mid:]
                    # Each column is fitted independently against its own
                    # (post-split, post-BUG-2-clamp) dimensions.
                    _fill_body(ph_map[1], _merge_boxes(left))
                    _fill_body(ph_map[2], _merge_boxes(right))
                else:
                    # BUG-2: single body column — clamp its right edge clear of the
                    # red wedge (only if it currently overruns it).
                    if ph_map[1].left + ph_map[1].width > _BODY_SAFE_RIGHT:
                        _place_ph(ph_map[1], ph_map[1].left, _BODY_SAFE_RIGHT - ph_map[1].left)

                    _fill_body(ph_map[1], _merge_boxes(body_text_boxes))
            except Exception as e:
                result.errors.append(f"Body transfer failed: {e}")

        elif body_text_boxes and 1 not in ph_map:
            # No body placeholder — add as floating text box
            try:
                slide_h = dest_prs.slide_height
                # BUG-2: keep the fallback textbox clear of the red wedge too.
                txBox = dest_slide.shapes.add_textbox(
                    Inches(0.5), Inches(1.5), _BODY_SAFE_RIGHT - Inches(0.5), slide_h - Inches(2)
                )
                all_paras = []
                for i, box in enumerate(body_text_boxes):
                    if i > 0:
                        from parser import Paragraph as Para, TextRun as TR
                        all_paras.append(Para(runs=[TR(text="")]))
                    all_paras.extend(box.paragraphs)
                _write_paragraphs_to_tf(txBox.text_frame, all_paras)
            except Exception as e:
                result.errors.append(f"Body fallback failed: {e}")

        # --- Transfer images ---
        for img in parsed.images:
            try:
                # Look for picture placeholder (idx >= 2 typically)
                img_ph = None
                for idx, ph in ph_map.items():
                    if idx >= 2 and str(ph.placeholder_format.type) in (
                        "PICTURE(18)", "OBJECT(14)", "MEDIA(16)"
                    ):
                        img_ph = ph
                        break
                _add_image_to_slide(dest_slide, img, placeholder=img_ph)
            except Exception as e:
                result.warnings.append(f"Image on slide {parsed.index + 1} could not be transferred: {e}")

        # --- Transfer tables ---
        for tbl in parsed.tables:
            try:
                # Look for content placeholder for table
                tbl_ph = ph_map.get(1) if 1 in ph_map and not body_text_boxes else None
                _add_table_to_slide(dest_slide, tbl, placeholder=tbl_ph)
            except Exception as e:
                result.warnings.append(f"Table on slide {parsed.index + 1} could not be transferred: {e}")

        # --- Unsupported content warnings ---
        if parsed.unsupported_shapes:
            slide_num = parsed.index + 1
            # De-dupe by shape kind, keeping label + V2 support level.
            seen = {}
            for us in parsed.unsupported_shapes:
                level = getattr(us, "support_level", "SKIPPED_WITH_WARNING")
                seen.setdefault(us.type, (us.label, level))

            for us_type, (us_label, level) in seen.items():
                if level == "PARTIALLY_SUPPORTED":
                    msg = (f"{us_label}: text content was salvaged; other elements "
                           f"could not be transferred.")
                else:
                    msg = f"{us_label} content could not be transferred automatically."
                result.content_warnings.append({
                    "slide": slide_num,
                    "type": us_type,
                    "support_level": level,
                    "message": msg,
                })

            skipped = [lbl for lbl, lvl in seen.values() if lvl != "PARTIALLY_SUPPORTED"]
            partial = [lbl for lbl, lvl in seen.values() if lvl == "PARTIALLY_SUPPORTED"]

            if skipped:
                has_transferred = (
                    parsed.title is not None
                    or any(b.full_text.strip() for b in parsed.body_boxes)
                    or parsed.images
                    or parsed.tables
                )
                if not has_transferred:
                    result.warnings.append(
                        f"⚠ Slide {slide_num} could not be fully transferred because it contains "
                        f"unsupported PowerPoint content: {', '.join(skipped)}."
                    )
                else:
                    result.warnings.append(
                        f"⚠ Slide {slide_num} contains unsupported content that was skipped: {', '.join(skipped)}."
                    )
            if partial:
                result.warnings.append(
                    f"⚠ Slide {slide_num}: text from grouped shapes was salvaged; "
                    f"other grouped elements were not transferred ({', '.join(partial)})."
                )

        results.append(result)

    _progress("saving")
    dest_prs.save(output_path)
    _dedupe_pptx(output_path)

    return results
