"""Core content transfer engine."""

import io
import hashlib
import zipfile
from collections import defaultdict
from typing import Optional
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn as _qn
from pptx.enum.shapes import MSO_SHAPE_TYPE


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

# Placeholder types, keyed by python-pptx's actual str(placeholder_format.type)
# repr — "NAME (value)" with a space before the parenthesis.
#
# PICTURE/MEDIA_CLIP/CHART/TABLE are type-locked: they can never hold typed
# bullet text. OBJECT is PowerPoint's generic "click an icon for a picture/
# chart/table, or just type" content placeholder — it's the sole content slot
# on the very common single-placeholder "Title and Content" layout, where it
# must stay eligible for body text. But a layout like "Content with Caption"
# pairs an OBJECT placeholder (the big picture area) with a separate, real
# BODY placeholder (the caption) — there, OBJECT is clearly the picture's
# slot and text must go in BODY instead, or the two end up stacked on top of
# each other (they're differently sized/positioned, not parallel columns).
# So OBJECT only counts as "text-capable" when no sibling BODY exists.
_STRICT_MEDIA_PH_TYPES = ("PICTURE (18)", "MEDIA_CLIP (10)", "CHART (8)", "TABLE (12)")
_MEDIA_PH_TYPES = _STRICT_MEDIA_PH_TYPES + ("OBJECT (7)",)  # types that can hold an image


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
    """Return pictures that are template *decoration* — the same image, in the
    same spot, repeated across most of the template's slides (a full-bleed
    background, watermark, or logo bar). Content photos, which differ slide to
    slide, are left alone; that stops a normal deck used as a template from
    stamping its slide-1 photos onto every output slide.
    """
    slides = list(template_prs.slides)
    if len(slides) < 2:
        return []

    seen = defaultdict(lambda: {"count": 0, "geom": None, "blob": None})
    for slide in slides:
        spTree = slide.element.find(f'{{{_NS_P}}}cSld/{{{_NS_P}}}spTree')
        if spTree is None:
            continue
        slide_keys = set()
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
                blob = slide.part.related_part(rId).blob
            except Exception:
                continue
            xfrm = child.find(f'.//{{{_NS_A}}}xfrm')
            if xfrm is None:
                continue
            off = xfrm.find(f'{{{_NS_A}}}off')
            ext = xfrm.find(f'{{{_NS_A}}}ext')
            if off is None or ext is None:
                continue
            key = (hashlib.md5(blob).hexdigest(),
                   off.get('x'), off.get('y'), ext.get('cx'), ext.get('cy'))
            if key in slide_keys:      # same pic twice on one slide — count once
                continue
            slide_keys.add(key)
            rec = seen[key]
            rec["count"] += 1
            if rec["geom"] is None:
                rec["geom"] = (int(off.get('x', 0)), int(off.get('y', 0)),
                               int(ext.get('cx', 0)), int(ext.get('cy', 0)))
                rec["blob"] = blob

    threshold = max(2, int(0.6 * len(slides)))
    return [(rec["blob"], *rec["geom"])
            for rec in seen.values() if rec["count"] >= threshold]


def _has_fullbleed_bg(design_images, slide_w, slide_h):
    """True when the template hides decoration in a (near) full-bleed background
    image — the CUCOM case, where body text must be clamped off it. A normal
    template has no such image and its placeholder geometry is used as-is.

    design_images entries are (blob, x, y, cx, cy) in EMU.
    """
    for _blob, x, y, cx, cy in design_images:
        if (x <= Inches(0.15) and y <= Inches(0.15)
                and cx >= slide_w * 0.98 and cy >= slide_h * 0.98):
            return True
    return False


def _add_background_images(dest_slide, design_images):
    """Inject background images behind all slide content. Returns the
    shape_ids of the images it added, so callers can tell decoration apart
    from actual transferred content (a decorative image is expected to sit
    behind everything — a small corner logo included — and isn't a
    candidate for the overlap audit the way a real content image is)."""
    spTree = dest_slide.element.find(f'{{{_NS_P}}}cSld/{{{_NS_P}}}spTree')
    if spTree is None:
        return []
    insert_idx = 2  # after nvGrpSpPr and grpSpPr
    shape_ids = []
    for blob, left, top, width, height in design_images:
        try:
            pic = dest_slide.shapes.add_picture(io.BytesIO(blob), left, top, width, height)
            el = pic.element
            spTree.remove(el)
            spTree.insert(insert_idx, el)
            insert_idx += 1
            shape_ids.append(pic.shape_id)
        except Exception:
            pass
    return shape_ids


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


# --- Image / body-text collision avoidance -------------------------------- #
#
# A source picture with no destination picture placeholder lands at its own
# (clamped-on-canvas) source coordinates. Canvas clamping only guarantees the
# image stays on the slide — it says nothing about whether that spot is also
# where the body placeholder now lives. An opaque screenshot dropped over the
# body text hides it in the rendered output even though the text is still
# present in the file (which is why the structural test suite didn't catch
# this — it checks text/image survival, not on-slide position).
#
# Threshold measured from the real regression (Chapter 09 "The Blood" slide 5
# on the CUCOM template): the screenshot covered 37% of its own area and 28%
# of the body placeholder's area, and visibly hid two bullet lines. 15% is
# comfortably below that real collision and above incidental edge contact
# (a shape clipping a corner of the body box typically overlaps <5%).
_COLLISION_THRESHOLD = 0.15


def _rect_intersection_area(a, b):
    """a, b = (left, top, width, height) in EMU. Overlap area, 0 if none."""
    ax1, ay1, ax2, ay2 = a[0], a[1], a[0] + a[2], a[1] + a[3]
    bx1, by1, bx2, by2 = b[0], b[1], b[0] + b[2], b[1] + b[3]
    iw = min(ax2, bx2) - max(ax1, bx1)
    ih = min(ay2, by2) - max(ay1, by1)
    return max(0, iw) * max(0, ih)


def _overlap_fraction(a, b):
    """max(intersection / area_a, intersection / area_b), 0 if disjoint."""
    area_a, area_b = a[2] * a[3], b[2] * b[3]
    if area_a <= 0 or area_b <= 0:
        return 0.0
    inter = _rect_intersection_area(a, b)
    return max(inter / area_a, inter / area_b) if inter > 0 else 0.0


def _overlaps_meaningfully(a, b, threshold=_COLLISION_THRESHOLD):
    """True once a and b share more than incidental edge contact."""
    return _overlap_fraction(a, b) >= threshold


def _find_slide_overlaps(dest_slide, slide_w, slide_h, background_shape_ids=()):
    """Generic end-of-slide audit: do any two placed elements (title, body,
    floating textbox, image, table) meaningfully overlap?

    This exists because per-content-type placement code (title vs. body vs.
    image) only ever checks the NEW element against what came before it --
    there was no single point that looked at the finished slide as a whole.
    A geometry mismatch between two placeholders used as parallel text
    columns (see the OBJECT/BODY placeholder-role fix) produced exactly this:
    text overlapping text, invisible to any single per-element check.

    background_shape_ids (from _add_background_images) are excluded outright
    -- template decoration is expected to sit behind, and be covered by,
    real content, whatever its size (a small corner logo counts just as much
    as CUCOM's full-bleed wedge art). Returns a list of short description
    strings, one per offending pair.
    """
    rects = []
    for sh in dest_slide.shapes:
        if sh.shape_id in background_shape_ids:
            continue
        if sh.left is None or sh.top is None or sh.width is None or sh.height is None:
            continue
        rect = (sh.left, sh.top, sh.width, sh.height)
        if sh.has_text_frame and sh.text_frame.text.strip():
            rects.append(("text", rect))
        elif sh.has_table:
            rects.append(("table", rect))
        elif sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            rects.append(("image", rect))

    conflicts = []
    for i in range(len(rects)):
        for j in range(i + 1, len(rects)):
            kind_a, rect_a = rects[i]
            kind_b, rect_b = rects[j]
            if _overlaps_meaningfully(rect_a, rect_b):
                conflicts.append(f"{kind_a} overlaps {kind_b}")
    return conflicts


def _clamp_rect(rect, slide_w, slide_h, margin):
    left, top, w, h = rect
    w = min(w, slide_w - 2 * margin)
    h = min(h, slide_h - 2 * margin)
    left = min(max(left, margin), slide_w - margin - w)
    top = min(max(top, margin), slide_h - margin - h)
    return (left, top, w, h)


def _adjust_image_for_content_collision(rect, protected_rects, slide_w, slide_h):
    """If `rect` (left, top, w, h EMU) meaningfully overlaps any protected
    content region (a filled title/body placeholder or textbox), find a
    placement that doesn't. Rects that don't collide come back unchanged —
    this never moves an image that's already fine.

    Least-destructive first: try moving (original size) into the open side of
    whatever it collides with; if nothing clears it, shrink (aspect-preserving)
    and retry; if it still can't fully clear protected content, keep the image
    — never drop it — at whichever attempted placement minimized total
    overlap.

    Returns (new_rect, moved: bool, still_colliding: bool).
    """
    margin = Inches(0.1)
    rect = _clamp_rect(rect, slide_w, slide_h, margin)
    colliding = [p for p in protected_rects if _overlaps_meaningfully(rect, p)]
    if not colliding:
        return rect, False, False

    left, top, w, h = rect
    aspect = (w / h) if h else 1.0

    def collides_any(r):
        return any(_overlaps_meaningfully(r, p) for p in protected_rects)

    def moves_for(cw, ch, base_left, base_top):
        # Slide the image clear of each colliding region on whichever side
        # (right/left/below/above it) has room, keeping the other axis put.
        out = []
        for (px, py, pw, ph_) in colliding:
            out += [(px + pw + margin, base_top), (px - margin - cw, base_top),
                    (base_left, py + ph_ + margin), (base_left, py - margin - ch)]
        return out

    tried = [rect]

    # Attempt 1 — move only, original size; smallest displacement wins.
    sized = [_clamp_rect((cl, ct, w, h), slide_w, slide_h, margin)
             for cl, ct in moves_for(w, h, left, top)]
    tried += sized
    clear = [c for c in sized if not collides_any(c)]
    if clear:
        best = min(clear, key=lambda c: abs(c[0] - left) + abs(c[1] - top))
        return best, True, False

    # Attempt 2 — shrink (aspect-preserving), retry the same moves at each size.
    scale = 1.0
    min_side = min(Inches(0.6), slide_w, slide_h)
    while True:
        scale *= 0.85
        nw, nh = int(w * scale), int(w * scale / aspect) if aspect else int(h * scale)
        if nw < min_side or nh < min_side:
            break
        for cl, ct in [(left, top)] + moves_for(nw, nh, left, top):
            cand = _clamp_rect((cl, ct, nw, nh), slide_w, slide_h, margin)
            tried.append(cand)
            if not collides_any(cand):
                return cand, True, False

    # Attempt 3 — best effort: keep the image, at whichever attempted
    # placement left the least overlap with protected content.
    best = min(tried, key=lambda c: sum(_overlap_fraction(c, p) for p in protected_rects))
    return best, True, True


def _add_image_to_slide(slide, img: ImageData, placeholder=None,
                        slide_w=None, slide_h=None, protected_rects=None):
    """Add an image to a slide, fitting into placeholder bounds if given.

    With no destination placeholder the image is placed at its source
    coordinates/size. The source slide is often larger than this template's, so
    when slide_w/slide_h are supplied the image is scaled down and nudged to
    stay fully on-canvas — and, if `protected_rects` names the regions this
    slide actually filled with title/body text, nudged clear of those too.

    Returns (left, top, width, height, still_colliding) so the caller can
    track this image as an obstacle for later ones and warn if it couldn't be
    fully cleared.
    """
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

    still_colliding = False
    if placeholder is None and slide_w and slide_h:
        margin = Inches(0.1)
        max_w, max_h = slide_w - 2 * margin, slide_h - 2 * margin
        if width > max_w or height > max_h:
            fit = min(max_w / width, max_h / height)
            width = int(width * fit)
            height = int(height * fit)
        left = min(max(left, margin), slide_w - margin - width)
        top = min(max(top, margin), slide_h - margin - height)

        if protected_rects:
            (left, top, width, height), _moved, still_colliding = \
                _adjust_image_for_content_collision(
                    (left, top, width, height), protected_rects, slide_w, slide_h)

    slide.shapes.add_picture(img_stream, left, top, width, height)
    return (left, top, width, height, still_colliding)


def _add_table_to_slide(slide, tbl: TableData, placeholder=None,
                        slide_w=None, slide_h=None):
    """Add a table to a slide."""
    if tbl.row_count == 0 or tbl.col_count == 0:
        return

    if placeholder is not None:
        left, top = placeholder.left, placeholder.top
        width, height = placeholder.width, placeholder.height
    else:
        left, top = tbl.left, tbl.top
        width, height = tbl.width, tbl.height
        # Source coords may exceed a smaller destination canvas — clamp width
        # and left so the table stays on-slide (row height handles itself).
        if slide_w and slide_h:
            margin = Inches(0.1)
            width = min(width, slide_w - 2 * margin)
            left = min(max(left, margin), slide_w - margin - width)
            top = max(min(top, slide_h - margin), margin)

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
    # CUCOM-style templates need body text clamped off a full-bleed background
    # wedge; plain templates don't — keep their native placeholder geometry.
    wedge_template = _has_fullbleed_bg(
        design_images, dest_prs.slide_width, dest_prs.slide_height)

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
        background_shape_ids = _add_background_images(dest_slide, design_images)

        # Map placeholders by idx
        ph_map = {ph.placeholder_format.idx: ph for ph in dest_slide.placeholders}

        # Regions this slide actually fills with title/body content — a
        # free-floating source image must not be dropped on top of these.
        # Populated as title/body get written below, in their final
        # (post wedge-clamp) position; checked against when images are placed.
        protected_rects = []

        # --- Transfer title ---
        if parsed.title and 0 in ph_map:
            title_ph = ph_map[0]
            if wedge_template and title_ph.left < _TITLE_MIN_LEFT:
                orig_right = title_ph.left + title_ph.width
                orig_top = title_ph.top       # read before xfrm override is created
                orig_height = title_ph.height
                title_ph.left = _TITLE_MIN_LEFT  # creates xfrm, zeros y
                title_ph.top = orig_top          # restore y
                title_ph.height = orig_height
                title_ph.width = max(orig_right - _TITLE_MIN_LEFT, Inches(4))
            try:
                _write_paragraphs_to_tf(title_ph.text_frame, parsed.title.paragraphs)
                protected_rects.append((title_ph.left, title_ph.top, title_ph.width, title_ph.height))
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
                protected_rects.append((txBox.left, txBox.top, txBox.width, txBox.height))
            except Exception as e:
                result.errors.append(f"Title fallback failed: {e}")

        # --- Transfer body text ---
        body_text_boxes = [b for b in parsed.body_boxes if b.full_text.strip()]

        # Placeholders that can actually take flowing bullet text: not the
        # title, never a type-locked media slot, and OBJECT only when there's
        # no sibling BODY placeholder claiming the real text role instead
        # (see _MEDIA_PH_TYPES/_STRICT_MEDIA_PH_TYPES above).
        has_real_body = any(str(ph.placeholder_format.type) == "BODY (2)"
                             for idx, ph in ph_map.items() if idx != 0)
        body_ph_items = [
            (idx, ph_map[idx]) for idx in sorted(ph_map) if idx != 0
            and str(ph_map[idx].placeholder_format.type) not in _STRICT_MEDIA_PH_TYPES
            and not (str(ph_map[idx].placeholder_format.type) == "OBJECT (7)" and has_real_body)
        ]
        body_phs = [ph for _, ph in body_ph_items]
        claimed_ph_indices = set()

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

        if body_text_boxes and body_phs:
            try:
                # Two-column source + a layout that actually has a 2nd text
                # placeholder → spread the boxes across both columns instead of
                # stacking them all in the (often half-width) first one.
                if len(body_text_boxes) >= 2 and len(body_phs) >= 2:
                    col1, col2 = body_phs[0], body_phs[1]
                    claimed_ph_indices.update((body_ph_items[0][0], body_ph_items[1][0]))
                    # BUG-2: if either column reaches into the red wedge, rebalance
                    # both into the text-safe band (0.92 .. 9.1in) as equal columns.
                    r1 = col1.left + col1.width
                    r2 = col2.left + col2.width
                    if wedge_template and (r1 > _BODY_SAFE_RIGHT or r2 > _BODY_SAFE_RIGHT):
                        col_w = (_BODY_SAFE_RIGHT - _BODY_SAFE_LEFT - _COL_GUTTER) // 2
                        _place_ph(col1, _BODY_SAFE_LEFT, col_w)
                        _place_ph(col2, _BODY_SAFE_LEFT + col_w + _COL_GUTTER, col_w)

                    mid = (len(body_text_boxes) + 1) // 2
                    left, right = body_text_boxes[:mid], body_text_boxes[mid:]
                    # Each column is fitted independently against its own
                    # (post-split, post-BUG-2-clamp) dimensions.
                    _fill_body(col1, _merge_boxes(left))
                    _fill_body(col2, _merge_boxes(right))
                    protected_rects.append((col1.left, col1.top, col1.width, col1.height))
                    protected_rects.append((col2.left, col2.top, col2.width, col2.height))
                else:
                    col1 = body_phs[0]
                    claimed_ph_indices.add(body_ph_items[0][0])
                    # BUG-2: single body column — clamp its right edge clear of the
                    # red wedge (only if it currently overruns it).
                    if wedge_template and col1.left + col1.width > _BODY_SAFE_RIGHT:
                        _place_ph(col1, col1.left, _BODY_SAFE_RIGHT - col1.left)

                    _fill_body(col1, _merge_boxes(body_text_boxes))
                    protected_rects.append((col1.left, col1.top, col1.width, col1.height))
            except Exception as e:
                result.errors.append(f"Body transfer failed: {e}")

        elif body_text_boxes and not body_phs:
            # No body placeholder — add as floating text box
            try:
                slide_h = dest_prs.slide_height
                # BUG-2: on a wedge template keep the fallback textbox clear of
                # the red wedge; otherwise use (almost) the full slide width.
                box_right = _BODY_SAFE_RIGHT if wedge_template else dest_prs.slide_width - Inches(0.5)
                txBox = dest_slide.shapes.add_textbox(
                    Inches(0.5), Inches(1.5), box_right - Inches(0.5), slide_h - Inches(2)
                )
                all_paras = []
                for i, box in enumerate(body_text_boxes):
                    if i > 0:
                        from parser import Paragraph as Para, TextRun as TR
                        all_paras.append(Para(runs=[TR(text="")]))
                    all_paras.extend(box.paragraphs)
                _write_paragraphs_to_tf(txBox.text_frame, all_paras)
                protected_rects.append((txBox.left, txBox.top, txBox.width, txBox.height))
            except Exception as e:
                result.errors.append(f"Body fallback failed: {e}")

        # --- Transfer images ---
        # Previously-placed free-floating images join protected_rects as we go,
        # so a second image on the same slide doesn't land on the first one
        # (still just bounding-box avoidance — no packing engine).
        placed_image_rects = []
        used_media_ph_indices = set(claimed_ph_indices)  # don't double-book a placeholder already filled with text
        for img in parsed.images:
            try:
                # Look for an unclaimed picture/media/object placeholder.
                img_ph = None
                for idx, ph in ph_map.items():
                    if idx not in used_media_ph_indices and str(ph.placeholder_format.type) in _MEDIA_PH_TYPES:
                        img_ph = ph
                        used_media_ph_indices.add(idx)
                        break
                left, top, width, height, still_colliding = _add_image_to_slide(
                    dest_slide, img, placeholder=img_ph,
                    slide_w=dest_prs.slide_width, slide_h=dest_prs.slide_height,
                    protected_rects=protected_rects + placed_image_rects)
                placed_image_rects.append((left, top, width, height))
                if still_colliding:
                    result.warnings.append(
                        f"⚠ Slide {parsed.index + 1} — an image could not be fully "
                        f"repositioned clear of the body text; some overlap remains.")
            except Exception as e:
                result.warnings.append(f"Image on slide {parsed.index + 1} could not be transferred: {e}")

        # --- Transfer tables ---
        for tbl in parsed.tables:
            try:
                # Look for content placeholder for table
                tbl_ph = ph_map.get(1) if 1 in ph_map and not body_text_boxes else None
                _add_table_to_slide(dest_slide, tbl, placeholder=tbl_ph,
                                    slide_w=dest_prs.slide_width,
                                    slide_h=dest_prs.slide_height)
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

        # A slide whose source held only untransferable content would otherwise
        # be an empty page — leave a visible marker instead.
        if len(dest_slide.shapes) == 0:
            kinds = ", ".join(sorted({u.label for u in parsed.unsupported_shapes})) \
                if parsed.unsupported_shapes else "content"
            try:
                note = dest_slide.shapes.add_textbox(
                    Inches(0.5), Inches(0.5),
                    dest_prs.slide_width - Inches(1), Inches(1))
                note.text_frame.text = (
                    f"[Original slide {parsed.index + 1}: {kinds} could not be "
                    f"transferred automatically — recreate manually.]")
            except Exception:
                pass

        overlaps = _find_slide_overlaps(dest_slide, dest_prs.slide_width, dest_prs.slide_height,
                                         background_shape_ids)
        if overlaps:
            result.warnings.append(
                f"⚠ Slide {parsed.index + 1} — layout conflict: {overlaps[0]} "
                f"(possible overlapping content).")

        results.append(result)

    _progress("saving")
    dest_prs.save(output_path)
    _dedupe_pptx(output_path)

    return results
