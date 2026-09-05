"""Core content transfer engine."""

import io
import zipfile
from typing import Optional
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn as _qn


def _dedupe_pptx(path: str):
    """Remove duplicate ZIP entries from a PPTX, keeping the last occurrence of each name."""
    buf = io.BytesIO()
    with zipfile.ZipFile(path, 'r') as src:
        # Dict preserves insertion order; duplicate keys overwrite → last one wins
        entries = {info.filename: (info, src.read(info.filename)) for info in src.infolist()}
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as dst:
        for fname, (info, data) in entries.items():
            dst.writestr(info, data)
    with open(path, 'wb') as f:
        f.write(buf.getvalue())

from parser import ParsedSlide, TextBox, Paragraph, TextRun, ImageData, TableData
from layout_matcher import TemplateLayout, SlideType, find_best_layout, analyze_template, classify_source_slide


OVERFLOW_THRESHOLD = 0.95  # flag slide if content uses >95% of placeholder height


class TransferResult:
    def __init__(self):
        self.slide_index: int = 0
        self.warnings: list[str] = []
        self.errors: list[str] = []
        self.overflow: bool = False
        self.content_warnings: list[dict] = []


def _apply_run_formatting(dest_run, src_run: TextRun):
    if src_run.bold is not None:
        dest_run.font.bold = src_run.bold
    if src_run.italic is not None:
        dest_run.font.italic = src_run.italic
    if src_run.font_size is not None:
        dest_run.font.size = Pt(src_run.font_size)
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


def _estimate_text_height(tf, paragraphs: list[Paragraph]) -> int:
    """Rough height estimate in EMU. ~914400 EMU per inch, ~12pt = ~152400 EMU/line."""
    line_height_emu = int(Pt(14).emu)  # 14pt per line as default
    total = 0
    for para in paragraphs:
        total += line_height_emu
        # Multi-line wrapping is hard to estimate without rendering — skip for now
    return total


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

    _progress("Reading template...")
    template_layouts = analyze_template(template_path)
    dest_prs = Presentation(template_path)

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

    _progress("Transferring slides...")
    for parsed in source_slides:
        result = TransferResult()
        result.slide_index = parsed.index

        # Classify and match layout
        slide_type = classify_source_slide(parsed)
        best_layout = find_best_layout(slide_type, template_layouts)
        dest_layout = dest_prs.slide_layouts[best_layout.index]
        dest_slide = dest_prs.slides.add_slide(dest_layout)

        # Map placeholders by idx
        ph_map = {ph.placeholder_format.idx: ph for ph in dest_slide.placeholders}

        # --- Transfer title ---
        if parsed.title and 0 in ph_map:
            title_ph = ph_map[0]
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

        if body_text_boxes and 1 in ph_map:
            body_ph = ph_map[1]
            # Combine all body boxes into one placeholder
            all_paras = []
            for i, box in enumerate(body_text_boxes):
                if i > 0:
                    # Separator between multiple text boxes
                    from parser import Paragraph as Para, TextRun as TR
                    all_paras.append(Para(runs=[TR(text="")]))
                all_paras.extend(box.paragraphs)

            try:
                _write_paragraphs_to_tf(body_ph.text_frame, all_paras)

                # Overflow detection
                est_h = _estimate_text_height(body_ph.text_frame, all_paras)
                if est_h > body_ph.height * OVERFLOW_THRESHOLD:
                    result.overflow = True
                    result.warnings.append(
                        f"⚠ Slide {parsed.index + 1} — content may exceed available space. Review recommended."
                    )
            except Exception as e:
                result.errors.append(f"Body transfer failed: {e}")

        elif body_text_boxes and 1 not in ph_map:
            # No body placeholder — add as floating text box
            try:
                slide_w = dest_prs.slide_width
                slide_h = dest_prs.slide_height
                txBox = dest_slide.shapes.add_textbox(
                    Inches(0.5), Inches(1.5), slide_w - Inches(1), slide_h - Inches(2)
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
            seen = {}
            for us in parsed.unsupported_shapes:
                seen.setdefault(us.type, us.label)

            labels = list(seen.values())
            slide_num = parsed.index + 1
            for us_type, us_label in seen.items():
                result.content_warnings.append({
                    "slide": slide_num,
                    "type": us_type,
                    "message": f"{us_label} content could not be transferred automatically.",
                })

            has_transferred = (
                parsed.title is not None
                or any(b.full_text.strip() for b in parsed.body_boxes)
                or parsed.images
                or parsed.tables
            )
            if not has_transferred:
                result.warnings.append(
                    f"⚠ Slide {slide_num} could not be fully transferred because it contains "
                    f"unsupported PowerPoint content: {', '.join(labels)}."
                )
            else:
                result.warnings.append(
                    f"⚠ Slide {slide_num} contains unsupported content that was skipped: {', '.join(labels)}."
                )

        results.append(result)

    _progress("Saving output...")
    dest_prs.save(output_path)
    _dedupe_pptx(output_path)

    return results
