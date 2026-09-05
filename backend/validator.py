"""Validate the output PPTX after transfer."""

import os
from dataclasses import dataclass, field
from pptx import Presentation


@dataclass
class ValidationReport:
    ok: bool = True
    slide_count_source: int = 0
    slide_count_output: int = 0
    slides_with_warnings: list[int] = field(default_factory=list)
    slides_with_errors: list[int] = field(default_factory=list)
    messages: list[str] = field(default_factory=list)

    @property
    def summary(self) -> str:
        if not self.ok:
            return "Transfer failed — output file is invalid."
        if self.slides_with_warnings:
            return f"Transfer completed with {len(self.slides_with_warnings)} slide(s) requiring review."
        return "Transfer successful."


def validate(output_path: str, source_slide_count: int, transfer_results) -> ValidationReport:
    report = ValidationReport(slide_count_source=source_slide_count)

    # File exists
    if not os.path.exists(output_path):
        report.ok = False
        report.messages.append("Output file was not created.")
        return report

    # Valid PPTX
    try:
        prs = Presentation(output_path)
    except Exception as e:
        report.ok = False
        report.messages.append(f"Output file cannot be opened as a valid PPTX: {e}")
        return report

    report.slide_count_output = len(prs.slides)

    # Slide count
    if report.slide_count_output != source_slide_count:
        report.messages.append(
            f"Slide count mismatch: source had {source_slide_count}, output has {report.slide_count_output}."
        )

    # Collect warnings/errors from transfer results
    for r in transfer_results:
        if r.overflow or r.warnings:
            report.slides_with_warnings.append(r.slide_index + 1)
            report.messages.extend(r.warnings)
        if r.errors:
            report.slides_with_errors.append(r.slide_index + 1)
            report.messages.extend(r.errors)

    # Check no slide is completely empty
    for i, slide in enumerate(prs.slides):
        has_text = any(
            shape.has_text_frame and shape.text_frame.text.strip()
            for shape in slide.shapes
        )
        has_shapes = len(slide.shapes) > 0
        if not has_shapes:
            report.messages.append(f"Slide {i + 1} appears to be completely empty.")
            if i + 1 not in report.slides_with_warnings:
                report.slides_with_warnings.append(i + 1)

    if report.slides_with_errors:
        report.ok = False

    return report
