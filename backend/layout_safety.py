"""V2 layout-safety — a predictable, conservative body-text fit estimator.

This is NOT a PowerPoint renderer. Real text layout depends on font metrics,
font availability, wrapping rules, theme settings, textbox insets, bullet
indentation and language. This module gives a *safety estimate* so the
transfer engine can decide, before it hands the deck to PowerPoint, whether
body text is likely to overflow and what the least-destructive adjustment is.

Prefer conservative (slightly pessimistic) estimates over falsely precise ones.

Fitting ladder (least destructive first):
  1. keep the intended font + normal spacing                      -> action "none"
  2. reduce line spacing within a safe bound                      -> "reduce_spacing"
  3. reduce font size, floored at MIN_FONT_SCALE (0.75 of intent) -> "reduce_font"
  4. still doesn't fit after 2+3 -> apply the floor as best effort
     and return a warning                                         -> "overflow"

Geometry note: `plan_fit()` and its default `margin_in=0.0` reproduce the exact
box geometry the V1 `_fill_body` heuristic used, so slides V1 left untouched
stay byte-identical. `estimate_content_height()` supports `margin_in` for
width-accuracy work and is unit-tested with it, but the transfer wiring passes
0.0 on purpose.
"""

from dataclasses import dataclass
from typing import Optional

EMU_PER_IN = 914400

# --- estimator constants (kept equal to the V1 _fill_body heuristic) ---
LINE_FACTOR = 1.15          # line box height as a multiple of font point size
AVG_GLYPH_EM = 0.46         # average glyph advance as a fraction of the em
BLANK_LINE_PT = 18.0        # height charged for an empty separator paragraph
DEFAULT_PT = 18.0           # assumed size when a run carries no explicit size
OVERFULL_TRIGGER = 1.15     # only intervene once content exceeds 115% of the box
MIN_CPL = 8                 # never assume fewer than 8 characters per line

# --- adjustment bounds ---
MIN_FONT_SCALE = 0.75           # V1 safety floor: never below 75% of intended size
SPACING_REDUCTION_MAX = 0.20    # at most a 20% line-spacing reduction
SPACING_REDUCTION_STEP = 0.05   # tried in 5% increments


@dataclass
class LayoutFitPlan:
    fits: bool                     # fits within the box after the chosen action
    action: str                    # "none" | "reduce_spacing" | "reduce_font" | "overflow"
    estimated_height_emu: int      # estimated content height AFTER the chosen action
    available_height_emu: int      # usable box height
    original_pt: float             # intended body font size
    final_pt: float                # font size after any reduction
    font_scale: float              # final_pt / original_pt  (1.0 == unchanged)
    line_spacing_reduction: float  # fraction 0.0 .. SPACING_REDUCTION_MAX
    overflow_emu: int              # how far content still exceeds the box (0 if it fits)
    warning: Optional[str]         # human-readable, only set when action == "overflow"


def _min_font_pt(paragraphs) -> float:
    sizes = [
        r.font_size
        for para in paragraphs
        for r in getattr(para, "runs", [])
        if getattr(r, "font_size", None)
    ]
    return float(min(sizes)) if sizes else DEFAULT_PT


def estimate_content_height(
    paragraphs,
    box_w_emu: int,
    *,
    font_scale: float = 1.0,
    line_spacing_reduction: float = 0.0,
    margin_in: float = 0.0,
) -> int:
    """Conservative wrapped-text height in EMU.

    Width matters: a narrower box packs fewer characters per line, so the same
    text needs more lines and more height.
    """
    usable_w_in = max(box_w_emu / EMU_PER_IN - 2 * margin_in, 0.5)
    spacing = LINE_FACTOR * max(0.0, 1.0 - line_spacing_reduction)
    total_in = 0.0
    for para in paragraphs:
        txt = para.full_text if hasattr(para, "full_text") else str(para)
        if not txt.strip():
            total_in += BLANK_LINE_PT * LINE_FACTOR / 72.0
            continue
        sizes = [r.font_size for r in getattr(para, "runs", []) if getattr(r, "font_size", None)]
        pt = (min(sizes) if sizes else DEFAULT_PT) * font_scale
        cpl = max(MIN_CPL, int(usable_w_in / (AVG_GLYPH_EM * pt / 72.0)))
        lines = max(1, -(-len(txt) // cpl))          # ceil division
        total_in += lines * (pt * spacing / 72.0)
    return int(total_in * EMU_PER_IN)


def plan_fit(
    paragraphs,
    box_w_emu: int,
    box_h_emu: int,
    *,
    min_font_scale: float = MIN_FONT_SCALE,
    margin_in: float = 0.0,
) -> LayoutFitPlan:
    """Decide the least-destructive way to fit `paragraphs` into the box."""
    orig_pt = _min_font_pt(paragraphs)
    avail_h = max(int(box_h_emu - 2 * margin_in * EMU_PER_IN), 1)

    def _plan(fits, action, est_h, scale, reduction, warning=None):
        return LayoutFitPlan(
            fits=fits,
            action=action,
            estimated_height_emu=int(est_h),
            available_height_emu=avail_h,
            original_pt=orig_pt,
            final_pt=round(orig_pt * scale, 2),
            font_scale=round(scale, 5),
            line_spacing_reduction=round(reduction, 5),
            overflow_emu=max(0, int(est_h - avail_h)),
            warning=warning,
        )

    # Step 1 — keep intended formatting. Mildly-full boxes are PowerPoint's job.
    need0 = estimate_content_height(paragraphs, box_w_emu, margin_in=margin_in)
    if need0 <= avail_h * OVERFULL_TRIGGER:
        return _plan(True, "none", need0, 1.0, 0.0)

    # Step 2 — reduce line spacing within a safe bound.
    reduction = 0.0
    while reduction < SPACING_REDUCTION_MAX - 1e-9:
        reduction = min(SPACING_REDUCTION_MAX, reduction + SPACING_REDUCTION_STEP)
        need = estimate_content_height(
            paragraphs, box_w_emu, line_spacing_reduction=reduction, margin_in=margin_in
        )
        if need <= avail_h:
            return _plan(True, "reduce_spacing", need, 1.0, reduction)

    # Step 3 — reduce font size, floored at min_font_scale (spacing already maxed).
    need_sp = estimate_content_height(
        paragraphs, box_w_emu,
        line_spacing_reduction=SPACING_REDUCTION_MAX, margin_in=margin_in,
    )
    scale = max(min_font_scale, min(1.0, avail_h / max(need_sp, 1)))
    need_final = estimate_content_height(
        paragraphs, box_w_emu, font_scale=scale,
        line_spacing_reduction=SPACING_REDUCTION_MAX, margin_in=margin_in,
    )
    if need_final <= avail_h:
        return _plan(True, "reduce_font", need_final, scale, SPACING_REDUCTION_MAX)

    # Step 4 — unavoidable overflow. Apply the floor as best effort, warn, still ship.
    return _plan(
        False, "overflow", need_final, min_font_scale, SPACING_REDUCTION_MAX,
        warning="body content exceeds the available layout area after safe fitting",
    )


def demo():
    """Self-check: the ladder and width-sensitivity behave as documented."""
    class _R:
        def __init__(self, t, s=None):
            self.text, self.font_size = t, s

    class _P:
        def __init__(self, runs):
            self.runs = runs

        @property
        def full_text(self):
            return "".join(r.text for r in self.runs)

    line = "word " * 12  # ~60 chars
    short = [_P([_R(line, 18)])]
    dense = [_P([_R(line, 18)]) for _ in range(40)]

    box_w, box_h = int(6 * EMU_PER_IN), int(4 * EMU_PER_IN)

    p = plan_fit(short, box_w, box_h)
    assert p.fits and p.action == "none" and p.font_scale == 1.0

    p = plan_fit(dense, box_w, box_h)
    assert p.action in ("reduce_spacing", "reduce_font", "overflow")
    assert p.font_scale >= MIN_FONT_SCALE
    if p.action == "overflow":
        assert p.warning and not p.fits

    wide = estimate_content_height(dense, int(8 * EMU_PER_IN))
    med = estimate_content_height(dense, int(5 * EMU_PER_IN))
    narrow = estimate_content_height(dense, int(3 * EMU_PER_IN))
    assert narrow > med > wide, (narrow, med, wide)

    print("[OK] layout_safety ladder + width sensitivity")


if __name__ == "__main__":
    demo()
