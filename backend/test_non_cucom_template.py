"""Self-contained — run: python test_non_cucom_template.py

The CUCOM safe-band clamps must fire only when the template carries a
full-bleed background image (the wedge). A plain template keeps its native
placeholder geometry.
"""
from pptx.util import Inches
from transfer import _has_fullbleed_bg

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)

fails = 0
def check(name, cond):
    global fails
    print(f"  [{'OK' if cond else 'FAIL'}] {name}")
    if not cond:
        fails += 1

# full-bleed image at origin -> wedge template
check("full-bleed 0,0 -> True",
      _has_fullbleed_bg([(b"", 0, 0, SLIDE_W, SLIDE_H)], SLIDE_W, SLIDE_H))

# CUCOM's is delivered slightly oversized/offset but still covers the slide
check("near-full-bleed -> True",
      _has_fullbleed_bg([(b"", -5000, -5000, SLIDE_W + 20000, SLIDE_H + 20000)],
                        SLIDE_W, SLIDE_H))

# the "compatible" template's inset decoration (L0.82 W11.84) -> not a wedge
check("inset bg 0.82in / 11.84in wide -> False",
      not _has_fullbleed_bg(
          [(b"", Inches(0.82), Inches(0.66), Inches(11.84), Inches(6.35))],
          SLIDE_W, SLIDE_H))

check("no images -> False", not _has_fullbleed_bg([], SLIDE_W, SLIDE_H))

# small logo in a corner -> not a wedge
check("corner logo -> False",
      not _has_fullbleed_bg([(b"", 0, 0, Inches(1.5), Inches(0.8))], SLIDE_W, SLIDE_H))

print(f"\nnon-CUCOM template: {'all passed' if not fails else f'{fails} FAILED'}")
raise SystemExit(1 if fails else 0)
