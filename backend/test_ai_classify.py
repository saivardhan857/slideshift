"""AI title/body reclassification (ai_classify.py).

No live API key or network access needed -- urllib.request.urlopen is
mocked. Covers: no-key fallback, a corrected response actually moving
fragments between .title/.body_boxes (preserving original body order), and
a malformed response falling back safely instead of raising.

Self-contained — no pytest.  Run:  python test_ai_classify.py
"""
import io
import json
import os
from contextlib import contextmanager
from unittest.mock import patch

from parser import ParsedSlide, TextBox, Paragraph, TextRun
import ai_classify
from ai_classify import classify_titles, _apply_roles, _slide_fragments

fails = 0
def check(name, cond, detail=""):
    global fails
    print(f"  [{'OK' if cond else 'FAIL'}] {name}{'' if cond else '  ' + str(detail)}")
    if not cond:
        fails += 1


def _tb(text, is_title=False, is_body=False):
    return TextBox(paragraphs=[Paragraph(runs=[TextRun(text=text)])],
                    is_title=is_title, is_body=is_body)


def _mk_slide(index, title_text, body_texts):
    title = _tb(title_text, is_title=True) if title_text else None
    return ParsedSlide(index=index, title=title,
                        body_boxes=[_tb(t, is_body=True) for t in body_texts])


@contextmanager
def _fake_response(payload_bytes):
    class _Resp:
        def read(self_inner):
            return payload_bytes
        def __enter__(self_inner):
            return self_inner
        def __exit__(self_inner, *a):
            return False
    yield _Resp()


# --------------------------------------------------------------------------- #
# Test 1 — no API key configured -> clean no-op, heuristic result untouched.
# --------------------------------------------------------------------------- #
slides1 = [_mk_slide(0, "Misdetected title fragment", ["Real body one", "Real body two"])]
with patch.dict(os.environ, {}, clear=False):
    os.environ.pop("ANTHROPIC_API_KEY", None)
    applied1 = classify_titles(slides1)
check("Test1: no key -> returns False", applied1 is False)
check("Test1: no key -> title untouched",
      slides1[0].title is not None and slides1[0].title.full_text == "Misdetected title fragment")
check("Test1: no key -> body untouched", len(slides1[0].body_boxes) == 2)

# --------------------------------------------------------------------------- #
# Test 2 — a corrected response demotes the misdetected title and promotes
# nothing new; original body order is preserved.
# --------------------------------------------------------------------------- #
slides2 = [_mk_slide(0, "Misdetected title fragment", ["Real body one", "Real body two"])]
fake_roles = [{"slide": 0, "fragment": 0, "role": "body"}]
fake_api_response = json.dumps({"content": [{"type": "text", "text": json.dumps(fake_roles)}]}).encode()

with patch.dict(os.environ, {"ANTHROPIC_API_KEY": "test-key"}):
    with patch("urllib.request.urlopen", return_value=_fake_response(fake_api_response).__enter__()):
        applied2 = classify_titles(slides2)

check("Test2: corrected response -> returns True", applied2 is True)
check("Test2: slide now has no title", slides2[0].title is None, slides2[0].title)
check("Test2: demoted fragment lands FIRST, original body order preserved after it",
      [b.full_text for b in slides2[0].body_boxes] ==
      ["Misdetected title fragment", "Real body one", "Real body two"],
      [b.full_text for b in slides2[0].body_boxes])

# --------------------------------------------------------------------------- #
# Test 3 — malformed response (not JSON) falls back safely, no exception,
# original heuristic result untouched.
# --------------------------------------------------------------------------- #
slides3 = [_mk_slide(0, "Some title", ["Some body"])]
garbage_response = json.dumps({"content": [{"type": "text", "text": "not valid json {{{"}]}).encode()

with patch.dict(os.environ, {"ANTHROPIC_API_KEY": "test-key"}):
    with patch("urllib.request.urlopen", return_value=_fake_response(garbage_response).__enter__()):
        applied3 = classify_titles(slides3)

check("Test3: malformed response -> returns False, no exception", applied3 is False)
check("Test3: malformed response -> title untouched",
      slides3[0].title is not None and slides3[0].title.full_text == "Some title")

# --------------------------------------------------------------------------- #
# Test 4 — _apply_roles directly: promoting a body fragment to title demotes
# the old title into body at ITS original relative position (not shoved to
# the front), and a role_map with no opinion on a fragment leaves it as-is.
# --------------------------------------------------------------------------- #
slides4 = [_mk_slide(0, "Old title", ["First body", "Should become title", "Last body"])]
# fragment 0 = "Old title" (demoted), fragment 2 = "Should become title"
# (body_boxes[1], since body fragments are indexed 1..n) promoted.
role_map = {(0, 0): "body", (0, 2): "title"}
_apply_roles(slides4, role_map)
check("Test4: new title promoted correctly",
      slides4[0].title is not None and slides4[0].title.full_text == "Should become title",
      slides4[0].title)
check("Test4: old title demoted into its original relative slot, other fragments untouched by role_map keep their role",
      [b.full_text for b in slides4[0].body_boxes] == ["Old title", "First body", "Last body"],
      [b.full_text for b in slides4[0].body_boxes])

print(f"\nAI title classification: {'all passed' if not fails else f'{fails} FAILED'}")
if __name__ == "__main__":
    raise SystemExit(1 if fails else 0)
