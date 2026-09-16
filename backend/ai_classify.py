"""AI-assisted title/body reclassification.

The regex/position heuristic in parser.py gets this wrong on real decks --
e.g. a body sentence fragment ("Articular surface the head of the femur,")
promoted to the slide title just because it's short and sits near the top.
This module sends the heuristic's existing title/body split for the whole
deck to Gemini in one batched call and lets it correct mistakes.

No API key, any network failure, a timeout, or a malformed response all fall
back to the heuristic's existing (unmodified) result -- this must never
block or break a conversion. classify_titles() never raises.
"""
import json
import logging
import os
import urllib.error
import urllib.request

logger = logging.getLogger("slideshift")

_MODEL = "gemini-flash-latest"  # Google's own stable alias -- avoids hardcoding a
                                 # specific dated version that gets retired/unstable
                                 # (gemini-2.0-flash was retired, then gemini-3.6-flash
                                 # showed intermittent 400/503 errors under real load)
_API_URL = f"https://generativelanguage.googleapis.com/v1beta/models/{_MODEL}:generateContent"

_SYSTEM_PROMPT = (
    "You are reviewing text fragments extracted from lecture slides. Each "
    "slide has one or more fragments; a heuristic guess has already picked "
    "one candidate TITLE (a short heading naming the slide's topic) and "
    "marked the rest BODY (bullets, explanations, labels, or continuations "
    "of a thought). The heuristic sometimes gets this wrong -- a fragment "
    "that continues a sentence, a list, or reads like a bullet is BODY even "
    "if it is short and sits near the top of the slide. A slide can "
    "legitimately have no real title at all.\n\n"
    "Review every fragment on every slide and decide the correct role for "
    "each. Respond with ONLY a JSON array, no prose, no markdown fences: "
    '[{"slide": <int>, "fragment": <int>, "role": "title"|"body"}, ...]'
)


def _slide_fragments(parsed):
    """[(fragment_index, TextBox, original_role)] for one ParsedSlide's
    non-empty text fragments -- title first (if any) at index 0, then
    body_boxes in original order. Empty-text fragments are skipped so the
    same indices are used consistently when building the request and when
    applying the response."""
    frags = []
    if parsed.title is not None and parsed.title.full_text.strip():
        frags.append((0, parsed.title, "title"))
    for i, b in enumerate(parsed.body_boxes):
        if b.full_text.strip():
            frags.append((i + 1, b, "body"))
    return frags


def _build_payload(slides):
    payload = []
    for s in slides:
        frags = _slide_fragments(s)
        if not frags:
            continue
        payload.append({
            "slide": s.index,
            "fragments": [{"fragment": idx, "text": tb.full_text.strip()}
                          for idx, tb, _role in frags],
        })
    return payload


def _apply_roles(slides, role_map):
    """role_map: {(slide_index, fragment_index): "title"|"body"}. Mutates
    each ParsedSlide's .title/.body_boxes based on role_map, defaulting any
    fragment role_map has no opinion on to its ORIGINAL heuristic role (a
    partial response should change as little as possible). At most one
    fragment per slide keeps the "title" role -- if the model marks more
    than one, the first (by original index) wins and the rest fall back to
    body, in their original relative order."""
    for s in slides:
        frags = _slide_fragments(s)
        if not frags:
            continue
        roles = [(idx, tb, role_map.get((s.index, idx), original_role))
                  for idx, tb, original_role in frags]
        title_frags = [(idx, tb) for idx, tb, role in roles if role == "title"]
        new_title_idx, new_title = title_frags[0] if title_frags else (None, None)
        s.title = new_title
        s.body_boxes = [tb for idx, tb, _role in roles if idx != new_title_idx]


def classify_titles(slides, timeout: float = 45.0) -> bool:
    """Ask Gemini to correct parser.py's title/body split for this deck.
    Returns True if AI classification was applied, False if it fell back to
    the heuristic's existing result (no key, request failed, bad response)."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return False

    payload = _build_payload(slides)
    if not payload:
        return False

    try:
        body = json.dumps({
            "systemInstruction": {"parts": [{"text": _SYSTEM_PROMPT}]},
            "contents": [{"parts": [{"text": json.dumps(payload)}]}],
            "generationConfig": {"maxOutputTokens": 4096, "responseMimeType": "application/json"},
        }).encode("utf-8")
        req = urllib.request.Request(
            _API_URL, data=body, method="POST",
            headers={
                "x-goog-api-key": api_key,
                "content-type": "application/json",
            },
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            response = json.loads(resp.read().decode("utf-8"))

        text = response["candidates"][0]["content"]["parts"][0]["text"]
        roles = json.loads(text)
        role_map = {
            (int(r["slide"]), int(r["fragment"])): r["role"]
            for r in roles if r.get("role") in ("title", "body")
        }
        if not role_map:
            return False

        _apply_roles(slides, role_map)
        return True
    except urllib.error.HTTPError as e:
        # Google's error body is diagnostic text (e.g. "API_KEY_INVALID",
        # "PERMISSION_DENIED") -- never the key itself -- safe to log.
        try:
            detail = e.read().decode("utf-8", errors="replace")[:300]
        except Exception:
            detail = ""
        logger.warning("AI title classification skipped (HTTPError %s): %s", e.code, detail)
        return False
    except Exception as e:
        logger.warning("AI title classification skipped (%s)", type(e).__name__)
        return False
