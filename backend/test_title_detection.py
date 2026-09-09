"""Regression: list-fragment guard keeps numbered/bulleted lines out of titles.

Self-contained — no pytest.  Run:  python test_title_detection.py

Real defect: Ch04 slide 22 had an empty content placeholder plus a stray
textbox "3.\tAs they have strands of DNA..." — the position heuristic promoted
that fragment to the slide title.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from parser import _looks_like_list_fragment


def run():
    cases = {
        "3.\tAs they have strands of DNA, they are capable of self-replication": True,
        "1) First point": True,
        "12. pathological cause number 12": True,
        "• bullet item": True,
        "- dash item": True,
        "MEMBRANE PROTEINS": False,
        "Chapter 4": False,
        "Cell Membrane": False,
        "": False,
    }
    bad = [t for t, exp in cases.items() if _looks_like_list_fragment(t) != exp]
    assert not bad, f"misclassified: {bad}"
    print("[OK] list-fragment detection")


if __name__ == "__main__":
    run()
