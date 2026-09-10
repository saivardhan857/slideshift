"""V2 Phase 4 — structured, deterministic transfer diagnostics.

A compact typed summary of what a transfer did, derived purely from the
TransferResult list + ValidationReport + measured timings. Independent of the
frontend and of any filesystem detail. Fed additively into the /api/transfer
completion event as `diagnostics`.
"""

from dataclasses import dataclass, field, asdict


@dataclass
class TransferDiagnostics:
    slides_processed: int = 0
    slides_output: int = 0
    slides_requiring_review: int = 0

    warning_count: int = 0
    error_count: int = 0

    unsupported_shape_count: int = 0          # chart / SmartArt / OLE / media / ...
    partially_supported_shape_count: int = 0  # groups where text was salvaged
    grouped_shape_salvage_count: int = 0      # subset: type == "group", salvaged

    overflow_count: int = 0                   # body columns flagged unfittable
    font_reduction_count: int = 0             # body columns with a font shrink
    spacing_reduction_count: int = 0          # body columns with a line-spacing cut

    transfer_seconds: float = 0.0
    validation_seconds: float = 0.0
    total_seconds: float = 0.0

    def to_dict(self) -> dict:
        return asdict(self)


# layout_safety.plan_fit actions and what each implies for the counters
_FONT_ACTIONS = {"reduce_font", "overflow"}
_SPACING_ACTIONS = {"reduce_spacing", "reduce_font", "overflow"}


def build_diagnostics(
    results,
    report,
    *,
    transfer_seconds: float = 0.0,
    validation_seconds: float = 0.0,
) -> TransferDiagnostics:
    """Build diagnostics from a completed transfer. Deterministic: same inputs
    (ignoring timings) always yield the same counts."""
    d = TransferDiagnostics()

    d.slides_processed = report.slide_count_source
    d.slides_output = report.slide_count_output
    d.slides_requiring_review = len(report.slides_with_warnings)
    d.warning_count = len(report.messages)
    d.error_count = sum(len(r.errors) for r in results)

    for r in results:
        for cw in r.content_warnings:
            level = cw.get("support_level", "SKIPPED_WITH_WARNING")
            if level == "PARTIALLY_SUPPORTED":
                d.partially_supported_shape_count += 1
                if cw.get("type") == "group":
                    d.grouped_shape_salvage_count += 1
            else:
                d.unsupported_shape_count += 1

        for action in getattr(r, "fit_actions", []):
            if action == "overflow":
                d.overflow_count += 1
            if action in _FONT_ACTIONS:
                d.font_reduction_count += 1
            if action in _SPACING_ACTIONS:
                d.spacing_reduction_count += 1

    d.transfer_seconds = round(float(transfer_seconds), 3)
    d.validation_seconds = round(float(validation_seconds), 3)
    d.total_seconds = round(d.transfer_seconds + d.validation_seconds, 3)
    return d


def demo():
    """Self-check with a hand-built result set."""
    class _R:
        def __init__(self, errors=None, cws=None, actions=None):
            self.errors = errors or []
            self.content_warnings = cws or []
            self.fit_actions = actions or []

    class _Rep:
        slide_count_source = 5
        slide_count_output = 5
        slides_with_warnings = [2, 4]
        messages = ["w1", "w2", "w3"]

    results = [
        _R(cws=[{"type": "group", "support_level": "PARTIALLY_SUPPORTED"}],
           actions=["reduce_spacing"]),
        _R(cws=[{"type": "chart", "support_level": "SKIPPED_WITH_WARNING"}],
           actions=["reduce_font"]),
        _R(actions=["overflow"]),
        _R(errors=["boom"], actions=["none"]),
        _R(),
    ]
    d = build_diagnostics(results, _Rep(), transfer_seconds=1.2, validation_seconds=0.3)
    assert d.slides_processed == 5 and d.slides_output == 5
    assert d.slides_requiring_review == 2
    assert d.warning_count == 3 and d.error_count == 1
    assert d.partially_supported_shape_count == 1 and d.grouped_shape_salvage_count == 1
    assert d.unsupported_shape_count == 1
    assert d.overflow_count == 1
    assert d.font_reduction_count == 2          # reduce_font + overflow
    assert d.spacing_reduction_count == 3       # reduce_spacing + reduce_font + overflow
    assert d.total_seconds == 1.5
    assert set(d.to_dict()) >= {"slides_processed", "overflow_count", "total_seconds"}
    print("[OK] diagnostics build")


if __name__ == "__main__":
    demo()
