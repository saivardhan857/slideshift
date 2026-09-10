"""Batch-convert a folder of source PPTX files through the SlideShift pipeline.

Reuses the exact pipeline main.py runs: parse_source -> transfer -> validate.
No changes to the conversion engine — this is just a loop + file handling.

Usage:
    python batch_convert.py [SOURCE_DIR] [TEMPLATE_PPTX] [OUTPUT_DIR]

Defaults target the Physiology folder / CUCOM template / "Physiology Converted".
"""

import re
import sys
import json
import traceback
from pathlib import Path

# Windows consoles default to cp1252; transfer warnings contain "⚠" etc.
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

from parser import parse_source
from transfer import transfer
from validator import validate

DESKTOP = Path.home() / "OneDrive" / "Desktop"
DEFAULT_SOURCE_DIR = DESKTOP / "Physiology"
DEFAULT_TEMPLATE = DESKTOP / "CUCOM Template.pptx"
DEFAULT_OUTPUT_DIR = DESKTOP / "Physiology Converted"


def out_name(src: Path) -> str:
    """Physiology_Chapter_NN_Converted.pptx, falling back to the sanitized stem."""
    m = re.search(r"chapter\s*0*(\d+)", src.stem, re.IGNORECASE)
    if m:
        return f"Physiology_Chapter_{int(m.group(1)):02d}_Converted.pptx"
    safe = re.sub(r"[^A-Za-z0-9]+", "_", src.stem).strip("_")
    return f"Physiology_{safe}_Converted.pptx"


def safe_path(dir_: Path, name: str) -> Path:
    """Never overwrite: append ' (2)', ' (3)', ... if the target exists."""
    p = dir_ / name
    if not p.exists():
        return p
    stem, suffix = p.stem, p.suffix
    n = 2
    while (dir_ / f"{stem} ({n}){suffix}").exists():
        n += 1
    return dir_ / f"{stem} ({n}){suffix}"


def main():
    source_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SOURCE_DIR
    template = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_TEMPLATE
    output_dir = Path(sys.argv[3]) if len(sys.argv) > 3 else DEFAULT_OUTPUT_DIR

    if not source_dir.is_dir():
        sys.exit(f"Source dir not found: {source_dir}")
    if not template.is_file():
        sys.exit(f"Template not found: {template}")

    output_dir.mkdir(parents=True, exist_ok=True)

    sources = sorted(p for p in source_dir.glob("*.pptx") if not p.name.startswith("~$"))
    print(f"Source dir : {source_dir}")
    print(f"Template   : {template}")
    print(f"Output dir : {output_dir}")
    print(f"Found {len(sources)} source .pptx file(s)\n")

    report = []
    for i, src in enumerate(sources, 1):
        entry = {"source": src.name, "output": None, "status": "failed",
                 "slides_source": None, "slides_output": None,
                 "warnings": [], "issues": []}
        print(f"[{i}/{len(sources)}] {src.name}")
        try:
            slides = parse_source(str(src))
            entry["slides_source"] = len(slides)
            if not slides:
                entry["issues"].append("Source presentation is empty (0 slides).")
                report.append(entry)
                print("      -> FAILED: empty source\n")
                continue

            dest = safe_path(output_dir, out_name(src))
            entry["output"] = dest.name

            results = transfer(
                source_slides=slides,
                source_path=str(src),
                template_path=str(template),
                output_path=str(dest),
                progress_callback=None,
            )
            vr = validate(str(dest), len(slides), results)
            entry["slides_output"] = vr.slide_count_output
            entry["warnings"] = list(vr.messages)

            # QC checks
            if not dest.exists():
                entry["issues"].append("Output file was not created.")
            if vr.slide_count_output != len(slides):
                entry["issues"].append(
                    f"Slide count mismatch: source {len(slides)}, output {vr.slide_count_output}.")
            if not vr.ok:
                entry["issues"].append("Validator flagged output as not OK (errors on some slides).")

            entry["status"] = "converted" if dest.exists() and vr.slide_count_output > 0 else "failed"
            print(f"      -> {entry['status'].upper()}: {dest.name} "
                  f"({vr.slide_count_source} -> {vr.slide_count_output} slides, "
                  f"{len(vr.slides_with_warnings)} need review)\n")
        except Exception as e:
            entry["issues"].append(f"{type(e).__name__}: {e}")
            entry["traceback"] = traceback.format_exc()
            print(f"      -> FAILED: {e}\n")
        report.append(entry)

    converted = [e for e in report if e["status"] == "converted"]
    failed = [e for e in report if e["status"] != "converted"]

    print("=" * 70)
    print("BATCH SUMMARY")
    print("=" * 70)
    print(f"Source PPTs processed : {len(report)}")
    print(f"Successfully converted : {len(converted)}")
    print(f"Failed                : {len(failed)}")
    print(f"Output folder         : {output_dir}")
    print("\nGenerated files:")
    for e in converted:
        flag = "  (review warnings)" if e["warnings"] else ""
        print(f"  - {e['output']}{flag}")
    if failed:
        print("\nFailed:")
        for e in failed:
            print(f"  - {e['source']}: {'; '.join(e['issues']) or 'unknown error'}")

    print("\nPer-file warnings / QC issues:")
    any_warn = False
    for e in report:
        lines = []
        for w in e["warnings"]:
            lines.append(f"      warn: {w}")
        for iss in e["issues"]:
            lines.append(f"      ISSUE: {iss}")
        if lines:
            any_warn = True
            print(f"  {e['source']}  ->  {e['output'] or '(none)'}")
            print("\n".join(lines))
    if not any_warn:
        print("  (none)")

    (output_dir / "_conversion_report.json").write_text(
        json.dumps(report, indent=2), encoding="utf-8")
    print(f"\nFull machine-readable report: {output_dir / '_conversion_report.json'}")


if __name__ == "__main__":
    main()
