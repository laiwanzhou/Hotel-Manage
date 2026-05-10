#!/usr/bin/env python3
"""Run CUIT thesis DOCX checks for one or more documents with a matrix summary."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
CHECKER = SCRIPT_DIR / "cuit_thesis_docx_format.py"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_folder_name(path: Path, used: set[str]) -> str:
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", path.stem).strip(" .")
    stem = stem or "document"
    candidate = stem
    counter = 2
    while candidate.casefold() in used:
        candidate = f"{stem}_{counter}"
        counter += 1
    used.add(candidate.casefold())
    return candidate


def load_report(report_path: Path) -> dict[str, object] | None:
    if not report_path.exists():
        return None
    try:
        return json.loads(report_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return None


def validate_outputs(input_path: Path, output_dir: Path, before_hash: str) -> dict[str, object]:
    stem = input_path.stem
    paths = {
        "comments": output_dir / f"{stem}_format_comments.docx",
        "fixed": output_dir / f"{stem}_format_fixed.docx",
        "json": output_dir / f"{stem}_format_report.json",
        "html": output_dir / f"{stem}_format_report.html",
        "run_output": output_dir / "run-output.json",
    }
    report = load_report(paths["json"])
    after_hash = sha256(input_path) if input_path.exists() else None
    checks = {
        "input_unchanged": before_hash == after_hash,
        "comments_exists": paths["comments"].exists(),
        "fixed_exists": paths["fixed"].exists(),
        "json_exists": paths["json"].exists(),
        "html_exists": paths["html"].exists(),
        "run_output_exists": paths["run_output"].exists(),
        "report_readable": report is not None,
        "issue_count_matches_modification_count": bool(
            report and report.get("issue_count") == report.get("modification_count")
        ),
    }
    return {
        "checks": checks,
        "passed": all(checks.values()),
        "paths": {key: str(value.resolve()) for key, value in paths.items()},
        "issue_count": report.get("issue_count") if report else None,
        "modification_count": report.get("modification_count") if report else None,
        "issue_summary_by_category": report.get("issue_summary_by_category") if report else None,
        "renderer_for_fixed": report.get("renderer_for_fixed") if report else None,
        "planned_renderer_for_layout_fixes": report.get("planned_renderer_for_layout_fixes") if report else None,
        "high_risk_layout_fixes_applied": report.get("high_risk_layout_fixes_applied") if report else None,
        "allow_ooxml_layout_fixes": report.get("allow_ooxml_layout_fixes") if report else None,
        "screenshot_status": report.get("screenshot_status") if report else None,
        "render_warning_count": report.get("render_qa", {}).get("warning_count") if report and report.get("render_qa") else None,
    }


def run_one(
    input_path: Path,
    output_dir: Path,
    renderer: str,
    screenshots: str,
    timeout: int,
    allow_ooxml_layout_fixes: bool,
) -> dict[str, object]:
    before_hash = sha256(input_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    command = [
        sys.executable,
        str(CHECKER),
        str(input_path),
        "--output-dir",
        str(output_dir),
        "--renderer",
        renderer,
        "--screenshots",
        screenshots,
    ]
    if allow_ooxml_layout_fixes:
        command.append("--allow-ooxml-layout-fixes")
    proc = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )
    run_output_path = output_dir / "run-output.json"
    if proc.stdout.strip():
        run_output_path.write_text(proc.stdout, encoding="utf-8")
    else:
        run_output_path.write_text(
            json.dumps(
                {
                    "error": "checker produced no stdout",
                    "returncode": proc.returncode,
                    "stderr": proc.stderr,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
    stderr_path = output_dir / "stderr.txt"
    if proc.stderr.strip():
        stderr_path.write_text(proc.stderr, encoding="utf-8")
    validation = validate_outputs(input_path, output_dir, before_hash)
    return {
        "input": str(input_path.resolve()),
        "output_dir": str(output_dir.resolve()),
        "command": command,
        "returncode": proc.returncode,
        "stderr_path": str(stderr_path.resolve()) if stderr_path.exists() else None,
        **validation,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Batch-run CUIT thesis DOCX format checks.")
    parser.add_argument("docx", nargs="+", help="Input .docx thesis file(s)")
    parser.add_argument("--output-dir", default="./resx", help="Batch output directory")
    parser.add_argument("--renderer", choices=["auto", "office", "wps", "ooxml"], default="auto")
    parser.add_argument("--screenshots", choices=["auto", "never", "require"], default="auto")
    parser.add_argument("--timeout", type=int, default=900, help="Per-document timeout in seconds")
    parser.add_argument(
        "--allow-ooxml-layout-fixes",
        action="store_true",
        help="Explicitly allow OOXML-only changes to headers, footers, page numbers, and section behavior.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output_root = Path(args.output_dir)
    output_root.mkdir(parents=True, exist_ok=True)
    used_names: set[str] = set()
    results: list[dict[str, object]] = []
    for raw in args.docx:
        input_path = Path(raw).resolve()
        if input_path.suffix.lower() != ".docx":
            results.append(
                {
                    "input": str(input_path),
                    "passed": False,
                    "error": "input is not a .docx file",
                }
            )
            continue
        if not input_path.exists():
            results.append(
                {
                    "input": str(input_path),
                    "passed": False,
                    "error": "input file does not exist",
                }
            )
            continue
        child_dir = output_root / safe_folder_name(input_path, used_names)
        try:
            results.append(
                run_one(
                    input_path,
                    child_dir,
                    args.renderer,
                    args.screenshots,
                    args.timeout,
                    args.allow_ooxml_layout_fixes,
                )
            )
        except subprocess.TimeoutExpired as exc:
            child_dir.mkdir(parents=True, exist_ok=True)
            timeout_record = {
                "input": str(input_path.resolve()),
                "output_dir": str(child_dir.resolve()),
                "passed": False,
                "error": f"timeout after {args.timeout} seconds",
                "stdout": exc.stdout,
                "stderr": exc.stderr,
            }
            (child_dir / "run-output.json").write_text(json.dumps(timeout_record, ensure_ascii=False, indent=2), encoding="utf-8")
            results.append(timeout_record)
    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "output_root": str(output_root.resolve()),
        "renderer": args.renderer,
        "screenshots": args.screenshots,
        "allow_ooxml_layout_fixes": args.allow_ooxml_layout_fixes,
        "total": len(results),
        "passed": sum(1 for result in results if result.get("passed")),
        "failed": sum(1 for result in results if not result.get("passed")),
        "results": results,
    }
    summary_path = output_root / "batch-summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary["failed"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
