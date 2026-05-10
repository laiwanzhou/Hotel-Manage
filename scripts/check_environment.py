#!/usr/bin/env python3
"""Check runtime support for CUIT thesis DOCX format workflows."""

from __future__ import annotations

import importlib.util
import json
import platform
import sys
from dataclasses import dataclass, asdict


@dataclass
class Check:
    name: str
    ok: bool
    detail: str
    impact_if_missing: str


def has_module(module: str) -> bool:
    return importlib.util.find_spec(module) is not None


def check_import(module: str, package: str, impact: str) -> Check:
    ok = has_module(module)
    return Check(
        name=package,
        ok=ok,
        detail="available" if ok else "missing",
        impact_if_missing=impact,
    )


def check_win32com() -> Check:
    if platform.system() != "Windows":
        return Check(
            name="pywin32 / win32com.client",
            ok=False,
            detail="not applicable outside Windows",
            impact_if_missing=(
                "Office/WPS COM automation is Windows-only. On non-Windows systems the tool can still run OOXML checks, "
                "but cannot ask Word/WPS to refresh pagination, fields, page numbers, or layout."
            ),
        )
    try:
        import win32com.client  # type: ignore

        return Check(
            name="pywin32 / win32com.client",
            ok=True,
            detail="available",
            impact_if_missing="",
        )
    except Exception as exc:
        return Check(
            name="pywin32 / win32com.client",
            ok=False,
            detail=f"missing or failed to import: {type(exc).__name__}: {exc}",
            impact_if_missing=(
                "Without win32com, Microsoft Word/WPS cannot be automated. The script must fall back to pure OOXML, "
                "which can create DOCX/JSON/HTML outputs but cannot reliably refresh page fields, pagination, TOC fields, "
                "or final Word/WPS layout. Page number drift and approximate screenshots are more likely."
            ),
        )


def check_com_app(prog_ids: list[str], label: str) -> Check:
    if platform.system() != "Windows":
        return Check(
            name=label,
            ok=False,
            detail="not applicable outside Windows",
            impact_if_missing="No Word/WPS COM rendering or compatibility save is available outside Windows.",
        )
    try:
        import win32com.client  # type: ignore
    except Exception:
        return Check(
            name=label,
            ok=False,
            detail="win32com unavailable",
            impact_if_missing=(
                f"{label} cannot be detected until pywin32/win32com is installed. "
                "OOXML fallback remains possible but visual layout, page numbers, and field refresh are less reliable."
            ),
        )
    errors: list[str] = []
    for prog_id in prog_ids:
        try:
            app = win32com.client.Dispatch(prog_id)
            try:
                app.Quit()
            except Exception:
                pass
            return Check(
                name=label,
                ok=True,
                detail=f"COM ProgID works: {prog_id}",
                impact_if_missing="",
            )
        except Exception as exc:
            errors.append(f"{prog_id}: {type(exc).__name__}: {exc}")
    return Check(
        name=label,
        ok=False,
        detail="; ".join(errors) if errors else "not found",
        impact_if_missing=(
            f"{label} automation is unavailable. If neither Word nor WPS COM works, the tool cannot perform a reliable "
            "compatibility save, field refresh, pagination refresh, or Word/WPS-based screenshot QA."
        ),
    )


def capability_level(checks: list[Check]) -> tuple[str, list[str]]:
    by_name = {check.name: check for check in checks}
    required = ["python-docx", "lxml", "PyMuPDF / fitz", "Pillow / PIL"]
    missing_required = [name for name in required if not by_name[name].ok]
    if missing_required:
        return "blocked", [f"Missing required package(s): {', '.join(missing_required)}"]
    word_ok = by_name["Microsoft Word COM"].ok
    wps_ok = by_name["WPS COM"].ok
    if word_ok or wps_ok:
        preferred = "Microsoft Word" if word_ok else "WPS"
        return (
            "full",
            [
                f"Full workflow available. {preferred} COM can be used for compatibility save, field refresh, pagination refresh, and more reliable screenshot QA.",
                "Use --renderer auto to prefer Word, then WPS, then OOXML fallback.",
            ],
        )
    return (
        "ooxml-only",
        [
            "OOXML-only workflow available.",
            "The tool can generate annotated DOCX, fixed DOCX, JSON, and HTML reports.",
            "High-risk layout features such as page numbers, headers/footers, section breaks, TOC fields, and final screenshots may be inaccurate until Word/WPS refreshes the document.",
        ],
    )


def main() -> int:
    checks = [
        Check(
            name="Python",
            ok=sys.version_info >= (3, 11),
            detail=sys.version.split()[0],
            impact_if_missing="Use Python 3.11 or newer for the supported runtime.",
        ),
        check_import("docx", "python-docx", "DOCX reading/writing cannot run without python-docx."),
        check_import("lxml", "lxml", "Some OOXML processing and python-docx dependencies may fail without lxml."),
        check_import("fitz", "PyMuPDF / fitz", "PDF-to-DOCX fixture generation and PDF text extraction helpers cannot run without PyMuPDF."),
        check_import("PIL", "Pillow / PIL", "Render QA image statistics are skipped without Pillow."),
        check_win32com(),
        check_com_app(["Word.Application"], "Microsoft Word COM"),
        check_com_app(["KWPS.Application", "WPS.Application", "et.Application"], "WPS COM"),
    ]
    level, notes = capability_level(checks)
    result = {
        "platform": platform.platform(),
        "python_executable": sys.executable,
        "capability_level": level,
        "notes": notes,
        "checks": [asdict(check) for check in checks],
        "install": {
            "pip": "python -m pip install -r requirements.txt",
            "conda": "conda env create -f environment.yml && conda activate cuit-thesis-docx-format",
            "windows_com_note": "On Windows, pywin32 plus installed Microsoft Word or WPS is required for reliable COM rendering and page/field refresh.",
        },
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if level in {"full", "ooxml-only"} else 1


if __name__ == "__main__":
    raise SystemExit(main())
