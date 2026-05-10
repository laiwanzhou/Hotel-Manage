#!/usr/bin/env python3
"""Check and fix CUIT bachelor thesis DOCX formatting."""

from __future__ import annotations

import argparse
from copy import deepcopy
import html
import json
import re
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass, replace
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterable
from xml.etree import ElementTree as ET

try:
    from docx import Document
    from docx.enum.section import WD_SECTION_START
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
except ModuleNotFoundError as exc:
    Document = None
    WD_SECTION_START = None
    WD_ALIGN_PARAGRAPH = None
    OxmlElement = None
    qn = None
    Cm = None
    Pt = None
    DOCX_IMPORT_ERROR = exc
else:
    DOCX_IMPORT_ERROR = None


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("", CONTENT_NS)


if WD_ALIGN_PARAGRAPH is None:
    ALIGN_LABELS = {}
    ALIGN_LABELS_ZH = {}
    ALIGNMENT_VALUES = {"left": None, "center": None, "right": None, "justified": None, None: None}
else:
    ALIGN_LABELS = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justified",
    }

    ALIGN_LABELS_ZH = {
        WD_ALIGN_PARAGRAPH.LEFT: "左对齐",
        WD_ALIGN_PARAGRAPH.CENTER: "居中",
        WD_ALIGN_PARAGRAPH.RIGHT: "右对齐",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "两端对齐",
    }

    ALIGNMENT_VALUES = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justified": WD_ALIGN_PARAGRAPH.JUSTIFY,
        None: None,
    }

ALIGN_LEFT = ALIGNMENT_VALUES["left"]
ALIGN_CENTER = ALIGNMENT_VALUES["center"]
ALIGN_RIGHT = ALIGNMENT_VALUES["right"]
ALIGN_JUSTIFIED = ALIGNMENT_VALUES["justified"]

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
RULE_CONFIG_PATH = SKILL_DIR / "references" / "rules.json"


def require_docx_dependencies() -> None:
    if DOCX_IMPORT_ERROR is None:
        return
    raise RuntimeError(
        "Missing required DOCX dependency. Run `python -m pip install -r requirements.txt` "
        "from the cuit-thesis-docx-format skill directory, then retry."
    ) from DOCX_IMPORT_ERROR


@dataclass(frozen=True)
class Rule:
    key: str
    label: str
    font_east_asia: str
    font_ascii: str
    size_pt: float
    bold: bool | None
    alignment: int | None
    line_spacing_pt: float | None = 20
    first_line_indent_cm: float | None = None
    first_line_indent_chars: float | None = None
    space_before_pt: float | None = None
    space_after_pt: float | None = None
    space_before_lines: float | None = None
    space_after_lines: float | None = None
    expected: str = ""
    category: str = "format"


@dataclass
class Issue:
    paragraph_index: int
    rule_key: str
    text_type: str
    text_excerpt: str
    current: str
    expected: str
    message: str
    category: str = "format"
    location: str | None = None
    after: str | None = None
    page: int | None = None
    before_screenshot: str | None = None
    after_screenshot: str | None = None
    screenshot_note: str | None = None


RULES: dict[str, Rule] = {
    "chapter": Rule(
        "chapter",
        "chapter title",
        "宋体",
        "Times New Roman",
        16,
        True,
        ALIGN_CENTER,
        space_before_lines=0.5,
        space_after_lines=0.5,
        expected="宋体三号16pt，加粗，居中，固定20磅行距，段前0.5行，段后0.5行，章序号与章名间空一格。",
    ),
    "heading2": Rule(
        "heading2",
        "second-level heading",
        "宋体",
        "Times New Roman",
        14,
        True,
        ALIGN_LEFT,
        space_before_lines=0.5,
        space_after_lines=0.5,
        expected="宋体四号14pt，加粗，左对齐，固定20磅行距，段前0.5行，段后0.5行，序号与题名间空一格。",
    ),
    "heading3": Rule(
        "heading3",
        "third-level heading",
        "宋体",
        "Times New Roman",
        12,
        True,
        ALIGN_LEFT,
        space_before_lines=0.5,
        space_after_lines=0.5,
        expected="宋体小四12pt，加粗，左对齐，固定20磅行距，段前0.5行，段后0.5行，序号与题名间空一格。",
    ),
    "body": Rule(
        "body",
        "body paragraph",
        "宋体",
        "Times New Roman",
        12,
        None,
        ALIGN_LEFT,
        first_line_indent_chars=2,
        space_before_pt=0,
        space_after_pt=0,
        expected="中文宋体小四12pt，英文和数字Times New Roman 12pt，左对齐，首行缩进2个汉字符，段前段后0磅，固定20磅行距。",
    ),
    "figure_caption": Rule(
        "figure_caption",
        "figure caption",
        "宋体",
        "Times New Roman",
        10.5,
        None,
        ALIGN_CENTER,
        first_line_indent_chars=0,
        space_before_lines=0,
        space_after_lines=0,
        expected="图题置于图下方，宋体五号10.5pt，居中，固定20磅行距，图序与图名之间空一格。",
    ),
    "table_caption": Rule(
        "table_caption",
        "table caption",
        "宋体",
        "Times New Roman",
        10.5,
        None,
        ALIGN_CENTER,
        first_line_indent_chars=0,
        space_before_lines=0,
        space_after_lines=0,
        expected="表题置于表上方，宋体五号10.5pt，居中，固定20磅行距，表序与表名之间空一格。",
    ),
    "abstract_title_zh": Rule(
        "abstract_title_zh",
        "Chinese abstract title",
        "宋体",
        "Times New Roman",
        16,
        True,
        ALIGN_CENTER,
        space_before_lines=0.5,
        space_after_lines=0.5,
        expected="摘要标题写作“摘 要”，宋体三号16pt，加粗，居中，段前0.5行，段后0.5行，固定20磅行距。",
    ),
    "abstract_title_en": Rule(
        "abstract_title_en",
        "English abstract title",
        "Times New Roman",
        "Times New Roman",
        14,
        True,
        ALIGN_CENTER,
        space_before_lines=0.5,
        space_after_lines=0.5,
        expected="英文摘要标题写作“ABSTRACT”，Times New Roman四号14pt，加粗，居中，段前0.5行，段后0.5行。",
    ),
    "keywords": Rule(
        "keywords",
        "keywords paragraph",
        "宋体",
        "Times New Roman",
        12,
        None,
        ALIGN_LEFT,
        expected="关键词段落宋体小四12pt，英文Times New Roman 12pt，关键词标签加粗，关键词之间用分号隔开。",
    ),
    "toc_title": Rule(
        "toc_title",
        "table of contents title",
        "宋体",
        "Times New Roman",
        16,
        True,
        ALIGN_CENTER,
        space_before_lines=0.5,
        space_after_lines=0.5,
        expected="目录标题写作“目 录”，宋体三号16pt，加粗，居中，段前0.5行，段后0.5行。",
    ),
}


def _rule_from_config(key: str, data: dict[str, object]) -> Rule:
    return Rule(
        key=key,
        label=str(data.get("label", key)),
        font_east_asia=str(data.get("font_east_asia", "宋体")),
        font_ascii=str(data.get("font_ascii", "Times New Roman")),
        size_pt=float(data.get("size_pt", 12)),
        bold=data.get("bold") if data.get("bold") is None else bool(data.get("bold")),
        alignment=ALIGNMENT_VALUES.get(data.get("alignment"), None),
        line_spacing_pt=data.get("line_spacing_pt") if data.get("line_spacing_pt") is None else float(data.get("line_spacing_pt")),
        first_line_indent_cm=data.get("first_line_indent_cm") if data.get("first_line_indent_cm") is None else float(data.get("first_line_indent_cm")),
        first_line_indent_chars=data.get("first_line_indent_chars") if data.get("first_line_indent_chars") is None else float(data.get("first_line_indent_chars")),
        space_before_pt=data.get("space_before_pt") if data.get("space_before_pt") is None else float(data.get("space_before_pt")),
        space_after_pt=data.get("space_after_pt") if data.get("space_after_pt") is None else float(data.get("space_after_pt")),
        space_before_lines=data.get("space_before_lines") if data.get("space_before_lines") is None else float(data.get("space_before_lines")),
        space_after_lines=data.get("space_after_lines") if data.get("space_after_lines") is None else float(data.get("space_after_lines")),
        expected=str(data.get("expected", "")),
        category=str(data.get("category", "format")),
    )


def load_rules() -> dict[str, Rule]:
    if not RULE_CONFIG_PATH.exists():
        return RULES
    data = json.loads(RULE_CONFIG_PATH.read_text(encoding="utf-8"))
    return {key: _rule_from_config(key, value) for key, value in data["rules"].items()}


RULES = load_rules()
if "keywords" in RULES:
    RULES.setdefault("keywords_zh", replace(RULES["keywords"], key="keywords_zh", label="Chinese keywords paragraph"))
    RULES.setdefault("keywords_en", replace(RULES["keywords"], key="keywords_en", label="English keywords paragraph"))
if "cover_title_zh" in RULES:
    RULES.setdefault("thesis_title_zh", replace(RULES["cover_title_zh"], key="thesis_title_zh", label="Chinese thesis title"))
if "cover_title_en" in RULES:
    RULES.setdefault("thesis_title_en", replace(RULES["cover_title_en"], key="thesis_title_en", label="English thesis title"))


def is_docx(path: Path) -> bool:
    return path.suffix.lower() == ".docx"


def paragraph_text(paragraph) -> str:
    return re.sub(r"\s+", " ", paragraph.text or "").strip()


SECTION_ORDER = {
    "cover": 0,
    "declaration": 1,
    "abstract_zh": 2,
    "abstract_en": 2,
    "toc": 3,
    "symbols": 4,
    "body": 5,
    "references": 6,
    "appendix": 7,
    "acknowledgement": 8,
}

SECTION_LABELS = {
    "cover": "封面",
    "declaration": "封二",
    "abstract": "摘要和关键词",
    "toc": "目录",
    "symbols": "符号说明",
    "body": "正文",
    "references": "参考文献",
    "appendix": "附录",
    "acknowledgement": "致谢",
}


def compact_text(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def detect_section_marker(text: str) -> str | None:
    compact = compact_text(text)
    if not compact:
        return None
    if "原创性声明" in compact or "版权使用授权书" in compact:
        return "declaration"
    if compact in {"摘要", "摘要:"} or text == "摘 要" or re.match(r"^摘要\s*[:：]", text):
        return "abstract_zh"
    if compact.upper() == "ABSTRACT" or re.match(r"^Abstract\s*[:：]", text, re.I):
        return "abstract_en"
    if compact == "目录" or text == "目 录":
        return "toc"
    if compact == "符号说明":
        return "symbols"
    if (
        re.match(r"^第[一二三四五六七八九十百\d]+章\S*", compact)
        or compact in {"引言", "绪论"}
        or re.match(r"^\d+(引言|绪论)$", compact)
    ):
        return "body"
    if compact == "参考文献":
        return "references"
    if compact in {"附录", "附錄"}:
        return "appendix"
    if compact in {"致谢", "致謝"}:
        return "acknowledgement"
    return None


def analyze_section_sequence(texts: list[str]) -> dict[str, object]:
    regions: list[str] = []
    warnings: list[str] = []
    present: set[str] = set()
    current = "cover"
    current_order = SECTION_ORDER[current]
    saw_nonempty = False

    for index, text in enumerate(texts):
        text = text.strip()
        if not text:
            regions.append(current)
            continue
        saw_nonempty = True
        marker = detect_section_marker(text)
        if marker is not None:
            if current == "toc" and marker == "body":
                # TOC lines can contain chapter patterns like "第1章 ... 1"; do not
                # advance to body until a real body heading appears outside TOC-entry shape.
                compact = compact_text(text)
                looks_like_toc_entry = (
                    ("." in text and re.search(r"\d+\s*$", text) is not None)
                    or ("……" in text or "..." in text)
                    or bool(re.match(r"^第[一二三四五六七八九十百\d]+章", compact) and re.search(r"\d+\s*$", compact))
                )
                if looks_like_toc_entry:
                    marker = None
            if marker is None:
                regions.append(current)
                present.add(current)
                continue
            marker_order = SECTION_ORDER[marker]
            if marker_order < current_order:
                warnings.append(
                    f"组成部分顺序疑似错位：段落 {index} 出现“{text[:40]}”，应属于{SECTION_LABELS.get(marker, marker)}，"
                    f"但当前已经进入{SECTION_LABELS.get(current, current)}。"
                )
            else:
                current = marker
                current_order = marker_order
        regions.append(current)
        present.add(current)

    if saw_nonempty:
        present.add("cover")
    if "abstract_zh" in present or "abstract_en" in present:
        present.add("abstract")

    required = ["cover", "declaration", "abstract", "toc", "body", "references", "acknowledgement"]
    optional = ["symbols", "appendix"]
    for key in required:
        if key not in present:
            warnings.append(f"缺失必需组成部分：{SECTION_LABELS[key]}。")
    for key in optional:
        if key not in present:
            warnings.append(f"未检测到非必需组成部分：{SECTION_LABELS[key]}；如论文需要该部分，请补充并按规范排版。")

    return {
        "regions": regions,
        "present": sorted(present),
        "warnings": warnings,
    }


def detect_abstract_thesis_titles(texts: list[str]) -> dict[int, str]:
    title_rules: dict[int, str] = {}
    previous_nonempty: int | None = None
    for index, text in enumerate(texts):
        stripped = text.strip()
        if not stripped:
            continue
        marker = detect_section_marker(stripped)
        if marker == "abstract_zh" and previous_nonempty is not None:
            title_rules.setdefault(previous_nonempty, "thesis_title_zh")
        elif marker == "abstract_en" and previous_nonempty is not None:
            title_rules.setdefault(previous_nonempty, "thesis_title_en")
        previous_nonempty = index
    return title_rules


def style_heading_key(paragraph, region: str) -> str | None:
    if region != "body":
        return None
    style_name = (paragraph.style.name if paragraph.style else "").strip().lower()
    style_id = (paragraph.style.style_id if paragraph.style else "").strip().lower()
    style_token = f"{style_name} {style_id}"
    if "heading 1" in style_token or "heading1" in style_token:
        return "chapter"
    if "heading 2" in style_token or "heading2" in style_token:
        return "heading2"
    if "heading 3" in style_token or "heading3" in style_token:
        return "heading3"
    return None


def classify_checked_paragraph(paragraph, idx: int, text: str, region: str, title_overrides: dict[int, str]) -> str | None:
    style_name = (paragraph.style.name if paragraph.style else "").lower()
    if style_name.startswith("toc") and text.replace(" ", "") != "目录":
        return "toc_level3" if style_name.endswith("3") else "toc_level2" if style_name.endswith("2") else "toc_level1"
    return title_overrides.get(idx) or style_heading_key(paragraph, region) or classify_paragraph(text, in_body=region == "body", region=region)


def classify_paragraph(text: str, in_body: bool, region: str = "front") -> str | None:
    compact = compact_text(text)
    if re.match(r"^论文总页数[:：]\s*\d+\s*页", text):
        return "toc_total_pages"
    if compact in {"摘要", "摘要:"} or text == "摘 要":
        return "abstract_title_zh"
    if compact.upper() == "ABSTRACT":
        return "abstract_title_en"
    if region == "abstract_zh" and re.match(r"^(关键词|关键字)\s*[:：]", text):
        return "keywords_zh"
    if region == "abstract_en" and re.match(r"^(Key\s*words?|Keywords?)\s*[:：]", text, re.I):
        return "keywords_en"
    if compact == "目录" or text == "目 录":
        return "toc_title"
    if re.match(r"^第[一二三四五六七八九十百\d]+章\s*\S+", text):
        return "chapter"
    if region == "body" and re.match(r"^\d+(引言|绪论)$", compact):
        return "chapter"
    if compact == "参考文献":
        return "reference_title"
    if compact == "符号说明":
        return "symbols_title"
    if compact in {"附录", "附錄"}:
        return "appendix_title"
    if compact in {"致谢", "致謝"}:
        return "acknowledgement_title"
    if "作者简历" in compact or "攻读学位期间发表" in compact:
        return "author_bio_title"
    if region == "toc":
        if re.match(r"^第[一二三四五六七八九十百\d]+章\s+\S+", text):
            return "toc_level1"
        if re.match(r"^\d+\.\d+\.\d+\s+\S+", text):
            return "toc_level3"
        if re.match(r"^\d+\.\d+\s+\S+", text):
            return "toc_level2"
    if re.match(r"^\d+\.\d+\.\d+\s*\S+", text):
        return "heading3"
    if re.match(r"^\d+\.\d+\s*\S+", text):
        return "heading2"
    if re.match(r"^图\s*\d+[\.\-－]\d+\s*\S+", text):
        return "figure_caption"
    if re.match(r"^表\s*\d+[\.\-－]\d+\s*\S+", text):
        return "table_caption"
    if re.match(r"^（?\(?\d+[-－]\d+）?\)?$", text):
        return "formula"
    if region == "references" and re.match(r"^\[?\d+\]?[\.、]?\s*\S+", text):
        return "reference_entry"
    if region == "symbols" and len(text) >= 3:
        return "symbols_body"
    if region == "appendix" and len(text) >= 10:
        return "appendix_body"
    if region == "author_bio" and len(text) >= 10:
        return "author_bio_body"
    if region == "acknowledgement" and len(text) >= 10:
        return "acknowledgement_body"
    if region == "abstract_zh" and len(text) >= 20:
        return "abstract_body_zh"
    if region == "abstract_en" and len(text) >= 20:
        return "abstract_body_en"
    if region == "cover":
        if "论文提交日期" in compact:
            return "cover_date"
        if re.search(r"论文题目|题目", text):
            return "cover_title_zh"
        if re.search(r"姓名|学号|学院|专业|研究方向|指导教师|辅助指导教师", text):
            return "cover_field_zh"
    if region == "body" and len(text) >= 1:
        return "body"
    return None


def next_region(text: str, current: str) -> str:
    compact = text.replace(" ", "")
    if compact == "摘要":
        return "abstract_zh"
    if compact.upper() == "ABSTRACT":
        return "abstract_en"
    if compact == "目录":
        return "toc"
    if re.match(r"^第[一二三四五六七八九十百\d]+章\s*\S+", text):
        return "body"
    if compact == "参考文献":
        return "references"
    if compact == "符号说明":
        return "symbols"
    if compact in {"附录", "附錄"}:
        return "appendix"
    if "作者简历" in compact or "攻读学位期间发表" in compact:
        return "author_bio"
    if compact in {"致谢", "致謝"}:
        return "acknowledgement"
    return current


def font_name(run, attr: str) -> str | None:
    rpr = run._element.rPr
    if rpr is None or rpr.rFonts is None:
        return run.font.name
    return rpr.rFonts.get(qn(f"w:{attr}")) or run.font.name


def _style_font_name(style, attr: str) -> str | None:
    while style is not None:
        rpr = style.element.rPr
        if rpr is not None and rpr.rFonts is not None:
            value = rpr.rFonts.get(qn(f"w:{attr}"))
            if value:
                return value
        if style.font and style.font.name:
            return style.font.name
        style = style.base_style
    return None


def _style_font_size(style) -> float | None:
    while style is not None:
        if style.font and style.font.size:
            return style.font.size.pt
        style = style.base_style
    return None


def _style_bold(style) -> bool | None:
    while style is not None:
        if style.font and style.font.bold is not None:
            return style.font.bold
        style = style.base_style
    return None


def common_run_format(paragraph) -> dict[str, object]:
    runs = [run for run in paragraph.runs if run.text.strip()]
    run = runs[0] if runs else None
    style = paragraph.style
    return {
        "eastAsia": font_name(run, "eastAsia") if run else _style_font_name(style, "eastAsia"),
        "ascii": font_name(run, "ascii") if run else _style_font_name(style, "ascii"),
        "size": run.font.size.pt if run and run.font.size else _style_font_size(style),
        "bold": run.bold if run and run.bold is not None else _style_bold(style),
    }


def run_effective_format(run, paragraph) -> dict[str, object]:
    style = paragraph.style
    return {
        "eastAsia": font_name(run, "eastAsia") or _style_font_name(style, "eastAsia"),
        "ascii": font_name(run, "ascii") or _style_font_name(style, "ascii"),
        "size": run.font.size.pt if run.font.size else _style_font_size(style),
        "bold": run.bold if run.bold is not None else _style_bold(style),
    }


def contains_cjk(text: str) -> bool:
    return bool(re.search(r"[\u3400-\u9fff]", text))


def contains_latin_or_digit(text: str) -> bool:
    return bool(re.search(r"[A-Za-z0-9]", text))


def run_format_summary(run, paragraph) -> str:
    fmt = run_effective_format(run, paragraph)
    return (
        f"文字片段={run.text[:40]!r}；中文字体={fmt['eastAsia'] or '继承/未直接设置'}；"
        f"西文字体={fmt['ascii'] or '继承/未直接设置'}；字号={fmt['size'] or '继承/未直接设置'}pt；"
        f"加粗={bold_label(fmt['bold'])}"
    )


def run_matches_rule(run, paragraph, rule: Rule) -> bool:
    text = run.text.strip()
    if not text:
        return True
    fmt = run_effective_format(run, paragraph)
    if not approx(fmt["size"], rule.size_pt):
        return False
    if rule.bold is not None and fmt["bold"] is not rule.bold:
        return False
    if contains_cjk(text) and fmt["eastAsia"] and fmt["eastAsia"] != rule.font_east_asia:
        return False
    if contains_latin_or_digit(text) and fmt["ascii"] and fmt["ascii"] != rule.font_ascii:
        return False
    return True


def collect_run_format_issues(paragraph, paragraph_index: int, rule: Rule, text: str) -> list[Issue]:
    issues: list[Issue] = []
    bad_runs = [
        run_format_summary(run, paragraph)
        for run in paragraph.runs
        if run.text.strip() and not run_matches_rule(run, paragraph, rule)
    ]
    if not bad_runs:
        return issues
    current = " | ".join(bad_runs[:6])
    if len(bad_runs) > 6:
        current += f" | ... +{len(bad_runs) - 6} more"
    issues.append(
        Issue(
            paragraph_index=paragraph_index,
            rule_key=f"{rule.key}_runs",
            text_type=f"{rule.label} run formatting",
            text_excerpt=text[:160],
            current=current,
            expected=rule.expected,
            message=f"文本类型：{rule.label} run formatting\n当前格式：{current}\n应改为：{rule.expected}",
            category="run-format",
            location=f"paragraph {paragraph_index}",
        )
    )
    return issues


def pt_value(value) -> float | None:
    return round(value.pt, 2) if value is not None else None


def value_label(value: object | None, unit: str = "") -> str:
    if value is None:
        return "继承/未直接设置"
    return f"{value}{unit}"


def bold_label(value: object | None) -> str:
    if value is True:
        return "是"
    if value is False:
        return "否"
    return "继承/未直接设置"


def line_spacing_pt(paragraph) -> float | None:
    pf = paragraph.paragraph_format
    style_pf = paragraph.style.paragraph_format if paragraph.style is not None else None
    line_spacing_value = pf.line_spacing if pf.line_spacing is not None else (style_pf.line_spacing if style_pf else None)
    return line_spacing_value.pt if hasattr(line_spacing_value, "pt") else None


def first_line_indent_cm(paragraph) -> float | None:
    pf = paragraph.paragraph_format
    style_pf = paragraph.style.paragraph_format if paragraph.style is not None else None
    first_indent_value = pf.first_line_indent if pf.first_line_indent is not None else (style_pf.first_line_indent if style_pf else None)
    return first_indent_value.cm if first_indent_value is not None else None


def first_line_indent_chars(paragraph) -> float | None:
    ppr = paragraph._p.pPr
    if ppr is not None:
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            value = ind.get(qn("w:firstLineChars"))
            if value is not None:
                return float(value) / 100
    if paragraph.style is not None:
        style_ppr = paragraph.style._element.pPr
        if style_ppr is not None:
            ind = style_ppr.find(qn("w:ind"))
            if ind is not None:
                value = ind.get(qn("w:firstLineChars"))
                if value is not None:
                    return float(value) / 100
    return None


def _set_first_line_indent_chars(paragraph, chars: float) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    for attr in ("w:firstLine", "w:firstLineChars"):
        qattr = qn(attr)
        if qattr in ind.attrib:
            del ind.attrib[qattr]
    ind.set(qn("w:firstLineChars"), str(int(round(chars * 100))))


def spacing_before_after_pt(paragraph) -> tuple[float | None, float | None]:
    pf = paragraph.paragraph_format
    style_pf = paragraph.style.paragraph_format if paragraph.style is not None else None
    before_value = pf.space_before if pf.space_before is not None else (style_pf.space_before if style_pf else None)
    after_value = pf.space_after if pf.space_after is not None else (style_pf.space_after if style_pf else None)
    return pt_value(before_value), pt_value(after_value)


def _spacing_element_from_paragraph(paragraph):
    ppr = paragraph._p.pPr
    if ppr is not None:
        spacing = ppr.find(qn("w:spacing"))
        if spacing is not None:
            return spacing
    if paragraph.style is not None:
        style_ppr = paragraph.style._element.pPr
        if style_ppr is not None:
            return style_ppr.find(qn("w:spacing"))
    return None


def spacing_before_after_lines(paragraph) -> tuple[float | None, float | None]:
    spacing = _spacing_element_from_paragraph(paragraph)
    if spacing is None:
        return None, None
    before = spacing.get(qn("w:beforeLines"))
    after = spacing.get(qn("w:afterLines"))
    return (
        float(before) / 100 if before is not None else None,
        float(after) / 100 if after is not None else None,
    )


def _set_line_spacing_units(paragraph, before: float | None, after: float | None) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    for attr in ("w:before", "w:after"):
        qattr = qn(attr)
        if qattr in spacing.attrib:
            del spacing.attrib[qattr]
    if before is not None:
        spacing.set(qn("w:beforeLines"), str(int(round(before * 100))))
    if after is not None:
        spacing.set(qn("w:afterLines"), str(int(round(after * 100))))


def paragraph_format_differences(paragraph, rule: Rule | None) -> list[str]:
    if rule is None:
        return []
    fmt = common_run_format(paragraph)
    style_pf = paragraph.style.paragraph_format if paragraph.style is not None else None
    differences: list[str] = []
    if not approx(fmt["size"], rule.size_pt):
        differences.append(f"字号：当前 {value_label(fmt['size'], 'pt')}，应为 {rule.size_pt}pt")
    if rule.bold is not None and fmt["bold"] is not rule.bold:
        differences.append(f"加粗：当前 {bold_label(fmt['bold'])}，应为 {bold_label(rule.bold)}")
    alignment = paragraph.alignment if paragraph.alignment is not None else (style_pf.alignment if style_pf else None)
    if rule.alignment is not None and alignment != rule.alignment:
        if not (rule.key in {"abstract_body_zh", "abstract_body_en", "acknowledgement_body"} and alignment in {WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.JUSTIFY}):
            differences.append(
                f"对齐方式：当前 {ALIGN_LABELS_ZH.get(alignment, '继承/未直接设置' if alignment is None else str(alignment))}，"
                f"应为 {ALIGN_LABELS_ZH.get(rule.alignment, str(rule.alignment))}"
            )
    actual_line_spacing = line_spacing_pt(paragraph)
    if rule.line_spacing_pt is not None and not approx(actual_line_spacing, rule.line_spacing_pt):
        differences.append(f"固定行距：当前 {value_label(round(actual_line_spacing, 2) if actual_line_spacing is not None else None, '磅')}，应为 {rule.line_spacing_pt}磅")
    actual_indent = first_line_indent_cm(paragraph)
    if rule.first_line_indent_cm is not None and not approx(actual_indent, rule.first_line_indent_cm, tolerance=0.15):
        differences.append(f"首行缩进：当前 {value_label(round(actual_indent, 2) if actual_indent is not None else None, 'cm')}，应为 {rule.first_line_indent_cm}cm")
    actual_indent_chars = first_line_indent_chars(paragraph)
    if rule.first_line_indent_chars is not None and not approx(actual_indent_chars, rule.first_line_indent_chars, tolerance=0.02):
        differences.append(f"首行缩进：当前 {value_label(actual_indent_chars, '字符')}，应为 {rule.first_line_indent_chars}字符")
    before, after = spacing_before_after_pt(paragraph)
    if rule.space_before_pt is not None and not approx(before, rule.space_before_pt):
        differences.append(f"段前：当前 {value_label(before, 'pt')}，应为 {rule.space_before_pt}pt")
    if rule.space_after_pt is not None and not approx(after, rule.space_after_pt):
        differences.append(f"段后：当前 {value_label(after, 'pt')}，应为 {rule.space_after_pt}pt")
    before_lines, after_lines = spacing_before_after_lines(paragraph)
    if rule.space_before_lines is not None and not approx(before_lines, rule.space_before_lines, tolerance=0.02):
        differences.append(f"段前：当前 {value_label(before_lines, '行')}，应为 {rule.space_before_lines}行")
    if rule.space_after_lines is not None and not approx(after_lines, rule.space_after_lines, tolerance=0.02):
        differences.append(f"段后：当前 {value_label(after_lines, '行')}，应为 {rule.space_after_lines}行")
    return differences


def current_format(paragraph, rule: Rule | None = None) -> str:
    fmt = common_run_format(paragraph)
    pf = paragraph.paragraph_format
    style_pf = paragraph.style.paragraph_format if paragraph.style is not None else None
    alignment = paragraph.alignment if paragraph.alignment is not None else (style_pf.alignment if style_pf else None)
    align = ALIGN_LABELS_ZH.get(alignment, "继承/未直接设置" if alignment is None else str(alignment))
    line_spacing_value = pf.line_spacing if pf.line_spacing is not None else (style_pf.line_spacing if style_pf else None)
    line_spacing = line_spacing_value.pt if hasattr(line_spacing_value, "pt") else None
    first_indent = first_line_indent_cm(paragraph)
    first_indent_chars = first_line_indent_chars(paragraph)
    before, after = spacing_before_after_pt(paragraph)
    before_lines, after_lines = spacing_before_after_lines(paragraph)
    parts = [
        f"段落样式：{paragraph.style.name if paragraph.style else '未知'}",
        f"中文字体：{fmt['eastAsia'] or '继承/未直接设置'}",
        f"西文字体：{fmt['ascii'] or '继承/未直接设置'}",
        f"字号：{value_label(fmt['size'], 'pt')}",
        f"加粗：{bold_label(fmt['bold'])}",
        f"对齐方式：{align}",
        f"固定行距：{value_label(round(line_spacing, 2) if line_spacing is not None else None, '磅')}",
        f"首行缩进：{value_label(first_indent_chars, '字符') if first_indent_chars is not None else value_label(round(first_indent, 2) if first_indent is not None else None, 'cm')}",
        f"段前：{value_label(before_lines, '行') if before_lines is not None else value_label(before, 'pt')}",
        f"段后：{value_label(after_lines, '行') if after_lines is not None else value_label(after, 'pt')}",
    ]
    differences = paragraph_format_differences(paragraph, rule)
    if differences:
        parts.insert(0, "疑似不符合项：" + "；".join(differences))
    return "; ".join(parts)


def approx(actual: float | None, expected: float | None, tolerance: float = 0.4) -> bool:
    if expected is None:
        return True
    if actual is None:
        return False
    return abs(actual - expected) <= tolerance


def paragraph_matches(paragraph, rule: Rule) -> bool:
    fmt = common_run_format(paragraph)
    pf = paragraph.paragraph_format
    style_pf = paragraph.style.paragraph_format if paragraph.style is not None else None
    size = fmt["size"]
    if not approx(size, rule.size_pt):
        return False
    if rule.bold is not None and fmt["bold"] is not rule.bold:
        return False
    alignment = paragraph.alignment if paragraph.alignment is not None else (style_pf.alignment if style_pf else None)
    if rule.alignment is not None and alignment != rule.alignment:
        if not (rule.key in {"abstract_body_zh", "abstract_body_en", "acknowledgement_body"} and alignment in {WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.JUSTIFY}):
            return False
    if rule.line_spacing_pt is not None:
        line_spacing_value = pf.line_spacing if pf.line_spacing is not None else (style_pf.line_spacing if style_pf else None)
        line_spacing = line_spacing_value.pt if hasattr(line_spacing_value, "pt") else None
        if not approx(line_spacing, rule.line_spacing_pt):
            return False
    if rule.first_line_indent_cm is not None:
        indent_value = pf.first_line_indent if pf.first_line_indent is not None else (style_pf.first_line_indent if style_pf else None)
        actual_indent = indent_value.cm if indent_value is not None else None
        if not approx(actual_indent, rule.first_line_indent_cm, tolerance=0.15):
            return False
    if rule.first_line_indent_chars is not None:
        actual_indent_chars = first_line_indent_chars(paragraph)
        if not approx(actual_indent_chars, rule.first_line_indent_chars, tolerance=0.02):
            return False
    before, after = spacing_before_after_pt(paragraph)
    if rule.space_before_pt is not None and not approx(before, rule.space_before_pt):
        return False
    if rule.space_after_pt is not None and not approx(after, rule.space_after_pt):
        return False
    before_lines, after_lines = spacing_before_after_lines(paragraph)
    if rule.space_before_lines is not None and not approx(before_lines, rule.space_before_lines, tolerance=0.02):
        return False
    if rule.space_after_lines is not None and not approx(after_lines, rule.space_after_lines, tolerance=0.02):
        return False
    return True


def set_run_font(run, rule: Rule) -> None:
    run.font.name = rule.font_ascii
    run._element.rPr.rFonts.set(qn("w:eastAsia"), rule.font_east_asia)
    run._element.rPr.rFonts.set(qn("w:ascii"), rule.font_ascii)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), rule.font_ascii)
    run.font.size = Pt(rule.size_pt)
    if rule.bold is not None:
        run.bold = rule.bold


def apply_rule(paragraph, rule: Rule) -> None:
    if rule.alignment is not None:
        paragraph.alignment = rule.alignment
    pf = paragraph.paragraph_format
    if rule.line_spacing_pt is not None:
        pf.line_spacing = Pt(rule.line_spacing_pt)
    if rule.first_line_indent_cm is not None:
        pf.first_line_indent = Cm(rule.first_line_indent_cm)
    if rule.first_line_indent_chars is not None:
        _set_first_line_indent_chars(paragraph, rule.first_line_indent_chars)
    if rule.space_before_pt is not None:
        pf.space_before = Pt(rule.space_before_pt)
    if rule.space_after_pt is not None:
        pf.space_after = Pt(rule.space_after_pt)
    if rule.space_before_lines is not None or rule.space_after_lines is not None:
        _set_line_spacing_units(paragraph, rule.space_before_lines, rule.space_after_lines)
    for run in paragraph.runs:
        set_run_font(run, rule)
    if not paragraph.runs and paragraph.text:
        set_run_font(paragraph.add_run(), rule)


def section_paragraph_ranges(document: Document) -> list[tuple[int, int]]:
    breaks: list[int] = []
    for idx, paragraph in enumerate(document.paragraphs):
        ppr = paragraph._p.pPr
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            breaks.append(idx)
    ranges: list[tuple[int, int]] = []
    start = 0
    for end in breaks:
        ranges.append((start, end))
        start = end + 1
    if start <= len(document.paragraphs) - 1:
        ranges.append((start, len(document.paragraphs) - 1))
    while len(ranges) < len(document.sections):
        ranges.append((len(document.paragraphs), len(document.paragraphs) - 1))
    return ranges[: len(document.sections)]


def section_text(document: Document, start: int, end: int) -> str:
    if start > end:
        return ""
    return "\n".join(paragraph_text(document.paragraphs[idx]) for idx in range(start, end + 1))


def first_body_paragraph_index(document: Document) -> int | None:
    texts = [paragraph_text(paragraph) for paragraph in document.paragraphs]
    regions = analyze_section_sequence(texts)["regions"]
    for idx, region in enumerate(regions):
        if region == "body":
            return idx
    return None


def ensure_body_starts_new_page_number_section(document: Document) -> bool:
    body_idx = first_body_paragraph_index(document)
    if body_idx is None or body_idx <= 0:
        return False
    ranges = section_paragraph_ranges(document)
    for section_idx, (start, end) in enumerate(ranges):
        if start <= body_idx <= end and start < body_idx:
            previous = document.paragraphs[body_idx - 1]
            ppr = previous._p.get_or_add_pPr()
            if ppr.find(qn("w:sectPr")) is not None:
                return False
            front_sect_pr = deepcopy(document.sections[section_idx]._sectPr)
            for child in list(front_sect_pr):
                if child.tag == qn("w:footerReference"):
                    front_sect_pr.remove(child)
            ppr.append(front_sect_pr)
            return True
    return False


def find_header_start_section(document: Document) -> tuple[int | None, str]:
    ranges = section_paragraph_ranges(document)
    declaration_pattern = re.compile(r"(声明|原创性声明|独创性声明|授权声明|版权声明|诚信声明)")
    for idx, (start, end) in enumerate(ranges):
        if declaration_pattern.search(section_text(document, start, end)):
            return idx, "declaration"
    fallback_pattern = re.compile(r"(摘\s*要|ABSTRACT|目\s*录|第一章|第\s*1\s*章)")
    for idx, (start, end) in enumerate(ranges):
        if fallback_pattern.search(section_text(document, start, end)):
            return idx, "fallback-front-matter"
    return None, "not-found"


def clear_header(header) -> None:
    for paragraph in header.paragraphs:
        paragraph.clear()


def apply_header(section, enabled: bool) -> None:
    section.different_first_page_header_footer = False
    section.header.is_linked_to_previous = False
    clear_header(section.header)
    clear_header(section.first_page_header)
    clear_header(section.even_page_header)
    if not enabled:
        return
    para = section.header.paragraphs[0] if section.header.paragraphs else section.header.add_paragraph()
    para.text = "成都信息工程大学学士学位论文"
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in para.runs:
        run.font.name = "Times New Roman"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        run.font.size = Pt(9)


def _set_section_page_number_format(section, fmt: str, start: int | None = None, clear_start: bool = False) -> None:
    if OxmlElement is None:
        return
    sect_pr = section._sectPr
    pg_num_type = sect_pr.find(qn("w:pgNumType"))
    if pg_num_type is None:
        pg_num_type = OxmlElement("w:pgNumType")
        cols = sect_pr.find(qn("w:cols"))
        if cols is not None:
            sect_pr.insert(list(sect_pr).index(cols), pg_num_type)
        else:
            sect_pr.append(pg_num_type)
    pg_num_type.set(qn("w:fmt"), fmt)
    if start is not None:
        pg_num_type.set(qn("w:start"), str(start))
    elif clear_start and qn("w:start") in pg_num_type.attrib:
        del pg_num_type.attrib[qn("w:start")]


def _roman_lower(number: int) -> str:
    values = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ]
    remaining = max(1, number)
    result: list[str] = []
    for value, token in values:
        while remaining >= value:
            result.append(token)
            remaining -= value
    return "".join(result)


def _add_simple_field(paragraph, instr: str, display: str) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run._r.append(begin)
    run = paragraph.add_run()
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = instr
    run._r.append(instr_text)
    run = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run._r.append(separate)
    paragraph.add_run(display)
    run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(end)


def _reset_footer_to_page_field(section, field_kind: str, display_number: int = 1) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    for child in list(footer._element):
        footer._element.remove(child)
    paragraph = footer.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if field_kind == "front":
        display = _roman_lower(display_number)
        instr = " PAGE  \\* roman  \\* MERGEFORMAT "
        _add_simple_field(paragraph, instr, display)
    else:
        paragraph.add_run("第")
        _add_simple_field(paragraph, " PAGE  \\* Arabic  \\* MERGEFORMAT ", str(display_number))
        paragraph.add_run("页 共")
        _add_simple_field(paragraph, " NUMPAGES  \\* Arabic  \\* MERGEFORMAT ", str(display_number))
        paragraph.add_run("页")
    for run in paragraph.runs:
        run.font.name = "Times New Roman"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        run.font.size = Pt(9)


def _rewrite_footer_page_field(section, field_kind: str, start: int = 1) -> None:
    _reset_footer_to_page_field(section, field_kind, display_number=start)
    return
    field_kind = field_kind.lower()
    visible = _roman_lower(start) if field_kind == "front" else str(start)
    instr = " PAGE  \\* roman  \\* MERGEFORMAT " if field_kind == "front" else " PAGE  \\* Arabic  \\* MERGEFORMAT "
    in_page_field = False
    seen_separator = False
    for node in section.footer._element.iter():
        if node.tag == qn("w:fldChar"):
            fld_type = node.get(qn("w:fldCharType"))
            if fld_type == "begin":
                in_page_field = True
                seen_separator = False
            elif fld_type == "separate" and in_page_field:
                seen_separator = True
            elif fld_type == "end":
                in_page_field = False
                seen_separator = False
        elif node.tag == qn("w:instrText") and in_page_field and node.text and "PAGE" in node.text.upper():
            node.text = instr
        elif node.tag == qn("w:t") and in_page_field and seen_separator:
            if node.text and re.fullmatch(r"\s*[\dIVXLCDMivxlcdm]+\s*", node.text):
                node.text = visible


def apply_page_number_formats(document: Document) -> None:
    texts = [paragraph_text(paragraph) for paragraph in document.paragraphs]
    regions = analyze_section_sequence(texts)["regions"]
    ranges = section_paragraph_ranges(document)
    front_started = False
    main_started = False
    decisions: list[tuple[int, str, int | None]] = []
    abstract_or_toc_sections: list[int] = []
    body_or_later_sections: list[int] = []
    for section_idx, _section in enumerate(document.sections):
        start, end = ranges[section_idx] if section_idx < len(ranges) else (0, -1)
        section_regions = set(regions[start : end + 1]) if start <= end else set()
        sec_text = section_text(document, start, end)
        sec_compact = compact_text(sec_text)
        looks_like_toc_section = (
            ("目录" in sec_compact)
            or ("目次" in sec_compact)
            or bool(re.search(r"\.{6,}", sec_text))
            or ("论文总页数" in sec_text)
        )
        if {"abstract_zh", "abstract_en", "toc"} & section_regions:
            abstract_or_toc_sections.append(section_idx)
        elif looks_like_toc_section:
            abstract_or_toc_sections.append(section_idx)
        if {"body", "references", "appendix", "acknowledgement"} & section_regions:
            body_or_later_sections.append(section_idx)

    first_front = min(abstract_or_toc_sections) if abstract_or_toc_sections else None
    last_known_front = max(abstract_or_toc_sections) if abstract_or_toc_sections else None
    main_candidates = [idx for idx in body_or_later_sections if last_known_front is None or idx > last_known_front]
    first_main = min(main_candidates) if main_candidates else None
    last_front = None
    if first_front is not None and first_main is not None and first_main >= first_front:
        last_front = first_main - 1
    elif first_front is not None:
        last_front = max(abstract_or_toc_sections)

    for section_idx, section in enumerate(document.sections):
        is_front = first_front is not None and last_front is not None and first_front <= section_idx <= last_front
        is_main = first_main is not None and section_idx >= first_main
        if is_front:
            _set_section_page_number_format(section, "lowerRoman", start=1 if not front_started else None, clear_start=front_started)
            decisions.append((section_idx, "front", None))
            front_started = True
        elif is_main:
            _set_section_page_number_format(section, "decimal", start=1 if not main_started else None, clear_start=main_started)
            decisions.append((section_idx, "main", 1 if not main_started else None))
            main_started = True
        else:
            footer = section.footer
            footer.is_linked_to_previous = False
            for child in list(footer._element):
                footer._element.remove(child)
    for section_idx, kind, display in decisions:
        if kind == "main":
            _rewrite_footer_page_field(document.sections[section_idx], kind, start=display or 1)
    for section_idx, kind, display in decisions:
        if kind == "front":
            _rewrite_footer_page_field(document.sections[section_idx], kind, start=display or 1)


def _xml_para_text(para_xml) -> str:
    return re.sub(r"\s+", " ", "".join(para_xml.itertext()) if para_xml is not None else "").strip()


def ensure_abstract_heading_split(document: Document) -> None:
    # Iterate over a snapshot because we may insert heading paragraphs before current ones.
    for paragraph in list(document.paragraphs):
        text = paragraph_text(paragraph)
        if not text:
            continue
        zh_match = re.match(r"^\s*摘要\s*[:：]\s*(.+)$", text)
        en_match = re.match(r"^\s*Abstract\s*[:：]\s*(.+)$", text, re.I)
        if zh_match:
            prev_text_compact = compact_text(_xml_para_text(paragraph._p.getprevious()))
            if prev_text_compact != "摘要":
                heading = paragraph.insert_paragraph_before("摘 要")
                apply_rule(heading, RULES["abstract_title_zh"])
            paragraph.text = zh_match.group(1).strip()
            apply_rule(paragraph, RULES["abstract_body_zh"])
            continue
        if en_match:
            prev_text_upper = compact_text(_xml_para_text(paragraph._p.getprevious())).upper()
            if prev_text_upper != "ABSTRACT":
                heading = paragraph.insert_paragraph_before("Abstract")
                apply_rule(heading, RULES["abstract_title_en"])
            paragraph.text = en_match.group(1).strip()
            apply_rule(paragraph, RULES["abstract_body_en"])


def _is_plain_blank_paragraph(paragraph) -> bool:
    if paragraph_text(paragraph):
        return False
    p = paragraph._p
    ppr = p.pPr
    if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
        return False
    for node in p.iter():
        if node.tag == qn("w:br"):
            return False
        if node.tag == qn("w:drawing"):
            return False
        if node.tag == qn("w:object"):
            return False
        if node.tag == qn("w:pict"):
            return False
    return True


def ensure_english_abstract_starts_new_page(document: Document) -> bool:
    changed = False
    texts = [paragraph_text(paragraph) for paragraph in document.paragraphs]
    title_overrides = detect_abstract_thesis_titles(texts)
    for idx, paragraph in enumerate(document.paragraphs):
        text = texts[idx]
        if not text:
            continue
        key = title_overrides.get(idx) or classify_paragraph(text, in_body=False, region="abstract_en")
        if key in {"thesis_title_en", "abstract_title_en"}:
            prev = document.paragraphs[idx - 1] if idx > 0 else None
            prev_has_sect_break = False
            if prev is not None and prev._p.pPr is not None and prev._p.pPr.find(qn("w:sectPr")) is not None:
                prev_has_sect_break = True
                # Keep section break for pagination, but normalize this placeholder paragraph.
                prev_pfmt = prev.paragraph_format
                if prev_pfmt.first_line_indent is not None:
                    prev_pfmt.first_line_indent = None
                    changed = True
                if prev_pfmt.line_spacing is not None:
                    prev_pfmt.line_spacing = None
                    changed = True
                if prev_pfmt.space_before is not None:
                    prev_pfmt.space_before = None
                    changed = True
                if prev_pfmt.space_after is not None:
                    prev_pfmt.space_after = None
                    changed = True
            if prev_has_sect_break:
                if paragraph.paragraph_format.page_break_before:
                    paragraph.paragraph_format.page_break_before = False
                    changed = True
            elif paragraph.paragraph_format.page_break_before is not True:
                paragraph.paragraph_format.page_break_before = True
                changed = True
            break
    return changed


def remove_redundant_front_blank_paragraphs(document: Document) -> bool:
    texts = [paragraph_text(paragraph) for paragraph in document.paragraphs]
    regions = analyze_section_sequence(texts)["regions"]
    front_regions = {"abstract_zh", "abstract_en", "toc"}
    removed = False
    for idx in range(len(document.paragraphs) - 2, 0, -1):
        if idx >= len(regions) or regions[idx] not in front_regions:
            continue
        if texts[idx]:
            continue
        paragraph = document.paragraphs[idx]
        if not _is_plain_blank_paragraph(paragraph):
            continue
        prev_region = regions[idx - 1] if idx - 1 < len(regions) else None
        next_region = regions[idx + 1] if idx + 1 < len(regions) else None
        if prev_region in front_regions and next_region in front_regions:
            paragraph._element.getparent().remove(paragraph._element)
            removed = True
    return removed


def remove_sectpr_only_blank_placeholders(document: Document) -> bool:
    removed = False
    for idx in range(len(document.paragraphs) - 1, 0, -1):
        paragraph = document.paragraphs[idx]
        if paragraph_text(paragraph):
            continue
        ppr = paragraph._p.pPr
        if ppr is None:
            continue
        sect_pr = ppr.find(qn("w:sectPr"))
        if sect_pr is None:
            continue
        # Keep true contentful blank lines untouched; only strip placeholder lines.
        has_content_signal = any(
            node.tag in {qn("w:br"), qn("w:drawing"), qn("w:object"), qn("w:pict")}
            for node in paragraph._p.iter()
        )
        if has_content_signal:
            continue
        # Move section properties to previous paragraph so sectioning remains valid.
        prev = document.paragraphs[idx - 1]
        prev_ppr = prev._p.get_or_add_pPr()
        old_prev_sect = prev_ppr.find(qn("w:sectPr"))
        if old_prev_sect is not None:
            prev_ppr.remove(old_prev_sect)
        prev_ppr.append(deepcopy(sect_pr))
        paragraph._element.getparent().remove(paragraph._element)
        removed = True
    return removed


def apply_page_setup(document: Document, allow_layout_fixes: bool) -> None:
    header_start, _header_reason = find_header_start_section(document)
    for section in document.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(3)
        if allow_layout_fixes:
            section.header_distance = Cm(1.5)
    if allow_layout_fixes:
        for idx, section in enumerate(document.sections):
            apply_header(section, enabled=header_start is not None and idx >= header_start)
        apply_page_number_formats(document)


def _section_format(section) -> str:
    return (
        f"page={round(section.page_width.cm, 2)}cm x {round(section.page_height.cm, 2)}cm; "
        f"margins top={round(section.top_margin.cm, 2)}cm, bottom={round(section.bottom_margin.cm, 2)}cm, "
        f"left={round(section.left_margin.cm, 2)}cm, right={round(section.right_margin.cm, 2)}cm"
    )


def _section_matches(section) -> bool:
    return (
        approx(section.page_width.cm, 21, 0.1)
        and approx(section.page_height.cm, 29.7, 0.1)
        and approx(section.top_margin.cm, 2.5, 0.1)
        and approx(section.bottom_margin.cm, 2.5, 0.1)
        and approx(section.left_margin.cm, 3, 0.1)
        and approx(section.right_margin.cm, 3, 0.1)
    )


def collect_section_issues(document: Document) -> list[Issue]:
    issues: list[Issue] = []
    header_start, header_reason = find_header_start_section(document)
    if header_reason == "not-found":
        rule = RULES["page_header"]
        issues.append(
            Issue(
                paragraph_index=-1,
                rule_key="page_header_start",
                text_type=rule.label,
                text_excerpt="document header start",
                current="未识别到声明页、摘要、目录或正文起点，无法可靠判断页眉起始页。",
                expected="页眉应从声明页开始到最后页，注明“成都信息工程大学学士学位论文”。",
                message="文本类型：page header\n当前情况：未识别到页眉起始页\n应改为：页眉从声明页开始到最后页。",
                category=rule.category,
                location="document header start",
            )
        )
    for sec_idx, section in enumerate(document.sections, start=1):
        if not _section_matches(section):
            current = _section_format(section)
            issues.append(
                Issue(
                    paragraph_index=-1,
                    rule_key="page_setup",
                    text_type="page setup",
                    text_excerpt=f"section {sec_idx}",
                    current=current,
                    expected="A4 21cm x 29.7cm; margins top/bottom 2.5cm, left/right 3cm.",
                    message=f"文本类型：page setup\n当前格式：{current}\n应改为：A4，页边距上/下2.5cm，左/右3cm。",
                    category="page",
                    location=f"section {sec_idx}",
                )
            )
        zero_idx = sec_idx - 1
        header_text = " ".join(paragraph_text(p) for p in section.header.paragraphs).strip()
        rule = RULES["page_header"]
        if header_start is not None and zero_idx < header_start and header_text:
            issues.append(
                Issue(
                    paragraph_index=-1,
                    rule_key="page_header_before_declaration",
                    text_type=rule.label,
                    text_excerpt=f"section {sec_idx} header",
                    current=header_text,
                    expected="页眉应从声明页开始；封面及声明页之前的部分不应添加该页眉。",
                    message=f"文本类型：{rule.label}\n当前格式：声明页之前已存在页眉：{header_text}\n应改为：封面及声明页之前不添加该页眉。",
                    category=rule.category,
                    location=f"section {sec_idx} header",
                )
            )
        if header_start is not None and zero_idx >= header_start and "成都信息工程大学学士学位论文" not in header_text:
            issues.append(
                Issue(
                    paragraph_index=-1,
                    rule_key="page_header",
                    text_type=rule.label,
                    text_excerpt=f"section {sec_idx} header",
                    current=header_text or "[empty header]",
                    expected=rule.expected,
                    message=f"文本类型：{rule.label}\n当前格式：{header_text or '[empty header]'}\n应改为：{rule.expected}",
                    category=rule.category,
                    location=f"section {sec_idx} header",
                )
            )
        issues.extend(collect_page_number_issues(section, sec_idx))
    return issues


def _section_page_number_format(section) -> tuple[str | None, str | None]:
    pg_num_type = section._sectPr.find(qn("w:pgNumType"))
    if pg_num_type is None:
        return None, None
    return pg_num_type.get(qn("w:fmt")), pg_num_type.get(qn("w:start"))


def _footer_field_info(section) -> tuple[str, str]:
    instr: list[str] = []
    texts: list[str] = []
    for node in section.footer._element.iter():
        if node.tag == qn("w:instrText") and node.text:
            instr.append(node.text)
        elif node.tag == qn("w:t") and node.text:
            texts.append(node.text)
    return "".join(texts).strip(), " ".join(instr).strip()


def validate_front_page_number(fmt: str | None, visible_text: str, field_text: str) -> str | None:
    if "PAGE" not in field_text.upper() and not visible_text:
        return None
    if fmt != "lowerRoman":
        return f"页码编号格式为 {fmt or '未设置'}，应为 lowerRoman 小写罗马数字。"
    compact = re.sub(r"\s+", "", visible_text)
    if compact and re.search(r"\d|[IVXLCDM]", compact):
        return f"页码显示为 {visible_text}，应为小写罗马数字 i, ii, iii, iv, v, ...。"
    return None


def validate_main_page_number(fmt: str | None, visible_text: str, field_text: str) -> str | None:
    if "PAGE" not in field_text.upper() and not visible_text:
        return None
    compact = re.sub(r"\s+", "", visible_text)
    field_upper = field_text.upper()
    has_main_pattern = bool(re.search(r"第.+页共.+页", compact)) or ("PAGE" in field_upper and "NUMPAGES" in field_upper)
    if fmt not in {None, "decimal"}:
        return f"页码编号格式为 {fmt}，正文至致谢应使用阿拉伯数字。"
    if not has_main_pattern:
        return f"页码显示为 {visible_text or field_text}，应为“第*页 共*页”。"
    return None


def collect_page_number_issues(section, sec_idx: int) -> list[Issue]:
    visible_text, field_text = _footer_field_info(section)
    if "PAGE" not in field_text.upper() and not visible_text:
        return []
    fmt, _start = _section_page_number_format(section)
    compact = re.sub(r"\s+", "", visible_text)
    is_main = bool(re.search(r"第.+页共.+页", compact)) or "NUMPAGES" in field_text.upper()
    if is_main:
        problem = validate_main_page_number(fmt, visible_text, field_text)
        rule_key = "main_page_number"
    else:
        problem = validate_front_page_number(fmt, visible_text, field_text)
        rule_key = "front_page_number"
    if problem is None:
        return []
    rule = RULES[rule_key]
    return [
        Issue(
            paragraph_index=-1,
            rule_key=rule_key,
            text_type=rule.label,
            text_excerpt=f"section {sec_idx} footer",
            current=problem,
            expected=rule.expected,
            message=f"文本类型：{rule.label}\n当前格式：{problem}\n应改为：{rule.expected}",
            category=rule.category,
            location=f"section {sec_idx} footer",
        )
    ]


def collect_abstract_title_issues(document: Document) -> list[Issue]:
    issues: list[Issue] = []
    texts = [paragraph_text(paragraph) for paragraph in document.paragraphs]
    has_zh_title = any(compact_text(text) == "摘要" or text.strip() == "摘 要" for text in texts)
    has_en_title = any(compact_text(text).upper() == "ABSTRACT" for text in texts)
    for idx, text in enumerate(texts):
        stripped = text.strip()
        if not stripped:
            continue
        if not has_zh_title and re.match(r"^摘要\s*[:：]", stripped):
            rule = RULES["abstract_title_zh"]
            issues.append(
                Issue(
                    paragraph_index=idx,
                    rule_key="abstract_title_zh_missing",
                    text_type=rule.label,
                    text_excerpt=stripped[:160],
                    current="摘要标题与摘要正文合并在同一段，且未检测到独立“摘 要”标题。",
                    expected=rule.expected,
                    message=f"文本类型：{rule.label}\n当前格式：摘要标题与正文合并\n应改为：{rule.expected}",
                    category=rule.category,
                    location=f"paragraph {idx}",
                )
            )
            break
    for idx, text in enumerate(texts):
        stripped = text.strip()
        if not stripped:
            continue
        if not has_en_title and re.match(r"^Abstract\s*[:：]", stripped, re.I):
            rule = RULES["abstract_title_en"]
            issues.append(
                Issue(
                    paragraph_index=idx,
                    rule_key="abstract_title_en_missing",
                    text_type=rule.label,
                    text_excerpt=stripped[:160],
                    current="英文摘要标题与摘要正文合并在同一段，且未检测到独立“ABSTRACT”标题。",
                    expected=rule.expected,
                    message=f"文本类型：{rule.label}\n当前格式：英文摘要标题与正文合并\n应改为：{rule.expected}",
                    category=rule.category,
                    location=f"paragraph {idx}",
                )
            )
            break
    return issues


def _issue_global(rule_key: str, text_type: str, current: str, expected: str, category: str, location: str, excerpt: str = "") -> Issue:
    return Issue(
        paragraph_index=-1,
        rule_key=rule_key,
        text_type=text_type,
        text_excerpt=excerpt or location,
        current=current,
        expected=expected,
        message=f"文本类型：{text_type}\n当前位置：{location}\n当前情况：{current}\n应改为：{expected}",
        category=category,
        location=location,
    )


def _document_xml_text(document: Document) -> str:
    return ET.tostring(document.element.body, encoding="unicode")


def paragraph_has_right_tab(paragraph) -> bool:
    ppr = paragraph._p.pPr
    if ppr is None:
        return False
    tabs = ppr.find(qn("w:tabs"))
    if tabs is None:
        return False
    for tab in tabs.findall(qn("w:tab")):
        if tab.get(qn("w:val")) == "right":
            return True
    return False


def collect_toc_issues(document: Document) -> list[Issue]:
    issues: list[Issue] = []
    xml = _document_xml_text(document)
    has_toc_field = "TOC" in xml and "instrText" in xml
    has_toc_title = any(paragraph_text(p).replace(" ", "") == "目录" for p in document.paragraphs)
    if has_toc_title and not has_toc_field:
        issues.append(
            _issue_global(
                "toc_field",
                "TOC field",
                "目录标题存在，但未检测到 Word TOC 域代码。",
                "目录应由 Word/WPS 目录域或等效可更新结构生成，便于页码刷新。",
                "toc",
                "document TOC",
            )
        )
    in_toc = False
    for idx, paragraph in enumerate(document.paragraphs):
        text = paragraph_text(paragraph)
        compact = text.replace(" ", "")
        if compact == "目录":
            in_toc = True
            continue
        if in_toc and re.match(r"^第[一二三四五六七八九十百\d]+章\s*\S+", text):
            in_toc = False
        if not in_toc or not text:
            continue
        style_name = (paragraph.style.name if paragraph.style else "").lower()
        looks_toc_entry = style_name.startswith("toc") or re.search(r"(\.{2,}|…+|\t)\s*\d+$", text) or re.match(r"^(\d+\.\d+|第[一二三四五六七八九十百\d]+章)", text)
        if looks_toc_entry and not paragraph_has_right_tab(paragraph):
            issues.append(
                Issue(
                    paragraph_index=idx,
                    rule_key="toc_tab_stop",
                    text_type="TOC page-number right tab",
                    text_excerpt=text[:160],
                    current="未检测到右对齐制表位。",
                    expected="目录页码应通过右对齐制表位或等效目录域右对齐。",
                    message=f"文本类型：TOC page-number right tab\n当前格式：未检测到右对齐制表位\n应改为：目录页码右对齐。",
                    category="toc",
                    location=f"paragraph {idx}",
                )
            )
    return issues


def iter_body_blocks(document: Document):
    para_by_el = {p._p: ("p", i, p) for i, p in enumerate(document.paragraphs)}
    table_by_el = {t._tbl: ("tbl", i, t) for i, t in enumerate(document.tables)}
    for child in document.element.body:
        if child in para_by_el:
            yield para_by_el[child]
        elif child in table_by_el:
            yield table_by_el[child]


def table_has_vertical_borders(table) -> bool:
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        return False
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        return False
    for name in ("insideV", "left", "right"):
        node = borders.find(qn(f"w:{name}"))
        if node is not None and node.get(qn("w:val")) not in {None, "nil", "none"}:
            return True
    return False


def collect_table_issues(document: Document) -> list[Issue]:
    issues: list[Issue] = []
    blocks = list(iter_body_blocks(document))
    for pos, (kind, table_idx, table) in enumerate(blocks):
        if kind != "tbl":
            continue
        prev_text = ""
        next_text = ""
        if pos > 0 and blocks[pos - 1][0] == "p":
            prev_text = paragraph_text(blocks[pos - 1][2])
        if pos + 1 < len(blocks) and blocks[pos + 1][0] == "p":
            next_text = paragraph_text(blocks[pos + 1][2])
        if not re.match(r"^表\s*\d+[-－]\d+\s*\S+", prev_text):
            issues.append(
                _issue_global(
                    "table_caption_position",
                    "table caption position",
                    f"表 {table_idx + 1} 上方未检测到“表X-X 表名”格式表题。",
                    "表序、表名、表注应置于表上方。",
                    "figure-table",
                    f"table {table_idx + 1}",
                    prev_text or next_text,
                )
            )
        if re.match(r"^表\s*\d+[-－]\d+\s*\S+", next_text):
            issues.append(
                _issue_global(
                    "table_caption_position",
                    "table caption position",
                    f"表 {table_idx + 1} 下方检测到表题，疑似位置错误。",
                    "表题应置于表上方，不应置于表下方。",
                    "figure-table",
                    f"table {table_idx + 1}",
                    next_text,
                )
            )
        style_name = table.style.name if table.style is not None else ""
        if "grid" in style_name.lower() or table_has_vertical_borders(table):
            issues.append(
                _issue_global(
                    "table_three_line",
                    "three-line table",
                    f"表 {table_idx + 1} 样式为 {style_name or 'unknown'}，可能包含全网格/竖线。",
                    "表的编排宜采用三线表；通常只保留顶线、表头下线和底线。",
                    "figure-table",
                    f"table {table_idx + 1}",
                )
            )
        if table_idx > 0 and prev_text.endswith("（续）"):
            first_row = " ".join(cell.text.strip() for cell in table.rows[0].cells) if table.rows else ""
            if not first_row:
                issues.append(
                    _issue_global(
                        "continued_table_header",
                        "continued table header",
                        f"续表 {table_idx + 1} 未检测到重复表头。",
                        "续表均应重复表头。",
                        "figure-table",
                        f"table {table_idx + 1}",
                    )
                )
    return issues


def collect_reference_issues(document: Document) -> list[Issue]:
    issues: list[Issue] = []
    regions: dict[str, int] = {}
    for idx, paragraph in enumerate(document.paragraphs):
        compact = paragraph_text(paragraph).replace(" ", "")
        if "结论" in compact and "参考文献" not in compact and "目录" not in compact:
            regions.setdefault("conclusion", idx)
        if compact == "参考文献":
            regions.setdefault("references", idx)
        if compact in {"附录", "附錄"}:
            regions.setdefault("appendix", idx)
    ref_idx = regions.get("references")
    if ref_idx is not None:
        conclusion_idx = regions.get("conclusion")
        appendix_idx = regions.get("appendix")
        if conclusion_idx is not None and ref_idx < conclusion_idx:
            issues.append(_issue_global("reference_position", "references position", "参考文献位于结论之前。", "参考文献通常应位于结论之后、附录之前。", "references", "references"))
        if appendix_idx is not None and ref_idx > appendix_idx:
            issues.append(_issue_global("reference_position", "references position", "参考文献位于附录之后。", "参考文献通常应位于结论之后、附录之前。", "references", "references"))
    else:
        return issues
    end_idx = regions.get("appendix", len(document.paragraphs))
    for idx in range(ref_idx + 1, end_idx):
        paragraph = document.paragraphs[idx]
        text = paragraph_text(paragraph)
        if not text:
            continue
        if not re.match(r"^\[?\d+\]?[\.、]?\s*\S+", text):
            continue
        pf = paragraph.paragraph_format
        first_indent = pf.first_line_indent.cm if pf.first_line_indent is not None else None
        if first_indent is None or first_indent >= 0:
            issues.append(
                Issue(
                    paragraph_index=idx,
                    rule_key="reference_hanging_indent",
                    text_type="reference hanging indent",
                    text_excerpt=text[:160],
                    current=f"first-line indent={first_indent if first_indent is not None else 'unset'}",
                    expected="参考文献续行缩进两个字符左对齐；通常应使用悬挂缩进。",
                    message="文本类型：reference hanging indent\n当前格式：未检测到悬挂缩进\n应改为：参考文献续行缩进两个字符左对齐。",
                    category="references",
                    location=f"paragraph {idx}",
                )
            )
    return issues


def collect_issues(document: Document) -> list[Issue]:
    issues: list[Issue] = []
    texts = [paragraph_text(paragraph) for paragraph in document.paragraphs]
    structure = analyze_section_sequence(texts)
    regions = structure["regions"]
    title_overrides = detect_abstract_thesis_titles(texts)
    for idx, paragraph in enumerate(document.paragraphs):
        text = texts[idx]
        if not text:
            continue
        key = classify_checked_paragraph(paragraph, idx, text, regions[idx], title_overrides)
        if key is None:
            continue
        rule = RULES[key]
        if not paragraph_matches(paragraph, rule):
            current = current_format(paragraph, rule)
            issues.append(
                Issue(
                    paragraph_index=idx,
                    rule_key=rule.key,
                    text_type=rule.label,
                    text_excerpt=text[:160],
                    current=current,
                    expected=rule.expected,
                    message=f"文本类型：{rule.label}\n当前格式：{current}\n应改为：{rule.expected}",
                    category=rule.category,
                    location=f"paragraph {idx}",
                )
            )
        issues.extend(collect_run_format_issues(paragraph, idx, rule, text))
    issues.extend(collect_section_issues(document))
    issues.extend(collect_abstract_title_issues(document))
    issues.extend(collect_toc_issues(document))
    issues.extend(collect_table_issues(document))
    issues.extend(collect_reference_issues(document))
    return issues


def apply_supported_rules(document: Document, allow_layout_fixes: bool) -> None:
    ensure_abstract_heading_split(document)
    apply_page_setup(document, allow_layout_fixes=allow_layout_fixes)
    texts = [paragraph_text(paragraph) for paragraph in document.paragraphs]
    structure = analyze_section_sequence(texts)
    regions = structure["regions"]
    title_overrides = detect_abstract_thesis_titles(texts)
    for idx, paragraph in enumerate(document.paragraphs):
        text = texts[idx]
        if not text:
            continue
        key = classify_checked_paragraph(paragraph, idx, text, regions[idx], title_overrides)
        if key is not None:
            apply_rule(paragraph, RULES[key])


def qname(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def ensure_comments_parts(workdir: Path) -> ET.ElementTree:
    comments_path = workdir / "word" / "comments.xml"
    if comments_path.exists():
        return ET.parse(comments_path)
    root = ET.Element(qname(W_NS, "comments"))
    tree = ET.ElementTree(root)
    comments_path.parent.mkdir(parents=True, exist_ok=True)
    tree.write(comments_path, encoding="utf-8", xml_declaration=True)
    return tree


def ensure_content_type(workdir: Path) -> None:
    path = workdir / "[Content_Types].xml"
    tree = ET.parse(path)
    root = tree.getroot()
    exists = any(child.attrib.get("PartName") == "/word/comments.xml" for child in root)
    if not exists:
        override = ET.SubElement(root, qname(CONTENT_NS, "Override"))
        override.set("PartName", "/word/comments.xml")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    tree.write(path, encoding="utf-8", xml_declaration=True)


def ensure_comments_relationship(workdir: Path) -> None:
    rels = workdir / "word" / "_rels" / "document.xml.rels"
    tree = ET.parse(rels)
    root = tree.getroot()
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    if any(child.attrib.get("Type") == rel_type for child in root):
        tree.write(rels, encoding="utf-8", xml_declaration=True)
        return
    used = {child.attrib.get("Id", "") for child in root}
    n = 1
    while f"rId{n}" in used:
        n += 1
    rel = ET.SubElement(root, qname(PKG_REL_NS, "Relationship"))
    rel.set("Id", f"rId{n}")
    rel.set("Type", rel_type)
    rel.set("Target", "comments.xml")
    tree.write(rels, encoding="utf-8", xml_declaration=True)


def add_comment_node(comments_root: ET.Element, comment_id: int, text: str) -> None:
    comment = ET.SubElement(comments_root, qname(W_NS, "comment"))
    comment.set(qname(W_NS, "id"), str(comment_id))
    comment.set(qname(W_NS, "author"), "Codex CUIT thesis format skill")
    comment.set(qname(W_NS, "date"), datetime.now(timezone.utc).isoformat())
    p = ET.SubElement(comment, qname(W_NS, "p"))
    for line_no, line in enumerate(text.splitlines()):
        if line_no:
            ET.SubElement(p, qname(W_NS, "br"))
        r = ET.SubElement(p, qname(W_NS, "r"))
        t = ET.SubElement(r, qname(W_NS, "t"))
        t.text = line


def add_comment_marker(paragraph: ET.Element, comment_id: int) -> None:
    start = ET.Element(qname(W_NS, "commentRangeStart"))
    start.set(qname(W_NS, "id"), str(comment_id))
    end = ET.Element(qname(W_NS, "commentRangeEnd"))
    end.set(qname(W_NS, "id"), str(comment_id))
    ref_run = ET.Element(qname(W_NS, "r"))
    ref = ET.SubElement(ref_run, qname(W_NS, "commentReference"))
    ref.set(qname(W_NS, "id"), str(comment_id))
    paragraph.insert(0, start)
    paragraph.append(end)
    paragraph.append(ref_run)


def _comment_title(issue: Issue) -> str:
    labels = {
        "keywords": "关键词段落",
        "keywords_zh": "中文关键词段落",
        "keywords_en": "英文关键词段落",
        "keywords_runs": "关键词段落",
        "keywords_zh_runs": "中文关键词段落",
        "keywords_en_runs": "英文关键词段落",
        "thesis_title_zh": "中文论文题目",
        "thesis_title_en": "英文论文题目",
        "abstract_title_zh": "中文摘要标题",
        "abstract_body_zh": "中文摘要正文",
        "abstract_title_en": "英文摘要标题",
        "abstract_body_en": "英文摘要正文",
        "chapter": "章标题",
        "heading2": "二级标题",
        "heading3": "三级标题",
        "body": "正文段落",
        "figure_caption": "图题",
        "table_caption": "表题",
        "reference_title": "参考文献标题",
        "reference_entry": "参考文献条目",
        "appendix_title": "附录标题",
        "appendix_body": "附录正文",
    }
    key = issue.rule_key.removesuffix("_runs")
    return labels.get(issue.rule_key) or labels.get(key) or issue.text_type


def _short_expected_items(issue: Issue) -> list[str]:
    key = issue.rule_key.removesuffix("_runs")
    if key in {"keywords", "keywords_zh", "keywords_en"}:
        return [
            "中文用宋体小四，英文/数字用 Times New Roman 小四",
            "关键词标签加粗",
            "关键词之间用分号分隔",
            "段落行距为固定 20 磅",
        ]
    if key == "thesis_title_zh":
        return [
            "中文论文题目使用宋体三号 16pt，加粗，居中",
            "行距为固定 20 磅",
        ]
    if key == "thesis_title_en":
        return [
            "英文论文题目使用 Times New Roman 三号 16pt，加粗，居中",
            "行距为固定 20 磅",
        ]
    body_items = {
        "body": [
            "中文用宋体小四 12pt，英文/数字用 Times New Roman 12pt",
            "左对齐",
            "首行缩进 2 个汉字符，段前 0 磅，段后 0 磅",
            "行距为固定 20 磅",
        ],
        "abstract_body_zh": [
            "中文用宋体小四 12pt",
            "行距为固定 20 磅",
        ],
        "abstract_body_en": [
            "英文用 Times New Roman 小四 12pt",
            "行距为固定 20 磅",
        ],
        "appendix_body": [
            "中文用宋体小四 12pt，英文/数字用 Times New Roman 12pt",
            "两端对齐",
            "首行缩进 2 个汉字符",
            "行距为固定 20 磅",
        ],
        "acknowledgement_body": [
            "中文用宋体小四 12pt",
            "行距为固定 20 磅",
        ],
    }
    if key in body_items:
        return body_items[key]
    heading_items = {
        "chapter": [
            "宋体三号 16pt，加粗，居中",
            "行距为固定 20 磅，段前 0.5 行，段后 0.5 行",
            "章序号与章名之间空一格",
        ],
        "heading2": [
            "宋体四号 14pt，加粗，左对齐",
            "行距为固定 20 磅，段前 0.5 行，段后 0.5 行",
            "序号与题名之间空一格",
        ],
        "heading3": [
            "宋体小四 12pt，加粗，左对齐",
            "行距为固定 20 磅，段前 0.5 行，段后 0.5 行",
            "序号与题名之间空一格",
        ],
        "abstract_title_zh": [
            "中文摘要标题使用宋体三号 16pt，加粗，居中",
            "行距为固定 20 磅",
        ],
        "abstract_title_en": [
            "英文摘要标题使用 Times New Roman 三号 16pt，加粗，居中",
            "行距为固定 20 磅",
        ],
        "toc_title": [
            "目录标题使用宋体三号 16pt，加粗，居中",
            "行距为固定 20 磅",
        ],
    }
    if key in heading_items:
        return heading_items[key]
    if key in {"figure_caption", "table_caption"}:
        return [
            "图表题使用宋体五号",
            "居中，行距为固定 20 磅",
            "编号和标题之间保留一个空格",
        ]
    if key.startswith("reference"):
        return [
            "“参考文献”标题使用宋体三号 16pt，加粗，按章标题格式排版",
            "条目中文用宋体小四 12pt，英文用 Times New Roman 12pt",
            "条目行距为固定 20 磅，续行缩进 2 个汉字符左对齐",
            "不在自动检查中判断 GB/T 7714 著录内容细节",
        ]
    return [issue.expected]


def comment_message_for_issue(issue: Issue) -> str:
    title = _comment_title(issue)
    lines = [f"{title}格式不符合要求。", "", "需调整："]
    for index, item in enumerate(_short_expected_items(issue), start=1):
        lines.append(f"{index}. {item}")
    return "\n".join(lines)


def issues_for_comments(issues: Iterable[Issue]) -> list[Issue]:
    selected: list[Issue] = []
    by_paragraph_and_rule: set[tuple[int, str]] = set()
    deferred_run_issues: list[Issue] = []

    for issue in issues:
        base_rule = issue.rule_key.removesuffix("_runs")
        key = (issue.paragraph_index, base_rule)
        if issue.rule_key.endswith("_runs"):
            deferred_run_issues.append(issue)
            continue
        selected.append(issue)
        by_paragraph_and_rule.add(key)

    for issue in deferred_run_issues:
        base_rule = issue.rule_key.removesuffix("_runs")
        key = (issue.paragraph_index, base_rule)
        if key not in by_paragraph_and_rule:
            selected.append(issue)
            by_paragraph_and_rule.add(key)

    return selected


def write_zip_from_dir(source_dir: Path, out_path: Path) -> None:
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in source_dir.rglob("*"):
            if file.is_file():
                zf.write(file, file.relative_to(source_dir).as_posix())


def create_annotated_docx(source: Path, target: Path, issues: Iterable[Issue]) -> None:
    issues = issues_for_comments(issues)
    if not issues:
        shutil.copyfile(source, target)
        return
    if target.exists():
        target.unlink()
    with TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        with zipfile.ZipFile(source, "r") as zf:
            zf.extractall(tmpdir)
        comments_tree = ensure_comments_parts(tmpdir)
        comments_root = comments_tree.getroot()
        existing_ids = [
            int(node.attrib.get(qname(W_NS, "id"), "0"))
            for node in comments_root.findall(qname(W_NS, "comment"))
            if node.attrib.get(qname(W_NS, "id"), "0").isdigit()
        ]
        next_id = (max(existing_ids) + 1) if existing_ids else 0

        document_xml = tmpdir / "word" / "document.xml"
        doc_tree = ET.parse(document_xml)
        doc_root = doc_tree.getroot()
        paragraphs = doc_root.findall(f".//{qname(W_NS, 'body')}/{qname(W_NS, 'p')}")
        for issue in issues:
            if issue.paragraph_index < 0:
                continue
            if issue.paragraph_index >= len(paragraphs):
                continue
            comment_id = next_id
            next_id += 1
            add_comment_node(comments_root, comment_id, comment_message_for_issue(issue))
            add_comment_marker(paragraphs[issue.paragraph_index], comment_id)

        comments_tree.write(tmpdir / "word" / "comments.xml", encoding="utf-8", xml_declaration=True)
        doc_tree.write(document_xml, encoding="utf-8", xml_declaration=True)
        ensure_content_type(tmpdir)
        ensure_comments_relationship(tmpdir)
        write_zip_from_dir(tmpdir, target)


def try_com_save(path: Path, renderer: str) -> str:
    if renderer == "ooxml":
        return "ooxml"
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        if renderer in {"office", "wps"}:
            raise RuntimeError(f"{renderer} renderer requested but win32com is unavailable: {exc}") from exc
        return "ooxml"

    attempts = []
    if renderer in {"auto", "office"}:
        attempts.append(("office", "Word.Application"))
    if renderer in {"auto", "wps"}:
        attempts.extend([("wps", "KWPS.Application"), ("wps", "WPS.Application")])

    last_error: Exception | None = None
    for name, prog_id in attempts:
        app = None
        doc = None
        try:
            app = win32com.client.DispatchEx(prog_id)
            app.Visible = False
            doc = app.Documents.Open(str(path.resolve()))
            doc.SaveAs2(str(path.resolve()), FileFormat=16)
            doc.Close(False)
            app.Quit()
            return name
        except Exception as exc:
            last_error = exc
            try:
                if doc is not None:
                    doc.Close(False)
                if app is not None:
                    app.Quit()
            except Exception:
                pass
            if renderer in {"office", "wps"}:
                raise RuntimeError(f"{renderer} renderer failed via {prog_id}: {exc}") from exc
    if renderer == "auto":
        return "ooxml"
    raise RuntimeError(f"Renderer failed: {last_error}")


def probe_com_renderer(renderer: str) -> str:
    if renderer == "ooxml":
        return "ooxml"
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        if renderer in {"office", "wps"}:
            raise RuntimeError(f"{renderer} renderer requested but win32com is unavailable: {exc}") from exc
        return "ooxml"
    attempts = []
    if renderer in {"auto", "office"}:
        attempts.append(("office", "Word.Application"))
    if renderer in {"auto", "wps"}:
        attempts.extend([("wps", "KWPS.Application"), ("wps", "WPS.Application")])
    last_error: Exception | None = None
    for name, prog_id in attempts:
        app = None
        try:
            app = win32com.client.DispatchEx(prog_id)
            app.Visible = False
            app.Quit()
            return name
        except Exception as exc:
            last_error = exc
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
            if renderer in {"office", "wps"}:
                raise RuntimeError(f"{renderer} renderer failed via {prog_id}: {exc}") from exc
    if renderer == "auto":
        return "ooxml"
    raise RuntimeError(f"Renderer probe failed: {last_error}")


def _dispatch_app(renderer_used: str):
    import win32com.client  # type: ignore

    prog_ids = {
        "office": ["Word.Application"],
        "wps": ["KWPS.Application", "WPS.Application"],
    }.get(renderer_used, [])
    last_error: Exception | None = None
    for prog_id in prog_ids:
        try:
            app = win32com.client.DispatchEx(prog_id)
            app.Visible = False
            return app
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"Cannot start {renderer_used} renderer: {last_error}")


def _collect_pages_with_com(docx_path: Path, issues: list[Issue], renderer_used: str) -> dict[int, int]:
    if renderer_used not in {"office", "wps"}:
        return {}
    app = None
    doc = None
    pages: dict[int, int] = {}
    try:
        app = _dispatch_app(renderer_used)
        doc = app.Documents.Open(str(docx_path.resolve()))
        # Word constant wdActiveEndPageNumber = 3.
        for issue in issues:
            if issue.paragraph_index < 0:
                continue
            paragraph_no = issue.paragraph_index + 1
            if paragraph_no <= doc.Paragraphs.Count:
                pages[issue.paragraph_index] = int(doc.Paragraphs(paragraph_no).Range.Information(3))
        doc.Close(False)
        app.Quit()
    except Exception:
        try:
            if doc is not None:
                doc.Close(False)
            if app is not None:
                app.Quit()
        except Exception:
            pass
        return {}
    return pages


def _export_pdf_with_com(docx_path: Path, pdf_path: Path, renderer_used: str) -> None:
    app = None
    doc = None
    try:
        app = _dispatch_app(renderer_used)
        doc = app.Documents.Open(str(docx_path.resolve()))
        # Prefer ExportAsFixedFormat when available; fall back to SaveAs2 PDF.
        try:
            doc.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
        except Exception:
            doc.SaveAs2(str(pdf_path.resolve()), FileFormat=17)
        doc.Close(False)
        app.Quit()
    except Exception:
        try:
            if doc is not None:
                doc.Close(False)
            if app is not None:
                app.Quit()
        except Exception:
            pass
        raise


def _render_pdf_pages(pdf_path: Path, pages: set[int], image_dir: Path, prefix: str) -> dict[int, Path]:
    import pymupdf  # type: ignore

    image_dir.mkdir(parents=True, exist_ok=True)
    rendered: dict[int, Path] = {}
    with pymupdf.open(pdf_path) as doc:
        for page_no in sorted(pages):
            if page_no < 1 or page_no > len(doc):
                continue
            page = doc[page_no - 1]
            pix = page.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5), alpha=False)
            out = image_dir / f"{prefix}_page_{page_no}.png"
            pix.save(out)
            rendered[page_no] = out
    return rendered


def _find_codex_render_docx() -> Path | None:
    candidates = sorted(
        (Path.home() / ".codex" / "plugins" / "cache" / "openai-primary-runtime" / "documents").glob(
            "*/skills/documents/render_docx.py"
        ),
        reverse=True,
    )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _render_docx_pages_with_artifact_tool(docx_path: Path, out_dir: Path) -> list[Path]:
    render_script = _find_codex_render_docx()
    if render_script is None:
        raise RuntimeError("Codex documents render_docx.py was not found.")
    out_dir.mkdir(parents=True, exist_ok=True)
    command = [
        sys.executable,
        str(render_script),
        str(docx_path.resolve()),
        "--output_dir",
        str(out_dir.resolve()),
        "--renderer",
        "artifact-tool",
    ]
    proc = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if proc.returncode != 0:
        raise RuntimeError(
            "artifact-tool DOCX render failed. "
            f"stdout={proc.stdout.strip()} stderr={proc.stderr.strip()}"
        )
    pages = sorted(
        out_dir.glob("page-*.png"),
        key=lambda p: int(re.search(r"page-(\d+)\.png$", p.name).group(1))
        if re.search(r"page-(\d+)\.png$", p.name)
        else 0,
    )
    if not pages:
        raise RuntimeError(f"artifact-tool did not produce page PNGs in {out_dir}")
    return pages


def _png_dimensions(path: Path) -> tuple[int | None, int | None]:
    try:
        with path.open("rb") as fh:
            header = fh.read(24)
        if header[:8] != b"\x89PNG\r\n\x1a\n":
            return None, None
        return int.from_bytes(header[16:20], "big"), int.from_bytes(header[20:24], "big")
    except Exception:
        return None, None


def _png_visual_stats(path: Path | None) -> dict[str, float | int | None]:
    if path is None or not path.exists():
        return {"ink_ratio": None, "dark_ratio": None, "width": None, "height": None}
    try:
        from PIL import Image
    except Exception:
        return {"ink_ratio": None, "dark_ratio": None, "width": None, "height": None}
    try:
        with Image.open(path) as image:
            gray = image.convert("L")
            width, height = gray.size
            histogram = gray.resize((max(1, width // 4), max(1, height // 4))).histogram()
            total = sum(histogram)
            if total == 0:
                return {"ink_ratio": None, "dark_ratio": None, "width": width, "height": height}
            ink = sum(histogram[:245])
            dark = sum(histogram[:80])
            return {
                "ink_ratio": round(ink / total, 4),
                "dark_ratio": round(dark / total, 4),
                "width": width,
                "height": height,
            }
    except Exception:
        return {"ink_ratio": None, "dark_ratio": None, "width": None, "height": None}


def _qa_page_records(before_pages: list[Path], after_pages: list[Path]) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    total = max(len(before_pages), len(after_pages))
    for idx in range(total):
        before = before_pages[idx] if idx < len(before_pages) else None
        after = after_pages[idx] if idx < len(after_pages) else None
        before_size = before.stat().st_size if before and before.exists() else 0
        after_size = after.stat().st_size if after and after.exists() else 0
        bw, bh = _png_dimensions(before) if before else (None, None)
        aw, ah = _png_dimensions(after) if after else (None, None)
        before_stats = _png_visual_stats(before)
        after_stats = _png_visual_stats(after)
        findings: list[str] = []
        if before is None:
            findings.append("missing before-page image")
        if after is None:
            findings.append("missing after-page image")
        if before_size and before_size < 1500:
            findings.append("before-page image is suspiciously small")
        if after_size and after_size < 1500:
            findings.append("after-page image is suspiciously small")
        if bw and aw and (bw, bh) != (aw, ah):
            findings.append("before/after rendered page dimensions differ")
        before_ink = before_stats.get("ink_ratio")
        after_ink = after_stats.get("ink_ratio")
        if isinstance(after_ink, float) and after_ink < 0.005:
            findings.append("after-page image appears blank or nearly blank")
        if isinstance(before_ink, float) and isinstance(after_ink, float) and abs(after_ink - before_ink) > 0.18:
            findings.append("before/after page ink density changed substantially")
        after_dark = after_stats.get("dark_ratio")
        if isinstance(after_dark, float) and after_dark > 0.25:
            findings.append("after-page has unusually large dark regions; inspect for overlap or rendering blocks")
        records.append(
            {
                "page": idx + 1,
                "before": str(before.resolve()) if before else None,
                "after": str(after.resolve()) if after else None,
                "before_size_bytes": before_size,
                "after_size_bytes": after_size,
                "before_dimensions": [bw, bh],
                "after_dimensions": [aw, ah],
                "before_visual_stats": before_stats,
                "after_visual_stats": after_stats,
                "findings": findings,
            }
        )
    return records


def render_full_document_qa(
    input_path: Path,
    fixed_path: Path,
    output_dir: Path,
    mode: str,
) -> tuple[str, dict[str, object] | None]:
    try:
        qa_dir = output_dir / "render_qa"
        before_dir = qa_dir / "before"
        after_dir = qa_dir / "after"
        before_pages = _render_docx_pages_with_artifact_tool(input_path, before_dir)
        after_pages = _render_docx_pages_with_artifact_tool(fixed_path, after_dir)
        records = _qa_page_records(before_pages, after_pages)
        finding_count = sum(len(record["findings"]) for record in records)
        status = (
            "Full-document render QA generated with Codex artifact-tool. "
            f"Rendered {len(before_pages)} before pages and {len(after_pages)} after pages; "
            f"basic image checks found {finding_count} warning(s). "
            "When the fixed DOCX is produced through OOXML fallback, these page images are QA previews only and may not match Word/WPS rendering exactly."
        )
        return status, {
            "renderer": "artifact-tool",
            "before_page_count": len(before_pages),
            "after_page_count": len(after_pages),
            "warning_count": finding_count,
            "pages": records,
            "note": (
                "Use these page images to inspect layout drift, overlap, clipping, and header/footer issues when Office/WPS COM is unavailable. "
                "If the fixed DOCX was produced through OOXML fallback, preview images may be inaccurate because Word/WPS pagination, fields, fonts, and compatibility layout are not refreshed by COM."
            ),
        }
    except Exception as exc:
        status = f"Full-document render QA unavailable: {exc}"
        if mode == "require":
            raise RuntimeError(status) from exc
        return status, None


def attach_screenshots(
    input_path: Path,
    fixed_path: Path,
    issues: list[Issue],
    output_dir: Path,
    renderer_used: str,
    mode: str,
) -> tuple[str, dict[str, object] | None]:
    if mode == "never":
        note = "Screenshots disabled by --screenshots never."
        for issue in issues:
            issue.screenshot_note = note
        return note, None
    if renderer_used not in {"office", "wps"}:
        note, qa = render_full_document_qa(input_path, fixed_path, output_dir, mode)
        issue_note = (
            "Paragraph-level screenshots require Microsoft Word or WPS COM rendering. "
            "Full-document before/after page render QA was generated instead."
            if qa
            else note
        )
        for issue in issues:
            issue.screenshot_note = issue_note
        return note, qa
    try:
        pages = _collect_pages_with_com(input_path, issues, renderer_used)
        for issue in issues:
            issue.page = pages.get(issue.paragraph_index)
        needed_pages = {page for page in pages.values() if page}
        if not needed_pages:
            note = "Screenshots skipped because paragraph page numbers could not be resolved."
            if mode == "require":
                raise RuntimeError(note)
            for issue in issues:
                issue.screenshot_note = note
            return note, None

        screenshot_dir = output_dir / "screenshots"
        before_pdf = output_dir / "_before_render.pdf"
        after_pdf = output_dir / "_after_render.pdf"
        _export_pdf_with_com(input_path, before_pdf, renderer_used)
        _export_pdf_with_com(fixed_path, after_pdf, renderer_used)
        before_images = _render_pdf_pages(before_pdf, needed_pages, screenshot_dir, "before")
        after_images = _render_pdf_pages(after_pdf, needed_pages, screenshot_dir, "after")
        before_pdf.unlink(missing_ok=True)
        after_pdf.unlink(missing_ok=True)
        for issue in issues:
            if issue.page in before_images:
                issue.before_screenshot = str(before_images[issue.page].resolve())
            if issue.page in after_images:
                issue.after_screenshot = str(after_images[issue.page].resolve())
        status, qa = render_full_document_qa(input_path, fixed_path, output_dir, "auto")
        return "Paragraph-level screenshots generated from Office/WPS-rendered DOCX pages. " + status, qa
    except Exception as exc:
        note = f"Screenshots unavailable: {exc}"
        if mode == "require":
            raise RuntimeError(note) from exc
        for issue in issues:
            issue.screenshot_note = note
        return note, None


def attach_after_formats(fixed_path: Path, issues: list[Issue]) -> None:
    fixed_doc = Document(str(fixed_path))
    for issue in issues:
        if 0 <= issue.paragraph_index < len(fixed_doc.paragraphs):
            rule_key = issue.rule_key.removesuffix("_runs")
            issue.after = current_format(fixed_doc.paragraphs[issue.paragraph_index], RULES.get(rule_key))


def issue_dict(issue: Issue) -> dict[str, object]:
    return {
        "paragraph_index": issue.paragraph_index,
        "rule_key": issue.rule_key,
        "category": issue.category,
        "text_type": issue.text_type,
        "text_excerpt": issue.text_excerpt,
        "location": issue.location,
        "page": issue.page,
        "current": issue.current,
        "expected": issue.expected,
        "after": issue.after,
        "message": issue.message,
        "before_screenshot": issue.before_screenshot,
        "after_screenshot": issue.after_screenshot,
        "screenshot_note": issue.screenshot_note,
    }


def summarize_issues(issues: list[Issue]) -> dict[str, int]:
    summary: dict[str, int] = {}
    for issue in issues:
        summary[issue.category] = summary.get(issue.category, 0) + 1
    return dict(sorted(summary.items()))


def write_html_report(report: dict[str, object], path: Path) -> None:
    issues = report.get("issues", [])
    assert isinstance(issues, list)
    category_summary = report.get("issue_summary_by_category") or {}
    category_html = ""
    if isinstance(category_summary, dict):
        rows_summary = "".join(
            f"<tr><td>{html.escape(str(category))}</td><td>{html.escape(str(count))}</td></tr>"
            for category, count in category_summary.items()
        )
        category_html = (
            "<h2>问题分类汇总</h2>"
            "<table class='summary-table'><thead><tr><th>类别</th><th>数量</th></tr></thead>"
            f"<tbody>{rows_summary}</tbody></table>"
        )
    render_qa = report.get("render_qa")
    qa_html = ""
    if isinstance(render_qa, dict):
        qa_rows = []
        for page_obj in render_qa.get("pages", []):
            page = page_obj if isinstance(page_obj, dict) else {}
            before = page.get("before")
            after = page.get("after")
            warnings = page.get("findings") or []
            warning_text = ", ".join(str(item) for item in warnings) if isinstance(warnings, list) else str(warnings)
            before_html = ""
            after_html = ""
            if before:
                before_html = f'<figure><figcaption>修改前 第 {html.escape(str(page.get("page")))} 页</figcaption><img src="{html.escape(Path(str(before)).resolve().as_uri())}"></figure>'
            if after:
                after_html = f'<figure><figcaption>修改后 第 {html.escape(str(page.get("page")))} 页</figcaption><img src="{html.escape(Path(str(after)).resolve().as_uri())}"></figure>'
            qa_rows.append(
                "<section class='qa-page'>"
                f"<h3>页面 {html.escape(str(page.get('page')))}</h3>"
                f"<p><strong>基础检查：</strong>{html.escape(warning_text or '未发现图片尺寸/文件异常')}</p>"
                f"<div class='shots'>{before_html}{after_html}</div>"
                "</section>"
            )
        qa_html = (
            "<h1>渲染截图 QA</h1>"
            "<div class='summary'>"
            f"<p><strong>渲染器：</strong>{html.escape(str(render_qa.get('renderer', '')))}</p>"
            f"<p><strong>修改前页数：</strong>{html.escape(str(render_qa.get('before_page_count', '')))}</p>"
            f"<p><strong>修改后页数：</strong>{html.escape(str(render_qa.get('after_page_count', '')))}</p>"
            f"<p><strong>基础警告数：</strong>{html.escape(str(render_qa.get('warning_count', '')))}</p>"
            f"<p>{html.escape(str(render_qa.get('note', '')))}</p>"
            "</div>"
            + "".join(qa_rows)
        )
    structure_warnings = report.get("structure_warnings") or []
    if structure_warnings:
        warning_items = "".join(f"<li>{html.escape(str(item))}</li>" for item in structure_warnings)
        structure_html = f"<section class='structure'><h2>组成部分提醒</h2><ul>{warning_items}</ul></section>"
    else:
        structure_html = ""
    rows = []
    for n, issue_obj in enumerate(issues, start=1):
        issue = issue_obj if isinstance(issue_obj, dict) else {}
        before_img = issue.get("before_screenshot")
        after_img = issue.get("after_screenshot")
        screenshots = html.escape(str(issue.get("screenshot_note") or ""))
        if before_img and after_img:
            before_uri = Path(str(before_img)).resolve().as_uri()
            after_uri = Path(str(after_img)).resolve().as_uri()
            screenshots = (
                f'<div class="shots"><figure><figcaption>修改前</figcaption>'
                f'<img src="{html.escape(before_uri)}"></figure>'
                f'<figure><figcaption>修改后</figcaption>'
                f'<img src="{html.escape(after_uri)}"></figure></div>'
            )
        rows.append(
            "<section class='issue'>"
            f"<h2>#{n} {html.escape(str(issue.get('text_type', 'unknown')))}</h2>"
            f"<p><strong>位置：</strong>段落 {html.escape(str(issue.get('paragraph_index')))}"
            f"{'，页 ' + html.escape(str(issue.get('page'))) if issue.get('page') else ''}</p>"
            f"<p><strong>文本：</strong>{html.escape(str(issue.get('text_excerpt') or ''))}</p>"
            f"<p><strong>修改前：</strong>{html.escape(str(issue.get('current') or ''))}</p>"
            f"<p><strong>应改为：</strong>{html.escape(str(issue.get('expected') or ''))}</p>"
            f"<p><strong>修改后：</strong>{html.escape(str(issue.get('after') or ''))}</p>"
            f"{screenshots}"
            "</section>"
        )
    doc = f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>CUIT thesis DOCX format report</title>
<style>
body {{ font-family: "Microsoft YaHei", "Segoe UI", sans-serif; margin: 32px; line-height: 1.6; color: #1f2937; }}
h1 {{ margin-bottom: 8px; }}
.summary {{ padding: 16px; background: #f3f4f6; border-left: 4px solid #2563eb; margin: 16px 0 24px; }}
.summary-table {{ border-collapse: collapse; margin: 12px 0 24px; }}
.summary-table th, .summary-table td {{ border: 1px solid #d1d5db; padding: 6px 10px; text-align: left; }}
.issue {{ border: 1px solid #d1d5db; border-radius: 6px; padding: 16px; margin: 18px 0; }}
.structure {{ border: 1px solid #f59e0b; background: #fffbeb; border-radius: 6px; padding: 16px; margin: 18px 0; }}
.qa-page {{ border: 1px solid #bfdbfe; border-radius: 6px; padding: 16px; margin: 18px 0; }}
.issue h2 {{ margin-top: 0; font-size: 18px; }}
.shots {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }}
figure {{ margin: 0; }}
img {{ max-width: 100%; border: 1px solid #d1d5db; }}
code {{ background: #f3f4f6; padding: 1px 4px; }}
</style>
</head>
<body>
<h1>成都信息工程大学学士学位论文 DOCX 格式检查报告</h1>
<div class="summary">
<p><strong>输入文件：</strong>{html.escape(str(report.get("input", "")))}</p>
<p><strong>批注版：</strong>{html.escape(str(report.get("annotated_docx", "")))}</p>
<p><strong>修正版：</strong>{html.escape(str(report.get("fixed_docx", "")))}</p>
<p><strong>修改/问题总数：</strong>{html.escape(str(report.get("issue_count", 0)))}</p>
<p><strong>渲染方式：</strong>fixed={html.escape(str(report.get("renderer_for_fixed", "")))}, comments={html.escape(str(report.get("renderer_for_comments", "")))}</p>
<p><strong>截图状态：</strong>{html.escape(str(report.get("screenshot_status", "")))}</p>
</div>
{category_html}
{structure_html}
{qa_html}
{''.join(rows)}
</body>
</html>
"""
    path.write_text(doc, encoding="utf-8")


def stdout_json(report: dict[str, object]) -> str:
    return json.dumps(report, ensure_ascii=True, indent=2)


def run(
    input_path: Path,
    output_dir: Path,
    renderer: str,
    screenshots: str,
    allow_ooxml_layout_fixes: bool = False,
) -> dict[str, object]:
    require_docx_dependencies()
    if not input_path.exists():
        raise FileNotFoundError(input_path)
    if not is_docx(input_path):
        raise ValueError("Input must be a .docx file.")
    output_dir.mkdir(parents=True, exist_ok=True)

    comments_path = output_dir / f"{input_path.stem}_format_comments.docx"
    fixed_path = output_dir / f"{input_path.stem}_format_fixed.docx"
    report_path = output_dir / f"{input_path.stem}_format_report.json"
    html_report_path = output_dir / f"{input_path.stem}_format_report.html"

    source_doc = Document(str(input_path))
    source_texts = [paragraph_text(paragraph) for paragraph in source_doc.paragraphs]
    structure_analysis = analyze_section_sequence(source_texts)
    issues = collect_issues(source_doc)
    create_annotated_docx(input_path, comments_path, issues)

    planned_renderer = probe_com_renderer(renderer)
    allow_layout_fixes = planned_renderer in {"office", "wps"} or allow_ooxml_layout_fixes
    fixed_doc = Document(str(input_path))
    apply_supported_rules(fixed_doc, allow_layout_fixes=allow_layout_fixes)
    fixed_doc.save(str(fixed_path))
    fixed_renderer = try_com_save(fixed_path, renderer)
    comments_renderer = try_com_save(comments_path, renderer)
    attach_after_formats(fixed_path, issues)
    screenshot_status, render_qa = attach_screenshots(
        input_path=input_path,
        fixed_path=fixed_path,
        issues=issues,
        output_dir=output_dir,
        renderer_used=fixed_renderer,
        mode=screenshots,
    )

    _header_start, header_reason = find_header_start_section(source_doc)
    advisory = [
        "Semantic correctness of cover fields is not checked.",
        "Bibliographic content validity is not fully checked.",
        "Dynamic page number fields are preserved where present; this script does not rebuild complex field codes.",
    ]
    if header_reason == "fallback-front-matter":
        advisory.append(
            "No declaration page was detected. Header insertion starts from the first recognized abstract/TOC/body section, so verify the required declaration page and header start manually."
        )
    if header_reason == "not-found":
        advisory.append(
            "No reliable header start page was detected. The script cannot safely decide where the declaration-page header should begin."
        )
    if fixed_renderer == "ooxml":
        advisory.append(
            "The fixed DOCX was produced through OOXML fallback. Render QA images are approximate previews and may not match Microsoft Word or WPS exactly; use Word/WPS for final visual confirmation when possible."
        )
    if fixed_renderer == "ooxml" and not allow_ooxml_layout_fixes:
        advisory.append(
            "OOXML mode did not automatically modify high-risk layout features: headers, footers, page numbers, header/footer distance, or section-start behavior. Add --allow-ooxml-layout-fixes only if the user explicitly accepts possible pagination drift."
        )
    if fixed_renderer == "ooxml" and allow_ooxml_layout_fixes:
        advisory.append(
            "OOXML high-risk layout fixes were explicitly enabled. Verify headers, footers, page numbers, section breaks, and pagination in Word/WPS before using the fixed document."
        )

    report = {
        "input": str(input_path.resolve()),
        "annotated_docx": str(comments_path.resolve()),
        "fixed_docx": str(fixed_path.resolve()),
        "json_report": str(report_path.resolve()),
        "html_report": str(html_report_path.resolve()),
        "issue_count": len(issues),
        "modification_count": len(issues),
        "issue_summary_by_category": summarize_issues(issues),
        "renderer_for_fixed": fixed_renderer,
        "renderer_for_comments": comments_renderer,
        "planned_renderer_for_layout_fixes": planned_renderer,
        "high_risk_layout_fixes_applied": allow_layout_fixes,
        "allow_ooxml_layout_fixes": allow_ooxml_layout_fixes,
        "screenshot_status": screenshot_status,
        "render_qa": render_qa,
        "structure_analysis": structure_analysis,
        "structure_warnings": structure_analysis["warnings"],
        "issues": [issue_dict(issue) for issue in issues],
        "unsupported_or_advisory": advisory,
    }
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    write_html_report(report, html_report_path)
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Check and fix CUIT bachelor thesis DOCX formatting.")
    parser.add_argument("docx", help="Input .docx thesis path")
    parser.add_argument("--output-dir", default="./format_results", help="Output directory")
    parser.add_argument("--renderer", choices=["auto", "office", "wps", "ooxml"], default="auto")
    parser.add_argument(
        "--screenshots",
        choices=["auto", "never", "require"],
        default="auto",
        help="Generate before/after page screenshots when Office/WPS rendering is available.",
    )
    parser.add_argument(
        "--allow-ooxml-layout-fixes",
        action="store_true",
        help=(
            "Explicitly allow OOXML-only changes to high-risk layout features "
            "(headers, footers, page numbers, header/footer distance, and section behavior)."
        ),
    )
    args = parser.parse_args()
    report = run(
        Path(args.docx),
        Path(args.output_dir),
        args.renderer,
        args.screenshots,
        allow_ooxml_layout_fixes=args.allow_ooxml_layout_fixes,
    )
    print(stdout_json(report))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
