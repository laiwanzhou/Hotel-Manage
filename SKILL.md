---
name: cuit-thesis-docx-format
description: Use when checking or fixing Chengdu University of Information Technology bachelor thesis DOCX formatting against the university specification, especially when the user needs an annotated review copy and a separately formatted output copy while preserving the original document.
---

# CUIT Thesis DOCX Format

## Purpose

Use this skill for `.docx` bachelor thesis format checks against the Chengdu University of Information Technology thesis specification. The workflow must preserve the input file and create two new `.docx` files:

1. an annotated copy that marks each detected formatting issue with Word comments
2. a formatted copy with supported rules applied

## Inputs

- A thesis file ending in `.docx`
- Multiple thesis files ending in `.docx` when batch checking is requested
- Optional output directory
- Optional renderer choice: `auto`, `office`, `wps`, or `ooxml`

Do not use this skill for `.pdf`, `.tex`, or `.typ` files.

## Required Output Contract

Always produce:

- `<stem>_format_comments.docx`
- `<stem>_format_fixed.docx`
- `<stem>_format_report.json`
- `<stem>_format_report.html`

Never overwrite or modify the input file.

Each comment in the annotated copy must state:

- text type, such as chapter title, second-level heading, body paragraph, figure caption, table caption, abstract title, or keywords
- current format
- expected format

## Execution

Run the bundled checker:

```powershell
python scripts/cuit_thesis_docx_format.py thesis.docx --output-dir ./format_results --renderer auto --screenshots auto
```

On Windows PowerShell, set UTF-8 console and Python IO encoding before commands that contain Chinese paths:

```powershell
$env:PYTHONIOENCODING='utf-8'
$OutputEncoding=[System.Text.Encoding]::UTF8
[Console]::OutputEncoding=[System.Text.Encoding]::UTF8
```

Always pass paths as quoted arguments or `-LiteralPath` values. Avoid building command strings that concatenate Chinese file names; prefer direct argument arrays or the bundled Python batch script.

When capturing the command-line JSON stdout for audit/debugging, write it inside the same output directory, for example:

```powershell
python scripts/cuit_thesis_docx_format.py thesis.docx --output-dir ./res2 --renderer ooxml --screenshots auto | Out-File -LiteralPath ./res2/run-output.json -Encoding utf8
```

Do not leave run-output files in the parent source folder when the user asked for results in a specific `resx` folder.

For multiple thesis files, run the batch matrix script:

```powershell
$env:PYTHONIOENCODING='utf-8'
$OutputEncoding=[System.Text.Encoding]::UTF8
[Console]::OutputEncoding=[System.Text.Encoding]::UTF8
python scripts/run_batch_matrix.py `
  "paper-format/论文一.docx" `
  "paper-format/论文二.docx" `
  --output-dir "paper-format/resx" `
  --renderer ooxml `
  --screenshots auto |
  Out-File -LiteralPath "paper-format/resx/batch-run-output.json" -Encoding utf8
```

Batch output structure must be:

```text
resx/
  batch-summary.json
  batch-run-output.json
  论文一/
    run-output.json
    论文一_format_comments.docx
    论文一_format_fixed.docx
    论文一_format_report.json
    论文一_format_report.html
    render_qa/
  论文二/
    run-output.json
    ...
```

Each child folder name must be based on the input thesis stem. If duplicate stems exist, append `_2`, `_3`, and so on. Never mix outputs from different theses in the same folder.

## Environment Gate

Before processing a user thesis, run the environment checker:

```powershell
python scripts/check_environment.py
```

The checker reports one of three capability levels:

- `full`: required Python packages are available and Microsoft Word or WPS COM automation works.
- `ooxml-only`: required Python packages are available, but Word/WPS COM automation is unavailable.
- `blocked`: required Python packages are missing.

If the capability level is `blocked`, install dependencies before continuing:

```powershell
python -m pip install -r requirements.txt
```

Conda users can instead create an isolated environment:

```powershell
conda env create -f environment.yml
conda activate cuit-thesis-docx-format
```

If the capability level is `ooxml-only`, pause before running the full task and explain:

- Without `win32com`, the script can still create the annotated DOCX, fixed DOCX, JSON report, and HTML report through pure OOXML.
- Without `win32com` plus working Microsoft Word/WPS COM, the script cannot reliably drive Microsoft Word or WPS to refresh layout, update fields, export PDFs, update TOC/page fields, or generate Word/WPS-accurate before/after screenshots.
- OOXML-only output can cause or fail to detect page-number drift, header/footer drift, section-break issues, TOC field staleness, and preview screenshots that differ from Word/WPS.
- With `win32com` plus installed Microsoft Word or WPS, `--renderer auto` can try Word first, then WPS, then fall back to OOXML.

Ask the user to choose one of these paths before continuing:

1. install/enable the Python COM dependency and Word/WPS COM automation, then run with `--renderer auto`
2. continue immediately with `--renderer ooxml`

When this skill is used by Codex, this pause-and-confirm gate is mandatory. The bundled script itself remains non-interactive by default so it can run in automated workflows; direct script users choose behavior through `--renderer`.

On Windows, the usual installation command is:

```powershell
python -m pip install pywin32
```

After installation, verify again with:

```powershell
python scripts/check_environment.py
```

Renderer behavior:

- `auto`: try Microsoft Office Word first, then WPS, then pure OOXML.
- `office`: require Microsoft Word COM automation for final compatibility save.
- `wps`: require WPS COM automation for final compatibility save.
- `ooxml`: use pure DOCX/OOXML processing only.
- In OOXML mode, high-risk layout features are report-only by default: headers, footers, page numbers, header/footer distance, and section-start behavior are not automatically rewritten.
- `--allow-ooxml-layout-fixes`: explicitly allow OOXML-only high-risk layout edits. Use only after the user accepts possible pagination drift and agrees to verify the result in Word/WPS.

If Office/WPS automation is unavailable, do not fail in `auto`; fall back to OOXML and report the fallback.

Screenshot behavior:

- `--screenshots auto`: generate paragraph-level before/after screenshots when Office or WPS COM rendering is available; otherwise render full-document before/after page images with the Codex `documents` artifact-tool renderer when available.
- `--screenshots require`: fail if neither Office/WPS nor artifact-tool rendering can generate screenshots.
- `--screenshots never`: skip screenshots.

When Office/WPS is unavailable, screenshots are full-page before/after QA images rather than paragraph-located screenshots because pure OOXML cannot reliably tell which rendered page contains a paragraph. The HTML/JSON report must state this fallback clearly.

When the fixed DOCX is produced with `--renderer ooxml` or through OOXML fallback, preview images are approximate QA images only. They may differ from Microsoft Word or WPS because pagination, field refresh, fonts, compatibility layout, and some drawing objects are not refreshed through Office/WPS COM. The result must tell the user not to treat these screenshots as final visual truth.

When capability is `ooxml-only`, do not automatically repair headers, footers, page numbers, or section breaks unless the user explicitly asks for `--allow-ooxml-layout-fixes`. Still report these issues in JSON/HTML and annotated comments where possible.

## Supported Rules

The script applies and checks the rules summarized in `references/cuit-format-spec.md`.
Machine-readable rules live in `references/rules.json`; update that file before changing Python code when the university specification changes.

## Thesis Section Boundaries

Treat every thesis as nine ordered, continuous sections. The sections appear in this order, do not repeat, and a later section marker means the previous section has ended:

1. Cover: the first page only; fields include classification number, UDC, security level, thesis title, name, student number, college, major, research direction, supervisor, assistant supervisor, and submission date.
2. Declaration page: originality statement and thesis copyright authorization. The specification does not define formatting rules for this part, so only report it when missing.
3. Abstract and keywords.
4. Table of contents.
5. Symbol explanation: optional; when absent, remind the user in JSON/HTML.
6. Body: introduction or preface, thesis body, and conclusion.
7. References.
8. Appendix: optional; when absent, remind the user in JSON/HTML.
9. Acknowledgement.

Do not classify cover fields by searching the whole thesis. Cover fields, including submission date, are only checked inside the cover section. Use `第一章`, `引言`, or `绪论` as auxiliary body-start boundaries when available.

Chinese `关键词` and English `Key words` are distinct paragraph types even when they share most formatting requirements.

When Office/WPS automation is available, header and page-number fixes may be applied automatically. When it is unavailable, report those issues and the OOXML fallback risk instead of claiming final layout certainty.

Page-number rules are section-specific:

- Declaration through table-of-contents pages use continuous lowercase Roman numerals: `i, ii, iii, iv, v, ...`, small five 9pt, centered. The university specification table may contain an old “Greek numeral” wording; treat that as corrected to lowercase Roman numerals.
- Body through acknowledgement pages restart at Arabic page 1 and display `第*页 共*页`, Times New Roman small five 9pt, centered.

Supported automated formatting includes:

- A4 page size
- top/bottom margins 2.5 cm and left/right margins 3 cm
- centered university thesis header
- cover title, cover identity fields, and submission date
- Chinese and English abstract titles, abstract body text, and keywords
- table-of-contents title, total-pages line, and first/second/third-level TOC entries
- chapter, second-level, and third-level heading font/alignment/spacing
- body paragraph font, first-line indent, spacing, and line height
- figure and table caption font/alignment/spacing
- formula number formatting
- reference/appendix/acknowledgement title formatting
- reference entries, symbol explanations, appendix body, author biography body, and acknowledgement body
- explicit page setup, header start-from-declaration, and footer/page-number issues in reports
- run-level formatting scan for every text run inside recognized paragraphs, so mixed-font or partially wrong text is reported even when the paragraph-level style looks correct
- table-of-contents structure checks for Word/WPS TOC fields and right-aligned page-number tab stops
- table checks for caption placement above the table, likely three-line table structure, and repeated headers for continued tables
- reference section checks for position after conclusion and before appendix, title formatting, entry font/line spacing, and hanging indent. Do not check GB/T 7714 bibliographic content details such as type markers, years, or whether numbering uses `[序号]`.
- full-page render QA warnings for blank/nearly blank pages, before/after page-size drift, large ink-density changes, and unusually dark rendered regions

The script must not claim full compliance for rules that require human review, such as research content quality, whether the cover information is semantically correct, or whether every reference item is bibliographically valid.

## Workflow

1. Validate the input path and ensure it is `.docx`.
2. Copy the input to an annotated output path and a fixed output path.
3. Load `references/rules.json`.
4. Inspect paragraphs, headers, footers, page setup, and recognized captions/headings.
5. Build the nine-section sequence first, then classify each checked paragraph inside its section. Report missing required sections and optional-section reminders in JSON/HTML.
6. Compare current formatting against the expected rule for that text type, resolving direct formatting and paragraph-style inheritance where possible.
7. Scan all runs in each recognized paragraph and report mixed direct/inherited run formatting that violates the rule.
8. Run structural checks for TOC fields/page-number alignment, table captions/three-line table signals/continued headers, and reference section position/indent/basic formatting.
9. Insert Word comments into the annotated copy for paragraph-level detected issues.
10. Apply supported formatting rules to the fixed copy. If the selected path is OOXML-only, skip automatic high-risk layout rewrites unless `--allow-ooxml-layout-fixes` was explicitly requested.
11. Compare the fixed copy and record the after-format for each issue.
12. Run renderer compatibility save:
   - prefer Microsoft Word
   - then try WPS
   - otherwise leave the pure OOXML output
13. Write JSON and HTML reports:
   - modification count
   - issue count by category
   - paragraph index and page number when available
   - text excerpt
   - current format
   - expected format
   - after-format
   - before/after screenshots when available
   - full-document render QA page images when paragraph-level screenshots are unavailable
   - fallback or unsupported-check notes
   - an OOXML preview accuracy warning whenever the fixed DOCX was produced without Office/WPS COM rendering
14. If stdout is captured, save the run-output JSON inside the selected output directory, such as `res1/run-output.json`, `res2/run-output.json`, or another user-specified `resx` folder.

## Batch Matrix Workflow

Use `scripts/run_batch_matrix.py` when two or more `.docx` theses are provided, or when the user wants a formal test matrix. The batch script must:

1. accept Chinese paths as normal Python arguments, not through manual PowerShell string concatenation
2. create one child output folder per input thesis under the requested `resx` output root
3. run the single-document checker for each thesis
4. save each thesis stdout as that child folder's `run-output.json`
5. verify the input hash is unchanged
6. verify comments/fixed/JSON/HTML outputs exist
7. verify the JSON report is readable and `issue_count` equals `modification_count`
8. write `batch-summary.json` in the root `resx` folder with pass/fail status, paths, issue counts, renderer, screenshot status, and warnings for each thesis

For batch runs, pass `--allow-ooxml-layout-fixes` only when every input document owner accepts the same OOXML layout risk.

If one thesis fails, continue checking the remaining theses and record the failure in `batch-summary.json`.

## Test Fixtures

`tests/fixtures/大模型生成文本检测技术研究_刘鹄鸣_minimal.docx` is a minimal executable DOCX fixture converted from the user-provided PDF by PyMuPDF text extraction. It includes the front matter, TOC-like pages, conclusion pages, and reference pages. Use it to exercise paragraph classification, run scanning, TOC/reference checks, and renderer QA. Do not use it to judge faithful PDF layout conversion because it is a text-extraction fixture, not a visual reproduction of the PDF.

## Common Mistakes

- Do not edit the original thesis file.
- Do not silently replace comments with inline text; comments must be Word comments when possible.
- Do not use PDF visual checks as a substitute for DOCX style checks.
- Do not claim Office or WPS was used unless the script reports that renderer as successful.
- Do not claim screenshots were generated unless the report includes screenshot paths.
- Do not claim paragraph-level screenshot precision when the renderer fell back to artifact-tool full-page QA.
- Do not auto-fix ambiguous cover-page content unless a paragraph type can be identified confidently.
- Do not auto-fix OOXML-only headers, footers, page numbers, or section behavior unless the user explicitly enabled `--allow-ooxml-layout-fixes`.

## Verification Scenarios

Use `tests/pressure-scenarios.md` when updating this skill or its script.
