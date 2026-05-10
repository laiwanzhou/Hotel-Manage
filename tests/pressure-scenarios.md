# Pressure Scenarios

Use these scenarios before changing the skill or script.

## Scenario 1: Preserve Original

Input: a `.docx` thesis with wrong margins and wrong heading fonts.

Expected:

- original file hash is unchanged
- `_format_comments.docx` exists
- `_format_fixed.docx` exists
- `_format_report.json` exists
- `_format_report.html` exists
- comments mention text type, current format, and expected format

## Scenario 2: Renderer Fallback

Input: a valid `.docx` on a machine without Microsoft Office or WPS COM automation.

Expected:

- command succeeds with `--renderer auto`
- report states OOXML fallback for editing/compatibility save
- both output files are created
- HTML and JSON reports include full-document render QA page images when Codex artifact-tool is available
- if artifact-tool is unavailable, HTML and JSON reports state why screenshots were skipped

## Scenario 3: Office Priority

Input: a valid `.docx` on a machine with Microsoft Word and WPS installed.

Expected:

- `--renderer auto` attempts Microsoft Word first
- WPS is not used if Word compatibility save succeeds
- paragraph-level screenshots are generated from Word-rendered pages when `--screenshots auto` is used
- full-document render QA images are also generated when artifact-tool is available

## Scenario 4: Ambiguous Paragraph

Input: cover-page metadata or a short standalone line that does not match a known heading/caption pattern.

Expected:

- script does not overclaim the text type
- script does not apply body formatting unless the paragraph is in a recognized body region

## Scenario 5: Comment Content

Input: a chapter title `第一章绪论` in 12 pt, unbold, left aligned.

Expected comment:

- text type: chapter title
- current format: includes font size, bold state, alignment, spacing
- expected format: Songti 16 pt, bold, centered, 20 pt line spacing, before/after 0.5 line, one space between chapter number and title

## Scenario 6: Report Entries

Input: a thesis with at least three formatting issues.

Expected:

- JSON `modification_count` equals the number of issue entries
- each issue entry contains `paragraph_index`, `text_excerpt`, `current`, `expected`, and `after`
- HTML report shows the same issue count and a readable entry for each issue
- JSON and HTML reports include issue counts grouped by category
- screenshot fields contain image paths only when screenshots were actually generated

## Scenario 7: No COM Render QA

Input: a valid `.docx` on a machine without `win32com`, Microsoft Word COM, or WPS COM.

Expected:

- `--screenshots auto` attempts Codex artifact-tool full-document rendering
- JSON report contains `render_qa` with before/after page image paths when artifact-tool succeeds
- HTML report includes a `渲染截图 QA` section
- each issue explains that paragraph-level screenshots require Office/WPS COM when only full-document QA is available

## Scenario 8: Full Rule Coverage

Input: a thesis containing cover fields, Chinese/English abstracts, TOC entries, body headings, figure/table captions, formulas, references, appendix, author biography, acknowledgement, headers, and page numbers.

Expected:

- all text types described by the university specification are represented in `references/rules.json`
- recognized instances are categorized in the JSON report
- page setup and header/footer issues appear as explicit report issues even when they cannot be attached to a paragraph comment

## Scenario 9: Style Inheritance

Input: a paragraph whose font size, font name, bold state, alignment, or line spacing comes from its Word paragraph style rather than direct run formatting.

Expected:

- report resolves inherited style values where possible
- script does not report `inherited` as an error when the inherited style satisfies the rule

## Scenario 10: Run-Level Mixed Formatting

Input: a recognized body paragraph where the paragraph style is correct, but one internal run uses a wrong font, wrong size, or wrong bold state.

Expected:

- JSON report contains a `run-format` issue for the paragraph
- issue `current` includes a concise summary of the offending run text and its effective direct/inherited format
- annotated DOCX comment states the text type, current run format, and expected paragraph rule

## Scenario 11: TOC, Table, and Reference Structure

Input: a thesis with a visible directory, tables, conclusion, references, and optional appendix.

Expected:

- TOC issues are reported when a visible directory lacks a Word/WPS TOC field or TOC entries lack right-aligned page-number tab stops
- table issues are reported when table captions are below tables or absent, when table style suggests full grid/vertical borders, or when a continued table lacks a repeated header
- reference issues are reported when the reference section appears before conclusion or after appendix, entries lack hanging indent, entries do not start with `[序号]`, or entries lack basic GB/T 7714 type/year markers

## Scenario 12: PDF-Derived Minimal Fixture

Input: `tests/fixtures/大模型生成文本检测技术研究_刘鹄鸣_minimal.docx`, generated from the user-provided PDF by PyMuPDF text extraction.

Expected:

- command succeeds without modifying the fixture
- report exercises front matter, TOC-like lines, conclusion, and reference entries
- fixture is treated as a text-extraction test set, not as a faithful PDF layout conversion
- artifact-tool render produces page PNGs suitable for visual QA

## Scenario 13: Render QA Visual Warnings

Input: a DOCX whose fixed output renders with a blank page, changed page dimensions, major before/after ink-density drift, or large dark blocks.

Expected:

- JSON `render_qa.pages[*]` includes `before_visual_stats` and `after_visual_stats`
- page-level `findings` flags likely blank pages, page-size changes, substantial ink-density changes, or unusually dark regions
- report describes these as QA warnings that require human confirmation, not as final visual proof

## Scenario 14: Batch Matrix Output Layout

Input: two or more `.docx` theses, including at least one Chinese file name.

Expected:

- `scripts/run_batch_matrix.py` creates one child folder per thesis under the requested `resx` directory
- each child folder name is based on the thesis file stem
- duplicate stems get `_2`, `_3`, and so on
- each child folder contains `run-output.json`, `_format_comments.docx`, `_format_fixed.docx`, `_format_report.json`, and `_format_report.html`
- root `resx/batch-summary.json` lists each input thesis, output folder, pass/fail checks, issue count, renderer, screenshot status, and render warning count
- outputs from different theses are never mixed in the same folder

## Scenario 15: Chinese Path Safety

Input: a thesis path containing Chinese characters, spaces, or punctuation.

Expected:

- PowerShell examples set `PYTHONIOENCODING`, `$OutputEncoding`, and `[Console]::OutputEncoding` to UTF-8
- scripts use Python `Path` and `subprocess.run([...])`, not string-built shell commands
- `run-output.json`, JSON reports, HTML reports, and batch summaries preserve Chinese file names
- instructions tell users to quote paths or use `-LiteralPath` when using PowerShell directly

## Scenario 16: OOXML High-Risk Layout Guard

Input: a DOCX with missing or wrong headers, footers, page numbers, or section-start behavior.

Expected:

- with `--renderer ooxml` and no `--allow-ooxml-layout-fixes`, the fixed DOCX does not rewrite headers, footers, page numbers, header/footer distance, or section-start behavior
- JSON/HTML reports still list high-risk layout issues and explain that OOXML mode skipped automatic layout rewrites
- report contains `high_risk_layout_fixes_applied: false`
- with `--renderer ooxml --allow-ooxml-layout-fixes`, report contains `high_risk_layout_fixes_applied: true` and advisory text warns that Word/WPS verification is required
- with working Word/WPS COM through `--renderer auto`, high-risk layout fixes may be applied because COM can refresh pagination and fields afterward
