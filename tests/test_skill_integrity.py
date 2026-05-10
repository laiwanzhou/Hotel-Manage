import importlib.util
import json
import subprocess
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class SkillIntegrityTests(unittest.TestCase):
    def load_checker_module(self):
        module_path = ROOT / "scripts" / "cuit_thesis_docx_format.py"
        spec = importlib.util.spec_from_file_location("cuit_thesis_docx_format", module_path)
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)
        return module

    def test_rules_json_is_utf8_and_contains_chinese_rules(self):
        rules_path = ROOT / "references" / "rules.json"
        data = json.loads(rules_path.read_text(encoding="utf-8"))

        self.assertEqual(data["source"], "成都信息工程大学学士学位论文规范.docx")
        self.assertEqual(data["rules"]["cover_title_zh"]["font_east_asia"], "宋体")
        self.assertIn("中文论文题目", data["rules"]["cover_title_zh"]["expected"])
        self.assertIn("参考文献", data["rules"]["reference_title"]["expected"])

    def test_checker_help_does_not_require_docx_dependencies(self):
        script = ROOT / "scripts" / "cuit_thesis_docx_format.py"

        proc = subprocess.run(
            [sys.executable, str(script), "--help"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("--renderer", proc.stdout)
        self.assertIn("--screenshots", proc.stdout)

    def test_stdout_json_is_ascii_safe_and_round_trips_chinese(self):
        module = self.load_checker_module()

        payload = {"input": "成都市育龄人口生育意愿.docx", "excerpt": "摘 要"}
        stdout = module.stdout_json(payload)

        stdout.encode("ascii")
        self.assertEqual(json.loads(stdout), payload)

    def test_comment_message_is_short_actionable_and_hides_run_details(self):
        module = self.load_checker_module()
        issue = module.Issue(
            paragraph_index=1,
            rule_key="keywords_runs",
            text_type="keywords paragraph run formatting",
            text_excerpt="关键词：数据可视化；爬虫；Mysql数据库；疫情监测系统",
            current="文字片段='关键词：'; 中文字体=Times New Roman Regular; 西文字体=Times New Roman Regular; 字号=12.0pt | 文字片段='Mysql'; 中文字体=Times New Roman Regular",
            expected="关键词段落宋体小四12pt，英文Times New Roman 12pt，关键词标签加粗，关键词之间用分号隔开。",
            message="",
            category="run-format",
        )

        message = module.comment_message_for_issue(issue)

        self.assertIn("关键词段落格式不符合要求。", message)
        self.assertIn("需调整：", message)
        self.assertIn("关键词标签加粗", message)
        self.assertLessEqual(len(message), 220)
        self.assertNotIn("文字片段", message)
        self.assertNotIn("中文字体=", message)

    def test_body_comment_uses_concrete_spacing_values(self):
        module = self.load_checker_module()
        issue = module.Issue(
            paragraph_index=1,
            rule_key="body",
            text_type="body paragraph",
            text_excerpt="正文段落",
            current="段前 16.35pt",
            expected="正文：中文宋体小四12pt，英文和数字Times New Roman 12pt，两端对齐或左对齐，首行缩进2个汉字符，段前段后0磅，固定20磅行距。",
            message="",
            category="body",
            location="paragraph 1",
        )

        message = module.comment_message_for_issue(issue)

        self.assertIn("首行缩进 2 个汉字符", message)
        self.assertIn("段前 0 磅，段后 0 磅", message)
        self.assertNotIn("按学校规范设置", message)

    def test_body_indent_is_written_as_two_character_units_not_centimeters(self):
        module = self.load_checker_module()
        document = module.Document()
        paragraph = document.add_paragraph("正文段落内容足够长，用于测试首行缩进字符单位。")
        rule = module.RULES["body"]

        module.apply_rule(paragraph, rule)

        ind = paragraph._p.get_or_add_pPr().find(module.qn("w:ind"))
        self.assertEqual(ind.get(module.qn("w:firstLineChars")), "200")
        self.assertIsNone(ind.get(module.qn("w:firstLine")))
        self.assertTrue(module.paragraph_matches(paragraph, rule))

    def test_comment_issues_merge_paragraph_and_run_format_for_same_paragraph(self):
        module = self.load_checker_module()
        paragraph_issue = module.Issue(
            paragraph_index=7,
            rule_key="keywords",
            text_type="keywords paragraph",
            text_excerpt="关键词：数据可视化；爬虫；Mysql数据库；疫情监测系统",
            current="固定行距不符合要求",
            expected="关键词段落宋体小四12pt，英文Times New Roman 12pt，关键词标签加粗，关键词之间用分号隔开。",
            message="",
            category="abstract",
        )
        run_issue = module.Issue(
            paragraph_index=7,
            rule_key="keywords_runs",
            text_type="keywords paragraph run formatting",
            text_excerpt="关键词：数据可视化；爬虫；Mysql数据库；疫情监测系统",
            current="文字片段='Mysql'; 中文字体=Times New Roman Regular",
            expected=paragraph_issue.expected,
            message="",
            category="run-format",
        )

        merged = module.issues_for_comments([paragraph_issue, run_issue])

        self.assertEqual(merged, [paragraph_issue])

    def test_body_date_is_not_classified_as_cover_submission_date(self):
        module = self.load_checker_module()

        key = module.classify_paragraph(
            "（2020 年 1 月 22 日）启动了全球第一个实时新冠病毒监测系统，使用 Python 进行数据处理。",
            in_body=True,
            region="body",
        )

        self.assertEqual(key, "body")

    def test_word_heading_styles_are_classified_even_when_number_text_is_missing_or_compact(self):
        module = self.load_checker_module()
        document = module.Document()
        heading2 = document.add_paragraph("课题背景")
        heading2.style = document.styles["Heading 2"]
        chapter = document.add_paragraph("2相关理论及技术")
        chapter.style = document.styles["Heading 1"]

        self.assertEqual(module.style_heading_key(heading2, "body"), "heading2")
        self.assertEqual(module.style_heading_key(chapter, "body"), "chapter")

    def test_heading_half_line_spacing_is_written_as_line_units_not_points(self):
        module = self.load_checker_module()
        document = module.Document()
        paragraph = document.add_paragraph("1.1 课题背景")
        rule = module.RULES["heading2"]

        module.apply_rule(paragraph, rule)

        spacing = paragraph._p.get_or_add_pPr().find(module.qn("w:spacing"))
        self.assertEqual(spacing.get(module.qn("w:beforeLines")), "50")
        self.assertEqual(spacing.get(module.qn("w:afterLines")), "50")
        self.assertIsNone(spacing.get(module.qn("w:before")))
        self.assertIsNone(spacing.get(module.qn("w:after")))
        self.assertTrue(module.paragraph_matches(paragraph, rule))

    def test_caption_spacing_is_zero_lines(self):
        module = self.load_checker_module()
        document = module.Document()
        for key, text in [("figure_caption", "图2-1 示例图"), ("table_caption", "表2-1 示例表")]:
            with self.subTest(key=key):
                paragraph = document.add_paragraph(text)
                rule = module.RULES[key]
                module.apply_rule(paragraph, rule)
                spacing = paragraph._p.get_or_add_pPr().find(module.qn("w:spacing"))
                ind = paragraph._p.get_or_add_pPr().find(module.qn("w:ind"))
                self.assertEqual(spacing.get(module.qn("w:beforeLines")), "0")
                self.assertEqual(spacing.get(module.qn("w:afterLines")), "0")
                self.assertEqual(ind.get(module.qn("w:firstLineChars")), "0")
                self.assertIsNone(ind.get(module.qn("w:firstLine")))
                self.assertTrue(module.paragraph_matches(paragraph, rule))

    def test_toc_indents_match_specification(self):
        module = self.load_checker_module()

        self.assertAlmostEqual(module.RULES["toc_level2"].first_line_indent_cm, 0.74, places=2)
        self.assertAlmostEqual(module.RULES["toc_level3"].first_line_indent_cm, 1.48, places=2)

    def test_back_matter_titles_follow_chapter_spacing(self):
        module = self.load_checker_module()

        for key in ["reference_title", "appendix_title", "author_bio_title", "acknowledgement_title"]:
            with self.subTest(key=key):
                self.assertEqual(module.RULES[key].space_before_lines, 0.5)
                self.assertEqual(module.RULES[key].space_after_lines, 0.5)

    def test_body_alignment_is_left_only_per_spec(self):
        module = self.load_checker_module()
        document = module.Document()
        paragraph = document.add_paragraph("正文段落内容足够长，用于测试对齐方式。")
        paragraph.alignment = module.WD_ALIGN_PARAGRAPH.JUSTIFY
        rule = module.RULES["body"]

        self.assertFalse(module.paragraph_matches(paragraph, rule))

    def test_short_body_paragraphs_are_checked(self):
        module = self.load_checker_module()

        self.assertEqual(
            module.classify_paragraph("主页：在主页里查看公告以及系统信息。", in_body=True, region="body"),
            "body",
        )

    def test_dot_numbered_figure_caption_is_not_body(self):
        module = self.load_checker_module()

        self.assertEqual(
            module.classify_paragraph("图3.5 后台管理模块-普通用户设计图", in_body=True, region="body"),
            "figure_caption",
        )

    def test_cover_submission_date_requires_cover_region_and_label(self):
        module = self.load_checker_module()

        self.assertEqual(
            module.classify_paragraph("论文提交日期：2023 年 6 月", in_body=False, region="cover"),
            "cover_date",
        )
        self.assertNotEqual(
            module.classify_paragraph("2023 年 6 月", in_body=False, region="cover"),
            "cover_date",
        )

    def test_structure_analysis_reports_missing_optional_sections(self):
        module = self.load_checker_module()
        texts = [
            "成都信息工程大学学位论文",
            "摘 要",
            "关键词：数据可视化；爬虫",
            "ABSTRACT",
            "Key words: Visualization; Spider",
            "目 录",
            "第一章 绪论",
            "1.1 研究背景",
            "参考文献",
            "[1] 作者. 题名[J]. 期刊, 2020.",
            "致 谢",
            "感谢老师。",
        ]

        analysis = module.analyze_section_sequence(texts)
        warnings = "\n".join(analysis["warnings"])

        self.assertEqual(analysis["regions"][7], "body")
        self.assertIn("符号说明", warnings)
        self.assertIn("附录", warnings)

    def test_structure_analysis_handles_inline_abstract_and_numbered_intro(self):
        module = self.load_checker_module()
        texts = [
            "成都信息工程大学学位论文",
            "摘要：本文研究疫情监测平台的设计与实现。",
            "关键词：数据可视化；爬虫",
            "Abstract: This paper studies an epidemic monitoring platform.",
            "Key words: Visualization; Spider",
            "1 引 言",
            "正文第一段包含足够长的文本，用于确认已经进入正文区域。",
            "参考文献",
            "1 作者. 题名. 期刊, 2020.",
            "致 谢",
        ]
        analysis = module.analyze_section_sequence(texts)
        self.assertEqual(analysis["regions"][1], "abstract_zh")
        self.assertEqual(analysis["regions"][3], "abstract_en")
        self.assertEqual(analysis["regions"][5], "body")
        self.assertEqual(analysis["regions"][6], "body")

    def test_inline_abstract_label_reports_missing_standalone_title(self):
        module = self.load_checker_module()
        document = module.Document()
        document.add_paragraph("摘要：本文研究疫情监测平台的设计与实现。")
        document.add_paragraph("关键词：数据可视化；爬虫")

        issues = module.collect_abstract_title_issues(document)

        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].rule_key, "abstract_title_zh_missing")
        self.assertIn("摘 要", issues[0].expected)

    def test_abstract_page_thesis_titles_are_kept_and_standalone_titles_are_required(self):
        module = self.load_checker_module()
        texts = [
            "基于Python的疫情监测平台的设计与实现",
            "摘要：随着新冠肺炎疫情的爆发和蔓延，疫情对人们的生活方式产生影响。",
            "关键词：数据可视化；爬虫",
            "Design and Implementation of an Epidemic Monitoring Platform Based on Python",
            "Abstract: With the outbreak and spread of the COVID-19, the epidemic has had a profound impact.",
            "Key words: Visualization; Spider",
        ]

        overrides = module.detect_abstract_thesis_titles(texts)

        self.assertEqual(overrides[0], "thesis_title_zh")
        self.assertEqual(overrides[3], "thesis_title_en")
        self.assertIn("thesis_title_zh", module.RULES)
        self.assertIn("thesis_title_en", module.RULES)

    def test_front_page_number_field_cache_uses_lowercase_roman(self):
        module = self.load_checker_module()
        source = ROOT.parent / "蒋若楸-20230516.docx"
        if not source.exists():
            self.skipTest("sample thesis not available")
        document = module.Document(str(source))

        module.apply_page_number_formats(document)

        visible_text, field_text = module._footer_field_info(document.sections[1])
        fmt, _start = module._section_page_number_format(document.sections[1])
        self.assertEqual(fmt, "lowerRoman")
        self.assertIn("roman", field_text)
        self.assertNotRegex(visible_text, r"\d|[IVXLCDM]")

    def test_page_number_formats_split_front_section_before_body(self):
        module = self.load_checker_module()
        source = ROOT.parent / "蒋若楸-20230516.docx"
        if not source.exists():
            self.skipTest("sample thesis not available")
        document = module.Document(str(source))

        module.ensure_body_starts_new_page_number_section(document)
        module.apply_page_number_formats(document)
        texts = [module.paragraph_text(paragraph) for paragraph in document.paragraphs]
        ranges = module.section_paragraph_ranges(document)
        regions = module.analyze_section_sequence(texts)["regions"]

        for section_idx, (start, end) in enumerate(ranges):
            if start > end:
                continue
            section_regions = set(regions[start : end + 1])
            fmt, _start_no = module._section_page_number_format(document.sections[section_idx])
            if "body" in section_regions:
                self.assertEqual(fmt, "decimal")
            elif {"abstract_zh", "abstract_en", "toc"} & section_regions:
                visible_text, field_text = module._footer_field_info(document.sections[section_idx])
                self.assertEqual(fmt, "lowerRoman")
                self.assertIn("roman", field_text)
                self.assertNotIn("Arabic", field_text)

    def test_reference_checks_only_cover_format_not_gbt_content_markers(self):
        module = self.load_checker_module()
        document = module.Document()
        document.add_paragraph("参考文献")
        document.add_paragraph("1 作者. 题名. 期刊, 2020.")

        issues = module.collect_reference_issues(document)
        rule_keys = {issue.rule_key for issue in issues}

        self.assertNotIn("reference_number_format", rule_keys)
        self.assertNotIn("reference_gbt7714_basic", rule_keys)

    def test_page_number_requirement_distinguishes_front_and_main_matter(self):
        module = self.load_checker_module()

        self.assertIsNotNone(module.validate_front_page_number("decimal", "I", "PAGE"))
        self.assertIsNone(module.validate_front_page_number("lowerRoman", "i", "PAGE"))
        self.assertIsNotNone(module.validate_main_page_number("lowerRoman", "i", "PAGE"))
        self.assertIsNone(module.validate_main_page_number("decimal", "第1页 共35页", "PAGE NUMPAGES"))

    def test_batch_folder_names_preserve_chinese_stems_and_dedupe(self):
        module_path = ROOT / "scripts" / "run_batch_matrix.py"
        spec = importlib.util.spec_from_file_location("run_batch_matrix", module_path)
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)

        used: set[str] = set()
        first = module.safe_folder_name(Path("成都市育龄人口生育意愿及其影响因素研究吴文雯.docx"), used)
        second = module.safe_folder_name(Path("成都市育龄人口生育意愿及其影响因素研究吴文雯.docx"), used)

        self.assertEqual(first, "成都市育龄人口生育意愿及其影响因素研究吴文雯")
        self.assertEqual(second, "成都市育龄人口生育意愿及其影响因素研究吴文雯_2")


if __name__ == "__main__":
    unittest.main()
