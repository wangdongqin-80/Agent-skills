import importlib.util
import pathlib
import sys
import unittest


MODULE_PATH = pathlib.Path(__file__).with_name("pdf_section_to_docx.py")
SPEC = importlib.util.spec_from_file_location("pdf_section_to_docx", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class PdfSectionToDocxTests(unittest.TestCase):
    def test_sanitize_filename_removes_illegal_characters(self):
        self.assertEqual(
            MODULE.sanitize_filename(' 1.2 \u7ecf\u8425/\u8d22\u52a1:\u6982\u89c8?* "\u5b63\u5ea6" '),
            "1.2 \u7ecf\u8425_\u8d22\u52a1_\u6982\u89c8 \u5b63\u5ea6",
        )

    def test_narrate_table_uses_logical_yearly_sentences(self):
        table = [
            ["\u5e74\u4efd", "1 \u7f8e\u5143\u5151\u6362\u5217\u514b", "1 \u6b27\u5143\u5151\u6362\u5217\u514b"],
            ["2021", "103.54", "122.44"],
            ["2022", "113.15", "118.92"],
            ["2023", "100.62", "108.75"],
        ]

        narration = MODULE.narrate_table(table, title="\u8868 1 2021-2023 \u5e74\u7f8e\u5143\u3001\u6b27\u5143\u5151\u5217\u514b\u7684\u5e73\u5747\u6c47\u7387")

        self.assertIn("\u8868 1 2021-2023 \u5e74\u7f8e\u5143\u3001\u6b27\u5143\u5151\u5217\u514b\u7684\u5e73\u5747\u6c47\u7387", narration)
        self.assertIn("2021 \u5e74 1 \u7f8e\u5143\u5151\u6362 103.54 \u5217\u514b\uff1b1 \u6b27\u5143\u5151\u6362 122.44 \u5217\u514b\u3002", narration)
        self.assertIn("2022 \u5e74 1 \u7f8e\u5143\u5151\u6362 113.15 \u5217\u514b\uff1b1 \u6b27\u5143\u5151\u6362 118.92 \u5217\u514b\u3002", narration)
        self.assertIn("2023 \u5e74 1 \u7f8e\u5143\u5151\u6362 100.62 \u5217\u514b\uff1b1 \u6b27\u5143\u5151\u6362 108.75 \u5217\u514b\u3002", narration)

    def test_filter_redundant_guide_lines_keeps_only_first_page_header(self):
        blocks = [
            MODULE.Block("paragraph", "\u4e2d\u56fd\u5c45\u6c11\u8d74\u963f\u5c14\u5df4\u5c3c\u4e9a\u6295\u8d44\u7a0e\u6536\u6307\u5357"),
            MODULE.Block("heading", "\u7b2c\u4e00\u7ae0", level=1),
            MODULE.Block("paragraph", "\u6b63\u6587\u7b2c\u4e00\u6bb5"),
            MODULE.Block("paragraph", "\u4e2d\u56fd\u5c45\u6c11\u8d74\u963f\u5c14\u5df4\u5c3c\u4e9a\u6295\u8d44\u7a0e\u6536\u6307\u5357"),
            MODULE.Block("paragraph", "\u6b63\u6587\u7b2c\u4e8c\u6bb5"),
        ]

        filtered = MODULE.filter_redundant_guide_lines(blocks)

        self.assertEqual(
            [block.text for block in filtered],
            [
                "\u4e2d\u56fd\u5c45\u6c11\u8d74\u963f\u5c14\u5df4\u5c3c\u4e9a\u6295\u8d44\u7a0e\u6536\u6307\u5357",
                "\u7b2c\u4e00\u7ae0",
                "\u6b63\u6587\u7b2c\u4e00\u6bb5",
                "\u6b63\u6587\u7b2c\u4e8c\u6bb5",
            ],
        )

    def test_find_guide_title_combines_split_title_lines(self):
        blocks = [
            MODULE.Block("paragraph", "\u4e2d\u56fd\u5c45\u6c11\u8d74\u963f\u5c14\u5df4\u5c3c\u4e9a"),
            MODULE.Block("paragraph", "\u6295\u8d44\u7a0e\u6536\u6307\u5357"),
            MODULE.Block("paragraph", "\u56fd\u5bb6\u7a0e\u52a1\u603b\u5c40\u8bfe\u9898\u7ec4"),
        ]

        self.assertEqual(
            MODULE.find_guide_title(blocks),
            "\u4e2d\u56fd\u5c45\u6c11\u8d74\u963f\u5c14\u5df4\u5c3c\u4e9a\u6295\u8d44\u7a0e\u6536\u6307\u5357",
        )

    def test_split_sections_carries_heading_hierarchy(self):
        blocks = [
            MODULE.Block("heading", "\u603b\u8bba", level=1),
            MODULE.Block("paragraph", "\u603b\u8bba\u6bb5\u843d"),
            MODULE.Block("heading", "\u7ecf\u8425\u60c5\u51b5", level=2),
            MODULE.Block("paragraph", "\u7ecf\u8425\u60c5\u51b5\u6bb5\u843d"),
            MODULE.Block("heading", "\u6536\u5165\u7ed3\u6784", level=3),
            MODULE.Block("paragraph", "\u6536\u5165\u7ed3\u6784\u6bb5\u843d"),
            MODULE.Block("heading", "\u98ce\u9669\u63d0\u793a", level=2),
            MODULE.Block("paragraph", "\u98ce\u9669\u63d0\u793a\u6bb5\u843d"),
        ]

        sections = MODULE.split_sections(blocks)

        self.assertEqual([section.title for section in sections], ["\u603b\u8bba", "\u7ecf\u8425\u60c5\u51b5", "\u6536\u5165\u7ed3\u6784", "\u98ce\u9669\u63d0\u793a"])
        self.assertEqual([heading.text for heading in sections[2].heading_path], ["\u603b\u8bba", "\u7ecf\u8425\u60c5\u51b5", "\u6536\u5165\u7ed3\u6784"])
        self.assertEqual([block.text for block in sections[1].content], ["\u7ecf\u8425\u60c5\u51b5\u6bb5\u843d"])
        self.assertEqual([heading.text for heading in sections[3].heading_path], ["\u603b\u8bba", "\u98ce\u9669\u63d0\u793a"])

    def test_split_sections_skips_leading_body_without_heading(self):
        blocks = [
            MODULE.Block("paragraph", "\u5c01\u9762\u8bf4\u660e"),
            MODULE.Block("heading", "\u7b2c\u4e00\u8282", level=1),
            MODULE.Block("paragraph", "\u6b63\u6587"),
        ]

        sections = MODULE.split_sections(blocks)

        self.assertEqual(len(sections), 1)
        self.assertEqual(sections[0].title, "\u7b2c\u4e00\u8282")
        self.assertEqual([block.text for block in sections[0].content], ["\u6b63\u6587"])

    def test_drop_name_only_subheading_under_references(self):
        blocks = [
            MODULE.Block("heading", "\u53c2\u8003\u6587\u732e", level=1),
            MODULE.Block("heading", "\u8c2d\u6620\u8377\u674e\u73c2", level=3),
            MODULE.Block("paragraph", "58"),
        ]

        filtered = MODULE.drop_name_only_headings(blocks)
        sections = MODULE.split_sections(filtered)

        self.assertEqual([section.title for section in sections], ["\u53c2\u8003\u6587\u732e"])
        self.assertEqual([block.text for block in sections[0].content], ["\u8c2d\u6620\u8377\u674e\u73c2", "58"])

    def test_merge_multiline_chapter_heading(self):
        blocks = [
            MODULE.Block("heading", "\u7b2c\u516d\u7ae0 \u5728\u963f\u5c14\u5df4\u5c3c\u4e9a\u6295\u8d44\u53ef\u80fd\u5b58\u5728\u7684", level=1),
            MODULE.Block("heading", "\u7a0e\u6536\u98ce\u9669", level=1),
            MODULE.Block("heading", "6.1 \u4fe1\u606f\u62a5\u544a\u98ce\u9669", level=2),
            MODULE.Block("paragraph", "\u6b63\u6587"),
        ]

        merged = MODULE.merge_multiline_headings(blocks)
        sections = MODULE.split_sections(merged)

        self.assertEqual(sections[0].title, "\u7b2c\u516d\u7ae0 \u5728\u963f\u5c14\u5df4\u5c3c\u4e9a\u6295\u8d44\u53ef\u80fd\u5b58\u5728\u7684 \u7a0e\u6536\u98ce\u9669")
        self.assertEqual(sections[1].title, "6.1 \u4fe1\u606f\u62a5\u544a\u98ce\u9669")
        self.assertEqual(
            [heading.text for heading in sections[1].heading_path],
            [
                "\u7b2c\u516d\u7ae0 \u5728\u963f\u5c14\u5df4\u5c3c\u4e9a\u6295\u8d44\u53ef\u80fd\u5b58\u5728\u7684 \u7a0e\u6536\u98ce\u9669",
                "6.1 \u4fe1\u606f\u62a5\u544a\u98ce\u9669",
            ],
        )

    def test_merge_multiline_chapter_heading_allows_mismatched_levels(self):
        blocks = [
            MODULE.Block("heading", "\u7b2c\u4e00\u7ae0", level=1),
            MODULE.Block("heading", "\u53f0\u6e7e\u5730\u533a\u7ecf\u6d4e\u6982\u51b5", level=2),
        ]

        merged = MODULE.merge_multiline_headings(blocks)

        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0].text, "\u7b2c\u4e00\u7ae0 \u53f0\u6e7e\u5730\u533a\u7ecf\u6d4e\u6982\u51b5")
        self.assertEqual(merged[0].level, 1)

    def test_merge_multiline_short_special_heading(self):
        blocks = [
            MODULE.Block("heading", "\u524d", level=1),
            MODULE.Block("heading", "\u8a00", level=1),
            MODULE.Block("paragraph", "\u6b63\u6587"),
        ]

        merged = MODULE.merge_multiline_headings(blocks)

        self.assertEqual(len(merged), 2)
        self.assertEqual(merged[0].text, "\u524d\u8a00")
        self.assertEqual(merged[0].level, 1)

    def test_merge_multiline_appendix_heading(self):
        blocks = [
            MODULE.Block("heading", "\u9644", level=1),
            MODULE.Block("heading", "\u5f55", level=1),
            MODULE.Block("paragraph", "\u6b63\u6587"),
        ]

        merged = MODULE.merge_multiline_headings(blocks)

        self.assertEqual(len(merged), 2)
        self.assertEqual(merged[0].text, "\u9644\u5f55")
        self.assertEqual(merged[0].level, 1)

    def test_merge_multiline_appendix_letter_heading(self):
        blocks = [
            MODULE.Block("heading", "\u9644\u5f55 A \u963f\u5c14\u53ca\u5229\u4e9a\u653f\u5e9c\u90e8\u95e8\u548c\u76f8\u5173\u673a\u6784\u4e00", level=1),
            MODULE.Block("heading", "\u89c8\u8868", level=1),
        ]

        merged = MODULE.merge_multiline_headings(blocks)

        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0].text, "\u9644\u5f55 A \u963f\u5c14\u53ca\u5229\u4e9a\u653f\u5e9c\u90e8\u95e8\u548c\u76f8\u5173\u673a\u6784\u4e00 \u89c8\u8868")

    def test_merge_multiline_appendix_chinese_number_heading(self):
        blocks = [
            MODULE.Block("heading", "\u9644\u5f55\u4e8c \u963f\u585e\u62dc\u7586\u7b7e\u8ba2\u7a0e\u6536\u6761\u7ea6\u4e00", level=1),
            MODULE.Block("heading", "\u89c8\u8868", level=1),
        ]

        merged = MODULE.merge_multiline_headings(blocks)

        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0].text, "\u9644\u5f55\u4e8c \u963f\u585e\u62dc\u7586\u7b7e\u8ba2\u7a0e\u6536\u6761\u7ea6\u4e00 \u89c8\u8868")

    def test_merge_appendix_heading_with_short_paragraph_fragment(self):
        blocks = [
            MODULE.Block("heading", "\u9644\u5f55 D \u5728\u963f\u5c14\u53ca\u5229\u4e9a\u6295\u8d44\u7684\u4e3b\u8981\u4e2d\u8d44\u4f01", level=1),
            MODULE.Block("paragraph", "\u4e1a"),
            MODULE.Block("paragraph", "\u5185\u5bb9"),
        ]

        merged = MODULE.merge_heading_fragments(blocks)

        self.assertEqual(merged[0].text, "\u9644\u5f55 D \u5728\u963f\u5c14\u53ca\u5229\u4e9a\u6295\u8d44\u7684\u4e3b\u8981\u4e2d\u8d44\u4f01\u4e1a")
        self.assertEqual(merged[1].text, "\u5185\u5bb9")

    def test_promote_appendix_items_under_appendix_heading(self):
        blocks = [
            MODULE.Block("heading", "\u9644\u5f55", level=1),
            MODULE.Block("paragraph", "\u4e00\u3001\u4e24\u5cb8\u7a0e\u6cd5\u672f\u8bed\u5bf9\u6bd4\u8868"),
            MODULE.Block("paragraph", "\u5185\u5bb9"),
        ]

        promoted = MODULE.promote_appendix_items(blocks)
        sections = MODULE.split_sections(promoted)

        self.assertEqual([section.title for section in sections], ["\u9644\u5f55", "\u4e00\u3001\u4e24\u5cb8\u7a0e\u6cd5\u672f\u8bed\u5bf9\u6bd4\u8868"])
        self.assertEqual([heading.text for heading in sections[1].heading_path], ["\u9644\u5f55", "\u4e00\u3001\u4e24\u5cb8\u7a0e\u6cd5\u672f\u8bed\u5bf9\u6bd4\u8868"])

    def test_promote_appendix_items_relevels_existing_heading(self):
        blocks = [
            MODULE.Block("heading", "\u9644\u5f55", level=1),
            MODULE.Block("heading", "\u4e00\u3001\u4e24\u5cb8\u7a0e\u6cd5\u672f\u8bed\u5bf9\u6bd4\u8868", level=1),
            MODULE.Block("paragraph", "\u5185\u5bb9"),
        ]

        promoted = MODULE.promote_appendix_items(blocks)
        sections = MODULE.split_sections(promoted)

        self.assertEqual(promoted[1].level, 2)
        self.assertEqual([heading.text for heading in sections[1].heading_path], ["\u9644\u5f55", "\u4e00\u3001\u4e24\u5cb8\u7a0e\u6cd5\u672f\u8bed\u5bf9\u6bd4\u8868"])

    def test_looks_like_heading_rejects_toc_entry(self):
        level = MODULE.looks_like_heading(
            "1.1 \u8fd1\u5e74\u7ecf\u6d4e\u53d1\u5c55\u60c5\u51b5............................................................1",
            14.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_chapter_summary_sentence(self):
        level = MODULE.looks_like_heading(
            "\u7b2c\u4e00\u7ae0\u4e3b\u8981\u4ecb\u7ecd\u4e86\u963f\u5c14\u5df4\u5c3c\u4e9a\u7684\u7ecf\u6d4e\u6982\u51b5\uff0c\u5305\u62ec\u8fd1\u5e74\u7ecf\u6d4e\u53d1\u5c55\u60c5\u51b5",
            14.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_long_wrapped_body_line(self):
        level = MODULE.looks_like_heading(
            "\u5728\u8fc7\u53bb\u7684\u6570\u5e74\u91cc\uff0c\u963f\u5c14\u5df4\u5c3c\u4e9a\u5728\u5386\u7ecf\u91cd\u5927\u653f\u6cbb\u53d8\u9769\u7684\u540c\u65f6\u4ea7\u751f\u4e86",
            14.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_parenthesized_enumeration(self):
        level = MODULE.looks_like_heading(
            "\uff085\uff09\u7ecf\u6d4e\u589e\u957f\u60c5\u51b5",
            14.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_date_led_body_sentence(self):
        level = MODULE.looks_like_heading(
            "2021 \u5e74 1 \u6708 1 \u65e5\u8d77\uff0c\u7eb3\u7a0e\u4e49\u52a1\u4eba\u4e3a\u5728\u963f\u5c14\u5df4\u5c3c\u4e9a\u5e74\u6536\u5165\u8d85\u8fc7 800",
            16.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_month_day_sentence_fragment(self):
        level = MODULE.looks_like_heading(
            "12 \u6708 31 \u65e5\u524d\u6240\u6301\u6709\u80a1\u4efd\u6709\u9650\u516c\u53f8\u80a1\u7968\u6216\u516c\u53f8\u503a\uff0c\u5176\u4ea4\u6613\u6240\u5f97\u989d\u4e2d\uff0c",
            18.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_numbered_measurement_fragment(self):
        level = MODULE.looks_like_heading(
            "601.4 \u4ebf\u7f8e\u5143\uff0c\u81ea\u53f0\u6e7e\u5730\u533a\u8fdb\u53e3\u4e3a2,006.6 \u4ebf\u7f8e\u5143\uff0c\u5bf9\u53f0\u8d38\u6613\u9006\u5dee",
            15.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_integer_led_body_fragment(self):
        level = MODULE.looks_like_heading(
            "10 \u5e74\u4e14\u672a\u4e0a\u5e02\uff08\u67dc\uff09\u7684\u7814\u53d1\u5236\u9020\u516c\u53f8\u6216\u6210\u7acb\u672a",
            16.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_keeps_decimal_section_heading(self):
        level = MODULE.looks_like_heading(
            "2.3.1.3 \u7a0e\u7387",
            16.0,
            12.0,
        )
        self.assertEqual(level, 4)

    def test_looks_like_heading_rejects_bracketed_note_line(self):
        level = MODULE.looks_like_heading(
            "\u3010\u6309\u53f0\u6e7e\u5730\u533a\u672f\u8bed\u62fc\u97f3\u6392\u5e8f\u3011",
            16.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_formula_like_fragment(self):
        level = MODULE.looks_like_heading(
            "1 + \u7a0e\u7387",
            16.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_keeps_appendix_letter_heading(self):
        level = MODULE.looks_like_heading(
            "\u9644\u5f55 A \u963f\u5c14\u53ca\u5229\u4e9a\u653f\u5e9c\u90e8\u95e8\u548c\u76f8\u5173\u673a\u6784\u4e00\u89c8\u8868",
            16.0,
            12.0,
        )
        self.assertEqual(level, 1)

    def test_looks_like_heading_keeps_appendix_chinese_number_heading(self):
        level = MODULE.looks_like_heading(
            "\u9644\u5f55\u4e8c \u963f\u585e\u62dc\u7586\u7b7e\u8ba2\u7a0e\u6536\u6761\u7ea6\u4e00\u89c8\u8868",
            16.0,
            12.0,
        )
        self.assertEqual(level, 1)

    def test_looks_like_heading_rejects_appendix_toc_entry(self):
        level = MODULE.looks_like_heading(
            "\u9644\u5f55 B \u963f\u5c14\u53ca\u5229\u4e9a\u7b7e\u8ba2\u7a0e\u6536\u6761\u7ea6\u4e00\u89c8\u8868.....................................................166",
            16.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_circled_enumeration(self):
        level = MODULE.looks_like_heading(
            "\u2461\u4e2a\u4eba",
            16.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_latin_enumeration(self):
        level = MODULE.looks_like_heading(
            "A.\u5229\u7528\u5b58\u6b3e\u8d26\u6237\u9000\u7a0e",
            16.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_short_unumbered_body_item(self):
        level = MODULE.looks_like_heading(
            "\u2462\u80a1\u606f",
            18.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_orphan_short_fragment(self):
        level = MODULE.looks_like_heading(
            "\u89c8\u8868",
            18.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_rejects_person_name_line(self):
        level = MODULE.looks_like_heading(
            "\u8c2d\u6620\u8377 \u674e\u73c2",
            18.0,
            12.0,
        )
        self.assertIsNone(level)

    def test_looks_like_heading_keeps_short_unnumbered_heading(self):
        level = MODULE.looks_like_heading(
            "\u524d\u8a00",
            18.0,
            12.0,
        )
        self.assertEqual(level, 1)

    def test_looks_like_heading_keeps_table_of_contents_heading(self):
        level = MODULE.looks_like_heading(
            "\u76ee\u5f55",
            18.0,
            12.0,
        )
        self.assertEqual(level, 1)

    def test_looks_like_heading_keeps_appendix_heading(self):
        level = MODULE.looks_like_heading(
            "\u9644\u5f55",
            18.0,
            12.0,
        )
        self.assertEqual(level, 1)


if __name__ == "__main__":
    unittest.main()
