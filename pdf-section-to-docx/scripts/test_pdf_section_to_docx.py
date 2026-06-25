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


if __name__ == "__main__":
    unittest.main()
