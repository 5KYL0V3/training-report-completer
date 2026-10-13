"""Font preference checks using synthetic XML and the existing synthetic DOCX fixture."""
import unittest
from lxml import etree as ET
import test_report as fixtures
from apply_report import put_text
from inspect_report import W, NS, read_package, region_nodes


class FontSizeTests(unittest.TestCase):
    def paragraph(self):
        return ET.fromstring(('<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr><w:p><w:pPr><w:spacing w:line="240"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="楷体"/><w:b/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:r></w:p></w:tc>').encode('utf-8'))

    def test_local_size_preserves_other_properties(self):
        cell = self.paragraph()
        before_cell = ET.tostring(cell.find(W + 'tcPr'))
        before_paragraph = ET.tostring(cell.find(W + 'p').find(W + 'pPr'))
        original_run = ET.tostring(cell.find(W + 'p').find(W + 'r'))
        put_text(cell, '中文 ABC 123', font_size_pt=12)
        run = cell.find(W + 'p').findall(W + 'r')[-1]
        for tag in ('sz', 'szCs'):
            self.assertEqual(run.find(W + 'rPr').find(W + tag).get(W + 'val'), '24')
        self.assertEqual(run.find(W + 'rPr').find(W + 'rFonts').get(W + 'eastAsia'), '楷体')
        self.assertIsNotNone(run.find(W + 'rPr').find(W + 'b'))
        self.assertEqual(ET.tostring(cell.find(W + 'tcPr')), before_cell)
        self.assertEqual(ET.tostring(cell.find(W + 'p').find(W + 'pPr')), before_paragraph)
        self.assertEqual(ET.tostring(cell.find(W + 'p').find(W + 'r')), original_run)

    def test_template_inheritance_and_explicit_exception(self):
        for kwargs, expected in (({}, '21'), ({'font_size_pt': None}, '21'), ({'font_size_pt': 10.5}, '21')):
            cell = self.paragraph()
            put_text(cell, '新正文', **kwargs)
            run = cell.find(W + 'p').findall(W + 'r')[-1]
            self.assertEqual(run.find(W + 'rPr').find(W + 'sz').get(W + 'val'), expected)
        for size in (0, -1, True, '12', 12.2, float('nan'), float('inf')):
            cell = self.paragraph()
            before = ET.tostring(cell)
            with self.assertRaises(ValueError):
                put_text(cell, '不应写入', font_size_pt=size)
            self.assertEqual(ET.tostring(cell), before)

    def test_cli_wires_size_for_text_and_auto_score(self):
        fixture = fixtures.ReportTests()
        fixture.setUp()
        try:
            comment, score = fixture.cell('教师评语'), fixture.cell('评分')
            result = fixture.run_patch([
                {'target': comment, 'field': 'teacher_comment', 'text': '报告过程清晰。', 'font_size_pt': 12},
                {'target': score, 'field': 'score', 'font_size_pt': 12}])
            _, _, trees = read_package(result['output'])
            nodes = region_nodes(trees)
            for target in (comment, score):
                self.assertEqual(nodes[target].xpath('./w:p/w:r[w:t]/w:rPr/w:sz/@w:val', namespaces=NS), ['24'])
            error = fixture.run_patch([{'target': comment, 'op': 'replace', 'old': 'x', 'text': 'y', 'font_size_pt': 12}],
                                      output=fixture.root / 'bad-font.docx', success=False)
            self.assertIn('only for set_text', error['error'])
        finally:
            fixture.doCleanups()


if __name__ == '__main__':
    unittest.main()
