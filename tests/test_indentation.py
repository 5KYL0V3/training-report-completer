"""Only report reflection/comment prose receives character-based indentation."""
import unittest
import copy
from lxml import etree as ET
import test_report as fixtures
from inspect_report import W, NS, read_package, region_nodes, text_of
from apply_report import indent_report_paragraphs

class IndentationTests(unittest.TestCase):
    def test_scoped_multiline_and_ordinary_text(self):
        f = fixtures.ReportTests(); f.setUp()
        try:
            nested = next(t for t in f.report['tables'] if t['parent_cell'])['rows'][0][1]
            comment, score = f.cell('教师评语'), f.cell('评分')
            result = f.run_patch([
                {'target': nested, 'field': 'reflection', 'text': '心得第一段。\n心得第二段。', 'font_size_pt': 12},
                {'target': comment, 'field': 'teacher_comment', 'text': '评语第一段。\n评语第二段。'},
                {'target': score, 'field': 'score', 'text': '良'}])
            _, _, trees = read_package(result['output'])
            nodes = region_nodes(trees)
            for target in (nested, comment):
                cell = nodes[target]
                ps = cell.findall(W+'p')
                self.assertEqual(len(ps),2)
                self.assertEqual(cell.xpath('./w:p/w:pPr/w:ind/@w:firstLineChars',namespaces=NS),['200','200'])
                self.assertFalse(cell.xpath('.//w:br',namespaces=NS))
                self.assertFalse(any(text_of(p).startswith((' ', '\t', '\u3000')) for p in ps))
            self.assertFalse(nodes[score].xpath('.//w:ind',namespaces=NS))
            for target in (nested,comment):
                before=f.nodes[target].find(W+'tcPr'); after=nodes[target].find(W+'tcPr')
                self.assertEqual(ET.tostring(before),ET.tostring(after))
                self.assertEqual(ET.tostring(f.nodes[target].getprevious()),ET.tostring(nodes[target].getprevious()))
        finally: f.doCleanups()

    def test_format_only_preserves_content_and_other_properties(self):
        p=ET.fromstring(('<w:p xmlns:w="'+NS['w']+'"><w:pPr><w:spacing w:line="240"/><w:ind w:left="120" w:right="60" w:hanging="240" w:hangingChars="100" w:firstLine="300"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>已有真实评语，内容保留。</w:t></w:r></w:p>').encode())
        text=text_of(p); runs=ET.tostring(p.find(W+'r')); spacing=ET.tostring(p.find(W+'pPr').find(W+'spacing'))
        indent_report_paragraphs(p)
        ind=p.find(W+'pPr').find(W+'ind')
        self.assertEqual(ind.get(W+'firstLineChars'),'200')
        self.assertEqual(ind.get(W+'left'),'120'); self.assertEqual(ind.get(W+'right'),'60')
        for name in ('hanging','hangingChars','firstLine'): self.assertNotIn(W+name,ind.attrib)
        self.assertEqual(text_of(p),text); self.assertEqual(ET.tostring(p.find(W+'r')),runs)
        self.assertEqual(ET.tostring(p.find(W+'pPr').find(W+'spacing')),spacing)
        once=ET.tostring(p); indent_report_paragraphs(p); self.assertEqual(ET.tostring(p),once)

    def test_format_cli_scope_and_signature_protection(self):
        f=fixtures.ReportTests(); f.setUp()
        try:
            comment=f.cell('教师评语')
            op={'target':comment,'field':'teacher_comment','op':'format_paragraphs'}
            result=f.run_patch([op,{'target':f.paragraph(f.cell('指导教师')),'field':'teacher_comment','op':'format_paragraphs'}])
            self.assertEqual(result['skipped'][0]['reason'],'teacher_signature_image_preserved')
            _,_,trees=read_package(result['output'])
            self.assertEqual(text_of(trees['word/document.xml']),text_of(f.trees['word/document.xml']))
            op['field']='score'
            err=f.run_patch([op],output=f.root/'blocked.docx',success=False)
            self.assertIn('limited to reflection',err['error'])
        finally: f.doCleanups()

    def test_skip_lists_and_pictures(self):
        for inner in ('<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>编号</w:t></w:r>', '<w:r><w:t>图片</w:t><w:drawing/></w:r>', '<w:r><w:t> </w:t></w:r>'):
            p=ET.fromstring(('<w:p xmlns:w="'+NS['w']+'">'+inner+'</w:p>').encode())
            before=ET.tostring(p); indent_report_paragraphs(p); self.assertEqual(ET.tostring(p),before)

if __name__ == '__main__': unittest.main()
