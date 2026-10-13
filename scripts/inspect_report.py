#!/usr/bin/env python3
"""Read-only DOCX inventory. Requires lxml (included with python-docx)."""
import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'v': 'urn:schemas-microsoft-com:vml'}
W = '{' + NS['w'] + '}'
PART = re.compile(r'word/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml$')
FILENAME_TOKEN = re.compile(r'学号[（(]最后\s*2\s*位[）)]|学号最后\s*2\s*位|专业班级|学号|姓名|班级')
PROMPT = re.compile(r'（写）|\(写\)|（截图）|\(截图\)|截图与写')
SHOT = re.compile(r'（截图）|\(截图\)|截图与写')


def resolve_input(value):
    """Resolve the supplied path without machine-specific redirects."""
    return Path(value).expanduser().resolve(strict=True)


def read_package(path):
    raw = Path(path).read_bytes()
    import io
    with ZipFile(io.BytesIO(raw)) as z:
        entries = [(i, z.read(i.filename)) for i in z.infolist()]
    parser = ET.XMLParser(resolve_entities=False, no_network=True)
    trees = {i.filename: ET.fromstring(data, parser) for i, data in entries if PART.fullmatch(i.filename)}
    if 'word/document.xml' not in trees:
        raise ValueError('Not a DOCX: missing word/document.xml')
    return raw, entries, trees


def region_nodes(trees):
    nodes = {}
    for part, tree in trees.items():
        for kind in ('tc', 'p'):
            for index, node in enumerate(tree.iter(W + kind)):
                nodes[f'{part}#{kind}{index}'] = node
    return nodes


def text_of(node):
    paragraphs = [node] if node.tag == W + 'p' else list(node.iter(W + 'p'))
    lines = []
    for p in paragraphs:
        bits = []
        for leaf in p.iter():
            if leaf.tag not in (W + 't', W + 'br', W + 'tab'):
                continue
            owner = next((a for a in leaf.iterancestors() if a.tag == W + 'p'), None)
            if owner is p:
                bits.append(leaf.text or '' if leaf.tag == W + 't' else ('\n' if leaf.tag == W + 'br' else '\t'))
        lines.append(''.join(bits))
    return '\n'.join(lines)


def has_image(node):
    return bool(node.xpath('.//w:drawing | .//w:pict | .//a:blip | .//v:imagedata', namespaces=NS))


def label_of(node):
    if node.tag == W + 'tc':
        prev = node.getprevious()
        if prev is not None and prev.tag == W + 'tc':
            return text_of(prev).strip()[:100]
    return text_of(node).split('：', 1)[0].split(':', 1)[0][:60]


def inventory(path, trees, raw, entries):
    nodes = region_nodes(trees)
    ids = {node: key for key, node in nodes.items()}
    regions, shots, tables = [], [], []
    for key, node in nodes.items():
        # Expose paragraph IDs for precise edits, including nested table cells.
        text = text_of(node)
        label = label_of(node)
        image = has_image(node)
        colors = node.xpath('.//w:color/@w:val', namespaces=NS)
        hints = {'highlight': node.xpath('.//w:highlight/@w:val', namespaces=NS),
                 'shading': node.xpath('.//w:shd/@w:fill', namespaces=NS),
                 'color': [c for c in colors if c.lower() not in ('auto', '000000')],
                 'prompts': PROMPT.findall(text)}
        hints = {k: list(dict.fromkeys(v)) for k, v in hints.items() if v}
        teacher = '指导教师' in label
        record = {'id': key, 'label': label, 'text': text, 'blank': not text.strip(),
                  'has_image': image}
        parent_cell = next((a for a in node.iterancestors() if a.tag == W + 'tc'), None)
        if parent_cell is not None:
            record['parent_cell'] = ids[parent_cell]
        if node.tag == W + 'tc':
            record['paragraph_ids'] = [ids[p] for p in node.findall(W + 'p')]
        if hints:
            record['hints'] = hints
        if teacher:
            record['teacher_region_has_image'] = image
        if node.xpath('.//w:tbl | .//w:txbxContent | .//w:fldChar', namespaces=NS):
            record['complex'] = True
        if record['blank'] and not image:
            record['candidate_fill'] = True  # Semantic decision remains with the agent.
        regions.append(record)
        # Markers belong to paragraphs; do not count them again in ancestor cells.
        for index, match in enumerate(SHOT.finditer(text) if node.tag == W + 'p' else []):
            shots.append({'id': key + f'@shot{index}', 'target': key, 'marker': match.group(), 'status': 'unverified'})
    for part, tree in trees.items():
        for index, table in enumerate(tree.iter(W + 'tbl')):
            rows = [[ids[c] for c in row.findall(W + 'tc')] for row in table.findall(W + 'tr')]
            parent_cell = next((a for a in table.iterancestors() if a.tag == W + 'tc'), None)
            tables.append({'id': f'{part}#tbl{index}', 'rows': rows,
                           'parent_cell': ids.get(parent_cell)})
    fields = [{k: r[k] for k in ('id', 'label', 'text', 'has_image')} for r in regions
              if re.search(r'姓名|学号|班级|专业|学院|学校|学生|指导教师|教师评语|日期|评分', r['label'])]
    return {'schema': 2, 'source': str(path), 'source_sha256': hashlib.sha256(raw).hexdigest(),
            'filename': path.name, 'filename_placeholders': FILENAME_TOKEN.findall(path.name),
            'filename_contains_plus': '+' in path.name,
            'image_count': sum(i.filename.startswith('word/media/') for i, _ in entries),
            'fields': fields, 'tables': tables, 'regions': regions, 'screenshots': shots,
            'note': 'blank/color are candidates, not proof of a required field; image authenticity requires agent QA'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx', nargs='+', help='Report and optional DOCX manuals in one call')
    parser.add_argument('--output', help='Cache complete JSON here; stdout remains compact')
    parser.add_argument('--text-limit', type=int, default=280, help='Per-region stdout text limit; 0 = full')
    parser.add_argument('--debug', action='store_true', help='Include package XML in JSON (large)')
    args = parser.parse_args()
    records = []
    for value in args.docx:
        path = resolve_input(value)
        raw, entries, trees = read_package(path)
        record = inventory(path, trees, raw, entries)
        if args.debug:
            record['xml'] = {part: ET.tostring(tree, encoding='unicode') for part, tree in trees.items()}
        records.append(record)
    result = records[0] if len(records) == 1 else {'documents': records}
    if args.output:
        Path(args.output).write_text(json.dumps(result, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    for record in records:
        for region in record['regions'] + record['fields']:
            if args.text_limit and len(region['text']) > args.text_limit:
                region['chars'] = len(region['text'])
                region['text'] = region['text'][:args.text_limit]
                region['truncated'] = True
    print(json.dumps(result, ensure_ascii=False, separators=(',', ':')))


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, KeyError, ET.XMLSyntaxError) as exc:
        print(json.dumps({'error': str(exc)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)
