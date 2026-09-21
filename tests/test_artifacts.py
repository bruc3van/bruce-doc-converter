"""Assertions against actual saved Office/PDF artifacts, not only mocks."""
import json
import os
import shutil
import subprocess
import tempfile
import unittest
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from zipfile import ZipFile

from bruce_doc_converter.converter import convert_document

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / 'tests' / 'fixtures'
SCRIPT = ROOT / 'bruce_doc_converter' / 'md_to_docx' / 'index.js'
W = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


class ArtifactTests(unittest.TestCase):
    def test_office_and_pdf_fixture_content(self):
        with tempfile.TemporaryDirectory() as td:
            for name, expected in [('report.docx', 'Quarterly Report'), ('budget.xlsx', '=SUM(B2:B3)'),
                                   ('slides.pptx', 'Launch checklist'), ('report.pdf', 'PDF fixture text')]:
                with self.subTest(name=name):
                    result = convert_document(FIXTURES / name, output_dir=td)
                    self.assertTrue(result['success'], result)
                    self.assertIn(expected, result['markdown_content'])
            result = convert_document(FIXTURES / 'budget.xlsx', output_dir=td)
            # B4 has no cached result; C4 has a known cached result of 30.
            self.assertIn('30', result['markdown_content'])
            self.assertEqual(['B4'], [d['cell'] for d in result['diagnostics']])

    def require_node(self):
        node = shutil.which('node')
        modules = SCRIPT.parent / 'node_modules'
        if node and (modules / 'markdown-it').is_dir():
            return node
        if os.environ.get('BDC_REQUIRE_NODE_TESTS') == '1':
            self.fail('Node.js and local npm dependencies are required')
        self.skipTest('run npm ci --ignore-scripts in bruce_doc_converter/md_to_docx')

    def test_saved_docx_semantics_and_concurrent_names(self):
        node = self.require_node()
        with tempfile.TemporaryDirectory() as td:
            def convert(_):
                completed = subprocess.run([node, str(SCRIPT), str(FIXTURES / 'semantic.md'), td],
                                           capture_output=True, text=True, encoding='utf-8', timeout=60)
                self.assertEqual(0, completed.returncode, completed.stderr)
                return json.loads(completed.stdout)
            with ThreadPoolExecutor(max_workers=3) as pool:
                results = list(pool.map(convert, range(3)))
            self.assertEqual(3, len({r['output_path'] for r in results}))
            for result in results:
                with ZipFile(result['output_path']) as archive:
                    document = ET.fromstring(archive.read('word/document.xml'))
                    texts = ''.join(document.itertext())
                    self.assertIn('foo_bar_baz', texts)
                    self.assertIn('*literal*', texts)
                    self.assertIn('const value = "<tag>";', texts)
                    self.assertIn('a|b', texts)
                    self.assertEqual(1, len(document.findall('.//w:tbl', W)))
                    self.assertGreaterEqual(len(document.findall('.//w:numPr', W)), 3)
                    self.assertTrue(document.findall('.//w:hyperlink//w:b', W))
                    self.assertIn('https://example.com/a_(b)', archive.read('word/_rels/document.xml.rels').decode())
                    paragraphs = [''.join(p.itertext()) for p in document.findall('.//w:p', W)]
                    self.assertIn('Continuation paragraph.', paragraphs)

    def test_strict_markdown_failure_writes_no_docx(self):
        node = self.require_node()
        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / 'remote.md'
            source.write_text('![remote](https://example.com/image.png)', encoding='utf-8')
            env = {**os.environ, 'BRUCE_DOC_CONVERTER_STRICT': '1'}
            completed = subprocess.run([node, str(SCRIPT), str(source), td], env=env,
                                       capture_output=True, text=True, encoding='utf-8', timeout=60)
            self.assertEqual(1, completed.returncode)
            self.assertEqual('CONTENT_INCOMPLETE', json.loads(completed.stdout)['error_code'])
            self.assertFalse(list(Path(td).glob('*.docx')))
