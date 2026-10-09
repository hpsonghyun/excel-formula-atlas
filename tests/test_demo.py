"""Check that the public map agrees with the editable synthetic workbook."""
from pathlib import Path
import json
import re
import unittest
from zipfile import ZipFile
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def normalized(formula):
    return re.sub(r"'([A-Za-z][A-Za-z0-9_]*)'!", r'\1!', formula.lstrip('='))

def workbook_cells(path):
    result = {}
    with ZipFile(path) as z:
        names = [x.attrib['name'] for x in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets', NS)]
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            shared = [''.join(x.itertext()) for x in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        for i, name in enumerate(names, 1):
            cells = {}
            for c in ET.fromstring(z.read(f'xl/worksheets/sheet{i}.xml')).findall('.//s:sheetData/s:row/s:c', NS):
                formula = c.find('s:f', NS)
                value = c.find('s:v', NS)
                v = value.text if value is not None else None
                if c.attrib.get('t') == 's' and v is not None:
                    v = shared[int(v)]
                elif c.attrib.get('t') == 'inlineStr':
                    v = ''.join(c.find('s:is', NS).itertext())
                cells[c.attrib['r']] = {'formula': formula.text if formula is not None else '', 'value': v, 'type': c.attrib.get('t')}
            result[name] = cells
    return result

class SyntheticDemoTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.spec = json.loads((ROOT / 'examples/model-spec.json').read_text(encoding='utf-8'))
        cls.data = json.loads((ROOT / 'examples/map-data.json').read_text(encoding='utf-8'))
        cls.cells = workbook_cells(ROOT / 'examples/OperationsModel.xlsx')

    def test_synthetic_and_english(self):
        self.assertTrue(self.spec['synthetic'])
        self.assertTrue(self.data['example'])
        self.assertEqual(list(self.cells), ['Summary', 'Assumptions', 'Sales', 'Costing', 'Purchasing', 'Inventory', 'Orders', 'Catalog'])
        self.assertIsNone(re.search('[가-힣]', json.dumps(self.data, ensure_ascii=False)))

    def test_workbook_is_plain_xlsx(self):
        for name in ['examples/OperationsModel.xlsx', 'docs/downloads/OperationsModel.xlsx']:
            with ZipFile(ROOT / name) as z:
                self.assertIsNone(z.testzip())
                self.assertFalse(any('vbaproject' in n.lower() or 'externallinks' in n.lower() or 'embeddings' in n.lower() for n in z.namelist()))
                for n in z.namelist():
                    if n.endswith('.rels'):
                        self.assertNotIn(b'TargetMode="External"', z.read(n))

    def test_example_download_is_identical(self):
        self.assertEqual((ROOT / 'examples/OperationsModel.xlsx').read_bytes(), (ROOT / 'docs/downloads/OperationsModel.xlsx').read_bytes())

    def test_all_formula_results(self):
        self.assertEqual(len(self.spec['outputs']), 197)
        for item in self.spec['outputs']:
            c = self.cells[item['sheet']][item['address']]
            self.assertEqual(normalized(c['formula']), normalized(item['formula']))
            expected = item['expected']
            if isinstance(expected, (int, float)):
                self.assertAlmostEqual(float(c['value']), expected, delta=max(.00001, abs(expected)*1e-10), msg=f"{item['sheet']}!{item['address']}")
            else:
                self.assertEqual(c['value'], str(expected))

    def test_snapshot_formulas_match_workbook(self):
        found = set()
        for sheet in self.data['worksheets']:
            for cell in sheet['cells']:
                if cell.get('formula'):
                    actual = self.cells[sheet['name']][cell['address']]['formula']
                    self.assertEqual(normalized(cell['formula']), normalized(actual))
                    found.add((sheet['name'], cell['address']))
        self.assertEqual(len(found), 197)

    def test_every_graph_endpoint_is_captured(self):
        nodes = {n['id']: n for n in self.data['nodes']}
        self.assertEqual(len(nodes), 334)
        self.assertEqual(len(self.data['edges']), 492)
        for e in self.data['edges']:
            self.assertIn(e['s'], nodes)
            self.assertIn(e['t'], nodes)
        for n in nodes.values():
            if n.get('formula'):
                self.assertEqual(normalized(n['formula']), normalized(self.cells[n['sheet']][n['address']]['formula']))

    def test_expected_cash_formula(self):
        headline = self.spec['headline']
        self.assertEqual((headline['sheet'], headline['address']), ('Summary', 'P30'))
        self.assertEqual(normalized(self.cells['Summary']['P30']['formula']), 'SUM(P23:P25)-SUM(P26:P29)')
        self.assertAlmostEqual(float(self.cells['Summary']['P30']['value']), headline['expected'])

if __name__ == '__main__':
    unittest.main()
