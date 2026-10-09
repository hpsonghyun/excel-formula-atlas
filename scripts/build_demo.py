"""Build the standalone public demo from the checked synthetic snapshot."""
from pathlib import Path
import json

ROOT = Path(__file__).resolve().parents[1]

def build():
    data = json.loads((ROOT / 'examples/map-data.json').read_text(encoding='utf-8'))
    assert data.get('example') is True
    template = (ROOT / 'src/viewer/template.html').read_text(encoding='utf-8-sig')
    assert template.count('__FORMULA_MAP_DATA__') == 1
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    for char, escaped in [('<', r'\u003c'), ('>', r'\u003e'), ('&', r'\u0026'), ('\u2028', r'\u2028'), ('\u2029', r'\u2029')]:
        payload = payload.replace(char, escaped)
    target = ROOT / 'docs/demo.html'
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(template.replace('__FORMULA_MAP_DATA__', payload), encoding='utf-8', newline='\n')
    print('Built docs/demo.html from the synthetic example snapshot.')

if __name__ == '__main__':
    build()
