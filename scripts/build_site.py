"""Embed the actual demo and screenshots in the standalone introduction page."""
from pathlib import Path
import base64
import re

ROOT = Path(__file__).resolve().parents[1]

def build():
    page = (ROOT / 'src/site/index.html').read_text(encoding='utf-8')
    marker = '__FORMULA_ATLAS_DEMO_BASE64__'
    assert page.count(marker) == 1
    page = page.replace(marker, base64.b64encode((ROOT / 'docs/demo.html').read_bytes()).decode('ascii'))
    embedded = set()
    def embed(match):
        relative = match.group(1)
        embedded.add(relative)
        content = (ROOT / 'docs' / relative).read_bytes()
        return 'src="data:image/png;base64,' + base64.b64encode(content).decode('ascii') + '"'
    page, count = re.subn(r'src="(assets/(?:overview|selection|loading)\.png)"', embed, page)
    assert count == 4 and len(embedded) == 3, f'Expected 4 images using 3 screenshots, got {count} / {len(embedded)}'
    (ROOT / 'docs/index.html').write_text(page, encoding='utf-8', newline='\n')
    print('Built standalone docs/index.html with its real demo and 3 screenshots in 4 images.')

if __name__ == '__main__':
    build()
