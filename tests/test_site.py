"""Check the public site's real embedded assets and reproducible local build.

These checks use only the standard library and never start a server or browser.
Run: python -m unittest discover -s tests -p test_site.py -v
"""
from collections import Counter
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urlsplit
import base64
import json
import re
import shutil
import subprocess
import sys
import unittest
import uuid


ROOT = Path(__file__).resolve().parents[1]


class HTMLInventory(HTMLParser):
    def __init__(self, source):
        super().__init__(convert_charrefs=True)
        self.ids = set()
        self.links = []
        self.images = []
        self.feed(source)

    def handle_starttag(self, tag, attributes):
        attributes = dict(attributes)
        if attributes.get("id"):
            self.ids.add(attributes["id"])
        if tag == "a" and attributes.get("name"):
            self.ids.add(attributes["name"])
        for attribute in ("href", "src"):
            if attributes.get(attribute):
                self.links.append(attributes[attribute])
        if tag == "img":
            self.images.append(attributes.get("src", ""))


def markdown_without_fenced_code(source):
    return re.sub(r"(?ms)^\s*(`{3,}|~{3,})[^\n]*\n.*?^\s*\1\s*$", "", source)


def markdown_links(source):
    source = markdown_without_fenced_code(source)
    pattern = r'!?\[[^\]\n]*\]\(\s*(<[^>\n]+>|[^\s)]+)(?:\s+["\'][^\n]*?["\'])?\s*\)'
    links = [match.group(1).strip("<>") for match in re.finditer(pattern, source)]
    links.extend(match.group(1).strip("<>") for match in re.finditer(r"(?m)^\s*\[[^\]]+\]:\s*(<[^>]+>|\S+)", source))
    return links


def markdown_anchors(source):
    anchors = set(HTMLInventory(source).ids)
    duplicates = Counter()
    for match in re.finditer(r"(?m)^#{1,6}\s+(.+?)\s*#*\s*$", markdown_without_fenced_code(source)):
        title = re.sub(r"<[^>]*>", "", match.group(1))
        title = re.sub(r"!?\[([^\]]+)\]\([^)]*\)", r"\1", title)
        title = unescape(title).lower().replace("`", "")
        slug = re.sub(r"[^\w\-\s]", "", title)
        slug = re.sub(r"\s", "-", slug)
        count = duplicates[slug]
        duplicates[slug] += 1
        anchors.add(slug + (f"-{count}" if count else ""))
    return anchors


class PublicSiteContract(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.page_bytes = (ROOT / "docs/index.html").read_bytes()
        cls.page = cls.page_bytes.decode("utf-8")
        cls.demo_bytes = (ROOT / "docs/demo.html").read_bytes()
        cls.demo = cls.demo_bytes.decode("utf-8")
        cls.site_template = (ROOT / "src/site/index.html").read_text(encoding="utf-8")

    def test_01_embedded_preview_is_the_exact_public_demo(self):
        match = re.search(r"const demoBase64='([A-Za-z0-9+/=]+)';", self.page)
        self.assertIsNotNone(match, "The final site must contain its built-in demo")
        self.assertEqual(base64.b64decode(match.group(1), validate=True), self.demo_bytes)
        self.assertNotIn("__FORMULA_ATLAS_DEMO_BASE64__", self.page)

    def test_02_four_embedded_pngs_match_the_three_real_screenshots(self):
        original_images = HTMLInventory(self.site_template).images
        embedded_images = HTMLInventory(self.page).images
        self.assertEqual(original_images, [
            "assets/overview.png", "assets/overview.png",
            "assets/selection.png", "assets/loading.png",
        ])
        self.assertEqual(len(embedded_images), 4)
        self.assertEqual(len(set(original_images)), 3)
        prefix = "data:image/png;base64,"
        for original, embedded in zip(original_images, embedded_images):
            self.assertTrue(embedded.startswith(prefix), original)
            decoded = base64.b64decode(embedded[len(prefix):], validate=True)
            self.assertTrue(decoded.startswith(b"\x89PNG\r\n\x1a\n"), original)
            self.assertEqual(decoded, (ROOT / "docs" / original).read_bytes(), original)

    def assert_local_link(self, source_path, raw_link, effective_parent=None):
        link = urlsplit(raw_link)
        if link.scheme in ("https", "http", "mailto", "data") or link.netloc:
            return
        self.assertFalse(link.scheme, f"Unexpected link scheme: {source_path}: {raw_link}")
        parent = effective_parent or source_path.parent
        target = (parent / unquote(link.path)).resolve() if link.path else source_path.resolve()
        self.assertTrue(target.is_relative_to(ROOT.resolve()), f"Link escapes repository: {raw_link}")
        self.assertTrue(target.exists(), f"Missing link target: {source_path.relative_to(ROOT)} -> {raw_link}")
        if link.fragment:
            if target.suffix.lower() == ".html":
                anchors = HTMLInventory(target.read_text(encoding="utf-8-sig")).ids
            elif target.suffix.lower() == ".md":
                anchors = markdown_anchors(target.read_text(encoding="utf-8-sig"))
            else:
                self.fail(f"Cannot resolve fragment on {target.relative_to(ROOT)}: {raw_link}")
            self.assertIn(unquote(link.fragment), anchors, f"Missing anchor: {source_path.relative_to(ROOT)} -> {raw_link}")

    def test_03_local_html_links_and_anchors_exist(self):
        checked = 0
        for path in sorted(ROOT.rglob("*.html")):
            if any(part.startswith(".") or part == "node_modules" for part in path.relative_to(ROOT).parts):
                continue
            effective_parent = ROOT / "docs" if path == ROOT / "src/site/index.html" else None
            for link in HTMLInventory(path.read_text(encoding="utf-8-sig")).links:
                with self.subTest(page=str(path.relative_to(ROOT)), link=link[:120]):
                    self.assert_local_link(path, link, effective_parent)
                checked += 1
        self.assertGreater(checked, 20)

    def test_04_local_markdown_links_and_anchors_exist(self):
        checked = 0
        for path in sorted(ROOT.rglob("*.md")):
            if any(part.startswith(".") or part == "node_modules" for part in path.relative_to(ROOT).parts):
                continue
            for link in markdown_links(path.read_text(encoding="utf-8-sig")):
                with self.subTest(page=str(path.relative_to(ROOT)), link=link):
                    self.assert_local_link(path, link)
                checked += 1
        self.assertGreater(checked, 10)

    def test_05_demo_json_is_the_checked_snapshot_with_safe_escapes(self):
        data = json.loads((ROOT / "examples/map-data.json").read_text(encoding="utf-8"))
        match = re.search(r'<script id="formulaMapData" type="application/json">(.*?)</script>', self.demo, re.S)
        self.assertIsNotNone(match)
        payload = match.group(1)
        self.assertEqual(json.loads(payload), data)
        expected = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
        for char, escaped in (("<", r"\u003c"), (">", r"\u003e"), ("&", r"\u0026"), ("\u2028", r"\u2028"), ("\u2029", r"\u2029")):
            expected = expected.replace(char, escaped)
            self.assertNotIn(char, payload)
        self.assertEqual(payload, expected)

    def test_06_demo_is_exactly_the_current_viewer_template(self):
        pattern = r'(<script id="formulaMapData" type="application/json">).*?(</script>)'
        reconstructed, count = re.subn(pattern, lambda match: match.group(1) + "__FORMULA_MAP_DATA__" + match.group(2), self.demo, flags=re.S)
        self.assertEqual(count, 1)
        self.assertEqual(reconstructed.encode("utf-8"), (ROOT / "src/viewer/template.html").read_bytes())

    def test_07_site_rebuild_is_byte_identical_in_an_isolated_checkout(self):
        isolated = ROOT / (".site-build-" + uuid.uuid4().hex)
        isolated.mkdir()
        try:
            for relative in ("src/site/index.html", "scripts/build_site.py", "docs/demo.html"):
                destination = isolated / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(ROOT / relative, destination)
            shutil.copytree(ROOT / "docs/assets", isolated / "docs/assets")
            subprocess.run([sys.executable, str(isolated / "scripts/build_site.py")], cwd=isolated, check=True, capture_output=True, text=True)
            self.assertEqual((isolated / "docs/index.html").read_bytes(), self.page_bytes)
        finally:
            resolved = isolated.resolve()
            if resolved.parent != ROOT.resolve() or not resolved.name.startswith(".site-build-"):
                raise RuntimeError("Refusing to clean an unexpected site-build directory")
            shutil.rmtree(resolved)


if __name__ == "__main__":
    unittest.main()
