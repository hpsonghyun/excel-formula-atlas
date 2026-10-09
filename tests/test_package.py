"""Static release-contract checks. These do not replace an Excel runtime test.

Run from the repository root: python -m unittest discover -s tests -v
Only Python's standard library is required. No listener or Excel process starts.
"""
from pathlib import Path
import base64
import hashlib
import json
import re
import shutil
import subprocess
import sys
import unittest
import uuid


ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist/ExcelFormulaAtlas.txt"


def embedded_payload(body: str, name: str) -> bytes:
    parts = re.findall(
        rf"Private Function {name}Part\d+\(\) As String(.*?)End Function",
        body,
        re.S,
    )
    encoded = "".join(re.findall(r's = s & "([A-Za-z0-9+/=]+)"', "".join(parts)))
    return base64.b64decode(encoded, validate=True)


class PackageContract(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.raw = DIST.read_bytes()
        cls.body = cls.raw.decode("utf-8-sig")
        cls.viewer = (ROOT / "src/viewer/template.html").read_text(encoding="utf-8")
        cls.bridge = (ROOT / "src/navigation/navigation_bridge.ps1").read_text(encoding="utf-8")
        cls.names = re.findall(
            r"^(?:Public |Private )?(?:Sub|Function) (\w+)", cls.body, re.M | re.I
        )

    def test_01_distribution_encoding_and_newlines(self):
        self.assertTrue(self.raw.startswith(b"\xef\xbb\xbf"))
        self.assertIn(b"\r\n", self.raw)
        self.assertNotIn(b"\n", self.raw.replace(b"\r\n", b""))

    def test_02_one_pasteable_standard_module(self):
        self.assertEqual(self.body.count("Option Explicit"), 1)
        self.assertNotIn("Attribute VB_Name", self.body)
        self.assertNotIn("VBProject", self.body)
        for entry in ("ExcelMap_Start", "ExcelMap_Stop"):
            self.assertEqual(self.body.count(f"Public Sub {entry}()"), 1)

    def test_03_no_duplicate_procedures(self):
        self.assertGreater(len(self.names), 50)
        lowered = [name.casefold() for name in self.names]
        self.assertEqual(len(lowered), len(set(lowered)))

    def test_04_procedure_boundaries_are_balanced(self):
        opened = None
        for line in self.body.splitlines():
            start = re.match(r"^(?:Public |Private )?(Sub|Function) (\w+)", line, re.I)
            end = re.match(r"^End (Sub|Function)\s*$", line, re.I)
            if start:
                self.assertIsNone(opened, f"Nested procedure near {line}")
                opened = start.group(1).lower()
            if end:
                self.assertEqual(opened, end.group(1).lower(), line)
                opened = None
        self.assertIsNone(opened)

    def test_05_vba_string_quotes_close_on_each_line(self):
        for number, line in enumerate(self.body.splitlines(), 1):
            quoted = False
            i = 0
            while i < len(line):
                char = line[i]
                if char == "'" and not quoted:
                    break
                if char == '"':
                    if quoted and i + 1 < len(line) and line[i + 1] == '"':
                        i += 2
                        continue
                    quoted = not quoted
                i += 1
            self.assertFalse(quoted, f"Unclosed VBA string on line {number}")

    def test_06_vbe_line_and_continuation_limits(self):
        continued = 0
        for line in self.body.splitlines():
            self.assertLess(len(line), 1023)
            continued = continued + 1 if line.rstrip().endswith(" _") else 0
            self.assertLessEqual(continued, 24)

    def test_07_vba_instructions_and_messages_are_english(self):
        self.assertIsNone(re.search("[가-힣]", self.body))
        self.assertIn("QUICK START (Windows desktop Excel)", self.body)
        self.assertIn("Choose the source workbook.", self.body)
        self.assertNotIn("automatic approval", self.body.lower())
        self.assertNotIn("C:/obsidian", self.body)
        self.assertNotIn("C:\\obsidian", self.body)

    def test_08_english_first_paint_and_app_default(self):
        for value in (
            '<html lang="en">',
            '<title>Excel Formula Atlas</title>',
            "timer,lang='en'",
            "nextLang='en'",
            "const S={lang:'en'",
            "en:{title:'Excel Formula Atlas'",
        ):
            self.assertIn(value, self.viewer)
        self.assertIn("Object.assign(text.ko", self.viewer)
        self.assertIn("Object.assign(text.en", self.viewer)

    def test_09_viewer_has_no_remote_runtime_assets(self):
        self.assertIsNone(re.search(r'<script\b[^>]*\bsrc\s*=', self.viewer, re.I))
        self.assertIsNone(re.search(r'<link\b[^>]*\bhref=["\']https?://', self.viewer, re.I))
        self.assertIsNone(re.search(r'@import\s+["\']?https?://', self.viewer, re.I))

    def test_10_data_is_inert_json_after_visible_loader(self):
        marker = '<script id="formulaMapData" type="application/json">__FORMULA_MAP_DATA__</script>'
        self.assertEqual(self.viewer.count(marker), 1)
        self.assertLess(self.viewer.index('id="mapLoading"'), self.viewer.index(marker))
        self.assertLess(self.viewer.index("window.FM_LOADING="), self.viewer.index(marker))
        self.assertIn("JSON.parse", self.viewer)

    def test_11_html_payload_is_exact_source_bytes(self):
        self.assertEqual(
            embedded_payload(self.body, "FM_Template"),
            (ROOT / "src/viewer/template.html").read_bytes(),
        )

    def test_12_navigation_payload_is_exact_source_bytes(self):
        self.assertEqual(
            embedded_payload(self.body, "FM_BridgeScript"),
            (ROOT / "src/navigation/navigation_bridge.ps1").read_bytes(),
        )

    def test_13_build_metadata_matches_current_files(self):
        metadata = json.loads((ROOT / "dist/build-metadata.json").read_text(encoding="utf-8"))
        self.assertEqual(metadata["version"], (ROOT / "VERSION").read_text(encoding="utf-8").strip())
        self.assertEqual(metadata["sha256"], hashlib.sha256(self.raw).hexdigest())
        self.assertEqual(metadata["bytes"], len(self.raw))
        self.assertEqual(metadata["procedures"], len(self.names))
        self.assertEqual(metadata["lines"], len(self.body.splitlines()))
        for relative, expected in metadata["sourceHashes"].items():
            self.assertFalse(Path(relative).is_absolute())
            self.assertEqual(hashlib.sha256((ROOT / relative).read_bytes()).hexdigest(), expected)

    def test_14_website_download_matches_distribution(self):
        self.assertEqual(self.raw, (ROOT / "docs/downloads/ExcelFormulaAtlas.txt").read_bytes())

    def test_15_isolated_build_is_reproducible(self):
        # Keep the disposable checkout inside the repository so this test also
        # works in hosts that allow writes only to their workspace directory.
        isolated = ROOT / (".package-build-" + uuid.uuid4().hex)
        isolated.mkdir()
        try:
            shutil.copytree(ROOT / "src", isolated / "src")
            (isolated / "scripts").mkdir()
            for name in ("build.py", "progress_instrumentation.py"):
                shutil.copy2(ROOT / "scripts" / name, isolated / "scripts" / name)
            for name in ("VERSION", "LICENSE"):
                shutil.copy2(ROOT / name, isolated / name)
            command = [sys.executable, str(isolated / "scripts/build.py")]
            for _ in range(2):
                subprocess.run(command, cwd=isolated, check=True, capture_output=True, text=True)
                self.assertEqual(self.raw, (isolated / "dist/ExcelFormulaAtlas.txt").read_bytes())
                self.assertEqual(self.raw, (isolated / "docs/downloads/ExcelFormulaAtlas.txt").read_bytes())
        finally:
            resolved = isolated.resolve()
            if resolved.parent != ROOT.resolve() or not resolved.name.startswith(".package-build-"):
                raise RuntimeError("Refusing to clean an unexpected build-test directory")
            shutil.rmtree(resolved)

    def test_16_vba_json_encoder_escapes_script_delimiters(self):
        encoder = re.search(r"Private Function FM_JSON\(.*?End Function", self.body, re.S).group()
        self.assertIn("Case 34:", encoder)
        self.assertIn("Case 92:", encoder)
        self.assertIn("Case 0 To 31, 60, 62, 38, &H2028, &H2029", encoder)
        self.assertIn('items(i) = "\\u"', encoder)
        self.assertIn('Hex$(n)', encoder)

    def test_17_navigation_is_bounded_and_local(self):
        self.assertIn("Net.Sockets.TcpListener([Net.IPAddress]::Loopback,$port)", self.bridge)
        self.assertIn('$origin="http://127.0.0.1:$port"', self.bridge)
        self.assertIn("$token=[Guid]::NewGuid().ToString('N')", self.bridge)
        self.assertIn("invalid_origin", self.bridge)
        self.assertIn("invalid_host", self.bridge)
        self.assertIn("$allowed", self.bridge)
        self.assertIn("[Math]::Min(7200", self.bridge)
        self.assertNotIn("ExecutionPolicy", self.bridge)
        self.assertNotIn("-ExecutionPolicy", self.body)
        self.assertNotIn("Set-ExecutionPolicy", self.body)
        self.assertIn("shell.Run command, 0, False", self.body)

    def test_18_new_source_files_open_read_only(self):
        self.assertIn("Application.AutomationSecurity = 3", self.body)
        self.assertIn("Filename:=fullPath, UpdateLinks:=0, ReadOnly:=True", self.body)
        self.assertIn("Application.AutomationSecurity = oldSecurity", self.body)
        self.assertNotIn(".Saved = True", self.body)
        self.assertNotIn(".Saved=True", self.body)

    def test_19_progress_has_real_counts_yields_and_cleanup(self):
        self.assertIn("percent = Fix(completed * 100# / total)", self.body)
        self.assertIn("Application.DisplayStatusBar = oldDisplayStatusBar", self.body)
        self.assertIn("mFMProgressBusy = False", self.body)
        self.assertIn("DoEvents", self.body)
        self.assertIn("await paintYield()", self.viewer)
        self.assertIn("await chunks(w.cells", self.viewer)
        self.assertIn("done/total*100", self.viewer)
        self.assertIn("loading.error(error)", self.viewer)

    def test_20_every_counted_loading_stage_is_localized(self):
        stages = re.findall(r"loading\.stage\(([^\n]+)", self.viewer)
        self.assertGreaterEqual(len(stages), 10)
        for stage in stages:
            self.assertTrue(stage.startswith("loadingText("), stage)
        self.assertIn("loadingText('지도 데이터 읽기','Reading map data')", self.viewer)

    def test_21_distributed_code_retains_the_full_mit_notice(self):
        license_text = (ROOT / "LICENSE").read_text(encoding="utf-8").strip()
        commented = "\n".join("' " + line if line else "'" for line in license_text.splitlines())
        self.assertIn(commented, self.body.replace("\r\n", "\n"))
        self.assertIn("<!--\n" + license_text + "\n-->", self.viewer)


if __name__ == "__main__":
    unittest.main()
