"""Offline tests: no Zotero needed.  Run:  python3 -m unittest discover tests"""
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

from docx import Document  # noqa: E402
import zotero_docx as zd  # noqa: E402


def run(*args):
    return subprocess.run([sys.executable] + list(args), cwd=ROOT, capture_output=True, text=True)


class Pipeline(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp()
        cls.cdir = os.path.join(cls.tmp, "citations")
        r = run("scripts/zotero_link.py", "--sources", "examples/sources.json", "--offline", "--out", cls.cdir)
        assert r.returncode == 0, r.stderr
        with open(os.path.join(cls.cdir, "zotero_map.json")) as f:
            cls.zmap = json.load(f)

    def document_xml(self, path):
        with zipfile.ZipFile(path) as z:
            return z.read("word/document.xml").decode("utf-8"), z.read("docProps/custom.xml").decode("utf-8")

    def test_markdown_to_live_fields(self):
        out = os.path.join(self.tmp, "draft.docx")
        r = run("scripts/zotero_docx.py", "examples/draft.md", out, "--map",
                os.path.join(self.cdir, "zotero_map.json"))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        xml, custom = self.document_xml(out)
        self.assertEqual(xml.count("ADDIN ZOTERO_ITEM CSL_CITATION"), 8)
        self.assertEqual(xml.count("ADDIN ZOTERO_BIBL"), 1)
        self.assertNotIn("{@", xml)
        self.assertIn("ZOTERO_PREF_1", custom)
        self.assertIn("Placeholder and Mock ", xml)              # narrative author text
        self.assertIn("\"suppress-author\": true", xml)
        self.assertIn("\"locator\": \"7\"", xml)

    def test_every_item_has_a_uri(self):
        # Without a URI the Word plugin fails ("Zotero experienced an error ...").
        instr, _ = zd.citation_instr([("P001", None, None), ("S0001", None, None)], self.zmap, 1)
        payload = json.loads(instr.split("CSL_CITATION", 1)[1])
        for item in payload["citationItems"]:
            self.assertTrue(re.match(r"^http://zotero\.org/users/local/[^/]+/items/[2-9A-Z]{8}$", item["id"]))
            self.assertEqual(item["uris"], [item["id"]])
            self.assertEqual(item["itemData"]["id"], item["id"])

    def test_placeholder_split_across_runs(self):
        doc = Document()
        p = doc.add_paragraph()
        for piece in ["As shown {@P0", "01, p.", " 5} and {@P002|n}."]:
            p.add_run(piece)
        src = os.path.join(self.tmp, "typed.docx")
        doc.save(src)
        out = os.path.join(self.tmp, "typed_out.docx")
        r = run("scripts/zotero_docx.py", src, out, "--map", os.path.join(self.cdir, "zotero_map.json"))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        xml, _ = self.document_xml(out)
        self.assertEqual(xml.count("ZOTERO_ITEM"), 2)
        self.assertNotIn("{@", xml)

    def test_unknown_id_is_reported_not_guessed(self):
        doc = Document()
        doc.add_paragraph("Claim {@NOPE123}.")
        src = os.path.join(self.tmp, "bad.docx")
        doc.save(src)
        r = run("scripts/zotero_docx.py", src, os.path.join(self.tmp, "bad_out.docx"),
                "--map", os.path.join(self.cdir, "zotero_map.json"))
        self.assertEqual(r.returncode, 1)
        self.assertIn("NOT CONVERTED: {@NOPE123}", r.stdout)

    def test_parse(self):
        self.assertEqual(zd.parse_placeholder("@A, p. 12; @B"), [("A", "page", "12"), ("B", None, None)])
        self.assertEqual(zd.parse_placeholder("@A, fig. 3"), [("A", "figure", "3")])
        self.assertIsNone(zd.parse_placeholder("not a citation"))


if __name__ == "__main__":
    unittest.main()
