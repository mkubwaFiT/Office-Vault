"""Headless tests for the non-UI core layers: TextExtractor, VaultStore, Indexer.

Run:  python -m unittest discover -s tests
These use only the standard library and never construct a Tk window.
"""
import os
import sys
import zipfile
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import trove as vt  # noqa: E402


def _write(path, data, mode="w"):
    with open(path, mode) as f:
        f.write(data)


def _docx(path, text):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr(
            "word/document.xml",
            '<?xml version="1.0"?><w:document xmlns:w="http://w"><w:body><w:p><w:r>'
            f"<w:t>{text}</w:t></w:r></w:p></w:body></w:document>",
        )


def _xlsx(path, text):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/sharedStrings.xml", f'<sst xmlns="http://s"><si><t>{text}</t></si></sst>')


def _xlsx_multisheet(path, shared, *inline_per_sheet):
    """Workbook with a shared-string table plus inline strings in extra sheets."""
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/sharedStrings.xml", f'<sst xmlns="http://s"><si><t>{shared}</t></si></sst>')
        z.writestr("xl/worksheets/sheet1.xml",
                   '<worksheet xmlns="http://s"><sheetData><row><c t="s"><v>0</v></c></row></sheetData></worksheet>')
        for i, txt in enumerate(inline_per_sheet, start=2):
            z.writestr(f"xl/worksheets/sheet{i}.xml",
                       f'<worksheet xmlns="http://s"><sheetData><row><c t="inlineStr">'
                       f'<is><t>{txt}</t></is></c></row></sheetData></worksheet>')


def _pptx(path, text):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("ppt/slides/slide1.xml",
                   f'<p:sld xmlns:p="http://p" xmlns:a="http://a"><a:t>{text}</a:t></p:sld>')


class TestExtractor(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def test_ooxml_extraction(self):
        d, x, p = (os.path.join(self.tmp, n) for n in ("a.docx", "a.xlsx", "a.pptx"))
        _docx(d, "Confidential merger plan")
        _xlsx(x, "Quarterly budget forecast")
        _pptx(p, "Roadmap slide deck")
        self.assertIn("merger", vt.TextExtractor.extract(d, ".docx"))
        self.assertIn("budget", vt.TextExtractor.extract(x, ".xlsx"))
        self.assertIn("Roadmap", vt.TextExtractor.extract(p, ".pptx"))

    def test_bad_file_is_graceful(self):
        bad = os.path.join(self.tmp, "broken.docx")
        with open(bad, "wb") as f:
            f.write(b"not a zip")
        self.assertEqual(vt.TextExtractor.extract(bad, ".docx"), "")

    def test_xlsx_extracts_all_worksheets(self):
        # A word that only appears as an inline string in sheet 2/3 must still be
        # extracted, so full-text search reaches every worksheet of a workbook.
        x = os.path.join(self.tmp, "book.xlsx")
        _xlsx_multisheet(x, "Revenue", "Zephyrium confidential", "Appendix notes")
        body = vt.TextExtractor.extract(x, ".xlsx")
        self.assertIn("Revenue", body)      # shared string (sheet 1)
        self.assertIn("Zephyrium", body)    # inline string, sheet 2
        self.assertIn("Appendix", body)     # inline string, sheet 3

    def test_dominant_keyword(self):
        self.assertEqual(vt.TextExtractor.dominant_keyword("budget budget report report report"), "Report")


class TestStore(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.st = vt.VaultStore(os.path.join(self.tmp, "vault.db"))

    def tearDown(self):
        self.st.close()

    def _add(self, name, body):
        p = os.path.join(self.tmp, name)
        return self.st.upsert({
            "original_path": p, "vault_path": None, "filename": name,
            "ext": os.path.splitext(name)[1], "category": "T", "source_dir": self.tmp,
            "editable": 0, "mtime": 1.0, "size": 1, "body": body,
        })

    def test_fulltext_search_across_types(self):
        self._add("d.docx", "Confidential merger plan alpha")
        self._add("a.txt", "finance report q3")
        self.st.commit()
        self.assertTrue(any("d.docx" in r["filename"] for r in self.st.search("merger")))
        self.assertTrue(any("a.txt" in r["filename"] for r in self.st.search("finance")))

    def test_prefix_and_filename_search(self):
        self._add("budget_2026.xlsx", "quarterly numbers")
        self.st.commit()
        self.assertTrue(any("budget" in r["filename"] for r in self.st.search("budg")))
        self.assertEqual(len(self.st.search_filename("budget")), 1)

    def test_delete_removes_from_index_and_fts(self):
        fid = self._add("gone.txt", "temporary secret token")
        self.st.commit()
        self.st.delete(fid)
        self.assertEqual(self.st.count(), 0)
        self.assertEqual(self.st.search("secret"), [])

    def test_extension_filter_and_grouping_queries(self):
        self._add("report.docx", "quarterly revenue merger")
        self._add("book.xlsx", "revenue zephyrium confidential")
        self._add("notes.txt", "revenue meeting notes")
        self.st.commit()
        # distinct extensions with counts
        self.assertEqual({r["ext"] for r in self.st.extensions()}, {".docx", ".xlsx", ".txt"})
        # full-text search restricted to one extension
        self.assertEqual(len(self.st.search("revenue")), 3)
        self.assertEqual([r["filename"] for r in self.st.search("revenue", exts=[".xlsx"])], ["book.xlsx"])
        # filename search restricted to one extension
        self.assertEqual([r["filename"] for r in self.st.search_filename("book", exts=[".xlsx"])], ["book.xlsx"])
        # grouping helpers
        self.assertEqual([r["filename"] for r in self.st.files_by_ext(".txt")], ["notes.txt"])
        self.assertEqual(len(self.st.files_by_source(self.tmp)), 3)


class TestIndexer(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.src = os.path.join(self.tmp, "src")
        os.makedirs(os.path.join(self.src, "sub"))
        _write(os.path.join(self.src, "a.txt"), "budget report finance")
        _docx(os.path.join(self.src, "sub", "b.docx"), "Confidential merger plan")
        _write(os.path.join(self.src, "c.exe"), b"MZ", "wb")          # danger -> flagged
        _write(os.path.join(self.src, "pic.png"), b"x", "wb")         # image: tracked (empty body w/o OCR)
        _write(os.path.join(self.src, "skip.zip"), b"x", "wb")        # genuinely untracked
        self.st = vt.VaultStore(os.path.join(self.tmp, "idx.db"))

    def tearDown(self):
        self.st.close()

    def _run(self, target):
        res = {}
        idx = vt.Indexer(self.st, os.path.join(self.tmp, "vault"),
                         on_progress=lambda s, a: None,
                         on_done=lambda s, a, sec, canc: res.update(scanned=s, added=a, sec=sec, canc=canc))
        idx._run(target, copy_to_vault=False)
        self.st.commit()
        return res

    def test_index_counts_flags_and_search(self):
        res = self._run(self.src)
        self.assertEqual(res["scanned"], 3)                 # a.txt + b.docx + pic.png (.zip untracked)
        self.assertEqual(res["added"], 3)
        self.assertEqual(res["sec"], [("Executable/Script", os.path.join(self.src, "c.exe"))])
        self.assertFalse(res["canc"])
        self.assertTrue(any("b.docx" in r["filename"] for r in self.st.search("merger")))
        # the image is catalogued (indexed by name) even with no OCR engine present
        self.assertTrue(any("pic.png" in r["filename"] for r in self.st.all_files()))

    def test_reindex_is_idempotent(self):
        self._run(self.src)
        self.assertEqual(self._run(self.src)["added"], 0)

    def test_cancellation_stops_scan(self):
        big = os.path.join(self.tmp, "big")
        os.makedirs(big)
        for i in range(30):
            _write(os.path.join(big, f"f{i}.txt"), "x")
        res = {}
        idx = vt.Indexer(self.st, os.path.join(self.tmp, "vault"),
                         on_progress=lambda s, a: None,
                         on_done=lambda s, a, sec, canc: res.update(added=a, canc=canc))
        idx.cancel()
        idx._run(big, False)
        self.assertTrue(res["canc"])
        self.assertEqual(res["added"], 0)


class TestChunkAndEngines(unittest.TestCase):
    def test_find_query_lines_and_best_line(self):
        body = "intro line\nthe secret budget is here\nfooter"
        hits = vt.find_query_lines(body, "secret budget")
        self.assertEqual(hits, [(2, "the secret budget is here")])
        self.assertEqual(vt.best_match_line(body, "secret"), "the secret budget is here")
        self.assertEqual(vt.find_query_lines("", "x"), [])
        self.assertEqual(vt.best_match_line("no match here", "zzz"), "")

    def test_optional_engines_degrade_gracefully(self):
        # On a machine without the ML extras these are simply disabled no-ops.
        ocr = vt.OcrEngine()
        if not ocr.available:
            self.assertEqual(ocr.read("/nonexistent.png"), "")
        sem = vt.SemanticRanker()
        rows = [{"body": "a"}, {"body": "b"}]
        if not sem.available:  # rerank must return candidates unchanged
            self.assertEqual(sem.rerank("q", rows, text_of=lambda r: r["body"]), rows)


class TestFolderWatcher(unittest.TestCase):
    def test_polling_watcher_detects_change(self):
        import time
        tmp = tempfile.mkdtemp()
        _write(os.path.join(tmp, "seed.txt"), "one")
        fired = []
        w = vt.FolderWatcher(on_change=lambda folder: fired.append(folder), interval=1)
        w.backend = "poll"  # force the stdlib fallback for a deterministic test
        w.start([tmp])
        try:
            time.sleep(0.3)
            _write(os.path.join(tmp, "new.txt"), "two")   # change after baseline
            deadline = time.time() + 6
            while not fired and time.time() < deadline:
                time.sleep(0.3)
        finally:
            w.stop()
        self.assertTrue(fired, "polling watcher did not detect the new file")
        self.assertEqual(fired[0], tmp)


if __name__ == "__main__":
    unittest.main(verbosity=2)
