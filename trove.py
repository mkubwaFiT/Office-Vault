"""
Trove — local document catalog, deep search, and cleanup toolkit.
(Repository: mkubwaFiT/Office-Vault. "Vault Toolkit" was the former name.)

Architecture (layered):
  * TextExtractor  — stdlib text extraction (.txt + OOXML) + optional PDF/OCR.
  * VaultStore     — SQLite catalog with FTS5 full-text search (LIKE fallback).
  * Indexer        — cancellable, streaming, background disk scanner.
  * FolderWatcher  — live re-indexing (watchdog if present, else stdlib polling).
  * SemanticRanker — optional hybrid re-rank (SentenceTransformers, if installed).
  * OcrEngine      — optional OCR for images/PDF (EasyOCR, if installed).
  * TroveApp       — Tkinter UI (browse / search / preview / security / purge).

Design notes:
  * The CORE is standard-library only: lean, fast, ~30 MB build. The heavy ML
    features (semantic search, OCR) are OPTIONAL — they auto-activate only if the
    extra packages are installed, and degrade gracefully to keyword search if not.
  * Files are indexed *in place* (no whole-disk copying) — a searchable catalog
    plus a home for in-app notes, not a duplicate store.
  * Everything heavy runs off the Tk thread; the UI marshals updates via after().
  * All deletes are recoverable (OS Recycle Bin / Trash, local fallback).
"""
import os
import sys
import re
import time
import json
import shutil
import sqlite3
import hashlib
import logging
import zipfile
import threading
import subprocess
import xml.etree.ElementTree as ET
from collections import Counter

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
except Exception:  # headless (e.g. CI unit tests of the core layers) — GUI unused
    tk = ttk = filedialog = messagebox = simpledialog = None

# winreg / explorer / Defender are Windows-only.
IS_WINDOWS = sys.platform.startswith("win")
if IS_WINDOWS:
    import winreg

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
STOP_WORDS = set(
    "the and to a of in it is for that on you this with as at by not be are from "
    "but have an which was or we can if your has will all".split()
)

MS_EXTENSIONS = {
    '.doc': 'Word Documents',
    '.docx': 'Word Documents',
    '.xls': 'Excel Spreadsheets',
    '.xlsx': 'Excel Spreadsheets',
    '.ppt': 'PowerPoint Presentations',
    '.pptx': 'PowerPoint Presentations',
}
OOXML_EXTS = {'.docx', '.xlsx', '.pptx'}          # can be text-extracted with stdlib
LEGACY_OLE_EXTS = {'.doc', '.xls', '.ppt'}         # pre-2007 binary — no stdlib parser
IMAGE_EXTS = {'.png', '.jpg', '.jpeg'}             # OCR-eligible (optional engine)
PDF_EXTS = {'.pdf'}                                # text + optional OCR (needs PyMuPDF)
# Plain-text-family formats read directly as UTF-8 (stdlib — fits the lean core).
PLAINTEXT_EXTS = {'.txt', '.csv', '.tsv', '.md', '.log', '.json', '.xml', '.ini'}
RTF_EXTS = {'.rtf'}                                # control-word stripping (stdlib)
TRACKED_EXTS = set(MS_EXTENSIONS) | PLAINTEXT_EXTS | RTF_EXTS | IMAGE_EXTS | PDF_EXTS
DANGER_EXTS = {'.exe', '.bat', '.ps1', '.vbs', '.scr', '.dll', '.js', '.wsf'}

# Directory names skipped entirely while crawling (system / VCS / build / junk).
# Prevents wasting time and index space on OS and toolchain internals.
EXCLUDE_DIR_NAMES = {
    '$recycle.bin', 'system volume information', 'windows', 'winsxs',
    '$windows.~ws', '$windows.~bt', 'recovery', 'appdata', 'programdata',
    'node_modules', '.git', '.svn', '.hg', '__pycache__', '.cache',
    'site-packages', '.venv', 'venv', '.trash', '_recyclebin',
}
MAX_FILE_BYTES = 50 * 1024 * 1024     # skip individual files larger than this
ZIP_MEMBER_CAP = 300 * 1024 * 1024    # skip OOXML zip members bigger than this (bomb guard)

# Human-readable type labels for grouping / the Type column.
TYPE_LABELS = dict(MS_EXTENSIONS)
TYPE_LABELS.update({
    '.txt': 'Text', '.md': 'Markdown', '.log': 'Logs', '.ini': 'Config',
    '.csv': 'CSV / Data', '.tsv': 'CSV / Data', '.json': 'JSON / Data', '.xml': 'XML / Data',
    '.rtf': 'Rich Text', '.pdf': 'PDF Documents',
    '.png': 'Images', '.jpg': 'Images', '.jpeg': 'Images',
})

INDEX_TEXT_CAP = 1_000_000     # max chars of extracted body stored per file
PREVIEW_CHUNK = 200_000        # bytes of a large .txt loaded per "page"
BATCH_COMMIT = 200             # files per DB transaction while indexing
SEARCH_LIMIT = 500            # max search results returned to the UI

log = logging.getLogger("trove")


# ===========================================================================
# Optional acceleration engines — the CORE never depends on these. Each detects
# its library lazily and degrades to a no-op / keyword fallback if it is absent,
# so the stdlib-only build behaves exactly as before.
#   pip install "sentence-transformers numpy"   # hybrid semantic search
#   pip install easyocr                          # OCR for images / scanned PDFs
#   pip install pymupdf                          # PDF text + embedded images
#   pip install watchdog                         # event-driven folder watching
# ===========================================================================
def _module_available(modname):
    # find_spec checks importability WITHOUT importing (so we never load torch at
    # startup just to detect it — the heavy import happens lazily on first use).
    try:
        import importlib.util
        return importlib.util.find_spec(modname) is not None
    except Exception:
        return False


class OcrEngine:
    """OCR via EasyOCR when installed; otherwise `available` is False and read()
    returns "" so image/PDF indexing simply stores no text (graceful)."""
    def __init__(self):
        self.available = _module_available("easyocr")
        self._reader = None

    def read(self, path):
        if not self.available:
            return ""
        try:
            if self._reader is None:
                import easyocr
                self._reader = easyocr.Reader(['en'], gpu=False, verbose=False)
            return " ".join(self._reader.readtext(path, detail=0))
        except Exception as e:
            log.warning("OCR failed for %s: %s", path, e)
            return ""

    def read_bytes(self, data):
        if not self.available:
            return ""
        try:
            import numpy as np  # noqa
            from PIL import Image  # noqa
            import io
            if self._reader is None:
                import easyocr
                self._reader = easyocr.Reader(['en'], gpu=False, verbose=False)
            img = np.array(Image.open(io.BytesIO(data)).convert("RGB"))
            return " ".join(self._reader.readtext(img, detail=0))
        except Exception as e:
            log.warning("OCR (bytes) failed: %s", e)
            return ""


class SemanticRanker:
    """Optional hybrid search: re-ranks keyword (FTS) candidates by embedding
    similarity using SentenceTransformers. If the library is absent, `available`
    is False and rerank() returns the candidates unchanged (pure keyword)."""
    MODEL = "all-MiniLM-L6-v2"

    def __init__(self):
        self.available = _module_available("sentence_transformers") and _module_available("numpy")
        self._model = None

    def _load(self):
        if self._model is None:
            from sentence_transformers import SentenceTransformer
            self._model = SentenceTransformer(self.MODEL)
        return self._model

    def rerank(self, query, rows, text_of):
        """Return `rows` reordered by semantic similarity of text_of(row) to query.
        Falls back to the original order on any error or if unavailable."""
        if not self.available or not rows:
            return rows
        try:
            import numpy as np
            model = self._load()
            texts = [text_of(r) or "" for r in rows]
            qv = model.encode([query], normalize_embeddings=True)[0]
            mv = model.encode(texts, normalize_embeddings=True)
            scores = np.asarray(mv) @ np.asarray(qv)
            return [r for _s, r in sorted(zip(scores, rows), key=lambda t: t[0], reverse=True)]
        except Exception as e:
            log.warning("semantic rerank failed: %s", e)
            return rows


class FolderWatcher:
    """Live folder watching. Uses `watchdog` for event-driven notifications when
    installed; otherwise falls back to a lightweight stdlib mtime-poll. Calls
    on_change(folder) (debounced) whenever a watched folder's contents change."""
    def __init__(self, on_change, interval=10):
        self.on_change = on_change
        self.interval = interval
        self.folders = []
        self._stop = threading.Event()
        self._thread = None
        self._observer = None
        self._sigs = {}
        self.backend = "watchdog" if _module_available("watchdog") else "poll"

    @property
    def running(self):
        return bool(self._observer) or (self._thread is not None and self._thread.is_alive())

    def start(self, folders):
        self.stop()
        self.folders = [f for f in folders if os.path.isdir(f)]
        if not self.folders:
            return
        self._stop.clear()
        if self.backend == "watchdog":
            self._start_watchdog()
        else:
            self._thread = threading.Thread(target=self._poll_loop, daemon=True)
            self._thread.start()

    def stop(self):
        self._stop.set()
        if self._observer:
            try:
                self._observer.stop()
                self._observer.join(timeout=2)
            except Exception:
                pass
            self._observer = None
        self._thread = None

    def _start_watchdog(self):
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler

        watcher = self

        class _Handler(FileSystemEventHandler):
            def on_any_event(self, event):
                if event.is_directory:
                    return
                ext = os.path.splitext(event.src_path)[1].lower()
                if ext in TRACKED_EXTS:
                    watcher.on_change(os.path.dirname(event.src_path))

        self._observer = Observer()
        for folder in self.folders:
            self._observer.schedule(_Handler(), folder, recursive=True)
        self._observer.start()

    def _signature(self, folder):
        sig = {}
        for root_dir, _dirs, files in os.walk(folder):
            for name in files:
                if os.path.splitext(name)[1].lower() in TRACKED_EXTS:
                    p = os.path.join(root_dir, name)
                    try:
                        sig[p] = os.stat(p).st_mtime
                    except OSError:
                        pass
        return sig

    def _poll_loop(self):
        self._sigs = {f: self._signature(f) for f in self.folders}
        while not self._stop.wait(self.interval):
            for folder in self.folders:
                try:
                    sig = self._signature(folder)
                    if sig != self._sigs.get(folder):
                        self._sigs[folder] = sig
                        self.on_change(folder)
                except Exception as e:
                    log.warning("watch poll failed for %s: %s", folder, e)


# Module-level optional engines (cheap to construct; heavy import is deferred).
OCR = OcrEngine()
SEMANTIC = SemanticRanker()


# ===========================================================================
# Text extraction
# ===========================================================================
class TextExtractor:
    """Text extraction for indexing and preview. Core formats are stdlib-only;
    PDF and image OCR use optional engines and yield "" when those are absent."""

    @staticmethod
    def extract(path, ext, cap=INDEX_TEXT_CAP):
        try:
            if ext in PLAINTEXT_EXTS:
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read(cap)
            if ext in RTF_EXTS:
                return TextExtractor._rtf(path, cap)
            if ext == '.docx':
                return TextExtractor._ooxml(path, cap, member='word/document.xml')
            if ext == '.xlsx':
                return TextExtractor._xlsx(path, cap)
            if ext == '.pptx':
                return TextExtractor._ooxml(path, cap, member_prefix='ppt/slides/slide')
            if ext in IMAGE_EXTS:
                return OCR.read(path)[:cap]                 # "" unless EasyOCR present
            if ext in PDF_EXTS:
                return TextExtractor._pdf(path, cap)        # PyMuPDF text (+OCR) if present
        except Exception as e:
            log.warning("extract failed for %s: %s", path, e)
        return ""

    @staticmethod
    def _rtf(path, cap):
        """Strip RTF control words / groups to plain text — stdlib, no dependency."""
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            raw = f.read(cap * 4)
        raw = re.sub(r"\\'[0-9a-fA-F]{2}", ' ', raw)   # hex escapes
        raw = re.sub(r'\\[a-zA-Z]+-?\d* ?', ' ', raw)  # control words
        raw = raw.replace('{', ' ').replace('}', ' ').replace('\\', ' ')
        return re.sub(r'\s+', ' ', raw).strip()[:cap]

    @staticmethod
    def _pdf(path, cap):
        """PDF text via PyMuPDF (fitz) when installed; OCR embedded/scanned image
        bytes if an OCR engine is available. Returns "" if PyMuPDF is absent."""
        if not _module_available("fitz"):
            return ""
        texts, total = [], 0
        try:
            import fitz
            with fitz.open(path) as doc:
                for page in doc:
                    t = page.get_text() or ""
                    if t:
                        texts.append(t)
                        total += len(t)
                    if total < cap and OCR.available:  # scanned page / embedded images
                        for img in page.get_images(full=True):
                            try:
                                data = doc.extract_image(img[0]).get("image")
                                if data:
                                    ocr = OCR.read_bytes(data)
                                    if ocr:
                                        texts.append(ocr)
                                        total += len(ocr)
                            except Exception:
                                continue
                    if total >= cap:
                        break
        except Exception as e:
            log.warning("pdf extract failed for %s: %s", path, e)
        return ' '.join(texts)[:cap]

    @staticmethod
    def _xlsx(path, cap):
        """Extract text from the shared-string table AND every worksheet's inline
        strings, so a word in *any* sheet of a multi-worksheet workbook is indexed
        and searchable. (Shared strings are workbook-global; inline `<is><t>` live
        in each sheet — both are collected here.)"""
        texts, total = [], 0
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            parts = (['xl/sharedStrings.xml'] if 'xl/sharedStrings.xml' in names else []) + sorted(
                n for n in names if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')
            )
            for name in parts:
                if z.getinfo(name).file_size > ZIP_MEMBER_CAP:   # decompression-bomb guard
                    continue
                with z.open(name) as fh:
                    for _event, elem in ET.iterparse(fh):
                        if elem.tag.rsplit('}', 1)[-1] == 't' and elem.text:
                            texts.append(elem.text)
                            total += len(elem.text)
                        elem.clear()
                        if total >= cap:
                            break
                if total >= cap:
                    break
        return ' '.join(texts)[:cap]

    @staticmethod
    def _ooxml(path, cap, member=None, member_prefix=None):
        """Pull all <...:t> text nodes out of one or more OOXML parts."""
        texts, total = [], 0
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            targets = ([member] if member and member in names else []) + (
                sorted(n for n in names if member_prefix and n.startswith(member_prefix) and n.endswith('.xml'))
            )
            for name in targets:
                if z.getinfo(name).file_size > ZIP_MEMBER_CAP:   # decompression-bomb guard
                    continue
                with z.open(name) as fh:
                    for _event, elem in ET.iterparse(fh):
                        if elem.tag.rsplit('}', 1)[-1] == 't' and elem.text:
                            texts.append(elem.text)
                            total += len(elem.text)
                        elem.clear()
                        if total >= cap:
                            break
                if total >= cap:
                    break
        return ' '.join(texts)[:cap]

    @staticmethod
    def dominant_keyword(text):
        words = [w for w in re.findall(r'\b[a-zA-Z]{4,}\b', text[:10000].lower()) if w not in STOP_WORDS]
        if not words:
            return "Uncategorized"
        return Counter(words).most_common(1)[0][0].capitalize()


# ===========================================================================
# SQLite catalog + FTS5 full-text search
# ===========================================================================
class VaultStore:
    def __init__(self, db_path):
        self._lock = threading.Lock()
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.fts = self._detect_fts5()
        self._init_schema()

    def _detect_fts5(self):
        try:
            self.conn.execute("CREATE VIRTUAL TABLE temp.__fts_probe USING fts5(x)")
            self.conn.execute("DROP TABLE temp.__fts_probe")
            return True
        except sqlite3.OperationalError:
            log.warning("FTS5 unavailable — falling back to LIKE search")
            return False

    def _init_schema(self):
        with self._lock:
            self.conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS files (
                    id INTEGER PRIMARY KEY,
                    original_path TEXT UNIQUE,
                    vault_path TEXT,
                    filename TEXT,
                    ext TEXT,
                    category TEXT,
                    source_dir TEXT,
                    editable INTEGER DEFAULT 0,
                    mtime REAL,
                    size INTEGER,
                    body TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_files_source ON files(source_dir);
                CREATE INDEX IF NOT EXISTS idx_files_ext ON files(ext);
                CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
                """
            )
            if self.fts:
                self.conn.execute(
                    "CREATE VIRTUAL TABLE IF NOT EXISTS files_fts USING fts5(filename, body)"
                )
            self.conn.commit()

    # -- meta / settings -----------------------------------------------------
    def get_meta(self, key, default=None):
        with self._lock:
            row = self.conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return json.loads(row["value"]) if row else default

    def set_meta(self, key, value):
        with self._lock:
            self.conn.execute(
                "INSERT INTO meta(key, value) VALUES(?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, json.dumps(value)),
            )
            self.conn.commit()

    def is_empty(self):
        with self._lock:
            return self.conn.execute("SELECT COUNT(*) c FROM files").fetchone()["c"] == 0

    def count(self):
        with self._lock:
            return self.conn.execute("SELECT COUNT(*) c FROM files").fetchone()["c"]

    # -- writes --------------------------------------------------------------
    def get_by_path(self, original_path):
        with self._lock:
            return self.conn.execute(
                "SELECT * FROM files WHERE original_path=?", (original_path,)
            ).fetchone()

    def upsert(self, rec):
        with self._lock:
            self.conn.execute(
                """INSERT INTO files
                   (original_path, vault_path, filename, ext, category, source_dir, editable, mtime, size, body)
                   VALUES (:original_path,:vault_path,:filename,:ext,:category,:source_dir,:editable,:mtime,:size,:body)
                   ON CONFLICT(original_path) DO UPDATE SET
                     vault_path=excluded.vault_path, filename=excluded.filename, ext=excluded.ext,
                     category=excluded.category, source_dir=excluded.source_dir, editable=excluded.editable,
                     mtime=excluded.mtime, size=excluded.size, body=excluded.body""",
                rec,
            )
            fid = self.conn.execute(
                "SELECT id FROM files WHERE original_path=?", (rec["original_path"],)
            ).fetchone()["id"]
            if self.fts:
                self.conn.execute("DELETE FROM files_fts WHERE rowid=?", (fid,))
                self.conn.execute(
                    "INSERT INTO files_fts(rowid, filename, body) VALUES(?,?,?)",
                    (fid, rec["filename"], rec["body"] or ""),
                )
            return fid

    def commit(self):
        with self._lock:
            self.conn.commit()

    def delete(self, file_id):
        with self._lock:
            self.conn.execute("DELETE FROM files WHERE id=?", (file_id,))
            if self.fts:
                self.conn.execute("DELETE FROM files_fts WHERE rowid=?", (file_id,))
            self.conn.commit()

    def get(self, file_id):
        with self._lock:
            return self.conn.execute("SELECT * FROM files WHERE id=?", (file_id,)).fetchone()

    # -- browse queries ------------------------------------------------------
    def source_dirs(self):
        with self._lock:
            return self.conn.execute(
                "SELECT source_dir, COUNT(*) n FROM files GROUP BY source_dir ORDER BY source_dir"
            ).fetchall()

    def categories(self, source_dir):
        with self._lock:
            return self.conn.execute(
                "SELECT category, COUNT(*) n FROM files WHERE source_dir=? GROUP BY category ORDER BY category",
                (source_dir,),
            ).fetchall()

    def files_in(self, source_dir, category):
        with self._lock:
            return self.conn.execute(
                "SELECT * FROM files WHERE source_dir=? AND category=? ORDER BY filename",
                (source_dir, category),
            ).fetchall()

    def all_files(self):
        with self._lock:
            return self.conn.execute("SELECT * FROM files ORDER BY filename").fetchall()

    def extensions(self):
        """Distinct extensions present, with counts — for the group view + filter."""
        with self._lock:
            return self.conn.execute(
                "SELECT ext, COUNT(*) n FROM files GROUP BY ext ORDER BY ext"
            ).fetchall()

    def files_by_ext(self, ext):
        with self._lock:
            return self.conn.execute(
                "SELECT * FROM files WHERE ext=? ORDER BY filename", (ext,)
            ).fetchall()

    def files_by_source(self, source_dir):
        with self._lock:
            return self.conn.execute(
                "SELECT * FROM files WHERE source_dir=? ORDER BY filename", (source_dir,)
            ).fetchall()

    @staticmethod
    def _ext_clause(exts, params):
        """Append an `AND ext IN (...)` filter to a query if exts is given."""
        if not exts:
            return ""
        params.extend(exts)
        return " AND ext IN (%s)" % ",".join("?" for _ in exts)

    def search_filename(self, query, limit=SEARCH_LIMIT, exts=None):
        params = [f"%{query}%"]
        sql = "SELECT * FROM files WHERE filename LIKE ?" + self._ext_clause(exts, params)
        sql += " ORDER BY filename LIMIT ?"
        params.append(limit)
        with self._lock:
            return self.conn.execute(sql, params).fetchall()

    # -- search --------------------------------------------------------------
    def search(self, query, limit=SEARCH_LIMIT, exts=None):
        query = (query or "").strip()
        if not query:
            return []
        if self.fts:
            tokens = re.findall(r"\w+", query)
            if not tokens:
                return []
            params = [" ".join(f"{t}*" for t in tokens)]
            sql = ("SELECT f.*, snippet(files_fts, 1, '«', '»', ' … ', 12) AS snippet "
                   "FROM files_fts JOIN files f ON f.id = files_fts.rowid "
                   "WHERE files_fts MATCH ?") + self._ext_clause(exts, params)
            sql += " ORDER BY rank LIMIT ?"
            params.append(limit)
            try:
                with self._lock:
                    return self.conn.execute(sql, params).fetchall()
            except sqlite3.OperationalError as e:
                log.warning("FTS query failed (%s) — using LIKE", e)
        # LIKE fallback
        params = [f"%{query}%", f"%{query}%"]
        sql = "SELECT * FROM files WHERE (filename LIKE ? OR body LIKE ?)" + self._ext_clause(exts, params)
        sql += " ORDER BY filename LIMIT ?"
        params.append(limit)
        with self._lock:
            return self.conn.execute(sql, params).fetchall()

    def close(self):
        with self._lock:
            self.conn.close()


# ===========================================================================
# Cancellable streaming indexer
# ===========================================================================
class Indexer:
    def __init__(self, store, vault_dir, on_progress, on_done):
        self.store = store
        self.vault_dir = os.path.normpath(vault_dir)
        self.on_progress = on_progress
        self.on_done = on_done
        self._cancel = threading.Event()
        self._thread = None

    @property
    def running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self, target_dir, copy_to_vault=False):
        if self.running:
            return False
        self._cancel.clear()
        self._thread = threading.Thread(
            target=self._run, args=(target_dir, copy_to_vault), daemon=True
        )
        self._thread.start()
        return True

    def cancel(self):
        self._cancel.set()

    def _run(self, target_dir, copy_to_vault):
        scanned = added = 0
        sec_flags = []
        try:
            for root_dir, dirs, files in os.walk(target_dir):
                if self._cancel.is_set():
                    break
                # Prune excluded system/VCS/build dirs in place (stops descent too).
                dirs[:] = [d for d in dirs if d.lower() not in EXCLUDE_DIR_NAMES]
                if os.path.normpath(root_dir).startswith(self.vault_dir):
                    dirs[:] = []          # never descend into the vault
                    continue
                for name in files:
                    if self._cancel.is_set():
                        break
                    ext = os.path.splitext(name)[1].lower()
                    src = os.path.join(root_dir, name)
                    if ext in DANGER_EXTS:
                        sec_flags.append(("Executable/Script", src))
                        continue
                    if ext not in TRACKED_EXTS:
                        continue
                    scanned += 1
                    try:
                        if self._index_one(src, ext, copy_to_vault):
                            added += 1
                            if added % BATCH_COMMIT == 0:
                                self.store.commit()
                                self.on_progress(scanned, added)
                    except Exception as e:
                        log.warning("index failed for %s: %s", src, e)
            self.store.commit()
        except Exception as e:
            log.exception("indexer crashed: %s", e)
        finally:
            self.on_done(scanned, added, sec_flags, self._cancel.is_set())

    def _index_one(self, src, ext, copy_to_vault):
        st = os.stat(src)
        if st.st_size > MAX_FILE_BYTES:      # skip oversized files (memory/time guard)
            return False
        mtime = st.st_mtime
        existing = self.store.get_by_path(src)
        if existing and existing["mtime"] == mtime and existing["body"] is not None:
            return False  # unchanged, already indexed

        body = TextExtractor.extract(src, ext)
        category = TextExtractor.dominant_keyword(body) if ext == '.txt' else TYPE_LABELS.get(ext, "Documents")

        vault_path = None
        if copy_to_vault:
            vault_path = self._copy_into_vault(src)

        self.store.upsert({
            "original_path": src,
            "vault_path": vault_path,
            "filename": os.path.basename(src),
            "ext": ext,
            "category": category,
            "source_dir": os.path.dirname(src),
            "editable": 0,
            "mtime": mtime,
            "size": st.st_size,
            "body": body,
        })
        return True

    def _copy_into_vault(self, src):
        dest = os.path.join(self.vault_dir, os.path.basename(src))
        base, e = os.path.splitext(dest)
        counter = 1
        while os.path.exists(dest):
            dest = f"{base}_{counter}{e}"
            counter += 1
        shutil.copy2(src, dest)
        return dest


# ===========================================================================
# OS integration helpers (recoverable delete, reveal, registry, defender)
# ===========================================================================
def send_to_trash(path):
    path = os.path.abspath(path)
    try:
        if IS_WINDOWS:
            import ctypes
            from ctypes import wintypes

            class SHFILEOPSTRUCTW(ctypes.Structure):
                _fields_ = [
                    ("hwnd", wintypes.HWND), ("wFunc", wintypes.UINT),
                    ("pFrom", wintypes.LPCWSTR), ("pTo", wintypes.LPCWSTR),
                    ("fFlags", ctypes.c_uint16), ("fAnyOperationsAborted", wintypes.BOOL),
                    ("hNameMappings", ctypes.c_void_p), ("lpszProgressTitle", wintypes.LPCWSTR),
                ]
            FO_DELETE, FOF_ALLOWUNDO, FOF_NOCONFIRMATION, FOF_SILENT, FOF_NOERRORUI = 3, 0x40, 0x10, 0x04, 0x400
            op = SHFILEOPSTRUCTW()
            op.wFunc = FO_DELETE
            op.pFrom = path + "\0"
            op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
            res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
            return res == 0 and not op.fAnyOperationsAborted
        elif sys.platform == "darwin":
            script = f'tell application "Finder" to delete POSIX file "{path}"'
            return subprocess.call(["osascript", "-e", script],
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL) == 0
        else:
            trash = os.path.join(os.path.expanduser("~"), ".local", "share", "Trash")
            files_dir, info_dir = os.path.join(trash, "files"), os.path.join(trash, "info")
            os.makedirs(files_dir, exist_ok=True)
            os.makedirs(info_dir, exist_ok=True)
            base = os.path.basename(path)
            dest = os.path.join(files_dir, base)
            counter = 1
            while os.path.exists(dest):
                dest = os.path.join(files_dir, f"{base}.{counter}")
                counter += 1
            with open(os.path.join(info_dir, os.path.basename(dest) + ".trashinfo"), "w") as f:
                f.write(f"[Trash Info]\nPath={path}\nDeletionDate={time.strftime('%Y-%m-%dT%H:%M:%S')}\n")
            shutil.move(path, dest)
            return True
    except Exception as e:
        log.warning("trash failed for %s: %s", path, e)
        return False


def safe_delete(path, fallback_dir):
    """Recoverable delete: OS trash, else relocate to fallback_dir. Never hard-removes."""
    if not path or not os.path.exists(path):
        return True
    if send_to_trash(path):
        return True
    try:
        os.makedirs(fallback_dir, exist_ok=True)
        dest = os.path.join(fallback_dir, f"{int(time.time())}_{os.path.basename(path)}")
        counter = 1
        while os.path.exists(dest):
            dest = os.path.join(fallback_dir, f"{int(time.time())}_{counter}_{os.path.basename(path)}")
            counter += 1
        shutil.move(path, dest)
        return True
    except Exception as e:
        log.warning("safe_delete fallback failed for %s: %s", path, e)
        return False


def reveal_in_file_manager(path, select=False):
    path = os.path.normpath(path)
    try:
        if IS_WINDOWS:
            subprocess.Popen(f'explorer /select,"{path}"' if select else f'explorer "{path}"')
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path] if select else ["open", path])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path) if select else path])
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open location:\n{e}")


def purge_registry_mru(target_filename):
    if not IS_WINDOWS:
        return
    try:
        mru = r"Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, mru, 0, winreg.KEY_READ | winreg.KEY_WRITE)
        to_delete = []
        for i in range(winreg.QueryInfoKey(key)[1]):
            try:
                name, data, _ = winreg.EnumValue(key, i)
                if isinstance(data, bytes) and target_filename.lower() in data.decode('utf-16le', 'ignore').lower():
                    to_delete.append(name)
            except Exception:
                continue
        for v in to_delete:
            winreg.DeleteValue(key, v)
        winreg.CloseKey(key)
    except Exception as e:
        log.info("registry MRU purge skipped: %s", e)


def hash_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def query_tokens(query):
    return [t.lower() for t in re.findall(r"\w+", query or "")]


def find_query_lines(text, query, limit=50):
    """VS Code-style: return [(line_no, line_text)] for lines that contain any
    query token. Used for result snippets and jumping the preview to the hit."""
    tokens = query_tokens(query)
    if not text or not tokens:
        return []
    hits = []
    for i, line in enumerate(text.splitlines(), 1):
        low = line.lower()
        if any(t in low for t in tokens):
            hits.append((i, line.strip()))
            if len(hits) >= limit:
                break
    return hits


def best_match_line(text, query, width=120):
    """The single most relevant line for a compact result snippet, prefixed with
    its line number (VS Code `L<n>:` style)."""
    lines = find_query_lines(text, query, limit=1)
    if not lines:
        return ""
    n, s = lines[0]
    s = (s[:width] + " …") if len(s) > width else s
    return f"L{n}: {s}"


# ===========================================================================
# Tkinter UI
# ===========================================================================
class TroveApp:
    NOTES_LABEL = "My Notes (Vault)"

    def __init__(self, root):
        self.root = root
        self.root.title("Trove — Document Search & OCR   ·   Assembler: KIMANI S.M.")
        self.root.geometry("1240x760")
        self._set_window_icon()

        self.vault_dir = os.path.join(os.path.expanduser("~"), "TextVault_Data")
        self.notes_dir = os.path.join(self.vault_dir, "Notes")
        self.trash_dir = os.path.join(self.vault_dir, "_RecycleBin")
        os.makedirs(self.notes_dir, exist_ok=True)

        self._init_logging()
        self.store = VaultStore(os.path.join(self.vault_dir, "vault.db"))
        self._migrate_legacy_json()

        self.indexed_folders = self.store.get_meta("indexed_folders", [])
        self.indexer = Indexer(self.store, self.vault_dir,
                               on_progress=self._progress_cb, on_done=self._done_cb)
        self.watcher = FolderWatcher(on_change=self._on_watch_change)

        # UI/runtime state
        self.current = None            # current file sqlite Row (or None)
        self.item_meta = {}            # tree item id -> ('file', id) | ('src', dir) | ('cat', dir, cat)
        self.loaded_nodes = set()      # lazily-populated node ids
        self.history = []              # navigation stack of file ids
        self.hist_pos = -1
        self._navigating = False
        self._search_timer = None
        self._autosave_timer = None
        self._watch_timer = None
        self._preview_path = None
        self._preview_offset = 0
        self._active_query = ""        # current search text, for hit highlighting

        self.setup_ui()
        self.populate_browse()
        self.root.after(200, self.auto_recall_indexed)

    def _set_window_icon(self):
        """Set the window icon from the bundled/source PNG if present (best-effort)."""
        try:
            bases = [getattr(sys, "_MEIPASS", None), os.path.dirname(os.path.abspath(__file__))]
            for base in bases:
                if not base:
                    continue
                png = os.path.join(base, "assets", "trove.png")
                if os.path.exists(png):
                    self._icon_img = tk.PhotoImage(file=png)
                    self.root.iconphoto(True, self._icon_img)
                    return
        except Exception as e:
            log.info("window icon not set: %s", e)

    def _capability_summary(self):
        return (f"Semantic: {'on' if SEMANTIC.available else 'off'}  ·  "
                f"OCR: {'on' if OCR.available else 'off'}  ·  "
                f"Watch: {self.watcher.backend}")

    def _init_logging(self):
        logging.basicConfig(
            filename=os.path.join(self.vault_dir, "vault.log"),
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(message)s",
        )

    def _migrate_legacy_json(self):
        legacy = os.path.join(self.vault_dir, "vault_metadata.json")
        if not self.store.is_empty() or not os.path.exists(legacy):
            return
        try:
            with open(legacy, encoding="utf-8") as f:
                data = json.load(f)
            self.store.set_meta("indexed_folders", data.get("indexed_folders", []))
            for _vname, meta in data.get("files", {}).items():
                op = meta.get("original_path")
                if not op:
                    continue
                ext = os.path.splitext(op)[1].lower()
                self.store.upsert({
                    "original_path": op, "vault_path": None,
                    "filename": os.path.basename(op), "ext": ext,
                    "category": MS_EXTENSIONS.get(ext, "Uncategorized"),
                    "source_dir": os.path.dirname(op), "editable": 0,
                    "mtime": meta.get("mtime", 0), "size": 0, "body": None,
                })
            self.store.commit()
            log.info("migrated %d legacy entries", len(data.get("files", {})))
        except Exception as e:
            log.warning("legacy migration skipped: %s", e)

    # -- UI construction -----------------------------------------------------
    def setup_ui(self):
        tb = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        tb.pack(side=tk.TOP, fill=tk.X)

        def btn(parent, text, cmd, side=tk.LEFT):
            b = tk.Button(parent, text=text, command=cmd)
            b.pack(side=side, padx=2, pady=2)
            return b

        btn(tb, "Index Drive/Folder", self.index_folder)
        self.btn_cancel = btn(tb, "Cancel", self.cancel_index)
        self.btn_cancel.config(state=tk.DISABLED)
        self.btn_watch = btn(tb, "👁 Watch: Off", self.toggle_watch)
        btn(tb, "New Note", self.new_note)
        self.btn_save = btn(tb, "Save", self.save_current)
        btn(tb, "Clear", self.clear_view)
        self.btn_back = btn(tb, "◀ Back", self.go_back)
        self.btn_fwd = btn(tb, "Forward ▶", self.go_forward)
        btn(tb, "Find Duplicates", self.find_duplicates)
        btn(tb, "Report", self.generate_report)
        btn(tb, "Open Vault", lambda: reveal_in_file_manager(self.vault_dir))

        # Search (right side): type filter + mode + entry + clear
        sf = tk.Frame(tb)
        sf.pack(side=tk.RIGHT, padx=5)
        tk.Label(sf, text="Type:").pack(side=tk.LEFT)
        self.ext_filter = tk.StringVar(value="All types")
        self.ext_combo = ttk.Combobox(sf, textvariable=self.ext_filter, values=["All types"],
                                      width=11, state="readonly")
        self.ext_combo.pack(side=tk.LEFT, padx=(2, 4))
        self.ext_combo.bind("<<ComboboxSelected>>", lambda e: self.run_search())
        modes = ["Full-text", "Filename"] + (["Hybrid"] if SEMANTIC.available else [])
        self.search_mode = tk.StringVar(value="Full-text")
        mode_cb = ttk.Combobox(sf, textvariable=self.search_mode, values=modes, width=9, state="readonly")
        mode_cb.pack(side=tk.LEFT, padx=(0, 4))
        mode_cb.bind("<<ComboboxSelected>>", lambda e: self.run_search())
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.on_search_changed)
        tk.Entry(sf, textvariable=self.search_var, width=24).pack(side=tk.LEFT)
        tk.Button(sf, text="✕", command=self.clear_search).pack(side=tk.LEFT, padx=(2, 0))

        # Status + progress bar row
        sb = tk.Frame(self.root)
        sb.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(sb, textvariable=self.status_var, anchor="w").pack(side=tk.LEFT, padx=6)
        self.progress = ttk.Progressbar(sb, mode="indeterminate", length=180)
        self.progress.pack(side=tk.RIGHT, padx=6, pady=2)
        tk.Label(sb, text=self._capability_summary(), anchor="e", fg="#888").pack(side=tk.RIGHT, padx=10)

        # Main split
        self.paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left = tk.Frame(self.paned)
        # Tree toolbar: group-by selector + multi-select delete
        tree_bar = tk.Frame(left)
        tree_bar.pack(side=tk.TOP, fill=tk.X)
        tk.Label(tree_bar, text="Group by:").pack(side=tk.LEFT, padx=(2, 0))
        self.group_mode = tk.StringVar(value="Extension")
        gcb = ttk.Combobox(tree_bar, textvariable=self.group_mode, values=["Extension", "Folder"],
                           width=10, state="readonly")
        gcb.pack(side=tk.LEFT, padx=4, pady=2)
        gcb.bind("<<ComboboxSelected>>", lambda e: self.populate_browse())
        tk.Button(tree_bar, text="🗑 Delete Selected", command=self.delete_selected).pack(side=tk.RIGHT, padx=2)

        cols = ("modified", "size", "type", "location")
        ys = ttk.Scrollbar(left, orient="vertical")
        xs = ttk.Scrollbar(left, orient="horizontal")
        self.tree = ttk.Treeview(left, columns=cols, selectmode="extended",
                                 yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.heading("#0", text="Name")
        self.tree.heading("modified", text="Date Modified")
        self.tree.heading("size", text="Size")
        self.tree.heading("type", text="Type")
        self.tree.heading("location", text="Location")
        self.tree.column("#0", width=280, stretch=True)
        self.tree.column("modified", width=130, anchor="w", stretch=False)
        self.tree.column("size", width=80, anchor="e", stretch=False)
        self.tree.column("type", width=150, anchor="w", stretch=False)
        self.tree.column("location", width=260, anchor="w", stretch=False)
        ys.config(command=self.tree.yview)
        xs.config(command=self.tree.xview)
        ys.pack(side=tk.RIGHT, fill=tk.Y)
        xs.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.tree.bind('<<TreeviewOpen>>', self.on_tree_open)
        self.tree.bind("<Button-3>", self.show_tree_menu)
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.tree_menu = tk.Menu(self.root, tearoff=0)
        self.tree_menu.add_command(label="Open Original Location", command=self.open_original_location)
        self.tree_menu.add_command(label="Delete Selected (Recycle Bin)", command=self.delete_selected)
        self.tree_menu.add_command(label="Remove from Index (keep file)", command=self.remove_from_index)
        self.paned.add(left, weight=1)

        self.notebook = ttk.Notebook(self.paned)
        self.paned.add(self.notebook, weight=3)
        self._build_viewer_tab()
        self._build_security_tab()
        self._build_purge_tab()

        # Shortcuts
        self.root.bind("<Control-f>", lambda e: self._focus_search())
        self.root.bind("<Control-s>", lambda e: self.save_current())
        self.root.bind("<Escape>", lambda e: self.clear_search())
        self.root.bind("<Alt-Left>", lambda e: self.go_back())
        self.root.bind("<Alt-Right>", lambda e: self.go_forward())
        self._update_nav_buttons()

    def _build_viewer_tab(self):
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text="Viewer")
        self.viewer_info = tk.StringVar(value="Select a file to preview.")
        tk.Label(frame, textvariable=self.viewer_info, anchor="w", fg="#3a5").pack(fill=tk.X, padx=4, pady=2)
        body = tk.Frame(frame)
        body.pack(fill=tk.BOTH, expand=True)
        ys = ttk.Scrollbar(body, orient="vertical")
        self.text = tk.Text(body, wrap=tk.WORD, font=("Consolas", 11), yscrollcommand=ys.set, undo=True)
        ys.config(command=self.text.yview)
        ys.pack(side=tk.RIGHT, fill=tk.Y)
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.text.tag_configure("hit", background="#fde047", foreground="#111827")
        self.text.bind("<KeyRelease>", self.on_edit)
        self.load_more_btn = tk.Button(frame, text="Load more ▼", command=self.load_more_preview)
        # packed only when a large .txt is truncated

    def _build_security_tab(self):
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text="Security & Irrelevant Scan")
        bar = tk.Frame(frame)
        bar.pack(fill=tk.X, pady=5)
        tk.Button(bar, text="Clear List", command=lambda: self.sec_tree.delete(*self.sec_tree.get_children())).pack(side=tk.LEFT)
        tk.Button(bar, text="Run Defender Scan on Selected", command=self.run_defender_scan).pack(side=tk.LEFT, padx=5)
        tk.Button(bar, text="Delete Selected File", command=self.delete_security_file).pack(side=tk.LEFT)
        self.sec_progress = ttk.Progressbar(bar, mode='indeterminate', length=180)
        self.sec_progress.pack(side=tk.LEFT, padx=10)
        self.sec_tree = ttk.Treeview(frame, columns=("Type", "Path"), show="headings", selectmode="extended")
        self.sec_tree.heading("Type", text="Flag Type")
        self.sec_tree.heading("Path", text="Original File Path")
        self.sec_tree.column("Type", width=110)
        self.sec_tree.column("Path", width=560)
        self.sec_tree.pack(fill=tk.BOTH, expand=True)

    def _build_purge_tab(self):
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text="Deep Purge & Registry")
        bar = tk.Frame(frame)
        bar.pack(fill=tk.X, pady=5)
        tk.Button(bar, text="Refresh List", command=self.refresh_purge_list).pack(side=tk.LEFT, padx=5)
        tk.Button(bar, text="Deep Delete Selected (Files + Registry)", command=self.execute_deep_purge).pack(side=tk.LEFT)
        self.purge_tree = ttk.Treeview(frame, columns=("File", "Original Path"), show="headings", selectmode="extended")
        self.purge_tree.heading("File", text="File Name")
        self.purge_tree.heading("Original Path", text="Original Path")
        self.purge_tree.column("File", width=200)
        self.purge_tree.column("Original Path", width=520)
        self.purge_tree.pack(fill=tk.BOTH, expand=True)
        self.purge_tree.bind("<Visibility>", lambda e: self.refresh_purge_list())

    # -- browse tree (lazy, columnar) ---------------------------------------
    def populate_browse(self):
        self.tree.delete(*self.tree.get_children())
        self.item_meta.clear()
        self.loaded_nodes.clear()
        self._active_query = ""
        self._refresh_ext_filter()
        if self.group_mode.get() == "Folder":
            for row in self.store.source_dirs():
                label = self.NOTES_LABEL if row["source_dir"] == self.notes_dir else row["source_dir"]
                node = self.tree.insert("", "end", text=f"{label}  ({row['n']})", open=False)
                self.item_meta[node] = ("src", row["source_dir"])
                self.tree.insert(node, "end", text="…")
        else:  # Extension — groups all files of a type across every subfolder
            for row in self.store.extensions():
                node = self.tree.insert("", "end", text=f"{self._ext_label(row['ext'])}  ({row['n']})", open=False)
                self.item_meta[node] = ("ext", row["ext"])
                self.tree.insert(node, "end", text="…")
        self.status_var.set(f"Ready — {self.store.count()} files indexed")

    def on_tree_open(self, _event):
        node = self.tree.focus()
        if node in self.loaded_nodes or node not in self.item_meta:
            return
        kind = self.item_meta[node]
        self.tree.delete(*self.tree.get_children(node))  # drop placeholder
        if kind[0] == "ext":
            for f in self.store.files_by_ext(kind[1]):
                self._insert_file_node(node, f)
        elif kind[0] == "src":
            for row in self.store.categories(kind[1]):
                cat = self.tree.insert(node, "end", text=f"{row['category']}  ({row['n']})", open=False)
                self.item_meta[cat] = ("cat", kind[1], row["category"])
                self.tree.insert(cat, "end", text="…")
        elif kind[0] == "cat":
            for f in self.store.files_in(kind[1], kind[2]):
                self._insert_file_node(node, f)
        self.loaded_nodes.add(node)

    # -- helpers: file rows, properties, labels ------------------------------
    def _insert_file_node(self, parent, f, extra_text=""):
        leaf = self.tree.insert(
            parent, "end", text=f"{f['filename']}{extra_text}",
            values=(self._fmt_time(f["mtime"]), self._human_size(f["size"]),
                    TYPE_LABELS.get(f["ext"], f["ext"] or "file"), f["source_dir"]),
        )
        self.item_meta[leaf] = ("file", f["id"])
        return leaf

    @staticmethod
    def _ext_label(ext):
        if ext in TYPE_LABELS:
            return f"{TYPE_LABELS[ext]} ({ext})"
        return f"{ext or '(no extension)'} files"

    @staticmethod
    def _fmt_time(mtime):
        try:
            return time.strftime("%Y-%m-%d %H:%M", time.localtime(mtime or 0))
        except Exception:
            return ""

    @staticmethod
    def _human_size(n):
        n = float(n or 0)
        for unit in ("B", "KB", "MB", "GB", "TB"):
            if n < 1024:
                return f"{int(n)} {unit}" if unit == "B" else f"{n:.1f} {unit}"
            n /= 1024
        return f"{n:.1f} PB"

    def _refresh_ext_filter(self):
        vals = ["All types"] + [r["ext"] for r in self.store.extensions()]
        self.ext_combo["values"] = vals
        if self.ext_filter.get() not in vals:
            self.ext_filter.set("All types")

    # -- search --------------------------------------------------------------
    def on_search_changed(self, *_):
        if self._search_timer:
            self.root.after_cancel(self._search_timer)
        self._search_timer = self.root.after(300, self.run_search)  # debounce

    def run_search(self):
        query = self.search_var.get().strip()
        self._active_query = query
        exts = None if self.ext_filter.get() == "All types" else [self.ext_filter.get()]

        # No text query: honour a standalone Type filter, else show the browse tree.
        if not query:
            if exts:
                self._show_type_listing(exts[0])
            else:
                self.populate_browse()
            return

        self.tree.delete(*self.tree.get_children())
        self.item_meta.clear()
        self.loaded_nodes.clear()

        mode = self.search_mode.get()
        if mode == "Filename":
            rows = self.store.search_filename(query, exts=exts)
        else:
            rows = self.store.search(query, exts=exts)          # keyword/FTS candidates
            if mode == "Hybrid":                                 # optional semantic re-rank
                rows = SEMANTIC.rerank(query, rows, text_of=lambda r: r["body"])

        scope = "" if not exts else f" in {exts[0]}"
        header = self.tree.insert("", "end", text=f"[{mode}] '{query}'{scope}  ({len(rows)})", open=True)
        self.item_meta[header] = ("results",)
        for f in rows:
            # VS Code-style: show the exact line the term was found on.
            snip = best_match_line(f["body"], query) if f["body"] else ""
            if not snip and mode != "Filename" and "snippet" in f.keys() and f["snippet"]:
                snip = f["snippet"]
            self._insert_file_node(header, f, extra_text=(f"   —   {snip}" if snip else ""))
        self.status_var.set(f"{len(rows)} match(es) · {mode} · '{query}'{scope}")

    def _show_type_listing(self, ext):
        self.tree.delete(*self.tree.get_children())
        self.item_meta.clear()
        self.loaded_nodes.clear()
        rows = self.store.files_by_ext(ext)
        header = self.tree.insert("", "end", text=f"All {self._ext_label(ext)}  ({len(rows)})", open=True)
        self.item_meta[header] = ("results",)
        for f in rows:
            self._insert_file_node(header, f)
        self.status_var.set(f"{len(rows)} {ext} file(s)")

    def clear_search(self):
        if self.search_var.get():
            self.search_var.set("")   # trace -> populate_browse
        else:
            self.populate_browse()

    def _focus_search(self):
        self.clear_search()
        self.status_var.set("Search focused — type to find text across all files")

    # -- selection / preview -------------------------------------------------
    def on_tree_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        meta = self.item_meta.get(sel[0])
        if meta and meta[0] == "file":
            self.open_file(meta[1])

    def open_file(self, file_id, add_history=True):
        row = self.store.get(file_id)
        if not row:
            return
        self.current = row
        if add_history and not self._navigating:
            self.history = self.history[: self.hist_pos + 1]
            self.history.append(file_id)
            self.hist_pos = len(self.history) - 1
        self._update_nav_buttons()
        self.render_preview(row)

    def render_preview(self, row):
        self.text.config(state=tk.NORMAL)
        self.text.delete(1.0, tk.END)
        self.load_more_btn.pack_forget()
        self._preview_path = None
        ext = row["ext"]
        path = row["original_path"]
        editable = bool(row["editable"]) and os.path.exists(path)

        if editable:  # in-app note
            self.viewer_info.set(f"✏  Editable note — {path}")
            try:
                with open(path, encoding="utf-8", errors="ignore") as f:
                    self.text.insert(tk.END, f.read())
            except Exception as e:
                self.text.insert(tk.END, f"[Could not read note: {e}]")
            self._highlight_hits()
            return

        # read-only preview
        if not os.path.exists(path):
            self.viewer_info.set(f"⚠  Original missing — {path}")
            self.text.insert(tk.END, row["body"] or "[File no longer at original location.]")
        elif ext in PLAINTEXT_EXTS:
            self._preview_path = path
            self._preview_offset = 0
            self.viewer_info.set(f"👁  Read-only preview — {path}")
            self.load_more_preview()
            return
        elif ext in OOXML_EXTS or ext in RTF_EXTS:
            self.viewer_info.set(f"👁  Extracted text preview — {path}")
            body = row["body"] if row["body"] is not None else TextExtractor.extract(path, ext)
            self.text.insert(tk.END, body or "[No extractable text found.]")
        elif ext in IMAGE_EXTS:
            note = "" if OCR.available else "   (install EasyOCR to read text inside images)"
            self.viewer_info.set(f"🖼  Image — OCR text{note} — {path}")
            self.text.insert(tk.END, (row["body"] or "").strip() or "[No OCR text — image indexed by name only.]")
        elif ext in PDF_EXTS:
            note = "" if _module_available("fitz") else "   (install PyMuPDF to read PDF text)"
            self.viewer_info.set(f"📄  PDF text{note} — {path}")
            body = row["body"] if row["body"] is not None else TextExtractor.extract(path, ext)
            self.text.insert(tk.END, (body or "").strip() or "[No extractable text.]")
        elif ext in LEGACY_OLE_EXTS:
            self.viewer_info.set(f"🔒  Legacy binary — {path}")
            self.text.insert(
                tk.END,
                f"Preview not available for legacy {ext} (pre-2007 OLE format).\n"
                "Right-click → Open Original Location to view in Microsoft Office."
            )
        self._highlight_hits()
        self.text.config(state=tk.DISABLED)

    def _highlight_hits(self):
        """Highlight all query-token occurrences in the preview and scroll to the
        first — the 'jump to the matching chunk' behaviour."""
        self.text.tag_remove("hit", "1.0", tk.END)
        first = None
        for tok in query_tokens(self._active_query):
            start = "1.0"
            while True:
                pos = self.text.search(tok, start, tk.END, nocase=True)
                if not pos:
                    break
                end = f"{pos}+{len(tok)}c"
                self.text.tag_add("hit", pos, end)
                if first is None:
                    first = pos
                start = end
        if first:
            self.text.see(first)

    def load_more_preview(self):
        if not self._preview_path:
            return
        self.text.config(state=tk.NORMAL)
        try:
            with open(self._preview_path, encoding="utf-8", errors="ignore") as f:
                f.seek(self._preview_offset)
                chunk = f.read(PREVIEW_CHUNK)
        except Exception as e:
            self.text.insert(tk.END, f"[Read error: {e}]")
            self.text.config(state=tk.DISABLED)
            return
        self.text.insert(tk.END, chunk)
        self._preview_offset += len(chunk.encode("utf-8", "ignore"))
        more = len(chunk) == PREVIEW_CHUNK
        self._highlight_hits()
        self.text.config(state=tk.DISABLED)
        if more:
            self.load_more_btn.pack(side=tk.BOTTOM, pady=3)
            self.viewer_info.set(f"👁  Read-only preview (showing first {self._preview_offset // 1000} KB) — {self._preview_path}")
        else:
            self.load_more_btn.pack_forget()

    # -- live folder watching ------------------------------------------------
    def toggle_watch(self):
        if self.watcher.running:
            self.watcher.stop()
            self.btn_watch.config(text="👁 Watch: Off")
            self.status_var.set("Folder watching stopped.")
            return
        folders = [f for f in self.indexed_folders if os.path.isdir(f)]
        if not folders:
            messagebox.showinfo("Watch", "Index a folder first, then enable watching.")
            return
        self.watcher.start(folders)
        self.btn_watch.config(text=f"👁 Watch: On ({self.watcher.backend})")
        self.status_var.set(f"Watching {len(folders)} folder(s) live via {self.watcher.backend}.")

    def _on_watch_change(self, folder):
        """Watcher thread → debounce → re-index the changed folder on the UI thread."""
        def kick():
            if not self.indexer.running:
                self._begin_index(folder, silent=True)
        try:
            if self._watch_timer:
                self.root.after_cancel(self._watch_timer)
            self._watch_timer = self.root.after(1500, kick)
        except Exception:
            pass

    # -- notes / editing -----------------------------------------------------
    def new_note(self):
        path = os.path.join(self.notes_dir, f"Draft_{int(time.time())}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write("New note...")
        fid = self.store.upsert({
            "original_path": path, "vault_path": path, "filename": os.path.basename(path),
            "ext": ".txt", "category": "Draft Notes", "source_dir": self.notes_dir,
            "editable": 1, "mtime": os.path.getmtime(path), "size": os.path.getsize(path),
            "body": "New note...",
        })
        self.store.commit()
        self.populate_browse()
        self.open_file(fid)

    def on_edit(self, event):
        if not self.current or not self.current["editable"]:
            return
        if event.keysym in ("Up", "Down", "Left", "Right", "Prior", "Next"):
            return
        if self._autosave_timer:
            self.root.after_cancel(self._autosave_timer)
        self._autosave_timer = self.root.after(600, self.save_current)

    def save_current(self):
        if not self.current or not self.current["editable"]:
            return
        path = self.current["original_path"]
        content = self.text.get(1.0, tk.END)
        if content.endswith("\n"):
            content = content[:-1]
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            self.store.upsert({
                "original_path": path, "vault_path": path, "filename": os.path.basename(path),
                "ext": ".txt", "category": TextExtractor.dominant_keyword(content) or "Draft Notes",
                "source_dir": self.notes_dir, "editable": 1, "mtime": os.path.getmtime(path),
                "size": len(content.encode("utf-8")), "body": content[:INDEX_TEXT_CAP],
            })
            self.store.commit()
            self.current = self.store.get_by_path(path)
            self.status_var.set(f"Saved {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))

    # -- navigation ----------------------------------------------------------
    def go_back(self):
        if self.hist_pos > 0:
            self.hist_pos -= 1
            self._navigate_to(self.history[self.hist_pos])

    def go_forward(self):
        if self.hist_pos < len(self.history) - 1:
            self.hist_pos += 1
            self._navigate_to(self.history[self.hist_pos])

    def _navigate_to(self, file_id):
        self._navigating = True
        try:
            self.open_file(file_id, add_history=False)
        finally:
            self._navigating = False
        self._update_nav_buttons()

    def _update_nav_buttons(self):
        self.btn_back.config(state=tk.NORMAL if self.hist_pos > 0 else tk.DISABLED)
        self.btn_fwd.config(state=tk.NORMAL if self.hist_pos < len(self.history) - 1 else tk.DISABLED)

    def clear_view(self):
        self.tree.selection_remove(self.tree.selection())
        self.text.config(state=tk.NORMAL)
        self.text.delete(1.0, tk.END)
        self.text.config(state=tk.DISABLED)
        self.load_more_btn.pack_forget()
        self.current = None
        self._preview_path = None
        self._active_query = ""
        self.viewer_info.set("Cleared. Select a file to preview.")
        self.status_var.set("Ready")

    # -- indexing (threaded) -------------------------------------------------
    def auto_recall_indexed(self):
        for folder in list(self.indexed_folders):
            if os.path.exists(folder) and not self.indexer.running:
                self._begin_index(folder, silent=True)

    def index_folder(self):
        target = filedialog.askdirectory(title="Select Drive or Folder to Index")
        if not target:
            return
        if self.indexer.running:
            messagebox.showinfo("Busy", "An index is already running. Cancel it first.")
            return
        if target not in self.indexed_folders:
            self.indexed_folders.append(target)
            self.store.set_meta("indexed_folders", self.indexed_folders)
        self._begin_index(target, silent=False)

    def _begin_index(self, target, silent):
        self._index_silent = silent
        if self.indexer.start(target):
            self.btn_cancel.config(state=tk.NORMAL)
            self.progress.start(12)
            self.status_var.set(f"Indexing {target} …")

    def cancel_index(self):
        if self.indexer.running:
            self.indexer.cancel()
            self.status_var.set("Cancelling…")

    def _progress_cb(self, scanned, added):
        self.root.after(0, lambda: self.status_var.set(f"Indexing… {added} added / {scanned} scanned"))

    def _done_cb(self, scanned, added, sec_flags, cancelled):
        self.root.after(0, lambda: self._index_finished(scanned, added, sec_flags, cancelled))

    def _index_finished(self, scanned, added, sec_flags, cancelled):
        self.progress.stop()
        self.btn_cancel.config(state=tk.DISABLED)
        for flag_type, path in sec_flags:
            self.sec_tree.insert("", "end", values=(flag_type, path))
        self.populate_browse()
        self.refresh_purge_list()
        state = "cancelled" if cancelled else "complete"
        self.status_var.set(f"Index {state} — {added} added, {scanned} scanned, {self.store.count()} total")
        if not getattr(self, "_index_silent", True):
            msg = f"Index {state}.\nAdded/updated {added} of {scanned} scanned files."
            if sec_flags:
                msg += f"\nFlagged {len(sec_flags)} executables/scripts (Security tab)."
            messagebox.showinfo("Indexing", msg)

    # -- context-menu actions ------------------------------------------------
    def show_tree_menu(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        # Keep an existing multi-selection; otherwise select what was clicked.
        if item not in self.tree.selection():
            self.tree.selection_set(item)
        self.tree_menu.tk_popup(event.x_root, event.y_root)

    # -- multi-select delete (files, subfolders, or whole extension groups) --
    def _collect_selected_file_ids(self):
        """Resolve the current selection (files AND group nodes) to a flat set of
        file ids, querying the store so lazily-unexpanded groups still resolve."""
        ids = set()
        for item in self.tree.selection():
            meta = self.item_meta.get(item)
            if not meta:
                continue
            if meta[0] == "file":
                ids.add(meta[1])
            elif meta[0] == "ext":
                ids.update(f["id"] for f in self.store.files_by_ext(meta[1]))
            elif meta[0] == "src":
                ids.update(f["id"] for f in self.store.files_by_source(meta[1]))
            elif meta[0] == "cat":
                ids.update(f["id"] for f in self.store.files_in(meta[1], meta[2]))
            elif meta[0] == "results":  # a search/type header: take its listed files
                for child in self.tree.get_children(item):
                    cm = self.item_meta.get(child)
                    if cm and cm[0] == "file":
                        ids.add(cm[1])
        return ids

    def delete_selected(self):
        ids = self._collect_selected_file_ids()
        if not ids:
            messagebox.showinfo("Delete", "Select one or more files, folders, or type groups first.")
            return
        if not messagebox.askyesno(
            "Confirm Delete",
            f"Send {len(ids)} file(s) to the Recycle Bin / Trash and remove them from the index?\n\n"
            "Originals are moved to the Recycle Bin (recoverable). File contents are never altered.",
            default="no", icon="warning",
        ):
            return
        cur_id = self.current["id"] if self.current else None
        removed = 0
        for fid in ids:
            row = self.store.get(fid)
            if not row:
                continue
            if safe_delete(row["original_path"], self.trash_dir):
                self.store.delete(fid)
                removed += 1
        if cur_id in ids:
            self.clear_view()
        self.populate_browse()
        self.refresh_purge_list()
        self.status_var.set(f"Deleted {removed} file(s) to the Recycle Bin")
        messagebox.showinfo("Deleted", f"Sent {removed} file(s) to the Recycle Bin / Trash.")

    def _selected_file_row(self):
        sel = self.tree.selection()
        if sel and self.item_meta.get(sel[0], ("",))[0] == "file":
            return self.store.get(self.item_meta[sel[0]][1])
        return None

    def open_original_location(self):
        row = self._selected_file_row()
        if not row:
            return
        if os.path.exists(row["original_path"]):
            reveal_in_file_manager(row["original_path"], select=True)
        else:
            messagebox.showwarning("Not Found", "The original file no longer exists.")

    def remove_from_index(self):
        row = self._selected_file_row()
        if not row:
            return
        if messagebox.askyesno("Remove", f"Remove '{row['filename']}' from the index?\n(The original file is NOT deleted.)"):
            self.store.delete(row["id"])
            self.populate_browse()
            self.refresh_purge_list()

    # -- duplicates ----------------------------------------------------------
    def find_duplicates(self):
        if getattr(self, "_dup_running", False):
            return
        self._dup_running = True
        self.status_var.set("Scanning for duplicates…")
        threading.Thread(target=self._dup_worker, daemon=True).start()

    def _dup_worker(self):
        # Hashing every file is I/O-heavy — never do it on the UI thread.
        hashes, dups = {}, []
        for f in self.store.all_files():
            p = f["original_path"]
            if not os.path.exists(p):
                continue
            try:
                h = hash_file(p)
            except Exception:
                continue
            if h in hashes:
                dups.append(f)
            else:
                hashes[h] = f
        self.root.after(0, lambda: self._dup_done(dups))

    def _dup_done(self, dups):
        self._dup_running = False
        self.status_var.set("Ready")
        if not dups:
            messagebox.showinfo("Duplicates", "No exact duplicate files found.")
            return
        if messagebox.askyesno("Remove Duplicates",
                               f"Found {len(dups)} duplicate(s). Send the redundant copies to the Recycle Bin?"):
            removed = 0
            for f in dups:
                if safe_delete(f["original_path"], self.trash_dir):
                    self.store.delete(f["id"])
                    removed += 1
            self.populate_browse()
            self.refresh_purge_list()
            messagebox.showinfo("Success", f"Removed {removed} duplicate file(s).")

    # -- report --------------------------------------------------------------
    def generate_report(self):
        path = os.path.join(self.vault_dir, f"Vault_Report_{int(time.time())}.txt")
        content = (
            "保管庫 (Vault) レポート\n============================\n作成者: KIMANI S.M.\n"
            f"インデックス済みファイル数: {self.store.count()}\n"
            f"インデックス済みフォルダ数: {len(self.indexed_folders)}\n\n対象フォルダ:\n"
            + "".join(f"- {d}\n" for d in self.indexed_folders)
        )
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        messagebox.showinfo("レポート作成", f"レポートが作成されました:\n{path}")

    # -- security tab --------------------------------------------------------
    def delete_security_file(self):
        sels = self.sec_tree.selection()
        if not sels:
            return
        if not messagebox.askyesno("Confirm Delete",
                                   f"Send {len(sels)} flagged file(s) to the Recycle Bin / Trash?\n"
                                   "They remain recoverable.", default="no", icon="warning"):
            return
        removed = 0
        for item in sels:
            path = self.sec_tree.item(item)["values"][1]
            if safe_delete(path, self.trash_dir):
                self.sec_tree.delete(item)
                removed += 1
        messagebox.showinfo("Done", f"Sent {removed} file(s) to the Recycle Bin / Trash.")

    def run_defender_scan(self):
        if not IS_WINDOWS:
            messagebox.showinfo("Unavailable", "Windows Defender scanning is only available on Windows.")
            return
        sels = self.sec_tree.selection()
        if not sels:
            return
        path = self.sec_tree.item(sels[0])["values"][1]
        self.sec_progress.start(10)
        threading.Thread(target=self._defender_thread, args=(path,), daemon=True).start()

    def _defender_thread(self, path):
        try:
            cmd = ["powershell", "-Command", f"Start-MpScan -ScanType CustomScan -ScanPath '{path}'"]
            subprocess.Popen(cmd, creationflags=0x08000000,
                             stdout=subprocess.PIPE, stderr=subprocess.PIPE).communicate()
            self.root.after(0, lambda: messagebox.showinfo("Scan Complete", f"Defender scan finished:\n{path}"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Defender failed:\n{e}"))
        finally:
            self.root.after(0, self.sec_progress.stop)

    # -- deep purge ----------------------------------------------------------
    def refresh_purge_list(self):
        self.purge_tree.delete(*self.purge_tree.get_children())
        for f in self.store.all_files():
            self.purge_tree.insert("", "end", iid=str(f["id"]), values=(f["filename"], f["original_path"]))

    def execute_deep_purge(self):
        sels = self.purge_tree.selection()
        if not sels:
            return
        prompt = (
            f"You are about to DEEP PURGE {len(sels)} file(s).\n\n"
            "This removes each file from the index AND its original location on disk, "
            "and scrubs Windows Recent-Docs registry traces.\n\n"
            "Deleted files go to the Recycle Bin / Trash where possible.\n\n"
            "Type  PURGE  (in capitals) to confirm:"
        )
        if simpledialog.askstring("Confirm Deep Purge", prompt, parent=self.root) != "PURGE":
            messagebox.showinfo("Cancelled", "Deep Purge cancelled — no files were deleted.")
            return
        purged = 0
        for iid in sels:
            row = self.store.get(int(iid))
            if not row:
                continue
            if row["vault_path"] and os.path.exists(row["vault_path"]):
                safe_delete(row["vault_path"], self.trash_dir)
            safe_delete(row["original_path"], self.trash_dir)
            purge_registry_mru(row["filename"])
            self.store.delete(row["id"])
            purged += 1
        self.populate_browse()
        self.refresh_purge_list()
        messagebox.showinfo("Deep Purge Complete", f"Deep purge executed on {purged} file(s).")


# Backwards-compatible alias (the app was formerly named "Vault Toolkit").
VaultToolkitApp = TroveApp


if __name__ == "__main__":
    root = tk.Tk()
    app = TroveApp(root)
    root.mainloop()
