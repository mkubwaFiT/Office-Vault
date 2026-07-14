<p align="center">
  <img src="assets/trove.png" width="120" alt="Trove icon">
</p>

# Trove

[![CI](https://github.com/mkubwaFiT/Office-Vault/actions/workflows/ci.yml/badge.svg)](https://github.com/mkubwaFiT/Office-Vault/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/mkubwaFiT/Office-Vault)](https://github.com/mkubwaFiT/Office-Vault/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

**Trove** — a single-window desktop tool that indexes, deep-searches, previews, and
cleans up the documents scattered across your drives: text notes, Microsoft Office
files, PDFs, and images. It catalogs folders you point it at (in place — no copying),
searches *inside* every file (all Excel worksheets included), and provides
security/cleanup utilities (Defender scan, duplicate removal, deep purge).

> Trove is the evolution of the earlier `TextVault` / `OfficeVault` / "Vault Toolkit"
> tools. The GitHub repository keeps its original name, **Office-Vault**.

The **core is standard-library only** — lean, fast, ~30 MB. Heavier capabilities
(semantic search, OCR) are **optional extras** that activate only if you install
them (see [Optional power features](#optional-power-features)).

---

## Download

**[Download the latest Windows build →](https://github.com/mkubwaFiT/Office-Vault/releases/latest)** (Windows 10/11, 64-bit)

Grab `Trove-vX.Y.Z-windows-x64.zip`, **extract the whole ZIP**, and run
`Trove/Trove.exe`. Keep the `.exe` and its `_internal/` folder together — it's a
fast-start *folder* build, not a lone `.exe`. A `.sha256` file is attached if you
want to verify the download. (First launch may show a SmartScreen notice because
the build is unsigned → *More info → Run anyway*.)

## Features

- **Full-text search across everything** — a SQLite **FTS5** index lets you find
  words *inside* `.txt` notes **and** Office documents (`.docx/.xlsx/.pptx` text is
  extracted at index time). Excel search covers **every worksheet** of a workbook
  (shared strings *and* per-sheet inline strings). Results show matching files with
  a highlighted snippet. Toggle full-text/filename search and restrict it to a
  single **extension type**; the box is debounced. (Falls back to `LIKE` if FTS5 is
  unavailable.)
- **Organize by type + file properties** — group the tree by **extension type**
  (all files of a kind across every folder and subfolder) or by folder, with
  **Date Modified, Size, Type, and Location** columns per file.
- **Multi-select cleanup** — select any mix of files, subfolders, or whole
  extension groups and send them to the Recycle Bin in one action (button,
  right-click, or `Delete`). Recoverable; contents are never altered.
- **Jump-to-match preview** — search results show the exact line the term was
  found on (VS Code style), and opening a hit scrolls to and **highlights** every
  occurrence in the preview.
- **Live folder watching** — toggle **👁 Watch** to keep the index current as
  files change. Uses `watchdog` if installed, else a built-in stdlib poll.
- **PDF & image awareness** — `.pdf`, `.png`, `.jpg`, `.jpeg` are catalogued and
  filterable; their text becomes searchable when the optional OCR/PDF extras are
  installed (see below).
- **Scales to large disks — index in place** — indexing catalogs files where they
  live instead of copying every document into the vault, and runs in a
  **cancellable** background scan with a live progress/count and a **Cancel** button.
- **Real previews** — `.docx/.xlsx/.pptx` show extracted text; large `.txt` files
  page in on demand (*Load more*) so nothing freezes; legacy `.doc/.xls/.ppt`
  (pre-2007 OLE) show a clear "open in Office" notice.
- **Navigation & editing** — lazy-loaded tree grouped by source → category,
  **Back/Forward** history, **Clear** / **Clear-search**, and shortcuts
  (`Ctrl+F`, `Ctrl+S`, `Esc`, `Alt+←/→`). In-app notes are editable with auto-save;
  indexed documents are read-only.
- **Duplicate finder** — chunked SHA-256 hashing flags and removes exact duplicates.
- **Security tab** — executables/scripts found while indexing (`.exe`, `.bat`,
  `.ps1`, `.vbs`, `.scr`, `.dll`, `.js`, `.wsf`) are flagged; run an on-demand
  Windows Defender scan or send them to the Recycle Bin.
- **Deep Purge** — remove a file from the index, its original location, **and**
  scrub its traces from the Windows Explorer `RecentDocs` MRU registry.
- **Vault report** — generate a summary report (Japanese-localized).

## Optional power features

The core needs nothing beyond the standard library. These extras are **auto-detected
at runtime** — install only what you want; without them Trove falls back to keyword
search / no-OCR and runs exactly as before. The status bar shows what's active.

```bash
pip install -r requirements-optional.txt     # or pick individual lines
```

| Extra | Unlocks | Fallback when absent |
|-------|---------|----------------------|
| `sentence-transformers`, `numpy` | **Hybrid search** — a *Hybrid* mode that re-ranks keyword hits by meaning (`all-MiniLM-L6-v2`) | Full-text (FTS5) keyword search |
| `easyocr` | **OCR** — read text inside `.png/.jpg/.jpeg` and scanned PDFs | Images indexed by filename only |
| `pymupdf` | **PDF text** (+ embedded-image OCR) | PDFs indexed by filename only |
| `watchdog` | **Event-driven** live folder watching | Built-in stdlib mtime-poll |

> These pull in PyTorch (~2 GB) and are **not** in the shipped `.exe`, which stays
> lean. Install them into a Python environment and run Trove from source to use them.

## Architecture

The app is layered (single file; core is stdlib-only):

| Layer | Responsibility |
|-------|----------------|
| `TextExtractor` | Text extraction for `.txt` + OOXML (`zipfile` + `xml.etree`); optional PDF/OCR. |
| `VaultStore` | SQLite catalog with an FTS5 virtual table (LIKE fallback). |
| `Indexer` | Cancellable, streaming, batched background disk scanner. |
| `FolderWatcher` | Live re-indexing — `watchdog` if present, else stdlib polling. |
| `SemanticRanker` / `OcrEngine` | Optional ML engines; no-ops if their libs are absent. |
| `TroveApp` | Tkinter UI: browse / search / preview / security / purge. |

Data lives in `~/TextVault_Data/`: a `vault.db` catalog, in-app notes under
`Notes/`, a `vault.log`, and a `_RecycleBin/` delete fallback. An existing
`vault_metadata.json` from v1 is imported automatically on first run.

## Performance

Startup and large-disk handling are addressed at two layers:

- **Build:** `Trove.spec` uses a **onedir, no-UPX** configuration instead
  of `--onefile` + UPX, avoiding a full runtime re-extraction (and Defender
  re-scan) on every launch.
- **Runtime:** indexing is **index-in-place**, **streamed** (batched SQLite
  commits), **cancellable**, and entirely off the UI thread; browsing lazy-loads
  the tree and search hits the DB instead of re-reading files — so neither a large
  index nor fast typing blocks the window.

## Requirements

- **Python 3.8+** with Tkinter (bundled with standard CPython on Windows/macOS).
- No third-party dependencies — standard library only.
- **PyInstaller** is needed only to build the `.exe` (`build_trove.py` installs it).

## Running from source

```bash
python trove.py
```

The vault lives in `~/TextVault_Data/` (kept for backward compatibility with
existing metadata).

## Usage

1. **Index** — click *Index Drive/Folder* and choose a folder or drive. Tracked
   files (`.txt` + Office docs) are catalogued **in place** (originals are not
   moved or copied) with a live count; hit **Cancel** to stop a long scan.
   Indexed folders are remembered and re-synced in the background on next launch.
2. **Browse** — pick **Group by: Extension** to see every file of a type across all
   folders/subfolders, or **Folder** for the source → category view. Each file shows
   Date Modified, Size, Type, and Location. Click a file to preview; **◀ Back /
   Forward ▶** (or `Alt+←/→`) retrace, **Clear** resets.
3. **Search** — type to search **inside** every indexed file (full-text, all Excel
   worksheets included) or switch to *Filename*. The **Type** dropdown restricts the
   search to one extension. Results list matching files with a snippet; the box is
   debounced. `Ctrl+F` focuses it, `Esc` / **✕** clears it.
4. **Search inside files** — the *Hybrid* mode (if the semantic extra is installed)
   re-ranks results by meaning; opening a hit **highlights** every match and jumps
   to it. Toggle **👁 Watch** to keep the index live as files change.
5. **Clean up junk** — select any mix of files, subfolders, or extension groups and
   click **🗑 Delete Selected** (or press `Delete`) to send them to the Recycle Bin.
6. **Preview & notes** — indexed files open **read-only** (large text pages in via
   *Load more*); right-click → *Open Original Location* to edit in the source app.
   **New Note** creates an editable note in the vault that auto-saves (`Ctrl+S`).
7. **Tidy up** — *Find Duplicates* recycles exact-duplicate copies; *Report* writes
   a summary; *Open Vault* opens the data folder.
8. **Security & Deep Purge tabs** — review flagged executables (Defender scan or
   recycle), or *Deep Delete* files from the index + disk + RecentDocs registry
   (type `PURGE` to confirm).

All deletions go to the Recycle Bin / Trash — see [Safety](#safety).

## Building the executable (Windows)

PyInstaller produces a binary for the OS it runs on, so build on **Windows**:

```bash
python build_trove.py
```

Output: `dist/Trove/Trove.exe`. Distribute the **entire** `dist/Trove/` folder
(zip it) — onedir builds are a folder, not a lone `.exe`. (Or just push a `v*`
tag and GitHub Actions builds and publishes it for you.)

## Platform notes

The toolkit targets **Windows** (registry MRU scrubbing, Explorer integration,
Defender scanning, and Office binary handling). Those Windows-only features are
guarded behind a platform check, so the app still **imports and runs on
macOS/Linux** for development — the unavailable features no-op or show a notice,
and "open location" falls back to `open`/`xdg-open`.

## Safety

Destructive actions are designed to be **recoverable** and **hard to trigger by accident**:

- **All deletions go to the Recycle Bin / Trash**, never a permanent wipe — vault
  deletes, duplicate removal, the Security-tab delete, and Deep Purge alike. If
  the OS trash is unavailable, the file is moved to a local
  `~/TextVault_Data/_RecycleBin/` backup instead. The app never hard-deletes data.
- **Deep Purge requires type-to-confirm.** Because it also removes the *original*
  file on disk and scrubs registry traces, you must type `PURGE` to proceed — a
  reflexive "Yes" click is not enough.
- **Confirmation dialogs default to "No"**, so pressing Enter or mis-clicking cancels.

The Recycle-Bin behavior is dependency-free: Windows `SHFileOperation` (`FOF_ALLOWUNDO`)
via `ctypes`, macOS Finder via AppleScript, Linux freedesktop trash.

User content is never committed to git — `TextVault_Data/`, `vault_metadata.json`,
build artifacts, and `*.exe` are all `.gitignore`d.

## Repository layout

| File | Purpose |
|------|---------|
| `trove.py` | The application (Tkinter, single file; stdlib core + optional engines). |
| `Trove.spec` | PyInstaller spec — fast-launch onedir/no-UPX config. |
| `build_trove.py` | Build script: cleans artifacts and builds from the spec. |
| `assets/` | App icon (`trove.ico`/`.png`) + its stdlib generator `make_icon.py`. |
| `requirements-optional.txt` | Optional ML/OCR/watch extras (auto-detected). |
| `tests/` | Headless unit tests for the core layers. |
| `.github/workflows/` | CI (tests) + auto Windows build/release. |
| `README.md` · `CHANGELOG.md` · `LICENSE` · `.gitignore` | Docs, history, license, ignores. |

## License

Released under the [MIT License](LICENSE) — © 2026 KIMANI S.M.

## Credits

Assembler: **KIMANI S.M.**
