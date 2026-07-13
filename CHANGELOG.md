# Changelog

All notable changes to this project are documented here.
This project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- **Group by extension type** — the browse tree can group *all* files of a type
  across every indexed folder and subfolder (toggle Extension ↔ Folder).
- **Type filter in search** — restrict a full-text or filename search to one
  extension (e.g. only `.xlsx`); with an empty query it lists all files of that type.
- **File property columns** — Date Modified, Size, Type, and Location now show
  next to each file in the tree.
- **Multi-select delete** — select any mix of files, subfolders, or whole
  extension groups and send them to the Recycle Bin (recoverable) in one action,
  via button, right-click, or the `Delete` key. File contents are never altered.

### Changed
- **Excel search now covers every worksheet** — `.xlsx` extraction reads the
  shared-string table *and* each sheet's inline strings, so a word in any sheet of
  a multi-worksheet workbook is indexed and searchable.

## [2.1.0] — 2026-07-06

### Fixed
- `find_duplicates` hashed every file on the UI thread (could freeze on large
  vaults) — moved to a background worker.

### Changed
- Filename search now runs as a SQL `LIKE` query instead of loading all rows.
- The `tkinter` import is guarded so the core layers import headlessly (for CI).

### Added
- `tests/` unit suite (extraction, catalog/FTS search, indexing, cancellation)
  and a GitHub Actions CI workflow running byte-compile + tests on 3.9 and 3.12.

[2.1.0]: https://github.com/mkubwaFiT/Office-Vault/releases/tag/v2.1.0

## [2.0.0] — 2026-07-01

Architecture rebuild for scale, search, and preview. **Breaking:** the data
model moved from a copy-everything vault + `vault_metadata.json` to an
index-in-place SQLite catalog (`vault.db`). The old JSON is imported
automatically on first run.

### Added
- **Full-text search (SQLite FTS5)** across `.txt` *and* Office documents, with
  highlighted result snippets and a filename/full-text toggle; `LIKE` fallback
  when FTS5 is unavailable.
- **OOXML text extraction** (`.docx/.xlsx/.pptx`) using only `zipfile` + `xml.etree`.
- **Real previews:** extracted-text preview for Office; paged (*Load more*) preview
  for large `.txt`; clear notice for legacy OLE `.doc/.xls/.ppt`.
- **Cancellable, streaming indexer** with live progress/count and a Cancel button.
- **Navigation/UX:** lazy-loaded tree, Back/Forward history, Clear, Clear-search,
  keyboard shortcuts (`Ctrl+F`, `Ctrl+S`, `Esc`, `Alt+←/→`), status bar, logging
  to `vault.log`.

### Changed
- Indexing is now **in place** (no whole-disk copying) — fixes hangs and storage
  bloat on large drives; the O(n²) metadata lookup is gone (indexed DB queries).
- Search is debounced and runs against the DB instead of re-reading files.
- Refactored the ~830-line god class into `TextExtractor` / `VaultStore` /
  `Indexer` / `VaultToolkitApp` layers; silent `except` blocks now log.

### Notes
- A packaged Windows build for v2 must be produced on Windows (`python build_vault.py`).

## [1.0.0] — 2026-06-25

First public release — the harmonized successor to the separate **TextVault**
and **OfficeVault** tools.

### Added
- Unified vault for `.txt` notes (editable, content-categorized, full-text
  search) and Microsoft Office binaries (read-only, corruption-safe) in one app.
- Background-threaded folder indexing so the window opens instantly even when
  whole drives are indexed.
- Recoverable deletes: every delete routes to the OS Recycle Bin / Trash
  (Windows `SHFileOperation`, macOS Finder, Linux freedesktop trash), with a
  local `_RecycleBin` fallback — the app never hard-`os.remove`s data.
- Type-to-confirm (`PURGE`) gate on Deep Purge; other delete dialogs default to "No".
- Chunked SHA-256 duplicate finder; on-demand Windows Defender scan of flagged files.
- Cross-platform "open location"; Windows-only features are platform-guarded.

### Changed
- Build switched from PyInstaller `--onefile` + UPX to **onedir, no-UPX** for
  fast startup and fewer antivirus false positives.

### Removed
- Legacy `txt_toolkit.py` / `office_toolkit.py` forks and their separate builders.

[1.0.0]: https://github.com/mkubwaFiT/Office-Vault/releases/tag/v1.0.0
