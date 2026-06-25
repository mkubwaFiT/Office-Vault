# Changelog

All notable changes to this project are documented here.
This project adheres to [Semantic Versioning](https://semver.org/).

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
