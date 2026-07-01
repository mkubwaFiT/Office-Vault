# Vault Toolkit

A single-window desktop tool for collecting, organizing, searching, and securely
purging text notes and Microsoft Office documents scattered across your drives.

It indexes folders you point it at, copies tracked files into a local vault,
groups them by source and type, lets you edit `.txt` notes inline, and provides
security/cleanup utilities (Defender scan, duplicate removal, deep purge with
Windows Recent-Docs registry scrubbing).

> This is the **harmonized** successor to the earlier `TextVault` and `OfficeVault`
> tools, merged into one codebase and rebuilt for fast startup.

---

## Download

**[Download the latest Windows build →](https://github.com/mkubwaFiT/Office-Vault/releases/latest)** (Windows 10/11, 64-bit)

Grab `Vault_Toolkit-vX.Y.Z-windows-x64.zip`, **extract the whole ZIP**, and run
`Vault_Toolkit/Vault_Toolkit.exe`. Keep the `.exe` and its `_internal/` folder
together — it's a fast-start *folder* build, not a lone `.exe`. A `.sha256` file
is attached if you want to verify the download. (First launch may show a
SmartScreen notice because the build is unsigned → *More info → Run anyway*.)

## Features

- **Full-text search across everything** — a SQLite **FTS5** index lets you find
  words *inside* `.txt` notes **and** Office documents (`.docx/.xlsx/.pptx` text is
  extracted at index time). Results show matching files with a highlighted snippet.
  Toggle between full-text and filename search; the box is debounced so typing
  never stalls. (Falls back to `LIKE` search if FTS5 is unavailable.)
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

## Architecture

The app is layered (single file, stdlib only):

| Layer | Responsibility |
|-------|----------------|
| `TextExtractor` | Stdlib text extraction for `.txt` + OOXML (`zipfile` + `xml.etree`). |
| `VaultStore` | SQLite catalog with an FTS5 virtual table (LIKE fallback). |
| `Indexer` | Cancellable, streaming, batched background disk scanner. |
| `VaultToolkitApp` | Tkinter UI: browse / search / preview / security / purge. |

Data lives in `~/TextVault_Data/`: a `vault.db` catalog, in-app notes under
`Notes/`, a `vault.log`, and a `_RecycleBin/` delete fallback. An existing
`vault_metadata.json` from v1 is imported automatically on first run.

## Performance

Startup and large-disk handling are addressed at two layers:

- **Build:** `Vault_Toolkit.spec` uses a **onedir, no-UPX** configuration instead
  of `--onefile` + UPX, avoiding a full runtime re-extraction (and Defender
  re-scan) on every launch.
- **Runtime:** indexing is **index-in-place**, **streamed** (batched SQLite
  commits), **cancellable**, and entirely off the UI thread; browsing lazy-loads
  the tree and search hits the DB instead of re-reading files — so neither a large
  index nor fast typing blocks the window.

## Requirements

- **Python 3.8+** with Tkinter (bundled with standard CPython on Windows/macOS).
- No third-party dependencies — standard library only.
- **PyInstaller** is needed only to build the `.exe` (`build_vault.py` installs it).

## Running from source

```bash
python vault_toolkit.py
```

The vault lives in `~/TextVault_Data/` (kept for backward compatibility with
existing metadata).

## Usage

1. **Index** — click *Index Drive/Folder* and choose a folder or drive. Tracked
   files (`.txt` + Office docs) are catalogued **in place** (originals are not
   moved or copied) with a live count; hit **Cancel** to stop a long scan.
   Indexed folders are remembered and re-synced in the background on next launch.
2. **Browse** — the left tree lazily loads by source folder → category. Click a
   file to preview it. Use **◀ Back / Forward ▶** (or `Alt+←/→`) to retrace, and
   **Clear** to reset the view.
3. **Search** — type to search **inside** every indexed file (full-text) or switch
   the dropdown to *Filename*. Results list matching files with a snippet; the box
   is debounced. `Ctrl+F` focuses it, `Esc` / **✕** clears it.
4. **Preview & notes** — Office and indexed `.txt` open **read-only** (large text
   pages in via *Load more*); right-click → *Open Original Location* to edit in the
   source app. **New Note** creates an editable note in the vault that auto-saves
   (`Ctrl+S` to save now).
5. **Tidy up** — *Find Duplicates* recycles exact-duplicate copies; *Report* writes
   a summary; *Open Vault* opens the data folder.
6. **Security tab** — review executables/scripts flagged during indexing; run a
   *Defender Scan*, or send a flagged file to the Recycle Bin.
7. **Deep Purge tab** — select files and *Deep Delete* to remove them from the
   index, their original location, and the Windows RecentDocs registry. This
   action asks you to type `PURGE` to confirm.

All deletions go to the Recycle Bin / Trash — see [Safety](#safety).

## Building the executable (Windows)

PyInstaller produces a binary for the OS it runs on, so build on **Windows**:

```bash
python build_vault.py
```

Output: `dist/Vault_Toolkit/Vault_Toolkit.exe`. Distribute the **entire**
`dist/Vault_Toolkit/` folder (zip it) — onedir builds are a folder, not a lone `.exe`.

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
| `vault_toolkit.py` | The application (Tkinter, single file). |
| `Vault_Toolkit.spec` | PyInstaller spec — fast-launch onedir/no-UPX config. |
| `build_vault.py` | Build script: cleans artifacts and builds from the spec. |
| `README.md` | This document. |
| `CHANGELOG.md` | Version history. |
| `LICENSE` | MIT license. |
| `.gitignore` | Excludes build output, caches, executables, and vault data. |

## License

Released under the [MIT License](LICENSE) — © 2026 KIMANI S.M.

## Credits

Assembler: **KIMANI S.M.**
