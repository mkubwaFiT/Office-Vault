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

## Features

- **Unified indexing** — point it at any folder or drive; it syncs both
  `.txt` notes and MS Office binaries (`.doc/.docx`, `.xls/.xlsx`, `.ppt/.pptx`)
  into a local vault and remembers indexed folders across sessions.
- **Smart organization** — files are grouped by source folder, then by category:
  `.txt` notes are auto-categorized by their dominant keyword; Office files by type.
- **Inline editor** — edit and auto-save `.txt` notes. Office binaries open
  **read-only** with a clear notice, so they can never be corrupted by the editor.
- **Search** — filter the tree live; `.txt` matches on filename *and* content,
  Office files match on filename.
- **Duplicate finder** — chunked SHA-256 hashing flags and removes exact duplicates.
- **Security tab** — executables/scripts found while indexing (`.exe`, `.bat`,
  `.ps1`, `.vbs`, `.scr`, `.dll`, `.js`, `.wsf`) are flagged; run an on-demand
  Windows Defender scan or delete them.
- **Deep Purge** — remove a file from the vault, its original location, **and**
  scrub its traces from the Windows Explorer `RecentDocs` MRU registry.
- **Vault report** — generate a summary report (Japanese-localized).

## Performance

The previous builds launched slowly. This version fixes that at two layers:

- **Build:** `Vault_Toolkit.spec` uses a **onedir, no-UPX** configuration instead
  of `--onefile` + UPX. One-file builds re-extract the entire ~11 MB Python/Tk
  runtime into `%TEMP%` on *every* launch, and UPX both decompresses at launch and
  frequently trips Windows Defender into re-scanning the binary. Onedir runs
  straight from its folder.
- **Runtime:** folder re-indexing runs on a **background thread after the window
  paints**, so startup is instant even when whole drives are indexed. All
  filesystem work happens off the UI thread; results are marshaled back safely.

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

Several actions are **destructive and irreversible**:

- **Deep Purge** deletes the original file from disk and edits the registry.
- **Delete Selected File** (Security tab) permanently removes files from the system.

Each is gated behind a confirmation dialog. User content is never committed to
git — `TextVault_Data/`, `vault_metadata.json`, build artifacts, and `*.exe` are
all `.gitignore`d.

## Repository layout

| File | Purpose |
|------|---------|
| `vault_toolkit.py` | The application (Tkinter, single file). |
| `Vault_Toolkit.spec` | PyInstaller spec — fast-launch onedir/no-UPX config. |
| `build_vault.py` | Build script: cleans artifacts and builds from the spec. |
| `.gitignore` | Excludes build output, caches, executables, and vault data. |

## Credits

Assembler: **KIMANI S.M.**
