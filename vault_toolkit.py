import os
import sys
import shutil
import re
import time
import json
import subprocess
import hashlib
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from collections import Counter

# winreg / explorer / Defender are Windows-only. Guard the import so the
# toolkit stays importable (and the non-Windows features degrade gracefully)
# on other platforms during development.
IS_WINDOWS = sys.platform.startswith("win")
if IS_WINDOWS:
    import winreg

# Common words to ignore when grouping .txt notes by relevance
STOP_WORDS = set([
    "the", "and", "to", "a", "of", "in", "it", "is", "for", "that", "on", "you",
    "this", "with", "as", "at", "by", "not", "be", "are", "from", "but", "have",
    "an", "which", "was", "or", "we", "can", "if", "your", "has", "will", "all"
])

# Microsoft Office Extensions Mapping (binary formats: read-only in the editor)
MS_EXTENSIONS = {
    '.doc': 'Word Documents',
    '.docx': 'Word Documents',
    '.xls': 'Excel Spreadsheets',
    '.xlsx': 'Excel Spreadsheets',
    '.ppt': 'PowerPoint Presentations',
    '.pptx': 'PowerPoint Presentations',
}

# File types the unified vault tracks: editable .txt notes + MS Office binaries
TRACKED_EXTS = set(MS_EXTENSIONS) | {'.txt'}

DANGER_EXTS = {'.exe', '.bat', '.ps1', '.vbs', '.scr', '.dll', '.js', '.wsf'}


class VaultToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vault Toolkit (Text + Office) - Assembler: KIMANI S.M.")
        self.root.geometry("1100x700")

        # Setup Vault Directory
        self.vault_dir = os.path.join(os.path.expanduser("~"), "TextVault_Data")
        if not os.path.exists(self.vault_dir):
            os.makedirs(self.vault_dir)

        # Caching and metadata memory
        self.metadata_file = os.path.join(self.vault_dir, "vault_metadata.json")
        self.metadata = self.load_metadata()
        self.file_cache = {}
        self.indexed_folders = self.metadata.get("indexed_folders", [])

        self.current_file = None
        self.file_index = {}

        # State for debouncing, security and thread-safe indexing
        self._autosave_timer = None
        self.suspicious_files = []
        self._meta_lock = threading.Lock()

        self.setup_ui()

        # Paint the window first, then re-sync indexed folders in the
        # background so startup is instant even when whole drives are indexed.
        self.refresh_index()
        self.refresh_purge_list()
        self.root.after(150, self.auto_recall_indexed)

    def load_metadata(self):
        if os.path.exists(self.metadata_file):
            try:
                with open(self.metadata_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {"files": {}, "indexed_folders": []}

    def save_metadata(self):
        self.metadata["indexed_folders"] = self.indexed_folders
        with open(self.metadata_file, 'w', encoding='utf-8') as f:
            json.dump(self.metadata, f, indent=4)

    def setup_ui(self):
        toolbar = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        tk.Button(toolbar, text="Index Drive/Folder", command=self.index_folder).pack(side=tk.LEFT, padx=2, pady=2)
        tk.Button(toolbar, text="New Note", command=self.new_note).pack(side=tk.LEFT, padx=2, pady=2)
        tk.Button(toolbar, text="Save Current", command=self.save_current).pack(side=tk.LEFT, padx=2, pady=2)

        tk.Button(toolbar, text="Find & Remove Duplicates", command=self.find_duplicates).pack(side=tk.LEFT, padx=2, pady=2)
        tk.Button(toolbar, text="Vault Report (日本語)", command=self.generate_japanese_report).pack(side=tk.LEFT, padx=2, pady=2)

        tk.Button(toolbar, text="Open Reports", command=self.open_report_location).pack(side=tk.LEFT, padx=2, pady=2)
        tk.Button(toolbar, text="Deep Reset Reports", command=self.deep_reset_reports).pack(side=tk.LEFT, padx=2, pady=2)

        # Status label (right of toolbar) for non-blocking background indexing
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(toolbar, textvariable=self.status_var, anchor="e").pack(side=tk.RIGHT, padx=8)

        search_frame = tk.Frame(toolbar)
        search_frame.pack(side=tk.RIGHT, padx=5)
        tk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.on_search)
        tk.Entry(search_frame, textvariable=self.search_var, width=30).pack(side=tk.LEFT)

        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left Pane: File Tree
        left_frame = tk.Frame(self.paned_window)
        tree_y_scroll = ttk.Scrollbar(left_frame, orient="vertical")
        tree_x_scroll = ttk.Scrollbar(left_frame, orient="horizontal")

        self.tree = ttk.Treeview(left_frame, yscrollcommand=tree_y_scroll.set, xscrollcommand=tree_x_scroll.set, selectmode="extended")

        tree_y_scroll.config(command=self.tree.yview)
        tree_x_scroll.config(command=self.tree.xview)

        tree_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        tree_x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_file_select)

        self.tree_menu = tk.Menu(self.root, tearoff=0)
        self.tree_menu.add_command(label="Delete Selected File(s)", command=self.delete_file)
        self.tree_menu.add_command(label="Open Original Location", command=self.open_original_location)
        self.tree.bind("<Button-3>", self.show_tree_menu)

        self.paned_window.add(left_frame, weight=1)

        # Right Pane Notebook
        self.right_notebook = ttk.Notebook(self.paned_window)
        self.paned_window.add(self.right_notebook, weight=3)

        editor_frame = tk.Frame(self.right_notebook)
        self.right_notebook.add(editor_frame, text="Editor")

        text_y_scroll = ttk.Scrollbar(editor_frame, orient="vertical")
        text_x_scroll = ttk.Scrollbar(editor_frame, orient="horizontal")

        self.text_editor = tk.Text(editor_frame, wrap=tk.NONE, font=("Consolas", 11),
                                   yscrollcommand=text_y_scroll.set, xscrollcommand=text_x_scroll.set, undo=True)

        text_y_scroll.config(command=self.text_editor.yview)
        text_x_scroll.config(command=self.text_editor.xview)

        text_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text_x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.text_editor.pack(fill=tk.BOTH, expand=True)
        self.text_editor.bind("<KeyRelease>", self.auto_save_draft)

        self.editor_menu = tk.Menu(self.root, tearoff=0)
        self.editor_menu.add_command(label="Copy", command=lambda: self.text_editor.event_generate("<<Copy>>"))
        self.editor_menu.add_command(label="Cut", command=lambda: self.text_editor.event_generate("<<Cut>>"))
        self.editor_menu.add_command(label="Paste", command=lambda: self.text_editor.event_generate("<<Paste>>"))
        self.editor_menu.add_separator()
        self.editor_menu.add_command(label="Undo", command=lambda: self.text_editor.event_generate("<<Undo>>"))
        self.editor_menu.add_command(label="Redo", command=lambda: self.text_editor.event_generate("<<Redo>>"))
        self.text_editor.bind("<Button-3>", self.show_editor_menu)

        # Security Scanner UI Tab
        self.security_frame = tk.Frame(self.right_notebook)
        self.right_notebook.add(self.security_frame, text="Security & Irrelevant Scan")

        sec_toolbar = tk.Frame(self.security_frame)
        sec_toolbar.pack(fill=tk.X, pady=5)
        tk.Button(sec_toolbar, text="Clear List", command=self.clear_security_list).pack(side=tk.LEFT)
        tk.Button(sec_toolbar, text="Run Defender Scan on Selected", command=self.run_defender_scan).pack(side=tk.LEFT, padx=5)
        tk.Button(sec_toolbar, text="Delete Selected File", command=self.delete_security_file).pack(side=tk.LEFT)

        self.scan_progress = ttk.Progressbar(sec_toolbar, mode='indeterminate', length=200)
        self.scan_progress.pack(side=tk.LEFT, padx=10)

        self.sec_tree = ttk.Treeview(self.security_frame, columns=("Type", "Path"), show="headings", selectmode="extended")
        self.sec_tree.heading("Type", text="Flag Type")
        self.sec_tree.heading("Path", text="Original File Path")
        self.sec_tree.column("Type", width=100)
        self.sec_tree.column("Path", width=500)
        self.sec_tree.pack(fill=tk.BOTH, expand=True)

        # Deep Purge & Registry Scan Tab
        self.deep_purge_frame = tk.Frame(self.right_notebook)
        self.right_notebook.add(self.deep_purge_frame, text="Deep Purge & Registry")

        purge_toolbar = tk.Frame(self.deep_purge_frame)
        purge_toolbar.pack(fill=tk.X, pady=5)
        tk.Button(purge_toolbar, text="Refresh Indexed List", command=self.refresh_purge_list).pack(side=tk.LEFT, padx=5)
        tk.Button(purge_toolbar, text="Deep Delete Selected (Files + Registry)", command=self.execute_deep_purge).pack(side=tk.LEFT)

        self.purge_tree = ttk.Treeview(self.deep_purge_frame, columns=("File", "Vault Path", "Original Path"), show="headings", selectmode="extended")
        self.purge_tree.heading("File", text="File Name")
        self.purge_tree.heading("Vault Path", text="Vault Path")
        self.purge_tree.heading("Original Path", text="Original Path")
        self.purge_tree.column("File", width=150)
        self.purge_tree.column("Vault Path", width=250)
        self.purge_tree.column("Original Path", width=250)
        self.purge_tree.pack(fill=tk.BOTH, expand=True)
        self.purge_tree.bind("<Visibility>", lambda e: self.refresh_purge_list())

    def show_tree_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            if item not in self.tree.selection():
                self.tree.selection_set(item)
            self.tree_menu.tk_popup(event.x_root, event.y_root)

    def show_editor_menu(self, event):
        self.editor_menu.tk_popup(event.x_root, event.y_root)

    def delete_file(self):
        selections = self.tree.selection()
        if not selections:
            return

        if messagebox.askyesno("Delete", f"Are you sure you want to delete {len(selections)} selected file(s) from the Vault?"):
            deleted_count = 0
            for item_id in selections:
                if item_id in self.file_index:
                    filepath = self.file_index[item_id]
                    filename = os.path.basename(filepath)
                    try:
                        if os.path.exists(filepath):
                            self._safe_delete(filepath)
                        if filename in self.metadata["files"]:
                            del self.metadata["files"][filename]
                        if filepath in self.file_cache:
                            del self.file_cache[filepath]

                        if self.current_file == filepath:
                            self.text_editor.config(state=tk.NORMAL)
                            self.text_editor.delete(1.0, tk.END)
                            self.current_file = None
                        deleted_count += 1
                    except Exception as e:
                        print(f"Failed to delete {filename}: {e}")
            self.save_metadata()
            self.refresh_index(self.search_var.get())
            if deleted_count > 0:
                messagebox.showinfo("Success", f"Deleted {deleted_count} file(s).")

    def open_original_location(self):
        selected = self.tree.selection()
        if not selected:
            return
        item_id = selected[0]
        if item_id in self.file_index:
            filepath = self.file_index[item_id]
            filename = os.path.basename(filepath)
            file_meta = self.metadata["files"].get(filename)

            if file_meta and "original_path" in file_meta:
                orig_path = file_meta["original_path"]
                if os.path.exists(orig_path):
                    self._reveal_in_file_manager(orig_path, select=True)
                else:
                    messagebox.showwarning("Not Found", "The original file no longer exists at its source location.")
            else:
                messagebox.showinfo("Info", "This file was created in the Vault or its original location is unknown.")

    def open_report_location(self):
        try:
            self._reveal_in_file_manager(self.vault_dir, select=False)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open location:\n{e}")

    def _reveal_in_file_manager(self, path, select=False):
        """Open the OS file manager at (or selecting) a path. Cross-platform."""
        path = os.path.normpath(path)
        try:
            if IS_WINDOWS:
                if select:
                    subprocess.Popen(f'explorer /select,"{path}"')
                else:
                    subprocess.Popen(f'explorer "{path}"')
            elif sys.platform == "darwin":
                args = ["open", "-R", path] if select else ["open", path]
                subprocess.Popen(args)
            else:
                target = os.path.dirname(path) if select else path
                subprocess.Popen(["xdg-open", target])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open location:\n{e}")

    # ------------------------------------------------------------------
    # Recoverable deletion: route everything through the OS Recycle Bin /
    # Trash so an accidental click is never an irreversible wipe. Falls back
    # to a local _RecycleBin folder; NEVER does a hard os.remove() of data.
    # Dependency-free (stdlib + ctypes on Windows).
    # ------------------------------------------------------------------
    def _safe_delete(self, path):
        """Move `path` to the OS trash; on failure, to ~/TextVault_Data/_RecycleBin.
        Returns True if the file is no longer at its original location."""
        if not path or not os.path.exists(path):
            return True
        if self._send_to_trash(path):
            return True
        # Safety net: relocate into a local recycle bin rather than deleting.
        try:
            backup_dir = os.path.join(self.vault_dir, "_RecycleBin")
            os.makedirs(backup_dir, exist_ok=True)
            dest = os.path.join(backup_dir, f"{int(time.time())}_{os.path.basename(path)}")
            counter = 1
            while os.path.exists(dest):
                dest = os.path.join(backup_dir, f"{int(time.time())}_{counter}_{os.path.basename(path)}")
                counter += 1
            shutil.move(path, dest)
            return True
        except Exception as e:
            print(f"Safe-delete fallback failed for {path}: {e}")
            return False

    def _send_to_trash(self, path):
        """Best-effort move to the OS Recycle Bin / Trash. Returns True on success."""
        path = os.path.abspath(path)
        try:
            if IS_WINDOWS:
                import ctypes
                from ctypes import wintypes

                FO_DELETE = 3
                FOF_ALLOWUNDO = 0x0040       # send to Recycle Bin instead of deleting
                FOF_NOCONFIRMATION = 0x0010  # we already confirmed in the UI
                FOF_SILENT = 0x0004
                FOF_NOERRORUI = 0x0400

                class SHFILEOPSTRUCTW(ctypes.Structure):
                    _fields_ = [
                        ("hwnd", wintypes.HWND),
                        ("wFunc", wintypes.UINT),
                        ("pFrom", wintypes.LPCWSTR),
                        ("pTo", wintypes.LPCWSTR),
                        ("fFlags", ctypes.c_uint16),
                        ("fAnyOperationsAborted", wintypes.BOOL),
                        ("hNameMappings", ctypes.c_void_p),
                        ("lpszProgressTitle", wintypes.LPCWSTR),
                    ]

                op = SHFILEOPSTRUCTW()
                op.wFunc = FO_DELETE
                op.pFrom = path + "\0"  # path list must be double-null terminated
                op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
                res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
                return res == 0 and not op.fAnyOperationsAborted

            elif sys.platform == "darwin":
                # Hand off to Finder so it lands in the user's Trash properly.
                script = f'tell application "Finder" to delete POSIX file "{path}"'
                return subprocess.call(
                    ["osascript", "-e", script],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                ) == 0

            else:
                # Freedesktop trash spec (~/.local/share/Trash)
                trash = os.path.join(os.path.expanduser("~"), ".local", "share", "Trash")
                files_dir = os.path.join(trash, "files")
                info_dir = os.path.join(trash, "info")
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
            print(f"Trash failed for {path}: {e}")
            return False

    def deep_reset_reports(self):
        if not messagebox.askyesno("Reset Reports", "Are you sure you want to permanently delete all generated Vault Reports?"):
            return

        deleted_count = 0
        for filename in os.listdir(self.vault_dir):
            if filename.startswith("Vault_Report_") and filename.endswith(".txt"):
                filepath = os.path.join(self.vault_dir, filename)
                try:
                    self._safe_delete(filepath)
                    deleted_count += 1
                    if filename in self.metadata["files"]:
                        del self.metadata["files"][filename]
                    if filepath in self.file_cache:
                        del self.file_cache[filepath]
                except Exception as e:
                    print(f"Error removing {filename}: {e}")

        self.save_metadata()
        self.refresh_index()
        messagebox.showinfo("Reset Complete", f"Successfully deleted {deleted_count} report(s).")

    def extract_dominant_keyword(self, content):
        words = re.findall(r'\b[a-zA-Z]{4,}\b', content[:10000].lower())
        filtered_words = [w for w in words if w not in STOP_WORDS]
        if not filtered_words:
            return "Uncategorized"
        counter = Counter(filtered_words)
        return counter.most_common(1)[0][0].capitalize()

    # ------------------------------------------------------------------
    # Indexing (background-threaded so the UI never blocks on os.walk)
    # ------------------------------------------------------------------
    def auto_recall_indexed(self):
        for folder in list(self.indexed_folders):
            if os.path.exists(folder):
                self._start_index_worker(folder, silent=True)

    def index_folder(self, target_dir=None, silent=False):
        if not target_dir:
            target_dir = filedialog.askdirectory(title="Select Drive or Folder to Index")
        if not target_dir:
            return

        if target_dir not in self.indexed_folders:
            self.indexed_folders.append(target_dir)
            self.save_metadata()

        self._start_index_worker(target_dir, silent=silent)

    def _start_index_worker(self, target_dir, silent):
        self.status_var.set(f"Indexing: {target_dir} ...")
        threading.Thread(target=self._index_worker, args=(target_dir, silent), daemon=True).start()

    def _index_worker(self, target_dir, silent):
        """Runs off the Tk thread. Does filesystem work only, then marshals
        all UI updates back to the main thread via root.after()."""
        try:
            found_count, sec_flags = self._scan_and_copy(target_dir)
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"Index error: {e}"))
            return
        self.root.after(0, lambda: self._index_done(found_count, sec_flags, silent))

    def _scan_and_copy(self, target_dir):
        """Filesystem-only scan + copy. Returns (found_count, [(flag_type, path), ...])."""
        found_count = 0
        sec_flags = []

        with self._meta_lock:
            for root_dir, dirs, files in os.walk(target_dir):
                # Skip the vault itself to avoid re-indexing our own copies
                if os.path.normpath(root_dir).startswith(os.path.normpath(self.vault_dir)):
                    continue

                for file in files:
                    source_path = os.path.join(root_dir, file)
                    ext = os.path.splitext(file)[1].lower()

                    if ext in DANGER_EXTS:
                        sec_flags.append(("Executable/Script", source_path))
                        continue

                    if ext not in TRACKED_EXTS:
                        continue

                    try:
                        mtime = os.stat(source_path).st_mtime

                        existing_vault_name = None
                        for v_name, data in self.metadata["files"].items():
                            if data.get("original_path") == source_path:
                                existing_vault_name = v_name
                                break

                        if existing_vault_name:
                            dest_path = os.path.join(self.vault_dir, existing_vault_name)
                            if os.path.exists(dest_path) and \
                               self.metadata["files"][existing_vault_name].get("mtime") == mtime:
                                continue  # unchanged, skip the copy
                        else:
                            dest_path = os.path.join(self.vault_dir, file)
                            counter = 1
                            while os.path.exists(dest_path):
                                name, e = os.path.splitext(file)
                                dest_path = os.path.join(self.vault_dir, f"{name}_{counter}{e}")
                                counter += 1

                        shutil.copy2(source_path, dest_path)
                        filename = os.path.basename(dest_path)
                        self.metadata["files"][filename] = {
                            "original_path": source_path,
                            "mtime": mtime,
                        }
                        found_count += 1
                    except Exception:
                        continue

            self.save_metadata()

        return found_count, sec_flags

    def _index_done(self, found_count, sec_flags, silent):
        for flag_type, path in sec_flags:
            self.sec_tree.insert("", "end", values=(flag_type, path))
        self.refresh_index(self.search_var.get())
        self.refresh_purge_list()
        self.status_var.set("Ready")
        if not silent:
            msg = f"Synced {found_count} new or modified file(s) (.txt + Office)."
            if sec_flags:
                msg += f"\nFlagged {len(sec_flags)} potentially dangerous/irrelevant files. Check Security tab."
            messagebox.showinfo("Indexing Complete", msg)

    # ------------------------------------------------------------------
    def refresh_index(self, search_query=""):
        self.tree.delete(*self.tree.get_children())
        self.file_index.clear()

        groups = {}
        search_query = search_query.lower()

        for filename in os.listdir(self.vault_dir):
            ext = os.path.splitext(filename)[1].lower()
            if ext not in TRACKED_EXTS:
                continue

            filepath = os.path.join(self.vault_dir, filename)
            content = ""  # only loaded for editable .txt

            try:
                mtime = os.path.getmtime(filepath)
                if ext in MS_EXTENSIONS:
                    # Binary Office file: categorize by type, never read content
                    category = MS_EXTENSIONS[ext]
                    self.file_cache[filepath] = {'mtime': mtime, 'category': category}
                else:
                    # .txt note: cache content + keyword categorization
                    cached = self.file_cache.get(filepath)
                    if cached and cached.get('mtime') == mtime and 'content' in cached:
                        content = cached['content']
                        category = cached['category']
                    else:
                        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                            content = f.read()
                        category = self.extract_dominant_keyword(content)
                        self.file_cache[filepath] = {'content': content, 'mtime': mtime, 'category': category}
            except Exception:
                continue

            if search_query:
                # MS files match on filename only; .txt also matches on content
                if search_query not in filename.lower() and search_query not in content.lower():
                    continue

            original_path = self.metadata["files"].get(filename, {}).get("original_path", "Vault Local")
            source_dir = os.path.dirname(original_path) if original_path != "Vault Local" else "Vault Local"

            groups.setdefault(source_dir, {}).setdefault(category, []).append((filename, filepath))

        for source_dir, categories in sorted(groups.items()):
            source_id = self.tree.insert("", "end", text=f"Source: {source_dir}", open=False)
            for category, files in sorted(categories.items()):
                cat_id = self.tree.insert(source_id, "end", text=f"{category} ({len(files)})", open=False)
                for filename, filepath in sorted(files, key=lambda x: x[0]):
                    item_id = self.tree.insert(cat_id, "end", text=f"{filename}")
                    self.file_index[item_id] = filepath

    def on_search(self, *args):
        self.refresh_index(self.search_var.get())

    def on_file_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        item_id = selected[0]
        if item_id in self.file_index:
            filepath = self.file_index[item_id]
            self.current_file = filepath
            self.load_file_to_editor(filepath)

    def load_file_to_editor(self, filepath):
        self.text_editor.config(state=tk.NORMAL)
        self.text_editor.delete(1.0, tk.END)

        ext = os.path.splitext(filepath)[1].lower()
        if ext in MS_EXTENSIONS:
            self.text_editor.insert(
                tk.END,
                "--- Binary Microsoft Office File ---\n\n"
                f"File: {os.path.basename(filepath)}\n"
                f"Type: {MS_EXTENSIONS[ext]}\n\n"
                "Direct text editing is disabled for this file type to prevent corruption.\n"
                "Please use 'Open Original Location' to edit in Microsoft Office."
            )
            self.text_editor.config(state=tk.DISABLED)  # Lock editor
        else:
            try:
                cached = self.file_cache.get(filepath)
                if cached and 'content' in cached:
                    self.text_editor.insert(tk.END, cached['content'])
                else:
                    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                        self.text_editor.insert(tk.END, f.read())
            except Exception as e:
                messagebox.showerror("Error", f"Could not read file:\n{e}")

    def new_note(self):
        filename = f"Draft_{int(time.time())}.txt"
        filepath = os.path.join(self.vault_dir, filename)

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write("New note...")

        self.current_file = filepath
        self.refresh_index()
        self.load_file_to_editor(filepath)

    def save_current(self):
        if not self.current_file:
            return
        ext = os.path.splitext(self.current_file)[1].lower()
        if ext in MS_EXTENSIONS:
            messagebox.showwarning("Read Only", "Cannot save edits to Microsoft Office binary formats directly from the Vault.")
            return

        content = self.text_editor.get(1.0, tk.END)
        if content.endswith('\n'):
            content = content[:-1]

        with open(self.current_file, 'w', encoding='utf-8') as f:
            f.write(content)

        category = self.extract_dominant_keyword(content)
        self.file_cache[self.current_file] = {
            'content': content,
            'mtime': os.path.getmtime(self.current_file),
            'category': category,
        }
        self.refresh_index(self.search_var.get())

    def auto_save_draft(self, event):
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Prior', 'Next']:
            return
        if self._autosave_timer:
            self.root.after_cancel(self._autosave_timer)
        self._autosave_timer = self.root.after(500, self.perform_auto_save)

    def perform_auto_save(self):
        if not self.current_file:
            return
        ext = os.path.splitext(self.current_file)[1].lower()
        if ext in MS_EXTENSIONS:
            return  # never auto-save binary files

        content = self.text_editor.get(1.0, tk.END)
        if content.endswith('\n'):
            content = content[:-1]
        try:
            with open(self.current_file, 'w', encoding='utf-8') as f:
                f.write(content)
            cat = self.file_cache.get(self.current_file, {}).get('category', 'Uncategorized')
            self.file_cache[self.current_file] = {
                'content': content,
                'mtime': os.path.getmtime(self.current_file),
                'category': cat,
            }
        except Exception:
            pass

    def find_duplicates(self):
        hashes = {}
        duplicates = []

        for item_id, filepath in self.file_index.items():
            if not os.path.exists(filepath):
                continue
            try:
                file_hash = self._hash_file(filepath)
            except Exception:
                continue
            if file_hash in hashes:
                duplicates.append((filepath, hashes[file_hash]))
            else:
                hashes[file_hash] = filepath

        if not duplicates:
            messagebox.showinfo("Duplicates", "No exact duplicate files found in the vault.")
            return

        msg = f"Found {len(duplicates)} duplicate(s).\nDo you want to automatically remove the redundant copies from the vault?"
        if messagebox.askyesno("Remove Duplicates", msg):
            removed = 0
            for dup_path, orig_path in duplicates:
                try:
                    self._safe_delete(dup_path)
                    filename = os.path.basename(dup_path)
                    if filename in self.metadata["files"]:
                        del self.metadata["files"][filename]
                    if dup_path in self.file_cache:
                        del self.file_cache[dup_path]
                    removed += 1
                except Exception:
                    pass
            self.save_metadata()
            self.refresh_index()
            self.refresh_purge_list()
            messagebox.showinfo("Success", f"Removed {removed} duplicate files.")

    @staticmethod
    def _hash_file(filepath):
        """Chunked SHA-256 so large Office files don't blow up memory."""
        h = hashlib.sha256()
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(65536), b''):
                h.update(chunk)
        return h.hexdigest()

    def generate_japanese_report(self):
        report_path = os.path.join(self.vault_dir, f"Vault_Report_{int(time.time())}.txt")
        file_count = len(self.file_index)
        folder_count = len(self.indexed_folders)

        report_content = "保管庫 (Vault) レポート\n"
        report_content += "============================\n"
        report_content += "作成者: KIMANI S.M.\n"
        report_content += f"インデックス済みファイル数: {file_count}\n"
        report_content += f"インデックス済みフォルダ数: {folder_count}\n\n"
        report_content += "対象フォルダ:\n"
        for folder in self.indexed_folders:
            report_content += f"- {folder}\n"

        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report_content)
        messagebox.showinfo("レポート作成", f"レポートが作成されました:\n{report_path}")
        self.refresh_index()

    def clear_security_list(self):
        self.sec_tree.delete(*self.sec_tree.get_children())

    def delete_security_file(self):
        selections = self.sec_tree.selection()
        if not selections:
            return

        # Default to "No" so a stray Enter/click cancels. Deletes are recoverable.
        if not messagebox.askyesno(
            "Confirm Delete",
            f"Send {len(selections)} flagged file(s) to the Recycle Bin / Trash?\n\n"
            "These are files on your system, outside the vault. They will remain "
            "recoverable from the Recycle Bin.",
            default="no", icon="warning",
        ):
            return

        removed = 0
        for item_id in selections:
            item = self.sec_tree.item(item_id)
            path = item['values'][1]
            if self._safe_delete(path):
                self.sec_tree.delete(item_id)
                removed += 1
            else:
                print(f"Could not delete {path}")
        messagebox.showinfo("Done", f"Sent {removed} file(s) to the Recycle Bin / Trash.")

    def run_defender_scan(self):
        if not IS_WINDOWS:
            messagebox.showinfo("Unavailable", "Windows Defender scanning is only available on Windows.")
            return
        selections = self.sec_tree.selection()
        if not selections:
            return
        path = self.sec_tree.item(selections[0])['values'][1]
        self.scan_progress.start(10)
        threading.Thread(target=self._execute_defender_thread, args=(path,), daemon=True).start()

    def _execute_defender_thread(self, path):
        try:
            CREATE_NO_WINDOW = 0x08000000  # Suppresses terminal flashing
            cmd = ["powershell", "-Command", f"Start-MpScan -ScanType CustomScan -ScanPath '{path}'"]
            process = subprocess.Popen(cmd, creationflags=CREATE_NO_WINDOW, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            process.communicate()
            self.root.after(0, lambda: messagebox.showinfo("Scan Complete", f"Defender scan finished for:\n{path}"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Could not execute Defender:\n{e}"))
        finally:
            self.root.after(0, self.scan_progress.stop)

    def refresh_purge_list(self):
        self.purge_tree.delete(*self.purge_tree.get_children())
        for filename in os.listdir(self.vault_dir):
            ext = os.path.splitext(filename)[1].lower()
            if ext in TRACKED_EXTS:
                vault_path = os.path.join(self.vault_dir, filename)
                original_path = self.metadata["files"].get(filename, {}).get("original_path", "Vault Local")
                self.purge_tree.insert("", "end", values=(filename, vault_path, original_path))

    def execute_deep_purge(self):
        selections = self.purge_tree.selection()
        if not selections:
            return

        # Type-to-confirm: Deep Purge deletes the ORIGINAL file on disk and edits
        # the registry, so a reflexive "Yes" click must never be enough.
        prompt = (
            f"You are about to DEEP PURGE {len(selections)} file(s).\n\n"
            "This removes each file from the Vault AND its original location on "
            "disk, and scrubs Windows Recent-Docs registry traces.\n\n"
            "Deleted files are sent to the Recycle Bin / Trash where possible, so "
            "they stay recoverable.\n\n"
            "Type  PURGE  (in capitals) to confirm:"
        )
        answer = simpledialog.askstring("Confirm Deep Purge", prompt, parent=self.root)
        if answer != "PURGE":
            messagebox.showinfo("Cancelled", "Deep Purge cancelled — no files were deleted.")
            return

        purged = 0
        for item_id in selections:
            item = self.purge_tree.item(item_id)
            filename, vault_path, original_path = item['values']

            if os.path.exists(vault_path):
                self._safe_delete(vault_path)

            if original_path != "Vault Local" and os.path.exists(original_path):
                self._safe_delete(original_path)

            if filename in self.metadata["files"]:
                del self.metadata["files"][filename]
            if vault_path in self.file_cache:
                del self.file_cache[vault_path]

            self._purge_registry_mru(filename)
            purged += 1

        self.save_metadata()
        self.refresh_index()
        self.refresh_purge_list()
        messagebox.showinfo("Deep Purge Complete", f"Deep purge executed on {purged} file(s).")

    def _purge_registry_mru(self, target_filename):
        # Cleans Explorer RecentDocs where file traces are heavily kept (Windows only)
        if not IS_WINDOWS:
            return
        try:
            mru_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs"
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, mru_path, 0, winreg.KEY_READ | winreg.KEY_WRITE)

            count = winreg.QueryInfoKey(key)[1]
            values_to_delete = []

            for i in range(count):
                try:
                    val_name, val_data, val_type = winreg.EnumValue(key, i)
                    if isinstance(val_data, bytes):
                        decoded_string = val_data.decode('utf-16le', errors='ignore')
                        if target_filename.lower() in decoded_string.lower():
                            values_to_delete.append(val_name)
                except Exception:
                    continue

            for v in values_to_delete:
                winreg.DeleteValue(key, v)

            winreg.CloseKey(key)
        except Exception as e:
            print(f"Registry purge skipped/failed for MRU: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = VaultToolkitApp(root)
    root.mainloop()
