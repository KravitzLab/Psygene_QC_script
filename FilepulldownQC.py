# -*- coding: utf-8 -*-
"""
Startup mode picker (FED vs BEAM)

FED mode:
- Folder picker
- Scans FED###_*.csv files
- Verifies 4 files per FED and presence of Bandit100, Bandit80, FR1, PR1
- Flags header-only (empty) CSVs (~1.5 KB or 0 data rows)
- Reports unique device counts (overall and per task)
- Moves files into task folders: Bandit_100, Bandit_80, FR1, PR1

BEAM mode (robust + date range filter):
- Single ZIP picker
- Asks for a START and END date to KEEP (YYYYMMDD, inclusive)
- Extracts to "_EXTRACT_<zipname>" (kept)
- Recursively finds ALL .csv anywhere inside (no folder-name assumptions)
- If none, auto-extracts nested ZIPs into sibling "_NESTED_<name>" folders and searches again
- Moves ONLY CSVs whose filenames match: BEAMXXX_YYYYMMDD??.csv and start<=YYYYMMDD<=end
- Places matched CSVs into sibling folder "BEAM" (same layer as the original zip)
- Deletes ONLY folders that contained moved CSVs and are empty afterward (keeps extraction root and anything non-empty)
- Prints a preview of zip contents and a shallow tree of extracted contents for debugging
"""

import os
import re
import sys
import shutil
import zipfile
import pandas as pd

# =========================
# Mode selection (GUI/console)
# =========================
def pick_mode_gui():
    try:
        import tkinter as tk
    except Exception:
        return None

    mode = {"value": None}

    def choose_fed():
        mode["value"] = "FED"
        root.destroy()

    def choose_beam():
        mode["value"] = "BEAM"
        root.destroy()

    def on_close():
        mode["value"] = None
        root.destroy()

    root = tk.Tk()
    root.title("Select Data Type")
    root.geometry("320x150")
    root.resizable(False, False)

    lbl = tk.Label(root, text="Are you working with FED or BEAM data?", font=("Segoe UI", 11))
    lbl.pack(pady=15)

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="FED", width=10, command=choose_fed).grid(row=0, column=0, padx=10)
    tk.Button(btn_frame, text="BEAM", width=10, command=choose_beam).grid(row=0, column=1, padx=10)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.attributes("-topmost", True)
    root.mainloop()
    return mode["value"]

def pick_mode_console():
    try:
        ans = input("Select mode: [F]ED or [B]EAM? ").strip().lower()
        if ans.startswith("f"): return "FED"
        if ans.startswith("b"): return "BEAM"
    except Exception:
        pass
    return None

# =========================
# Shared helpers
# =========================
def pick_folder():
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        print("tkinter not available; using current directory.")
        return os.getcwd()
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="Select folder")
    root.destroy()
    if not folder:
        print("No folder selected. Exiting."); sys.exit(0)
    return folder

def pick_zip_file():
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        path = input("Enter path to a .zip file: ").strip()
        if not (path and os.path.isfile(path) and path.lower().endswith(".zip")):
            print("Invalid .zip path. Exiting."); sys.exit(1)
        return path
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    zpath = filedialog.askopenfilename(
        title="Select BEAM .zip",
        filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
    )
    root.destroy()
    if not zpath:
        print("No .zip selected. Exiting."); sys.exit(0)
    if not zpath.lower().endswith(".zip"):
        print("Selected file is not a .zip. Exiting."); sys.exit(1)
    return zpath

def ask_date_range_to_keep():
    """
    Prompt for a START and END date (YYYYMMDD). Inclusive.
    GUI first; console fallback. Validates both are 8 digits and start<=end.
    """
    pattern = re.compile(r"^\d{8}$")

    # GUI path
    try:
        import tkinter as tk
        from tkinter import simpledialog, messagebox
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        while True:
            s = simpledialog.askstring("Filter by date range", "Enter START date (YYYYMMDD):")
            if s is None:
                messagebox.showinfo("Cancelled", "No date entered. Exiting.")
                sys.exit(0)
            e = simpledialog.askstring("Filter by date range", "Enter END date (YYYYMMDD):")
            if e is None:
                messagebox.showinfo("Cancelled", "No date entered. Exiting.")
                sys.exit(0)
            s = s.strip(); e = e.strip()
            if not (pattern.fullmatch(s) and pattern.fullmatch(e)):
                messagebox.showerror("Invalid date", "Dates must be exactly 8 digits (YYYYMMDD).")
                continue
            if s > e:
                messagebox.showerror("Invalid range", "START date must be <= END date.")
                continue
            root.destroy()
            return s, e
    except Exception:
        # Console fallback
        while True:
            s = input("Enter START date (YYYYMMDD): ").strip()
            e = input("Enter END date (YYYYMMDD): ").strip()
            if not (pattern.fullmatch(s) and pattern.fullmatch(e)):
                print("Dates must be exactly 8 digits (YYYYMMDD). Try again.")
                continue
            if s > e:
                print("START must be <= END. Try again.")
                continue
            return s, e

def ensure_dir(path):
    os.makedirs(path, exist_ok=True); return path

def unique_dest_path(dest_dir, filename):
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(dest_dir, filename); i = 1
    while os.path.exists(candidate):
        candidate = os.path.join(dest_dir, f"{base} ({i}){ext}"); i += 1
    return candidate

def is_numeric_dirname(name):
    return bool(re.fullmatch(r"\d+", name))

# =========================
# FED implementation
# =========================
REQUIRED_TYPES = {"Bandit100", "FR1", "Bandit80", "PR1"}
TASK_DIRNAME = {"Bandit100": "Bandit_100", "Bandit80": "Bandit_80", "FR1": "FR1", "PR1": "PR1"}
FED_PATTERN = re.compile(r"(?i)\bFED(\d{3})_", re.ASCII)
EMPTY_SIZE_BYTES = 1500  # ~1KB header-only files

def extract_fed_digits(filename):
    m = FED_PATTERN.search(filename); return m.group(1) if m else None

def normalize_session_type(value):
    if value is None: return None
    v = str(value).strip()
    for canonical in REQUIRED_TYPES:
        if v.lower() == canonical.lower(): return canonical
    return None

def is_likely_empty(csv_path):
    try: size_ok = os.path.getsize(csv_path) <= EMPTY_SIZE_BYTES
    except OSError: size_ok = False
    empty_by_rows = False
    try:
        head = pd.read_csv(csv_path, dtype=str, nrows=5, encoding_errors="ignore")
        if head.shape[0] == 0: empty_by_rows = True
        else:
            nonempty = head.apply(lambda s: s.astype(str).str.strip()).replace({"": pd.NA}).dropna(how="all")
            empty_by_rows = nonempty.empty
    except Exception: empty_by_rows = False
    return size_ok, empty_by_rows

def read_session_type(csv_path):
    try:
        head = pd.read_csv(csv_path, dtype=str, nrows=1000, encoding_errors="ignore")
        cols_map = {c.lower(): c for c in head.columns}
        if "session_type" not in cols_map:
            df = pd.read_csv(csv_path, dtype=str, encoding_errors="ignore")
            cols_map = {c.lower(): c for c in df.columns}
            if "session_type" not in cols_map: return None
            series = df[cols_map["session_type"]]
        else:
            series = head[cols_map["session_type"]]
        series = series.dropna().astype(str).str.strip()
        norm = series.map(normalize_session_type).dropna()
        if norm.empty: return None
        return norm.mode().iloc[0]
    except Exception:
        return None

def move_file_to_task_folder(folder, filename, session_type):
    canonical = normalize_session_type(session_type)
    if canonical not in TASK_DIRNAME: return False
    dest_dir = ensure_dir(os.path.join(folder, TASK_DIRNAME[canonical]))
    src_path = os.path.join(folder, filename)
    dest_path = unique_dest_path(dest_dir, filename)
    try:
        shutil.move(src_path, dest_path); return True
    except Exception as e:
        print(f"  ! Could not move {filename} → {dest_dir}: {e}"); return False

def run_fed_pipeline():
    folder = pick_folder()

    all_files = [f for f in os.listdir(folder) if f.lower().endswith(".csv")]
    csv_files = [f for f in all_files if FED_PATTERN.search(f)]
    if not csv_files:
        print("No CSV files with pattern 'FED###_' found in the selected folder."); return

    per_fed, all_types_global, file_to_type, file_is_empty = {}, {}, {}, set()

    for fname in csv_files:
        fed = extract_fed_digits(fname)
        if not fed: continue
        fullpath = os.path.join(folder, fname)
        empty_by_size, empty_by_rows = is_likely_empty(fullpath)
        stype = None if (empty_by_size or empty_by_rows) else read_session_type(fullpath)

        fed_info = per_fed.setdefault(fed, {"files": [], "types": set(), "by_type": {}, "empties": []})
        fed_info["files"].append(fname)

        if empty_by_size or empty_by_rows:
            fed_info["empties"].append(fname); file_is_empty.add(fname)

        if stype:
            canonical = normalize_session_type(stype)
            if canonical:
                fed_info["types"].add(canonical)
                fed_info["by_type"].setdefault(canonical, []).append(fname)
                all_types_global.setdefault(canonical, set()).add(fed)

        file_to_type[fname] = normalize_session_type(stype) if stype else None

    rows, any_issue = [], False
    for fed in sorted(per_fed.keys()):
        info = per_fed[fed]
        file_count = len(info["files"])
        present_types = sorted(info["types"])
        missing_types = sorted(REQUIRED_TYPES - info["types"])
        duplicates = {t: files for t, files in info["by_type"].items() if len(files) > 1}
        count_ok = (file_count == 4); types_ok = (len(missing_types) == 0)
        has_empties = len(info["empties"]) > 0
        overall_ok = count_ok and types_ok and not has_empties
        if not overall_ok or duplicates: any_issue = True
        rows.append({
            "FED": f"FED{fed}", "File_Count": file_count, "Count_OK(=4)": count_ok,
            "Present_Types": ";".join(present_types), "Missing_Types": ";".join(missing_types),
            "Duplicate_Types": "; ".join(f"{t} x{len(files)}" for t, files in info["by_type"].items()) if duplicates else "",
            "Empty_File_Count": len(info["empties"]), "Empty_File_Names": "; ".join(sorted(info["empties"])) if info["empties"] else "",
            "All_Types_OK": types_ok, "Overall_OK": overall_ok
        })

    report = pd.DataFrame(rows)
    unique_devices_total = len(per_fed)
    per_task_counts = {task: len(feds) for task, feds in sorted(all_types_global.items())}

    summary_rows = [{"FED": "SUMMARY", "File_Count": unique_devices_total, "Count_OK(=4)": "", "Present_Types": "",
                     "Missing_Types": "", "Duplicate_Types": "", "Empty_File_Count": "", "Empty_File_Names": "",
                     "All_Types_OK": "", "Overall_OK": ""}]
    for task, count in per_task_counts.items():
        summary_rows.append({"FED": f"→ {task}", "File_Count": count, "Count_OK(=4)": "", "Present_Types": "",
                             "Missing_Types": "", "Duplicate_Types": "", "Empty_File_Count": "", "Empty_File_Names": "",
                             "All_Types_OK": "", "Overall_OK": ""})
    report = pd.concat([report, pd.DataFrame(summary_rows)], ignore_index=True)

    out_csv = os.path.join(folder, "fed_session_check_report.csv")
    report.to_csv(out_csv, index=False)

    moved_counts = {d: 0 for d in TASK_DIRNAME.values()}
    skipped_unclassified, skipped_empty = [], []
    for dirname in TASK_DIRNAME.values(): ensure_dir(os.path.join(folder, dirname))

    for fname in csv_files:
        if fname in file_is_empty:
            skipped_empty.append(fname); continue
        stype = file_to_type.get(fname)
        if stype in TASK_DIRNAME:
            if move_file_to_task_folder(folder, fname, stype):
                moved_counts[TASK_DIRNAME[stype]] += 1
        else:
            skipped_unclassified.append(fname)

    print("\n=== FED Session Completeness Report ===")
    print(f"Folder: {folder}")
    print(f"Required session types: {sorted(REQUIRED_TYPES)}")
    print(f"Report saved to: {out_csv}\n")
    print(f"Total unique devices: {unique_devices_total}")
    for task, count in per_task_counts.items(): print(f"  {task}: {count} devices")
    print("\nFile sorting:")
    for dirname, n in moved_counts.items(): print(f"  Moved into {dirname}: {n} files")
    if skipped_empty: print(f"  Skipped empty files (left in place): {len(skipped_empty)}")
    if skipped_unclassified: print(f"  Skipped unclassified files (no recognized Session_type): {len(skipped_unclassified)}")

    if any_issue:
        print("\nIssues detected:")
        for _, r in report.iterrows():
            if str(r["FED"]).startswith("FED") and (not r["Overall_OK"]):
                print(f"- {r['FED']}:")
                if not r["Count_OK(=4)"]: print(f"  * Has {r['File_Count']} files (expected 4).")
                if r["Missing_Types"]: print(f"  * Missing types: {r['Missing_Types']}")
                if r["Duplicate_Types"]: print(f"  * Duplicates: {r['Duplicate_Types']}")
                if r["Empty_File_Count"] > 0: print(f"  * Likely empty files ({r['Empty_File_Count']}): {r['Empty_File_Names']}")
        print("\nOpen the CSV report for the full table.")
    else:
        print("All FEDs complete, non-empty, and contain all required session types. ✅")

# =========================
# BEAM implementation (robust + date range)
# =========================
# Matches: BEAM + 3 digits + '_' + 8-digit date + 2 trailing digits (device minute or similar)
BEAM_NAME_RE = re.compile(r"(?i)\bBEAM(\d{3})_(\d{8})(\d{2})\b")

def _print_tree(root, max_entries=60):
    """Print a shallow tree of extracted contents for quick debugging."""
    print(f"\n[Preview of extracted contents under: {root}]")
    shown = 0
    for dirpath, dirnames, filenames in os.walk(root):
        rel = os.path.relpath(dirpath, root)
        indent = "  " * (0 if rel == "." else rel.count(os.sep))
        node_name = os.path.basename(dirpath) if rel != "." else os.path.basename(root)
        print(f"{indent}{node_name}")
        for d in sorted(dirnames):
            print(f"{indent}  /{d}")
            shown += 1
            if shown >= max_entries:
                print("  ... (truncated)"); return
        for f in sorted(filenames):
            print(f"{indent}   - {f}")
            shown += 1
            if shown >= max_entries:
                print("  ... (truncated)"); return

def run_beam_pipeline():
    """
    BEAM mode (robust + date range filter):
    - Prompt for a single .zip and a date range to KEEP (YYYYMMDD..YYYYMMDD, inclusive)
    - Extract to _EXTRACT_<zipname> (kept)
    - Recursively find ALL .csv (case-insensitive) anywhere inside
      * If none found, auto-extract nested .zip files into sibling _NESTED_<name> folders and search again
    - Keep ONLY CSVs whose filenames match BEAMXXX_YYYYMMDD??.csv and start<=YYYYMMDD<=end
    - Move matched CSVs to sibling 'BEAM' folder next to the original zip
    - Delete ONLY source folders that are empty after moving
    """
    zip_path = pick_zip_file()
    start_date, end_date = ask_date_range_to_keep()  # YYYYMMDD each
    parent_folder = os.path.dirname(zip_path)
    zip_stem = os.path.splitext(os.path.basename(zip_path))[0]
    extract_root = os.path.join(parent_folder, f"_EXTRACT_{zip_stem}")
    ensure_dir(extract_root)

    print(f"\nSelected ZIP: {zip_path}")
    print(f"Date range to KEEP: {start_date} → {end_date} (inclusive)")

    # Show zip contents and extract
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            names = zf.namelist()
            print("\nZIP contains (first 30 entries):")
            for n in names[:30]:
                print(f"  - {n}")
            if len(names) > 30:
                print("  ... (truncated)")
            zf.extractall(extract_root)
        print(f"\nExtracted to: {extract_root}")
    except Exception as e:
        print(f"Failed to extract zip: {e}")
        sys.exit(1)

    _print_tree(extract_root, max_entries=60)

    def find_csvs(root):
        csv_paths = []
        for r, _, files in os.walk(root):
            for fn in files:
                if fn.lower().endswith(".csv"):
                    csv_paths.append(os.path.join(r, fn))
        return csv_paths

    # Pass 1: find CSVs
    csv_paths = find_csvs(extract_root)

    # If none, extract nested zips and search again
    if not csv_paths:
        nested_zips = []
        for r, _, files in os.walk(extract_root):
            for fn in files:
                if fn.lower().endswith(".zip"):
                    nested_zips.append(os.path.join(r, fn))
        if nested_zips:
            print(f"\nNo CSVs found initially. Found {len(nested_zips)} nested ZIP(s); extracting those...")
            for nz in nested_zips:
                nz_stem = os.path.splitext(os.path.basename(nz))[0]
                nested_out = os.path.join(parent_folder, f"_NESTED_{nz_stem}")
                ensure_dir(nested_out)
                try:
                    with zipfile.ZipFile(nz, 'r') as zf:
                        zf.extractall(nested_out)
                    print(f"  Extracted nested zip: {nz} → {nested_out}")
                except Exception as e:
                    print(f"  ! Failed to extract nested zip {nz}: {e}")

            # Search again in both extract_root and all _NESTED_* folders
            csv_paths = find_csvs(extract_root)
            for d in os.listdir(parent_folder):
                if d.startswith("_NESTED_"):
                    csv_paths.extend(find_csvs(os.path.join(parent_folder, d)))

    # Filter CSVs by BEAM naming + date range
    def in_range(path):
        name_no_ext = os.path.splitext(os.path.basename(path))[0]
        m = BEAM_NAME_RE.search(name_no_ext)
        if not m:
            return False
        date_str = m.group(2)  # YYYYMMDD
        return (start_date <= date_str <= end_date)

    matching_csvs = [p for p in csv_paths if in_range(p)]
    nonmatching = len(csv_paths) - len(matching_csvs)

    if not matching_csvs:
        print("\nNo CSV files matched BEAMXXX_YYYYMMDD?? within your date range.")
        print(f"Left extraction folder intact: {extract_root}")
        return

    # Destination: BEAM folder next to the original zip
    dest_beam = ensure_dir(os.path.join(parent_folder, "BEAM"))

    # Move only matching CSVs and track their source folders (for cleanup)
    moved = 0
    skipped = 0
    source_dirs_for_deletion = set()

    for src in matching_csvs:
        try:
            dest = unique_dest_path(dest_beam, os.path.basename(src))
            shutil.move(src, dest)
            moved += 1
            source_dirs_for_deletion.add(os.path.dirname(src))
        except Exception as e:
            print(f"  ! Could not move {src} → {dest_beam}: {e}")
            skipped += 1

    # Delete only the source folders that are now empty (but never the extract_root)
    deleted_dirs = []
    for sdir in sorted(source_dirs_for_deletion, key=lambda p: len(p.split(os.sep)), reverse=True):
        if os.path.abspath(sdir) == os.path.abspath(extract_root):
            continue
        try:
            if not any(os.scandir(sdir)):
                shutil.rmtree(sdir)
                deleted_dirs.append(sdir)
        except Exception as e:
            print(f"  ! Failed to delete folder {sdir}: {e}")

    print("\n=== BEAM Flattening + Date-Range Filter Summary ===")
    print(f"Date range: {start_date} → {end_date}")
    print(f"All CSVs discovered: {len(csv_paths)}")
    print(f"Matched CSVs (kept/moved): {moved}")
    if skipped:
        print(f"  Skipped (errors during move): {skipped}")
    print(f"Non-matching CSVs (left in place): {nonmatching}")
    if deleted_dirs:
        print(f"Deleted emptied source folders: {len(deleted_dirs)}")
        for d in deleted_dirs:
            print(f"  - {d}")
    else:
        print("No emptied source folders were deleted (they may still contain non-CSV files or non-matching files).")
    print(f"\nMatched CSVs are in: {dest_beam}")
    print(f"Kept extraction folder intact: {extract_root}")
    print("Done.")

# =========================
# Main
# =========================
def main():
    mode = pick_mode_gui()
    if mode is None:
        mode = pick_mode_console()

    if mode == "FED":
        run_fed_pipeline()
    elif mode == "BEAM":
        run_beam_pipeline()
    else:
        print("No mode selected. Exiting.")
        sys.exit(0)

if __name__ == "__main__":
    main()
