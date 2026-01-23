# excel_compare_gui_flex.py
import os
import glob
import tempfile
from datetime import datetime
import pandas as pd

import tkinter as tk
from tkinter import ttk, messagebox

# ---------------- Excel column helpers ----------------
def col_to_index(col: str) -> int:
    col = col.strip().upper()
    n = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Ungültige Spalte: {col}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1

def index_to_col(idx: int) -> str:
    idx += 1
    res = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        res = chr(r + ord("A")) + res
    return res

def parse_cols_spec(spec: str) -> list[str]:
    """
    Accepts:
      - "D:I" or "D-I"
      - "D,E,F" or "D;E;F"
      - "AA:AD"
    Returns list of column labels.
    """
    s = spec.strip().upper().replace(" ", "")
    if not s:
        raise ValueError("Spaltenbereich ist leer.")

    s = s.replace("-", ":")
    if ":" in s:
        start, end = s.split(":", 1)
        if not start or not end:
            raise ValueError("Ungültiger Bereich. Beispiel: D:I")
        a, b = col_to_index(start), col_to_index(end)
        if b < a:
            a, b = b, a
        return [index_to_col(i) for i in range(a, b + 1)]

    parts = []
    for token in s.replace(";", ",").split(","):
        if token:
            parts.append(token)
    if not parts:
        raise ValueError("Ungültige Spaltenliste. Beispiel: D,E,F oder D:I")
    for p in parts:
        _ = col_to_index(p)
    return parts

# ---------------- Normalization ----------------
def normalize_value(v):
    if pd.isna(v):
        return ""
    if isinstance(v, str):
        return " ".join(v.split())  # trim + collapse whitespace
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (pd.Timestamp, datetime)):
        return pd.Timestamp(v).date().isoformat()
    return str(v)

# ---------------- Sheet resolution with validation ----------------
def resolve_sheet_name(xl: pd.ExcelFile, sheet_spec: str | None) -> str:
    names = xl.sheet_names
    if not sheet_spec or not sheet_spec.strip():
        return names[0]
    s = sheet_spec.strip()
    if s.isdigit():
        idx = int(s) - 1  # user uses 1-based
        if idx < 0 or idx >= len(names):
            raise ValueError(f"Blatt-Nummer {s} existiert nicht. Vorhanden: 1..{len(names)}")
        return names[idx]
    if s not in names:
        raise ValueError(f"Blatt '{s}' nicht gefunden. Vorhanden: {names}")
    return s

# ---------------- Row mapping (header=None!) ----------------
def excel_row_to_iloc(row_excel: int) -> int:
    """
    header=None => pandas row 0 corresponds to Excel row 1
    => iloc = excel_row - 1
    """
    return row_excel - 1

def iloc_to_excel_row(iloc: int) -> int:
    return iloc + 1

# ---------------- Read block from file (header=None) ----------------
def read_block(path: str, sheet_spec: str, key_col: str, compare_cols: list[str],
               row_start_excel: int, row_end_excel: int):
    xl = pd.ExcelFile(path)
    sheet_name = resolve_sheet_name(xl, sheet_spec)

    # IMPORTANT: header=None => we do not care about any header rows
    df = pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")

    needed = [key_col] + compare_cols
    idxs = [col_to_index(c) for c in needed]
    max_idx = max(idxs)
    if df.shape[1] <= max_idx:
        raise ValueError(
            f"{os.path.basename(path)} ({sheet_name}) hat nur {df.shape[1]} Spalten, "
            f"aber du verlangst bis {index_to_col(max_idx)}."
        )

    rs, re = int(row_start_excel), int(row_end_excel)
    if re < rs:
        rs, re = re, rs

    start_i = excel_row_to_iloc(rs)
    end_i = excel_row_to_iloc(re)
    if start_i < 0:
        start_i = 0
    if end_i >= len(df):
        end_i = len(df) - 1
    if start_i > end_i:
        raise ValueError(
            f"{os.path.basename(path)} ({sheet_name}): Zeilenbereich {row_start_excel}-{row_end_excel} "
            f"passt nicht zur Datei (zu wenig Zeilen)."
        )

    block = df.iloc[start_i:end_i + 1].copy().reset_index(drop=False).rename(columns={"index": "_iloc"})
    block["_excel_row"] = block["_iloc"].apply(iloc_to_excel_row)

    # select by position
    vals = block.iloc[:, [block.columns.get_loc("_excel_row")] + [block.columns.get_loc("_iloc")]].copy()
    data = block.iloc[:, idxs].copy()
    data.columns = needed

    # normalize
    key_col = key_col.strip().upper()
    data[key_col] = data[key_col].apply(normalize_value)
    for c in compare_cols:
        data[c] = data[c].apply(normalize_value)

    out = pd.concat([block[["_excel_row"]].copy(), data], axis=1)
    out = out[out[key_col] != ""].copy()

    out.attrs["sheet_name"] = sheet_name
    out.attrs["file_name"] = os.path.basename(path)
    return out

# ---------------- Compare blocks ----------------
def make_payload(row, compare_cols):
    return "|".join(str(row[c]) for c in compare_cols)

def compare_blocks(a, b, key_col, compare_cols):
    # duplicates handled by occurrence index (#1, #2, ...)
    a["_dup"] = a.groupby(key_col).cumcount() + 1
    b["_dup"] = b.groupby(key_col).cumcount() + 1
    a["_key2"] = a[key_col].astype(str) + "#" + a["_dup"].astype(str)
    b["_key2"] = b[key_col].astype(str) + "#" + b["_dup"].astype(str)

    a["_payload"] = a.apply(lambda r: make_payload(r, compare_cols), axis=1)
    b["_payload"] = b.apply(lambda r: make_payload(r, compare_cols), axis=1)

    m = a.merge(b, on="_key2", how="outer", suffixes=("_A", "_B"), indicator=True)

    def status(row):
        if row["_merge"] == "left_only":
            return "FEHLT_IN_B"
        if row["_merge"] == "right_only":
            return "FEHLT_IN_A"
        return "OK" if row["_payload_A"] == row["_payload_B"] else "ABWEICHUNG"

    m["STATUS"] = m.apply(status, axis=1)

    # Per-column diff flags (only meaningful if both)
    for c in compare_cols:
        m[f"DIFF_{c}"] = (m[f"{c}_A"].astype(str) != m[f"{c}_B"].astype(str)) & (m["_merge"] == "both")

    return m

# ---------------- Produce text report ----------------
def safe_write_path(preferred_dir: str, filename: str) -> str:
    """Try preferred_dir; if not writable, fall back to temp."""
    out_path = os.path.join(preferred_dir, filename)
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("")  # test write
        return out_path
    except Exception:
        return os.path.join(tempfile.gettempdir(), filename)

def write_text_report(m, compare_cols, out_txt_path,
                      fileA, sheetA, fileB, sheetB, key_col):
    # Determine if identical
    any_problem = (m["STATUS"] != "OK").any()

    lines = []
    lines.append("PRÜFPROTOKOLL")
    lines.append(f"A: {fileA} | Blatt: {sheetA}")
    lines.append(f"B: {fileB} | Blatt: {sheetB}")
    lines.append(f"Schlüsselspalte: {key_col} | Prüfspalten: {','.join(compare_cols)}")
    lines.append("")

    if not any_problem:
        lines.append("Beide Datenbereiche sind identisch.")
        with open(out_txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return

    # Missing rows
    missA = m[m["STATUS"] == "FEHLT_IN_A"].copy()
    missB = m[m["STATUS"] == "FEHLT_IN_B"].copy()

    if not missA.empty:
        lines.append("FEHLT_IN_A (existiert nur in B):")
        for _, r in missA.iterrows():
            key = r.get(f"{key_col}_B", "")
            row_b = r.get("_excel_row_B", "")
            occ = str(r.get("_key2","")).split("#")[-1]
            lines.append(f"  Key={key} (#{occ}): {fileB} {sheetB} Zeile {row_b}")
        lines.append("")

    if not missB.empty:
        lines.append("FEHLT_IN_B (existiert nur in A):")
        for _, r in missB.iterrows():
            key = r.get(f"{key_col}_A", "")
            row_a = r.get("_excel_row_A", "")
            occ = str(r.get("_key2","")).split("#")[-1]
            lines.append(f"  Key={key} (#{occ}): {fileA} {sheetA} Zeile {row_a}")
        lines.append("")

    # Cell-level differences
    diffs = m[m["STATUS"] == "ABWEICHUNG"].copy()
    if not diffs.empty:
        lines.append("ABWEICHUNGEN (Datei Blatt Zelle: Wert / Datei Blatt Zelle: Wert):")
        for _, r in diffs.iterrows():
            key = r.get(f"{key_col}_A", "")
            occ = str(r.get("_key2","")).split("#")[-1]
            row_a = int(r.get("_excel_row_A")) if pd.notna(r.get("_excel_row_A")) else None
            row_b = int(r.get("_excel_row_B")) if pd.notna(r.get("_excel_row_B")) else None

            for c in compare_cols:
                if bool(r.get(f"DIFF_{c}", False)):
                    cell_a = f"{c}{row_a}" if row_a is not None else f"{c}?"
                    cell_b = f"{c}{row_b}" if row_b is not None else f"{c}?"
                    va = "" if pd.isna(r.get(f"{c}_A", "")) else str(r.get(f"{c}_A", ""))
                    vb = "" if pd.isna(r.get(f"{c}_B", "")) else str(r.get(f"{c}_B", ""))
                    lines.append(
                        f"  Key={key} (#{occ}) | "
                        f"{fileA} {sheetA} {cell_a}: {va} / "
                        f"{fileB} {sheetB} {cell_b}: {vb}"
                    )

    with open(out_txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

# ---------------- Run compare with GUI inputs ----------------
def pick_two_xlsx(folder):
    files = sorted(glob.glob(os.path.join(folder, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).lower().startswith("pruefprotokoll")]
    if len(files) != 2:
        raise RuntimeError(f"Im Ordner müssen genau 2 XLSX-Dateien liegen (gefunden: {len(files)}).")
    return files[0], files[1]

def run_compare(fileA_path, fileB_path,
                sheetA_spec, cols_spec, startA, endA, keyA,
                sheetB_spec, cols_specB, startB, endB, keyB):
    # Key must be same on both
    keyA = keyA.strip().upper()
    keyB = keyB.strip().upper()
    if keyA != keyB:
        raise ValueError("Schlüsselspalte muss in beiden Dateien gleich sein (z.B. beide 'C').")
    key_col = keyA
    _ = col_to_index(key_col)

    compare_colsA = parse_cols_spec(cols_spec)
    compare_colsB = parse_cols_spec(cols_specB)
    if compare_colsA != compare_colsB:
        raise ValueError("Prüfspalten müssen in beiden Dateien identisch sein (z.B. überall D:I).")
    compare_cols = compare_colsA

    # validate row numbers
    for v, name in [(startA, "Startzeile A"), (endA, "Endzeile A"), (startB, "Startzeile B"), (endB, "Endzeile B")]:
        if not str(v).strip().isdigit():
            raise ValueError(f"{name} muss eine Zahl sein.")
    rsA, reA = int(startA), int(endA)
    rsB, reB = int(startB), int(endB)

    A = read_block(fileA_path, sheetA_spec, key_col, compare_cols, rsA, reA)
    B = read_block(fileB_path, sheetB_spec, key_col, compare_cols, rsB, reB)

    m = compare_blocks(A, B, key_col, compare_cols)

    # output path: try working dir, else temp
    preferred_dir = os.getcwd()
    out_txt = safe_write_path(preferred_dir, "pruefprotokoll.txt")

    sheetA_name = A.attrs.get("sheet_name", sheetA_spec)
    sheetB_name = B.attrs.get("sheet_name", sheetB_spec)

    write_text_report(
        m=m,
        compare_cols=compare_cols,
        out_txt_path=out_txt,
        fileA=os.path.basename(fileA_path),
        sheetA=sheetA_name,
        fileB=os.path.basename(fileB_path),
        sheetB=sheetB_name,
        key_col=key_col
    )

    return out_txt, os.path.basename(fileA_path), sheetA_name, os.path.basename(fileB_path), sheetB_name

# ---------------- GUI ----------------
def list_sheets(path: str) -> str:
    xl = pd.ExcelFile(path)
    names = xl.sheet_names
    out = [f"{os.path.basename(path)} (Blätter: {len(names)})"]
    for i, n in enumerate(names, start=1):
        out.append(f"  {i}: {n}")
    return "\n".join(out)

def main_gui():
    folder = os.getcwd()
    try:
        f1, f2 = pick_two_xlsx(folder)
    except Exception as e:
        messagebox.showerror("Fehler", str(e))
        return

    root = tk.Tk()
    root.title("Excel Blockvergleich (flexibel, header-egal)")

    frm = ttk.Frame(root, padding=12)
    frm.grid()

    fileA_path = tk.StringVar(value=f1)
    fileB_path = tk.StringVar(value=f2)

    ttk.Label(frm, text="Datei A:").grid(column=0, row=0, sticky="w")
    ttk.Label(frm, text=os.path.basename(f1)).grid(column=1, row=0, sticky="w")

    ttk.Label(frm, text="Datei B:").grid(column=0, row=1, sticky="w")
    ttk.Label(frm, text=os.path.basename(f2)).grid(column=1, row=1, sticky="w")

    def swap_files():
        a, b = fileA_path.get(), fileB_path.get()
        fileA_path.set(b); fileB_path.set(a)
        lblA.config(text=os.path.basename(fileA_path.get()))
        lblB.config(text=os.path.basename(fileB_path.get()))

    lblA = ttk.Label(frm, text=os.path.basename(fileA_path.get()))
    lblB = ttk.Label(frm, text=os.path.basename(fileB_path.get()))
    lblA.grid_forget(); lblB.grid_forget()

    # rewrite with actual labels so swap works
    for w in frm.grid_slaves():
        if isinstance(w, ttk.Label) and w.cget("text") == os.path.basename(f1):
            w.destroy()
        if isinstance(w, ttk.Label) and w.cget("text") == os.path.basename(f2):
            w.destroy()
    lblA = ttk.Label(frm, text=os.path.basename(fileA_path.get()))
    lblA.grid(column=1, row=0, sticky="w")
    lblB = ttk.Label(frm, text=os.path.basename(fileB_path.get()))
    lblB.grid(column=1, row=1, sticky="w")

    ttk.Button(frm, text="A ↔ B tauschen", command=swap_files).grid(column=2, row=0, rowspan=2, padx=(10,0))

    def show_sheets():
        try:
            a = list_sheets(fileA_path.get())
            b = list_sheets(fileB_path.get())
            messagebox.showinfo("Blätter anzeigen", a + "\n\n" + b)
        except Exception as e:
            messagebox.showerror("Fehler", str(e))

    ttk.Button(frm, text="Blätter anzeigen", command=show_sheets).grid(column=2, row=2, padx=(10,0), pady=(6,0))

    ttk.Separator(frm, orient="horizontal").grid(column=0, row=3, columnspan=3, sticky="ew", pady=10)

    ttk.Label(frm, text="Schlüsselspalte (für beide):").grid(column=0, row=4, sticky="w")
    key_var = tk.StringVar(value="C")
    ttk.Entry(frm, width=10, textvariable=key_var).grid(column=1, row=4, sticky="w")

    ttk.Label(frm, text="Prüfspalten (für beide, z.B. D:I):").grid(column=0, row=5, sticky="w")
    cols_var = tk.StringVar(value="D:I")
    ttk.Entry(frm, width=18, textvariable=cols_var).grid(column=1, row=5, sticky="w")

    ttk.Separator(frm, orient="horizontal").grid(column=0, row=6, columnspan=3, sticky="ew", pady=10)

    ttk.Label(frm, text="Einstellungen Datei A").grid(column=0, row=7, sticky="w")
    ttk.Label(frm, text="Einstellungen Datei B").grid(column=1, row=7, sticky="w")

    ttk.Label(frm, text="Blatt (Nr oder Name):").grid(column=0, row=8, sticky="w")
    sheetA = tk.StringVar(value="1")
    ttk.Entry(frm, width=18, textvariable=sheetA).grid(column=0, row=9, sticky="w")

    ttk.Label(frm, text="Blatt (Nr oder Name):").grid(column=1, row=8, sticky="w")
    sheetB = tk.StringVar(value="1")
    ttk.Entry(frm, width=18, textvariable=sheetB).grid(column=1, row=9, sticky="w")

    ttk.Label(frm, text="Startzeile:").grid(column=0, row=10, sticky="w")
    startA = tk.StringVar(value="45")
    ttk.Entry(frm, width=10, textvariable=startA).grid(column=0, row=11, sticky="w")

    ttk.Label(frm, text="Endzeile:").grid(column=0, row=12, sticky="w")
    endA = tk.StringVar(value="86")
    ttk.Entry(frm, width=10, textvariable=endA).grid(column=0, row=13, sticky="w")

    ttk.Label(frm, text="Startzeile:").grid(column=1, row=10, sticky="w")
    startB = tk.StringVar(value="47")
    ttk.Entry(frm, width=10, textvariable=startB).grid(column=1, row=11, sticky="w")

    ttk.Label(frm, text="Endzeile:").grid(column=1, row=12, sticky="w")
    endB = tk.StringVar(value="88")
    ttk.Entry(frm, width=10, textvariable=endB).grid(column=1, row=13, sticky="w")

    status = tk.StringVar(value="Start drücken → pruefprotokoll.txt (oder im TEMP-Ordner) wird erstellt.")
    ttk.Label(frm, textvariable=status, foreground="gray").grid(column=0, row=16, columnspan=3, sticky="w", pady=(10,0))

    def on_start():
        try:
            out_txt, a_name, a_sheet, b_name, b_sheet = run_compare(
                fileA_path.get(), fileB_path.get(),
                sheetA.get(), cols_var.get(), startA.get(), endA.get(), key_var.get(),
                sheetB.get(), cols_var.get(), startB.get(), endB.get(), key_var.get()
            )
            messagebox.showinfo(
                "Fertig",
                f"A: {a_name} | Blatt: {a_sheet} | Zeilen {startA.get()}-{endA.get()} | Spalten {cols_var.get()} | Key {key_var.get().upper()}\n"
                f"B: {b_name} | Blatt: {b_sheet} | Zeilen {startB.get()}-{endB.get()} | Spalten {cols_var.get()} | Key {key_var.get().upper()}\n\n"
                f"Ausgabe:\n{out_txt}"
            )
            status.set(f"OK. Protokoll geschrieben: {out_txt}")
        except Exception as e:
            messagebox.showerror("Fehler", str(e))
            status.set(f"Fehler: {e}")

    ttk.Button(frm, text="Start Vergleich", command=on_start).grid(column=0, row=15, sticky="w")
    ttk.Button(frm, text="Beenden", command=root.destroy).grid(column=1, row=15, sticky="w")

    root.mainloop()

if __name__ == "__main__":
    main_gui()


