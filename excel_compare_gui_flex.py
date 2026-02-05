import os
import sys
import glob
import tempfile
from datetime import datetime
import configparser

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

__version__ = "2.0.0"
__build_date__ = "2026-02-04"


# ---------------- Excel column helpers ----------------
def col_to_index(col: str) -> int:
    col = col.strip().upper()
    if not col:
        raise ValueError("Leere Spaltenangabe.")
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
      - "D:K" or "D-K"
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
            raise ValueError("Ungültiger Bereich. Beispiel: D:K")
        a, b = col_to_index(start), col_to_index(end)
        if b < a:
            a, b = b, a
        return [index_to_col(i) for i in range(a, b + 1)]

    parts = [p for p in s.replace(";", ",").split(",") if p]
    if not parts:
        raise ValueError("Ungültige Spaltenliste. Beispiel: D,E,F oder D:K")
    for p in parts:
        _ = col_to_index(p)
    return parts


def normalize_value(v):
    if pd.isna(v):
        return ""
    if isinstance(v, str):
        return " ".join(v.split())
    if isinstance(v, (int, float)):
        # float-normalization helps avoid 12 vs 12.0 issues
        return float(v)
    if isinstance(v, (pd.Timestamp, datetime)):
        return pd.Timestamp(v).date().isoformat()
    return str(v)


# ---------------- Sheet resolution ----------------
def resolve_sheet_name(xl: pd.ExcelFile, sheet_spec: str | None) -> str:
    names = xl.sheet_names
    if not sheet_spec or not str(sheet_spec).strip():
        return names[0]
    s = str(sheet_spec).strip()
    if s.isdigit():
        idx = int(s) - 1
        if idx < 0 or idx >= len(names):
            raise ValueError(f"Blatt-Nummer {s} existiert nicht. Vorhanden: 1..{len(names)}")
        return names[idx]
    if s not in names:
        raise ValueError(f"Blatt '{s}' nicht gefunden. Vorhanden: {names}")
    return s


def excel_row_to_iloc(row_excel: int) -> int:
    return row_excel - 1


def iloc_to_excel_row(iloc: int) -> int:
    return iloc + 1


# ---------------- Read block (header=None) ----------------
def read_block(
    path: str,
    sheet_spec: str,
    key_col: str,
    compare_cols: list[str],
    row_start_excel: int,
    row_end_excel: int,
) -> pd.DataFrame:
    key_col = key_col.strip().upper()
    compare_cols = [c.strip().upper() for c in compare_cols]

    xl = pd.ExcelFile(path)
    sheet_name = resolve_sheet_name(xl, sheet_spec)

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

    data = block.iloc[:, idxs].copy()
    data.columns = needed

    # normalize
    data[key_col] = data[key_col].apply(normalize_value)
    for c in compare_cols:
        data[c] = data[c].apply(normalize_value)

    out = pd.concat([block[["_excel_row"]].copy(), data], axis=1)

    # drop empty key
    out = out[out[key_col] != ""].copy()

    # occurrence per key (order-sensitive!)
    out["_occ"] = out.groupby(key_col).cumcount() + 1
    out["_key2"] = out[key_col].astype(str) + "#" + out["_occ"].astype(str)

    # store values as VAL_0..VAL_n-1 for position-based comparison
    for i, c in enumerate(compare_cols):
        out[f"VAL_{i}"] = out[c]

    # keep only what we need
    keep = ["_excel_row", key_col, "_occ", "_key2"] + [f"VAL_{i}" for i in range(len(compare_cols))]
    out = out[keep].copy()

    out.attrs["sheet_name"] = sheet_name
    out.attrs["file_name"] = os.path.basename(path)
    out.attrs["key_col"] = key_col
    out.attrs["compare_cols"] = compare_cols
    return out


# ---------------- Compare blocks (order-sensitive duplicates) ----------------
def compare_blocks(
    A: pd.DataFrame,
    B: pd.DataFrame,
    nvals: int
) -> pd.DataFrame:
    # Merge on key+occurrence: "Zusammen#1", "Zusammen#2", ...
    m = A.merge(B, on="_key2", how="outer", suffixes=("_A", "_B"), indicator=True)

    def status(row):
        if row["_merge"] == "left_only":
            return "FEHLT_IN_B"
        if row["_merge"] == "right_only":
            return "FEHLT_IN_A"
        # both: check all VAL_i
        for i in range(nvals):
            va = row.get(f"VAL_{i}_A", "")
            vb = row.get(f"VAL_{i}_B", "")
            if str(va) != str(vb):
                return "ABWEICHUNG"
        return "OK"

    m["STATUS"] = m.apply(status, axis=1)

    # per-position diffs
    for i in range(nvals):
        m[f"DIFF_{i}"] = (
            (m["_merge"] == "both") &
            (m.get(f"VAL_{i}_A").astype(str) != m.get(f"VAL_{i}_B").astype(str))
        )

    return m


# ---------------- Reporting ----------------
def get_app_dir() -> str:
    """Directory where the running app/exe resides (not the working directory)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_default_output_dir() -> str:
    """Prefer app dir if writable; otherwise fall back to Documents\ExcelBlockvergleich."""
    app_dir = get_app_dir()

    # Try app dir
    try:
        test = os.path.join(app_dir, "_write_test.tmp")
        with open(test, "w", encoding="utf-8") as f:
            f.write("x")
        os.remove(test)
        return app_dir
    except Exception:
        pass

    docs = os.path.join(os.path.expanduser("~"), "Documents", "ExcelBlockvergleich")
    os.makedirs(docs, exist_ok=True)
    return docs


def safe_write_path(filename: str) -> str:
    out_dir = get_default_output_dir()
    return os.path.join(out_dir, filename)


def make_report_filename(prefix: str = "pruefprotokoll", ext: str = "txt") -> str:
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    return f"{prefix}_{ts}.{ext}"


def write_text_report(
    m: pd.DataFrame,
    out_txt_path: str,
    fileA: str, sheetA: str, keyA: str, colsA: list[str], rsA: int, reA: int,
    fileB: str, sheetB: str, keyB: str, colsB: list[str], rsB: int, reB: int,
):
    any_problem = (m["STATUS"] != "OK").any()

    lines = []
    lines.append("PRÜFPROTOKOLL")
    lines.append(f"Version: {__version__} | Build: {__build_date__} | Lauf: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Protokollpfad: {out_txt_path}")
    lines.append(f"A: {fileA} | Blatt: {sheetA} | Key: {keyA} | Spalten: {','.join(colsA)} | Zeilen: {rsA}-{reA}")
    lines.append(f"B: {fileB} | Blatt: {sheetB} | Key: {keyB} | Spalten: {','.join(colsB)} | Zeilen: {rsB}-{reB}")
    lines.append("")

    if not any_problem:
        lines.append("Beide Datenbereiche sind identisch.")
        with open(out_txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return

    missA = m[m["STATUS"] == "FEHLT_IN_A"].copy()
    missB = m[m["STATUS"] == "FEHLT_IN_B"].copy()

    if not missA.empty:
        lines.append("FEHLT_IN_A (existiert nur in B):")
        for _, r in missA.iterrows():
            key_val = r.get(f"{keyB}_B", "")
            occ = str(r.get("_key2", "")).split("#")[-1]
            row_b = r.get("_excel_row_B", "")
            lines.append(f"  Key={key_val} (#{occ}): {fileB} {sheetB} Zeile {row_b}")
        lines.append("")

    if not missB.empty:
        lines.append("FEHLT_IN_B (existiert nur in A):")
        for _, r in missB.iterrows():
            key_val = r.get(f"{keyA}_A", "")
            occ = str(r.get("_key2", "")).split("#")[-1]
            row_a = r.get("_excel_row_A", "")
            lines.append(f"  Key={key_val} (#{occ}): {fileA} {sheetA} Zeile {row_a}")
        lines.append("")

    diffs = m[m["STATUS"] == "ABWEICHUNG"].copy()
    if not diffs.empty:
        lines.append("ABWEICHUNGEN (Datei Blatt Zelle: Wert / Datei Blatt Zelle: Wert):")
        for _, r in diffs.iterrows():
            key_val = r.get(f"{keyA}_A", r.get(f"{keyB}_B", ""))
            occ = str(r.get("_key2", "")).split("#")[-1]
            row_a = int(r.get("_excel_row_A")) if pd.notna(r.get("_excel_row_A")) else None
            row_b = int(r.get("_excel_row_B")) if pd.notna(r.get("_excel_row_B")) else None

            for i in range(min(len(colsA), len(colsB))):
                if bool(r.get(f"DIFF_{i}", False)):
                    colA = colsA[i]
                    colB = colsB[i]
                    cell_a = f"{colA}{row_a}" if row_a is not None else f"{colA}?"
                    cell_b = f"{colB}{row_b}" if row_b is not None else f"{colB}?"
                    va = "" if pd.isna(r.get(f"VAL_{i}_A", "")) else str(r.get(f"VAL_{i}_A", ""))
                    vb = "" if pd.isna(r.get(f"VAL_{i}_B", "")) else str(r.get(f"VAL_{i}_B", ""))
                    lines.append(
                        f"  Key={key_val} (#{occ}) | "
                        f"{fileA} {sheetA} {cell_a}: {va} / "
                        f"{fileB} {sheetB} {cell_b}: {vb}"
                    )

    with open(out_txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


# ---------------- File helpers ----------------
def pick_two_xlsx(folder: str) -> tuple[str, str]:
    files = sorted(glob.glob(os.path.join(folder, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).lower().startswith("pruefprotokoll")]
    if len(files) == 2:
        return files[0], files[1]
    # fallback: empty
    return "", ""


def list_sheets(path: str) -> str:
    xl = pd.ExcelFile(path)
    names = xl.sheet_names
    out = [f"{os.path.basename(path)} (Blätter: {len(names)})"]
    for i, n in enumerate(names, start=1):
        out.append(f"  {i}: {n}")
    return "\n".join(out)


# ---------------- INI Presets ----------------
INI_NAME = "excel_compare.ini"


def load_ini(ini_path: str) -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    if os.path.exists(ini_path):
        cfg.read(ini_path, encoding="utf-8")
    return cfg


def save_ini(cfg: configparser.ConfigParser, ini_path: str):
    with open(ini_path, "w", encoding="utf-8") as f:
        cfg.write(f)


def preset_sections(cfg: configparser.ConfigParser) -> list[str]:
    # all non-default sections
    return [s for s in cfg.sections()]


def apply_preset_to_vars(cfg: configparser.ConfigParser, section: str, vars_map: dict[str, tk.StringVar]):
    if section not in cfg:
        raise ValueError(f"Preset '{section}' nicht gefunden.")
    sec = cfg[section]
    for k, var in vars_map.items():
        if k in sec:
            var.set(sec.get(k, ""))


def write_vars_to_preset(cfg: configparser.ConfigParser, section: str, vars_map: dict[str, tk.StringVar]):
    if section not in cfg:
        cfg.add_section(section)
    for k, var in vars_map.items():
        cfg[section][k] = var.get().strip()


# ---------------- Compare runner ----------------
def run_compare(
    fileA_path: str, fileB_path: str,
    sheetA_spec: str, keyA: str, colsA_spec: str, startA: str, endA: str,
    sheetB_spec: str, keyB: str, colsB_spec: str, startB: str, endB: str,
):
    if not fileA_path or not os.path.exists(fileA_path):
        raise ValueError("Datei A fehlt oder existiert nicht.")
    if not fileB_path or not os.path.exists(fileB_path):
        raise ValueError("Datei B fehlt oder existiert nicht.")

    keyA = keyA.strip().upper()
    keyB = keyB.strip().upper()
    _ = col_to_index(keyA)
    _ = col_to_index(keyB)

    colsA = parse_cols_spec(colsA_spec)
    colsB = parse_cols_spec(colsB_spec)
    if len(colsA) != len(colsB):
        raise ValueError("Vergleichsspalten müssen gleich viele Spalten haben (A und B).")

    for v, name in [(startA, "Startzeile A"), (endA, "Endzeile A"), (startB, "Startzeile B"), (endB, "Endzeile B")]:
        if not str(v).strip().isdigit():
            raise ValueError(f"{name} muss eine Zahl sein.")

    rsA, reA = int(startA), int(endA)
    rsB, reB = int(startB), int(endB)

    A = read_block(fileA_path, sheetA_spec, keyA, colsA, rsA, reA)
    B = read_block(fileB_path, sheetB_spec, keyB, colsB, rsB, reB)

    nvals = len(colsA)
    m = compare_blocks(A, B, nvals=nvals)

    out_txt = safe_write_path(make_report_filename())

    sheetA_name = A.attrs.get("sheet_name", sheetA_spec)
    sheetB_name = B.attrs.get("sheet_name", sheetB_spec)

    write_text_report(
        m=m,
        out_txt_path=out_txt,
        fileA=os.path.basename(fileA_path), sheetA=sheetA_name, keyA=keyA, colsA=colsA, rsA=rsA, reA=reA,
        fileB=os.path.basename(fileB_path), sheetB=sheetB_name, keyB=keyB, colsB=colsB, rsB=rsB, reB=reB,
    )

    return out_txt


# ---------------- GUI ----------------
def main_gui():
    folder = os.getcwd()
    ini_path = os.path.join(folder, INI_NAME)
    cfg = load_ini(ini_path)

    root = tk.Tk()
    root.title("Excel Blockvergleich (key-basiert, Duplikate order-sensitiv, Presets)")

    frm = ttk.Frame(root, padding=12)
    frm.grid(sticky="nsew")

    # Default file pick: if exactly 2 xlsx in folder
    f1, f2 = pick_two_xlsx(folder)

    # Variables
    fileA_var = tk.StringVar(value=f1)
    fileB_var = tk.StringVar(value=f2)

    preset_var = tk.StringVar(value="")

    sheetA_var = tk.StringVar(value="1")
    keyA_var = tk.StringVar(value="C")
    colsA_var = tk.StringVar(value="D:K")
    startA_var = tk.StringVar(value="14")
    endA_var = tk.StringVar(value="59")

    sheetB_var = tk.StringVar(value="1")
    keyB_var = tk.StringVar(value="C")
    colsB_var = tk.StringVar(value="D:K")
    startB_var = tk.StringVar(value="16")
    endB_var = tk.StringVar(value="61")

    # Map keys for INI
    vars_map = {
        "fileA": fileA_var,
        "fileB": fileB_var,
        "sheetA": sheetA_var,
        "keyA": keyA_var,
        "colsA": colsA_var,
        "startA": startA_var,
        "endA": endA_var,
        "sheetB": sheetB_var,
        "keyB": keyB_var,
        "colsB": colsB_var,
        "startB": startB_var,
        "endB": endB_var,
    }

    # --- Presets row ---
    ttk.Label(frm, text="Preset:").grid(column=0, row=0, sticky="w")
    presets = [""] + preset_sections(cfg)  # empty = none
    preset_combo = ttk.Combobox(frm, textvariable=preset_var, values=presets, width=30, state="readonly")
    preset_combo.grid(column=1, row=0, sticky="w")

    def refresh_presets():
        preset_combo["values"] = [""] + preset_sections(cfg)

    def load_preset():
        name = preset_var.get().strip()
        if not name:
            messagebox.showinfo("Preset", "Bitte ein Preset auswählen.")
            return
        apply_preset_to_vars(cfg, name, vars_map)
        messagebox.showinfo("Preset", f"Preset '{name}' geladen.")

    def save_preset_as():
        name = simpledialog.askstring("Preset speichern", "Name für das Preset (z.B. Tabelle-1):", parent=root)
        if not name:
            return
        name = name.strip()
        write_vars_to_preset(cfg, name, vars_map)
        save_ini(cfg, ini_path)
        refresh_presets()
        preset_var.set(name)
        messagebox.showinfo("Preset", f"Preset '{name}' gespeichert in {INI_NAME}.")

    ttk.Button(frm, text="Preset laden", command=load_preset).grid(column=2, row=0, padx=(10, 0))
    ttk.Button(frm, text="Preset speichern…", command=save_preset_as).grid(column=3, row=0, padx=(6, 0))

    ttk.Separator(frm, orient="horizontal").grid(column=0, row=1, columnspan=4, sticky="ew", pady=8)

    # --- Files row ---
    ttk.Label(frm, text="Datei A:").grid(column=0, row=2, sticky="w")
    ttk.Entry(frm, textvariable=fileA_var, width=55).grid(column=1, row=2, columnspan=2, sticky="w")

    def browse_a():
        p = filedialog.askopenfilename(title="Datei A wählen", filetypes=[("Excel", "*.xlsx")])
        if p:
            fileA_var.set(p)

    ttk.Button(frm, text="…", width=3, command=browse_a).grid(column=3, row=2, sticky="w")

    ttk.Label(frm, text="Datei B:").grid(column=0, row=3, sticky="w")
    ttk.Entry(frm, textvariable=fileB_var, width=55).grid(column=1, row=3, columnspan=2, sticky="w")

    def browse_b():
        p = filedialog.askopenfilename(title="Datei B wählen", filetypes=[("Excel", "*.xlsx")])
        if p:
            fileB_var.set(p)

    ttk.Button(frm, text="…", width=3, command=browse_b).grid(column=3, row=3, sticky="w")

    def swap_files():
        a, b = fileA_var.get(), fileB_var.get()
        fileA_var.set(b)
        fileB_var.set(a)

    ttk.Button(frm, text="A ↔ B tauschen", command=swap_files).grid(column=2, row=4, sticky="w", pady=(6, 0))

    def show_sheets():
        try:
            a = list_sheets(fileA_var.get())
            b = list_sheets(fileB_var.get())
            messagebox.showinfo("Blätter anzeigen", a + "\n\n" + b)
        except Exception as e:
            messagebox.showerror("Fehler", str(e))

    ttk.Button(frm, text="Blätter anzeigen", command=show_sheets).grid(column=3, row=4, sticky="w", pady=(6, 0))

    ttk.Separator(frm, orient="horizontal").grid(column=0, row=5, columnspan=4, sticky="ew", pady=10)

    # --- Settings columns ---
    ttk.Label(frm, text="Einstellungen Datei A", font=("Segoe UI", 9, "bold")).grid(column=0, row=6, sticky="w")
    ttk.Label(frm, text="Einstellungen Datei B", font=("Segoe UI", 9, "bold")).grid(column=2, row=6, sticky="w")

    # A fields
    ttk.Label(frm, text="Blatt (Nr/Name):").grid(column=0, row=7, sticky="w")
    ttk.Entry(frm, textvariable=sheetA_var, width=10).grid(column=1, row=7, sticky="w")

    ttk.Label(frm, text="Schlüsselspalte:").grid(column=0, row=8, sticky="w")
    ttk.Entry(frm, textvariable=keyA_var, width=10).grid(column=1, row=8, sticky="w")

    ttk.Label(frm, text="Vergleichsspalten:").grid(column=0, row=9, sticky="w")
    ttk.Entry(frm, textvariable=colsA_var, width=10).grid(column=1, row=9, sticky="w")

    ttk.Label(frm, text="Startzeile:").grid(column=0, row=10, sticky="w")
    ttk.Entry(frm, textvariable=startA_var, width=10).grid(column=1, row=10, sticky="w")

    ttk.Label(frm, text="Endzeile:").grid(column=0, row=11, sticky="w")
    ttk.Entry(frm, textvariable=endA_var, width=10).grid(column=1, row=11, sticky="w")

    # B fields
    ttk.Label(frm, text="Blatt (Nr/Name):").grid(column=2, row=7, sticky="w")
    ttk.Entry(frm, textvariable=sheetB_var, width=10).grid(column=3, row=7, sticky="w")

    ttk.Label(frm, text="Schlüsselspalte:").grid(column=2, row=8, sticky="w")
    ttk.Entry(frm, textvariable=keyB_var, width=10).grid(column=3, row=8, sticky="w")

    ttk.Label(frm, text="Vergleichsspalten:").grid(column=2, row=9, sticky="w")
    ttk.Entry(frm, textvariable=colsB_var, width=10).grid(column=3, row=9, sticky="w")

    ttk.Label(frm, text="Startzeile:").grid(column=2, row=10, sticky="w")
    ttk.Entry(frm, textvariable=startB_var, width=10).grid(column=3, row=10, sticky="w")

    ttk.Label(frm, text="Endzeile:").grid(column=2, row=11, sticky="w")
    ttk.Entry(frm, textvariable=endB_var, width=10).grid(column=3, row=11, sticky="w")

    ttk.Separator(frm, orient="horizontal").grid(column=0, row=12, columnspan=4, sticky="ew", pady=10)

    status = tk.StringVar(value=f"Ausgabeordner: {get_default_output_dir()} (Protokollname mit Zeitstempel). Presets: {INI_NAME}")
    ttk.Label(frm, textvariable=status, foreground="gray").grid(column=0, row=13, columnspan=4, sticky="w")

def show_about():
    exe_path = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
    out_dir = get_default_output_dir()
    info = (
        f"Excel Blockvergleich\n"
        f"Version: {__version__} (Build: {__build_date__})\n\n"
        f"Programm: {exe_path}\n"
        f"INI: {ini_path}\n"
        f"Ausgabeordner: {out_dir}\n"
    )
    messagebox.showinfo("Über", info)

    def on_start():
        try:
            out_txt = run_compare(
                fileA_var.get(), fileB_var.get(),
                sheetA_var.get(), keyA_var.get(), colsA_var.get(), startA_var.get(), endA_var.get(),
                sheetB_var.get(), keyB_var.get(), colsB_var.get(), startB_var.get(), endB_var.get(),
            )
            messagebox.showinfo("Fertig", f"Protokoll:\n{out_txt}")
            status.set(f"OK: {out_txt}")
        except Exception as e:
            messagebox.showerror("Fehler", str(e))
            status.set(f"Fehler: {e}")

    ttk.Button(frm, text="Start Vergleich", command=on_start).grid(column=0, row=14, sticky="w")
    ttk.Button(frm, text="Beenden", command=root.destroy).grid(column=1, row=14, sticky="w")
    ttk.Button(frm, text="Über…", command=show_about).grid(column=2, row=14, sticky="w")

    root.mainloop()


if __name__ == "__main__":
    main_gui()
