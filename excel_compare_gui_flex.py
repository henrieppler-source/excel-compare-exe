from __future__ import annotations

import os
import sys
import glob
import tempfile
import configparser
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog


APP_VERSION = "5.0.3"
INI_NAME = "excel_compare.ini"
REPORT_NAME = "pruefprotokoll.txt"


# ---------------- paths (CRITICAL FIX) ----------------
def app_dir() -> str:
    # When frozen by PyInstaller, sys.executable is the EXE path.
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def ini_path() -> str:
    return os.path.join(app_dir(), INI_NAME)


def report_path_preferred() -> str:
    return os.path.join(app_dir(), REPORT_NAME)


def safe_write_path(preferred: str) -> str:
    # Try preferred dir; fallback to TEMP
    try:
        with open(preferred, "w", encoding="utf-8") as f:
            f.write("")
        return preferred
    except Exception:
        return os.path.join(tempfile.gettempdir(), REPORT_NAME)


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
    s = (spec or "").strip().upper().replace(" ", "").replace("-", ":")
    if not s:
        raise ValueError("Spaltenbereich ist leer.")
    if ":" in s:
        a, b = s.split(":", 1)
        ia, ib = col_to_index(a), col_to_index(b)
        if ib < ia:
            ia, ib = ib, ia
        return [index_to_col(i) for i in range(ia, ib + 1)]
    parts = [p for p in s.replace(";", ",").split(",") if p]
    for p in parts:
        _ = col_to_index(p)
    return parts


def normalize_value(v):
    if pd.isna(v):
        return ""
    if isinstance(v, str):
        return " ".join(v.split())
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return str(v).strip()


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


# ---------------- file resolution (AUTOMATIK) ----------------
def resolve_basename_in_appdir(name_or_path: str) -> str:
    """
    If absolute/path -> return as-is.
    If bare filename -> resolve relative to EXE folder.
    """
    s = (name_or_path or "").strip()
    if not s:
        return ""
    if os.path.isabs(s) or (":" in s) or ("/" in s) or ("\\" in s):
        return s
    return os.path.join(app_dir(), s)


def auto_find_in_appdir(filename: str) -> str:
    """
    Find exact filename in EXE folder. If not found, try glob with '*' around.
    """
    if not filename:
        return ""
    exact = os.path.join(app_dir(), filename)
    if os.path.exists(exact):
        return exact

    # try glob fallback: e.g. if INI stores Tabelle-1-Land_*.xlsx
    pattern = os.path.join(app_dir(), filename)
    hits = glob.glob(pattern)
    if hits:
        hits.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        return hits[0]

    # last resort: contains match
    base = filename.replace("*", "")
    if base:
        hits2 = glob.glob(os.path.join(app_dir(), f"*{base}*"))
        hits2 = [p for p in hits2 if p.lower().endswith(".xlsx")]
        if hits2:
            hits2.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            return hits2[0]
    return ""


def list_sheets(path: str) -> str:
    xl = pd.ExcelFile(path)
    names = xl.sheet_names
    out = [f"{os.path.basename(path)} (Blätter: {len(names)})"]
    for i, n in enumerate(names, start=1):
        out.append(f"  {i}: {n}")
    return "\n".join(out)


# ---------------- Read block ----------------
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

    start_i = rs - 1
    end_i = re - 1
    if start_i < 0:
        start_i = 0
    if end_i >= len(df):
        end_i = len(df) - 1
    if start_i > end_i:
        raise ValueError(
            f"{os.path.basename(path)} ({sheet_name}): Zeilenbereich {row_start_excel}-{row_end_excel} ist ungültig."
        )

    df_slice = df.iloc[start_i:end_i + 1, :].copy().reset_index(drop=True)
    df_slice["_excel_row"] = list(range(rs, rs + len(df_slice)))

    data = df_slice.iloc[:, idxs].copy()
    data.columns = needed

    data[key_col] = data[key_col].apply(normalize_value)
    for c in compare_cols:
        data[c] = data[c].apply(normalize_value)

    out = pd.concat([df_slice[["_excel_row"]], data], axis=1)
    out = out[out[key_col] != ""].copy()

    # Duplikate: occurrence in Block-Reihenfolge
    out["_occ"] = out.groupby(key_col).cumcount() + 1
    out["_key2"] = out[key_col].astype(str) + "#" + out["_occ"].astype(str)

    for i, c in enumerate(compare_cols):
        out[f"VAL_{i}"] = out[c]

    keep = ["_excel_row", key_col, "_occ", "_key2"] + [f"VAL_{i}" for i in range(len(compare_cols))]
    out = out[keep].copy()

    out.attrs["sheet_name"] = sheet_name
    return out


def compare_blocks(A: pd.DataFrame, B: pd.DataFrame, nvals: int) -> pd.DataFrame:
    m = A.merge(B, on="_key2", how="outer", suffixes=("_A", "_B"), indicator=True)

    def status(row):
        if row["_merge"] == "left_only":
            return "FEHLT_IN_B"
        if row["_merge"] == "right_only":
            return "FEHLT_IN_A"
        for i in range(nvals):
            if str(row.get(f"VAL_{i}_A", "")) != str(row.get(f"VAL_{i}_B", "")):
                return "ABWEICHUNG"
        return "OK"

    m["STATUS"] = m.apply(status, axis=1)
    for i in range(nvals):
        m[f"DIFF_{i}"] = (
            (m["_merge"] == "both")
            & (m.get(f"VAL_{i}_A").astype(str) != m.get(f"VAL_{i}_B").astype(str))
        )
    return m


def write_text_report(
    m: pd.DataFrame,
    out_txt_path: str,
    meta: dict,
):
    lines = []
    lines.append("PRÜFPROTOKOLL")
    lines.append(f"A: {meta['fileA']} | Blatt: {meta['sheetA']} | Key: {meta['keyA']} | Spalten: {','.join(meta['colsA'])} | Zeilen: {meta['rsA']}-{meta['reA']}")
    lines.append(f"B: {meta['fileB']} | Blatt: {meta['sheetB']} | Key: {meta['keyB']} | Spalten: {','.join(meta['colsB'])} | Zeilen: {meta['rsB']}-{meta['reB']}")
    lines.append("")

    probs = m[m["STATUS"] != "OK"]
    if probs.empty:
        lines.append("Beide Datenbereiche sind identisch.")
        with open(out_txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return

    # Missing
    missA = probs[probs["STATUS"] == "FEHLT_IN_A"]
    missB = probs[probs["STATUS"] == "FEHLT_IN_B"]
    if not missA.empty:
        lines.append("FEHLT_IN_A (existiert nur in B):")
        for _, r in missA.iterrows():
            key_val = r.get(f"{meta['keyB']}_B", "")
            occ = str(r["_key2"]).split("#")[-1]
            rowb = r.get("_excel_row_B", "?")
            lines.append(f"  Key={key_val} (#{occ}) | {meta['fileB']} {meta['sheetB']} Zeile {rowb}")
        lines.append("")
    if not missB.empty:
        lines.append("FEHLT_IN_B (existiert nur in A):")
        for _, r in missB.iterrows():
            key_val = r.get(f"{meta['keyA']}_A", "")
            occ = str(r["_key2"]).split("#")[-1]
            rowa = r.get("_excel_row_A", "?")
            lines.append(f"  Key={key_val} (#{occ}) | {meta['fileA']} {meta['sheetA']} Zeile {rowa}")
        lines.append("")

    diffs = probs[probs["STATUS"] == "ABWEICHUNG"]
    if not diffs.empty:
        lines.append("ABWEICHUNGEN (Datei Blatt Zelle: Wert / Datei Blatt Zelle: Wert):")
        n = min(len(meta["colsA"]), len(meta["colsB"]))
        for _, r in diffs.iterrows():
            key_val = r.get(f"{meta['keyA']}_A", r.get(f"{meta['keyB']}_B", ""))
            occ = str(r["_key2"]).split("#")[-1]
            rowa = r.get("_excel_row_A", None)
            rowb = r.get("_excel_row_B", None)
            for i in range(n):
                if bool(r.get(f"DIFF_{i}", False)):
                    ca = meta["colsA"][i]
                    cb = meta["colsB"][i]
                    va = r.get(f"VAL_{i}_A", "")
                    vb = r.get(f"VAL_{i}_B", "")
                    lines.append(
                        f"  Key={key_val} (#{occ}) | "
                        f"{meta['fileA']} {meta['sheetA']} {ca}{rowa}: {va} / "
                        f"{meta['fileB']} {meta['sheetB']} {cb}{rowb}: {vb}"
                    )

    with open(out_txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


# ---------------- INI ----------------
def load_ini() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    p = ini_path()
    if os.path.exists(p):
        cfg.read(p, encoding="utf-8")
    return cfg


def save_ini(cfg: configparser.ConfigParser):
    with open(ini_path(), "w", encoding="utf-8") as f:
        cfg.write(f)


# ---------------- GUI ----------------
def main_gui():
    cfg = load_ini()

    root = tk.Tk()
    root.title(f"Excel Blockvergleich (Presets) v{APP_VERSION}")

    # Full paths kept internally; show basenames only
    fileA_path = tk.StringVar(value="")
    fileB_path = tk.StringVar(value="")
    fileA_disp = tk.StringVar(value="")
    fileB_disp = tk.StringVar(value="")

    preset = tk.StringVar(value="")

    sheetA = tk.StringVar(value="1")
    keyA = tk.StringVar(value="C")
    colsA = tk.StringVar(value="D:K")
    rA1 = tk.StringVar(value="14")
    rA2 = tk.StringVar(value="59")

    sheetB = tk.StringVar(value="1")
    keyB = tk.StringVar(value="C")
    colsB = tk.StringVar(value="D:K")
    rB1 = tk.StringVar(value="16")
    rB2 = tk.StringVar(value="61")

    status = tk.StringVar(value=f"INI: {INI_NAME} (neben EXE) | Ausgabe: {REPORT_NAME}")

    def setA(p: str):
        fileA_path.set(p or "")
        fileA_disp.set(os.path.basename(p) if p else "")

    def setB(p: str):
        fileB_path.set(p or "")
        fileB_disp.set(os.path.basename(p) if p else "")

    def refresh_presets():
        combo["values"] = [""] + list(cfg.sections())

    frm = ttk.Frame(root, padding=12)
    frm.grid(sticky="nsew")

    ttk.Label(frm, text="Preset:").grid(row=0, column=0, sticky="w")
    combo = ttk.Combobox(frm, textvariable=preset, values=[""] + list(cfg.sections()), width=28, state="readonly")
    combo.grid(row=0, column=1, sticky="w")

    def load_preset():
        name = preset.get().strip()
        if not name:
            messagebox.showinfo("Preset", "Bitte ein Preset auswählen.")
            return
        if name not in cfg:
            messagebox.showerror("Preset", f"Preset '{name}' nicht gefunden.")
            return
        s = cfg[name]

        # load basenames; resolve in EXE dir
        a_name = s.get("fileA", "").strip()
        b_name = s.get("fileB", "").strip()
        a_path = auto_find_in_appdir(a_name) if a_name else ""
        b_path = auto_find_in_appdir(b_name) if b_name else ""

        setA(a_path if a_path else resolve_basename_in_appdir(a_name))
        setB(b_path if b_path else resolve_basename_in_appdir(b_name))

        sheetA.set(s.get("sheetA", sheetA.get()))
        keyA.set(s.get("keyA", keyA.get()))
        colsA.set(s.get("colsA", colsA.get()))
        rA1.set(s.get("startA", rA1.get()))
        rA2.set(s.get("endA", rA2.get()))

        sheetB.set(s.get("sheetB", sheetB.get()))
        keyB.set(s.get("keyB", keyB.get()))
        colsB.set(s.get("colsB", colsB.get()))
        rB1.set(s.get("startB", rB1.get()))
        rB2.set(s.get("endB", rB2.get()))

        messagebox.showinfo("Preset", f"Preset '{name}' geladen.\nDateien werden im EXE-Ordner automatisch gesucht.")

    def save_preset():
        name = simpledialog.askstring("Preset speichern", "Name (z.B. Tabelle-1):", parent=root)
        if not name:
            return
        name = name.strip()
        if name not in cfg:
            cfg.add_section(name)

        cfg[name]["fileA"] = os.path.basename(fileA_disp.get().strip()) if fileA_disp.get().strip() else ""
        cfg[name]["fileB"] = os.path.basename(fileB_disp.get().strip()) if fileB_disp.get().strip() else ""

        cfg[name]["sheetA"] = sheetA.get().strip()
        cfg[name]["keyA"] = keyA.get().strip()
        cfg[name]["colsA"] = colsA.get().strip()
        cfg[name]["startA"] = rA1.get().strip()
        cfg[name]["endA"] = rA2.get().strip()

        cfg[name]["sheetB"] = sheetB.get().strip()
        cfg[name]["keyB"] = keyB.get().strip()
        cfg[name]["colsB"] = colsB.get().strip()
        cfg[name]["startB"] = rB1.get().strip()
        cfg[name]["endB"] = rB2.get().strip()

        save_ini(cfg)
        refresh_presets()
        preset.set(name)
        messagebox.showinfo("Preset", f"Preset '{name}' gespeichert in {ini_path()}")

    ttk.Button(frm, text="Preset laden", command=load_preset).grid(row=0, column=2, padx=(10, 0))
    ttk.Button(frm, text="Preset speichern…", command=save_preset).grid(row=0, column=3, padx=(6, 0))

    ttk.Separator(frm, orient="horizontal").grid(row=1, column=0, columnspan=4, sticky="ew", pady=8)

    ttk.Label(frm, text="Datei A:").grid(row=2, column=0, sticky="w")
    ttk.Entry(frm, textvariable=fileA_disp, width=55, state="readonly").grid(row=2, column=1, columnspan=2, sticky="w")

    def browse_a():
        p = filedialog.askopenfilename(title="Datei A wählen", filetypes=[("Excel", "*.xlsx")])
        if p:
            setA(p)

    ttk.Button(frm, text="…", width=3, command=browse_a).grid(row=2, column=3, sticky="w")

    ttk.Label(frm, text="Datei B:").grid(row=3, column=0, sticky="w")
    ttk.Entry(frm, textvariable=fileB_disp, width=55, state="readonly").grid(row=3, column=1, columnspan=2, sticky="w")

    def browse_b():
        p = filedialog.askopenfilename(title="Datei B wählen", filetypes=[("Excel", "*.xlsx")])
        if p:
            setB(p)

    ttk.Button(frm, text="…", width=3, command=browse_b).grid(row=3, column=3, sticky="w")

    def swap_files():
        a, b = fileA_path.get(), fileB_path.get()
        setA(b)
        setB(a)

    ttk.Button(frm, text="A ↔ B tauschen", command=swap_files).grid(row=4, column=2, sticky="w", pady=(6, 0))

    def show_sheets():
        try:
            a = fileA_path.get()
            b = fileB_path.get()
            if not a or not os.path.exists(a):
                raise ValueError("Datei A fehlt/ungültig.")
            if not b or not os.path.exists(b):
                raise ValueError("Datei B fehlt/ungültig.")
            messagebox.showinfo("Blätter anzeigen", list_sheets(a) + "\n\n" + list_sheets(b))
        except Exception as e:
            messagebox.showerror("Fehler", str(e))

    ttk.Button(frm, text="Blätter anzeigen", command=show_sheets).grid(row=4, column=3, sticky="w", pady=(6, 0))

    def auto_find_files():
        # If display has a basename (e.g. after preset), try to resolve again
        a_name = fileA_disp.get().strip()
        b_name = fileB_disp.get().strip()
        if a_name:
            found = auto_find_in_appdir(a_name) or resolve_basename_in_appdir(a_name)
            setA(found if os.path.exists(found) else "")
            if not fileA_path.get():
                fileA_disp.set(a_name)
        if b_name:
            found = auto_find_in_appdir(b_name) or resolve_basename_in_appdir(b_name)
            setB(found if os.path.exists(found) else "")
            if not fileB_path.get():
                fileB_disp.set(b_name)
        messagebox.showinfo("Automatik", "Dateisuche im EXE-Ordner wurde ausgeführt.\n(Erwartung: Excel-Dateien liegen neben der EXE.)")

    ttk.Button(frm, text="Automatisch finden", command=auto_find_files).grid(row=4, column=1, sticky="w", pady=(6, 0))

    ttk.Separator(frm, orient="horizontal").grid(row=5, column=0, columnspan=4, sticky="ew", pady=10)

    ttk.Label(frm, text="Einstellungen Datei A", font=("Segoe UI", 9, "bold")).grid(row=6, column=0, sticky="w")
    ttk.Label(frm, text="Einstellungen Datei B", font=("Segoe UI", 9, "bold")).grid(row=6, column=2, sticky="w")

    def row(label, var, r, c):
        ttk.Label(frm, text=label).grid(row=r, column=c, sticky="w")
        ttk.Entry(frm, textvariable=var, width=10).grid(row=r, column=c + 1, sticky="w")

    row("Blatt (Nr/Name):", sheetA, 7, 0)
    row("Schlüsselspalte:", keyA, 8, 0)
    row("Vergleichsspalten:", colsA, 9, 0)
    row("Startzeile:", rA1, 10, 0)
    row("Endzeile:", rA2, 11, 0)

    row("Blatt (Nr/Name):", sheetB, 7, 2)
    row("Schlüsselspalte:", keyB, 8, 2)
    row("Vergleichsspalten:", colsB, 9, 2)
    row("Startzeile:", rB1, 10, 2)
    row("Endzeile:", rB2, 11, 2)

    ttk.Separator(frm, orient="horizontal").grid(row=12, column=0, columnspan=4, sticky="ew", pady=10)
    ttk.Label(frm, textvariable=status, foreground="gray").grid(row=13, column=0, columnspan=4, sticky="w")

    def start_compare():
        try:
            a = fileA_path.get()
            b = fileB_path.get()
            if not a or not os.path.exists(a):
                raise ValueError("Datei A nicht gefunden. (Tipp: Dateien neben die EXE legen oder Dialog benutzen.)")
            if not b or not os.path.exists(b):
                raise ValueError("Datei B nicht gefunden. (Tipp: Dateien neben die EXE legen oder Dialog benutzen.)")

            colsA_list = parse_cols_spec(colsA.get())
            colsB_list = parse_cols_spec(colsB.get())
            if len(colsA_list) != len(colsB_list):
                raise ValueError("Vergleichsspalten müssen gleich viele Spalten haben (A und B).")

            rsA, reA = int(rA1.get()), int(rA2.get())
            rsB, reB = int(rB1.get()), int(rB2.get())

            A = read_block(a, sheetA.get(), keyA.get(), colsA_list, rsA, reA)
            B = read_block(b, sheetB.get(), keyB.get(), colsB_list, rsB, reB)

            m = compare_blocks(A, B, nvals=len(colsA_list))

            out = safe_write_path(report_path_preferred())
            write_text_report(
                m, out,
                meta={
                    "fileA": os.path.basename(a),
                    "fileB": os.path.basename(b),
                    "sheetA": A.attrs.get("sheet_name", sheetA.get()),
                    "sheetB": B.attrs.get("sheet_name", sheetB.get()),
                    "keyA": keyA.get().strip().upper(),
                    "keyB": keyB.get().strip().upper(),
                    "colsA": colsA_list,
                    "colsB": colsB_list,
                    "rsA": rsA, "reA": reA,
                    "rsB": rsB, "reB": reB,
                }
            )

            status.set(f"OK: {out}")
            messagebox.showinfo("Fertig", f"Prüfprotokoll erstellt:\n{out}")

        except Exception as e:
            status.set(f"Fehler: {e}")
            messagebox.showerror("Fehler", str(e))

    ttk.Button(frm, text="Start Vergleich", command=start_compare).grid(row=14, column=0, sticky="w")
    ttk.Button(frm, text="Beenden", command=root.destroy).grid(row=14, column=1, sticky="w")

    root.mainloop()


if __name__ == "__main__":
    main_gui()
