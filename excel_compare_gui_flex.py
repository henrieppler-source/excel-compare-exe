import os
import sys
import glob
from datetime import datetime
import configparser

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# ================== META ==================
__version__ = "2.0.1"
__build_date__ = "2026-02-05"

# ================== PATH / LOG ==================
def get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def bootlog(msg: str):
    try:
        with open(os.path.join(get_base_dir(), "startup.log"), "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat()} {msg}\n")
    except Exception:
        pass


# ================== EXCEL COLUMN HELPERS ==================
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
    s = spec.strip().upper().replace(" ", "").replace("-", ":")
    if ":" in s:
        a, b = s.split(":")
        ia, ib = col_to_index(a), col_to_index(b)
        if ib < ia:
            ia, ib = ib, ia
        return [index_to_col(i) for i in range(ia, ib + 1)]
    return [p for p in s.replace(";", ",").split(",") if p]


def normalize_value(v):
    if pd.isna(v):
        return ""
    if isinstance(v, str):
        return " ".join(v.split())
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (pd.Timestamp, datetime)):
        return pd.Timestamp(v).date().isoformat()
    return str(v)


# ================== SHEET ==================
def resolve_sheet_name(xl: pd.ExcelFile, spec: str) -> str:
    names = xl.sheet_names
    if spec.isdigit():
        return names[int(spec) - 1]
    if spec not in names:
        raise ValueError(f"Blatt '{spec}' nicht gefunden.")
    return spec


# ================== READ BLOCK (FIXED!) ==================
def read_block(path, sheet_spec, key_col, compare_cols, row_start, row_end):
    key_col = key_col.upper()
    compare_cols = [c.upper() for c in compare_cols]

    xl = pd.ExcelFile(path)
    sheet_name = resolve_sheet_name(xl, sheet_spec)

    df = pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")

    needed = [key_col] + compare_cols
    idxs = [col_to_index(c) for c in needed]

    rs, re = int(row_start) - 1, int(row_end)
    block_df = df.iloc[rs:re].copy()

    data = block_df.iloc[:, idxs].copy()
    data.columns = needed

    data[key_col] = data[key_col].apply(normalize_value)
    for c in compare_cols:
        data[c] = data[c].apply(normalize_value)

    data["_excel_row"] = range(row_start, row_start + len(data))
    data = data[data[key_col] != ""].copy()

    data["_occ"] = data.groupby(key_col).cumcount() + 1
    data["_key2"] = data[key_col].astype(str) + "#" + data["_occ"].astype(str)

    for i, c in enumerate(compare_cols):
        data[f"VAL_{i}"] = data[c]

    keep = ["_excel_row", key_col, "_key2"] + [f"VAL_{i}" for i in range(len(compare_cols))]
    out = data[keep].copy()

    out.attrs["sheet_name"] = sheet_name
    return out


# ================== COMPARE ==================
def compare_blocks(A, B, nvals):
    m = A.merge(B, on="_key2", how="outer", suffixes=("_A", "_B"), indicator=True)

    def status(r):
        if r["_merge"] == "left_only":
            return "FEHLT_IN_B"
        if r["_merge"] == "right_only":
            return "FEHLT_IN_A"
        for i in range(nvals):
            if str(r[f"VAL_{i}_A"]) != str(r[f"VAL_{i}_B"]):
                return "ABWEICHUNG"
        return "OK"

    m["STATUS"] = m.apply(status, axis=1)
    return m


# ================== REPORT ==================
def sanitize(s):
    for ch in '\\/:*?"<>| ':
        s = s.replace(ch, "-")
    return s.strip("-")


def make_report_name(sheet):
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    return f"Pruefprotokoll_{sanitize(sheet)}_{ts}.txt"


def write_report(m, out_path, infoA, infoB):
    lines = [
        "PRÜFPROTOKOLL",
        f"Version {__version__} ({__build_date__})",
        "",
        infoA,
        infoB,
        ""
    ]

    if not (m["STATUS"] != "OK").any():
        lines.append("Beide Bereiche sind identisch.")
    else:
        for _, r in m[m["STATUS"] == "ABWEICHUNG"].iterrows():
            for c in [c for c in r.index if c.startswith("VAL_")]:
                if str(r[f"{c}_A"]) != str(r[f"{c}_B"]):
                    lines.append(
                        f"{r['_key2']} | {c}: {r[f'{c}_A']} <> {r[f'{c}_B']}"
                    )

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


# ================== GUI ==================
INI_NAME = "excel_compare.ini"


def main_gui():
    bootlog("START GUI")
    root = tk.Tk()
    root.title(f"Excel Blockvergleich v{__version__}")

    frm = ttk.Frame(root, padding=12)
    frm.grid()

    fileA = tk.StringVar()
    fileB = tk.StringVar()
    sheetA = tk.StringVar(value="1")
    sheetB = tk.StringVar(value="1")
    keyA = tk.StringVar(value="A")
    keyB = tk.StringVar(value="A")
    colsA = tk.StringVar(value="B:K")
    colsB = tk.StringVar(value="B:K")
    startA = tk.StringVar(value="1")
    endA = tk.StringVar(value="10")
    startB = tk.StringVar(value="1")
    endB = tk.StringVar(value="10")

    def run():
        A = read_block(fileA.get(), sheetA.get(), keyA.get(),
                       parse_cols_spec(colsA.get()), startA.get(), endA.get())
        B = read_block(fileB.get(), sheetB.get(), keyB.get(),
                       parse_cols_spec(colsB.get()), startB.get(), endB.get())
        m = compare_blocks(A, B, len(parse_cols_spec(colsA.get())))
        out = make_report_name(B.attrs["sheet_name"])
        write_report(m, out,
                     f"A: {fileA.get()} Blatt {sheetA.get()}",
                     f"B: {fileB.get()} Blatt {sheetB.get()}")
        messagebox.showinfo("Fertig", out)

    ttk.Button(frm, text="Start", command=run).grid()

    root.mainloop()


if __name__ == "__main__":
    main_gui()
