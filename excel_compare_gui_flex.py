import os
import sys
import glob
from datetime import datetime
import configparser

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

import fnmatch
import time
from dataclasses import dataclass
from typing import Optional

__version__ = "1.0.0"
__build_date__ = "2026-02-10"


# ================== PATH / LOG ==================
def get_base_dir() -> str:
    # In EXE: Ordner der EXE. Im Script: Ordner der .py
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def bootlog(msg: str):
    try:
        p = os.path.join(get_base_dir(), "startup.log")
        with open(p, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat()} {msg}\n")
    except Exception:
        pass


# ================== EXCEL COLUMN HELPERS ==================
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


# ================== SHEET RESOLUTION ==================
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


# ================== READ BLOCK (FIXED) ==================
def read_block(
    path: str,
    sheet_spec: str,
    key_col: str,
    compare_cols: list[str],
    row_start_excel: int,
    row_end_excel: int,
    key_fallback: str = "none",  # none | excel_row | block_row
) -> pd.DataFrame:
    """
    Fixes:
    1) Kein Spaltenversatz (keine _iloc-Spalte vor den Daten bei Spaltenauswahl).
    2) Kein Index-Mismatch beim concat (block_df wird reset_index(drop=True) + Excel-Zeilen separat).
    3) Optional: key_fallback=excel_row -> wenn Schlüsselzelle leer ist, wird ROW:<ExcelZeile> als Schlüssel gesetzt.
    """
    key_col = key_col.strip().upper()
    compare_cols = [c.strip().upper() for c in compare_cols]
    key_fallback = (key_fallback or "none").strip().lower()

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

    rs, re_ = int(row_start_excel), int(row_end_excel)
    if re_ < rs:
        rs, re_ = re_, rs

    start_i = excel_row_to_iloc(rs)
    end_i = excel_row_to_iloc(re_)

    if start_i < 0:
        start_i = 0
    if end_i >= len(df):
        end_i = len(df) - 1
    if start_i > end_i:
        raise ValueError(
            f"{os.path.basename(path)} ({sheet_name}): Zeilenbereich {row_start_excel}-{row_end_excel} "
            f"passt nicht zur Datei (zu wenig Zeilen)."
        )

    # Originalblock aus Excel
    block_df = df.iloc[start_i:end_i + 1].copy()

    # Excel-Zeilennummern robust (unabhängig vom Pandas-Index)
    excel_rows = list(range(rs, re_ + 1))

    # Wichtig: Index auf 0..n-1 setzen, damit concat sauber ausrichtet
    block_df = block_df.reset_index(drop=True)

    # Datenspalten exakt aus block_df ziehen (keine Verschiebung)
    data = block_df.iloc[:, idxs].copy()
    data.columns = needed

    # normalize
    data[key_col] = data[key_col].apply(normalize_value)
    for c in compare_cols:
        data[c] = data[c].apply(normalize_value)

    # Zusammenführen
    out = pd.concat([pd.Series(excel_rows, name="_excel_row"), data], axis=1)

    # key fallback (optional)
    fallback_rows_excel: list[int] = []
    fallback_rows_rel: list[int] = []
    # backward compatible aliases
    if key_fallback in ("row_number", "rownumber", "row"):
        key_fallback = "excel_row"

    if key_fallback in ("excel_row", "block_row"):
        mask = out[key_col] == ""
        if mask.any():
            fallback_rows_excel = out.loc[mask, "_excel_row"].astype(int).tolist()
            if key_fallback == "excel_row":
                out.loc[mask, key_col] = "ROW:" + out.loc[mask, "_excel_row"].astype(int).astype(str)
            else:
                # relative to the configured block start row (1..n within start..end)
                rel = (out.loc[mask, "_excel_row"].astype(int) - int(rs) + 1)
                fallback_rows_rel = rel.astype(int).tolist()
                out.loc[mask, key_col] = "R:" + rel.astype(int).astype(str)

    # drop empty key (only if fallback disabled)
    if key_fallback not in ("excel_row", "block_row"):

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
    out.attrs["key_fallback"] = key_fallback
    out.attrs["fallback_rows"] = fallback_rows_excel
    out.attrs["fallback_rows_excel"] = fallback_rows_excel
    out.attrs["fallback_rows_rel"] = fallback_rows_rel
    return out


# ================== COMPARE BLOCKS ==================
def compare_blocks(A: pd.DataFrame, B: pd.DataFrame, nvals: int) -> pd.DataFrame:
    # Merge on key+occurrence: "Zusammen#1", "Zusammen#2", ...
    m = A.merge(B, on="_key2", how="outer", suffixes=("_A", "_B"), indicator=True)

    def status(row):
        if row["_merge"] == "left_only":
            return "FEHLT_IN_B"
        if row["_merge"] == "right_only":
            return "FEHLT_IN_A"
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


# ================== REPORTING ==================
def get_default_output_dir() -> str:
    """Prefer app dir if writable; otherwise fall back to Documents\\ExcelBlockvergleich."""
    app_dir = get_base_dir()

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


def sanitize_filename(s: str, max_len: int = 80) -> str:
    s = (s or "").strip()
    bad = '\\/:*?"<>|'
    for ch in bad:
        s = s.replace(ch, "-")
    s = s.replace(" ", "-")
    while "--" in s:
        s = s.replace("--", "-")
    s = s.strip("-")
    if not s:
        s = "Sheet"
    return s[:max_len]


def make_report_filename(sheet_b_name: str, prefix: str = "Pruefprotokoll", ext: str = "txt") -> str:
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    tag = sanitize_filename(sheet_b_name)
    return f"{prefix}_{tag}_{ts}.{ext}"


def write_text_report(
    m: pd.DataFrame,
    out_txt_path: str,
    fileA: str, sheetA: str, keyA: str, colsA: list[str], rsA: int, reA: int,
    fileB: str, sheetB: str, keyB: str, colsB: list[str], rsB: int, reB: int,
    key_fallback_mode: str = "none",
    fallback_rows_a_excel: Optional[list[int]] = None,
    fallback_rows_b_excel: Optional[list[int]] = None,
    fallback_abw_a_excel: Optional[list[int]] = None,
    fallback_abw_b_excel: Optional[list[int]] = None,
):
    any_problem = (m["STATUS"] != "OK").any()

    lines = []
    lines.append("PRÜFPROTOKOLL")
    lines.append(f"Version: {__version__} | Build: {__build_date__} | Lauf: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Protokollpfad: {out_txt_path}")
    lines.append(f"A: {fileA} | Blatt: {sheetA} | Key: {keyA} | Spalten: {','.join(colsA)} | Zeilen: {rsA}-{reA}")
    lines.append(f"B: {fileB} | Blatt: {sheetB} | Key: {keyB} | Spalten: {','.join(colsB)} | Zeilen: {rsB}-{reB}")
    lines.append("")
    # Key-Fallback info (optional)
    mode = (key_fallback_mode or "none").strip().lower()
    if mode in ("row_number", "rownumber", "row"):
        mode = "excel_row"
    a_rows = fallback_rows_a_excel or []
    b_rows = fallback_rows_b_excel or []
    if mode != "none":
        lines.append(f"Key-Fallback: {mode}")
        lines.append(f"Fallback-Key-Zeilen A: {len(a_rows)}, B: {len(b_rows)}")
        if any_problem:
            a_abw = fallback_abw_a_excel or []
            b_abw = fallback_abw_b_excel or []
            if a_abw or b_abw:
                # limit list length to keep reports readable
                def _fmt(nums):
                    nums = sorted(set(int(x) for x in nums))
                    if len(nums) > 30:
                        return ", ".join(map(str, nums[:30])) + ", ..."
                    return ", ".join(map(str, nums))
                lines.append(f"Fallback-Zeilen mit ABW (Excel): A[{_fmt(a_abw)}] | B[{_fmt(b_abw)}]")
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


# ================== FILE HELPERS ==================
def pick_two_xlsx(folder: str) -> tuple[str, str]:
    files = sorted(glob.glob(os.path.join(folder, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).lower().startswith("pruefprotokoll")]
    if len(files) == 2:
        return files[0], files[1]
    return "", ""


def list_sheets(path: str) -> str:
    xl = pd.ExcelFile(path)
    names = xl.sheet_names
    out = [f"{os.path.basename(path)} (Blätter: {len(names)})"]
    for i, n in enumerate(names, start=1):
        out.append(f"  {i}: {n}")
    return "\n".join(out)


# ================== INI PRESETS ==================
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


# ================== COMPARE RUNNER ==================

# ================== AUTOMATION (INI / BATCH) ==================

def _split_list(s: str) -> list[str]:
    if not s:
        return []
    return [p.strip() for p in s.split(",") if p.strip()]

def list_automation_sections(cfg: configparser.ConfigParser) -> list[str]:
    out: list[str] = []
    for sec in cfg.sections():
        if sec.upper().startswith("AUTOMATION:"):
            enabled = cfg.get(sec, "enabled", fallback="1").strip().lower()
            if enabled in ("1", "true", "yes", "on", ""):
                out.append(sec)
    return out

def automation_display_name(cfg: configparser.ConfigParser, sec: str) -> str:
    return cfg.get(sec, "name", fallback=sec)

def list_rule_sections(cfg: configparser.ConfigParser) -> list[str]:
    return [s for s in cfg.sections() if s.upper().startswith("RULE:")]

def parse_rule_section_name(sec: str) -> tuple[str, str]:
    # "RULE:GROUP:ID"
    parts = sec.split(":", 2)
    if len(parts) != 3:
        raise ValueError(f"Ungültiger Rule-Section-Name: {sec}")
    return parts[1].strip(), parts[2].strip()

def group_token(group: str) -> str:
    g = group.upper()
    if "INTERN" in g:
        return "_INTERN"
    if g.endswith("_G") or "_G" in g:
        return "_g"
    return ""

def parse_prefixes(s: str) -> list[str]:
    parts = [p.strip() for p in (s or "").split(";")]
    return [p for p in parts if p]

def sanitize_period_folder_name(name: str) -> str:
    return (name or "").strip()

def detect_period_inso_style(subfolder: str) -> tuple[str, str]:
    """
    Returns (period_type, PERIOD_string)

    Beispiele:
      "1 Monat_12-2025" -> ("MONTH", "2025-12")
      "3 Quartal_4-2025" -> ("QUARTER", "2025-Q4")
      "4 Halbjahr_2-2025" -> ("HALFYEAR", "2025-H2")
      "5 Jahr_2025" -> ("YEAR", "2025")
    """
    s = sanitize_period_folder_name(subfolder)

    if s.startswith("1 Monat_"):
        tail = s[len("1 Monat_"):]  # "12-2025"
        mm, yyyy = tail.split("-", 1)
        mm_i = int(mm)
        yyyy_i = int(yyyy)
        return "MONTH", f"{yyyy_i:04d}-{mm_i:02d}"

    if s.startswith("3 Quartal_"):
        tail = s[len("3 Quartal_"):]  # "4-2025"
        q, yyyy = tail.split("-", 1)
        q_i = int(q)
        yyyy_i = int(yyyy)
        return "QUARTER", f"{yyyy_i:04d}-Q{q_i}"

    if s.startswith("4 Halbjahr_"):
        tail = s[len("4 Halbjahr_"):]  # "2-2025"
        h, yyyy = tail.split("-", 1)
        h_i = int(h)
        yyyy_i = int(yyyy)
        return "HALFYEAR", f"{yyyy_i:04d}-H{h_i}"

    if s.startswith("5 Jahr_"):
        tail = s[len("5 Jahr_"):]  # "2025"
        yyyy_i = int(tail)
        return "YEAR", f"{yyyy_i:04d}"

    raise ValueError(f"Unbekanntes INSO-Unterordnerformat: '{subfolder}'")

def template_vars_for_rule(template: str, period: str, token: str) -> dict[str, str]:
    # BASENAME: bis vor {PERIOD} (oder fallback: Template ohne Extension)
    t = template
    if "{PERIOD}" in t:
        base = t.split("{PERIOD}", 1)[0]
    else:
        base = os.path.splitext(t)[0]
    return {
        "TEMPLATE": t.replace("{PERIOD}", period),
        "BASENAME": base.replace("{PERIOD}", ""),
        "PERIOD": period,
        "TOKEN": token or "",
    }

@dataclass(frozen=True)
class ResolveResult:
    status: str  # EXACT | GLOB_HIT | MISSING | AMBIGUOUS
    path: Optional[str]
    hits: list[str]
    pattern: str

def resolve_file(
    path_dir: str,
    template_name: str,
    period: str,
    token: str,
    mode: str,
    policy: str,
    patterns: list[str],
) -> ResolveResult:
    """
    mode: exact_only | exact_then_glob
    policy: unique_only
    patterns: Pattern-Templates mit {TEMPLATE},{BASENAME},{PERIOD},{TOKEN}
    """
    vars_ = template_vars_for_rule(template_name, period, token)
    exact_name = vars_["TEMPLATE"]
    exact_path = os.path.join(path_dir, exact_name)

    if os.path.exists(exact_path):
        return ResolveResult("EXACT", exact_path, [], exact_name)

    if mode.strip().lower() != "exact_then_glob":
        return ResolveResult("MISSING", None, [], exact_name)

    if not os.path.isdir(path_dir):
        return ResolveResult("MISSING", None, [], exact_name)

    all_files = [f for f in os.listdir(path_dir) if os.path.isfile(os.path.join(path_dir, f))]

    for pat_t in patterns:
        pat = pat_t
        for k, v in vars_.items():
            pat = pat.replace("{" + k + "}", v)

        hits = sorted([f for f in all_files if fnmatch.fnmatch(f, pat)])
        if len(hits) == 0:
            continue
        if len(hits) == 1:
            return ResolveResult("GLOB_HIT", os.path.join(path_dir, hits[0]), hits, pat)

        # ambiguous
        return ResolveResult("AMBIGUOUS", None, hits, pat)

    return ResolveResult("MISSING", None, [], exact_name)

def make_automation_report_filename(profile_name: str) -> str:
    # YYYYMMDD_HHMMSS_... for better sorting
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    tag = sanitize_filename(profile_name)
    return f"{ts}_Pruefprotokoll_Automatik_{tag}.txt"


def read_report_lines(path: str) -> list[str]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read().splitlines()
    except Exception:
        return []

def extract_detail_from_single_report(lines: list[str], max_lines: int) -> list[str]:
    """
    Nimmt bei ABWEICHUNG die relevanten Blöcke aus dem Einzelreport.
    Bei OK wird [] zurückgegeben.

    Das ist bewusst simpel und robust: wenn dein Einzelreport-Format später
    leicht variiert, fällt es nicht sofort auseinander.
    """
    if not lines:
        return ["(Konnte Einzelprotokoll nicht lesen)"]

    # OK-Shortcut
    for ln in lines:
        if "Beide Datenbereiche sind identisch." in ln:
            return []

    # Ab erstem Problem-Block
    start_idx = None
    for i, ln in enumerate(lines):
        if ln.startswith("FEHLT_IN_A") or ln.startswith("FEHLT_IN_B") or ln.startswith("ABWEICHUNGEN"):
            start_idx = i
            break

    if start_idx is None:
        start_idx = 0

    detail = lines[start_idx:]
    if len(detail) > max_lines:
        detail = detail[:max_lines] + ["... (Detail gekürzt)"]
    return detail

def run_automation(cfg: configparser.ConfigParser, ini_path: str, automation_sec: str, start_root: str) -> str:
    if automation_sec not in cfg:
        raise ValueError(f"Automation-Profil nicht gefunden: {automation_sec}")

    asec = cfg[automation_sec]
    profile_name = asec.get("name", automation_sec)

    left_root = asec.get("left_root", "").strip()
    right_root = asec.get("right_root", "").strip()
    if not left_root or not right_root:
        raise ValueError("left_root/right_root fehlen im Automation-Profil.")

    subfolder_mode = asec.get("subfolder_mode", "prefix_match").strip().lower()
    subfolder_prefixes = parse_prefixes(asec.get("subfolder_prefixes", ""))

    period_mode = asec.get("period_mode", "inso_style").strip().lower()

    rule_groups = _split_list(asec.get("rule_groups", ""))
    if not rule_groups:
        raise ValueError("rule_groups ist leer im Automation-Profil.")

    report_max_detail_lines = int(asec.get("report_max_detail_lines", "200"))
    report_mode = asec.get("report_mode", "aggregate_only").strip().lower()  # aggregate_only supported
    if report_mode != "aggregate_only":
        raise ValueError(f"Unsupported report_mode: {report_mode}")

    # Report output: always in <Startordner>\
    out_path = os.path.join(start_root, make_automation_report_filename(profile_name))

    file_resolve_mode = asec.get("file_resolve_mode", "exact_only").strip().lower()
    glob_apply_to = asec.get("glob_apply_to", "both").strip().lower()
    glob_policy = asec.get("glob_policy", "unique_only").strip().lower()
    glob_patterns = [p.strip() for p in asec.get("glob_patterns", "").split(";") if p.strip()]
    if not glob_patterns:
        glob_patterns = ["{TEMPLATE}"]

    key_fallback_profile = asec.get("key_fallback", "none").strip().lower()  # none | excel_row | block_row

    # locate roots
    left_abs = os.path.join(start_root, left_root)
    right_abs = os.path.join(start_root, right_root)
    if not os.path.isdir(left_abs):
        raise ValueError(f"left_root nicht gefunden: {left_abs}")
    if not os.path.isdir(right_abs):
        raise ValueError(f"right_root nicht gefunden: {right_abs}")

    # ensure report can be created
    try:
        os.makedirs(start_root, exist_ok=True)
    except Exception:
        pass

    t0 = time.time()
    ok_cnt = 0
    abw_cnt = 0
    skip_cnt = 0
    skip_reasons = {"FOLDER_MISSING": 0, "FILE_MISSING": 0, "AMBIGUOUS": 0, "RULE_DISABLED": 0, "RULE_SKIPPED": 0, "ERROR": 0}

    def w(line=""):
        with open(out_path, "a", encoding="utf-8") as f:
            f.write(line + "\n")

    # header
    w("SAMMEL-PRÜFPROTOKOLL (AUTOMATIK)")
    w(f"Profil: {profile_name} ({automation_sec})")
    w(f"Version: {__version__} | Build: {__build_date__} | Lauf: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    w(f"Startordner: {start_root}")
    w(f"Links: {left_root} | Rechts: {right_root}")
    w(f"Rule-Gruppen: {', '.join(rule_groups)}")
    w(f"Key-Fallback: {key_fallback_profile}")
    w("")

    # subfolders
    left_subs: list[str] = []
    if subfolder_mode == "single_folder":
        left_subs = ["."]
    else:
        for name in sorted(os.listdir(left_abs)):
            p = os.path.join(left_abs, name)
            if os.path.isdir(p):
                left_subs.append(name)

    # filter by prefixes if needed
    if subfolder_mode == "prefix_match":
        left_subs = [s for s in left_subs if any(s.startswith(px) for px in subfolder_prefixes)]

    # scan rules
    all_rule_secs = list_rule_sections(cfg)

    # helper: get rules for groups
    def rules_for_groups(groups: list[str]) -> list[str]:
        out = []
        for rsec in all_rule_secs:
            g, _id = parse_rule_section_name(rsec)
            if g in groups:
                out.append(rsec)
        return out

    selected_rules = rules_for_groups(rule_groups)
    if not selected_rules:
        w("WARNUNG: Keine Rules gefunden für die angegebenen Rule-Gruppen.")
        w("")
        return out_path

    # run
    for sub in left_subs:
        # match right folder
        if subfolder_mode == "single_folder":
            left_dir = left_abs
            right_dir = right_abs
            sub_label = "(single_folder)"
        else:
            left_dir = os.path.join(left_abs, sub)
            right_dir = os.path.join(right_abs, sub)
            sub_label = sub

        if not os.path.isdir(right_dir):
            skip_cnt += 1
            skip_reasons["FOLDER_MISSING"] += 1
            w("=" * 70)
            w(f"SKIP Unterordner fehlt rechts | {sub_label}")
            w(f"Rechts erwartet: {right_dir}")
            continue

        # period detection
        if period_mode == "none":
            period_type, period_val = "", ""
        elif period_mode == "inso_style":
            try:
                period_type, period_val = detect_period_inso_style(sub if sub != "." else "")
            except Exception as e:
                skip_cnt += 1
                skip_reasons["ERROR"] += 1
                w("=" * 70)
                w(f"SKIP Zeitraum konnte nicht erkannt werden | {sub_label}")
                w(f"Fehler: {e}")
                continue
        else:
            raise ValueError(f"period_mode unbekannt: {period_mode}")

        # rules loop
        for rsec in selected_rules:
            r = cfg[rsec]
            group, rid = parse_rule_section_name(rsec)

            # Rule enabled?
            enabled = r.get("enabled", "1").strip().lower() in ("1", "true", "yes", "on", "")
            if not enabled:
                skip_cnt += 1
                skip_reasons["RULE_DISABLED"] += 1
                w("=" * 70)
                w(f"SKIP {rsec} | deaktiviert (enabled=0) | Ordner: {sub_label}")
                continue

            skip_auto = r.get("skip_in_automation", "0").strip().lower() in ("1", "true", "yes", "on")
            if skip_auto:
                skip_cnt += 1
                skip_reasons["RULE_SKIPPED"] += 1
                w("=" * 70)
                w(f"SKIP {rsec} | skip_in_automation=1 | Ordner: {sub_label}")
                continue

            r_period_type = r.get("period_type", "").strip().upper()
            if period_mode != "none" and r_period_type and r_period_type != period_type:
                continue

            token = r.get("token", "").strip() or group_token(group)
            key_fallback = r.get("key_fallback", key_fallback_profile).strip().lower()
            if key_fallback in ("", "default"):
                key_fallback = key_fallback_profile
            key_fallback = (key_fallback or "none").strip().lower()
            if key_fallback in ("row_number", "rownumber", "row"):
                key_fallback = "excel_row"
            if key_fallback not in ("none", "excel_row", "block_row"):
                key_fallback = "none"


            filea_t = r.get("filea", "").strip()
            fileb_t = r.get("fileb", "").strip()
            if not filea_t or not fileb_t:
                skip_cnt += 1
                skip_reasons["ERROR"] += 1
                w("=" * 70)
                w(f"SKIP Rule unvollständig | {rsec}")
                continue

            # resolver settings per side
            apply_a = (glob_apply_to in ("a_only", "both"))
            apply_b = (glob_apply_to in ("b_only", "both"))

            mode_a = file_resolve_mode if apply_a else "exact_only"
            mode_b = file_resolve_mode if apply_b else "exact_only"

            ra = resolve_file(left_dir, filea_t, period_val, token, mode_a, glob_policy, glob_patterns)
            rb = resolve_file(right_dir, fileb_t, period_val, token, mode_b, glob_policy, glob_patterns)

            if ra.status in ("MISSING", "AMBIGUOUS") or rb.status in ("MISSING", "AMBIGUOUS"):
                skip_cnt += 1
                if ra.status == "AMBIGUOUS" or rb.status == "AMBIGUOUS":
                    skip_reasons["AMBIGUOUS"] += 1
                else:
                    skip_reasons["FILE_MISSING"] += 1

                w("=" * 70)
                w(f"SKIP {rsec} | Zeitraum: {period_type} {period_val} | Ordner: {sub_label}")
                w(f"A resolve: {ra.status} | pattern: {ra.pattern}")
                if ra.hits:
                    w(f"A hits: {', '.join(ra.hits[:10])}" + (" ..." if len(ra.hits) > 10 else ""))
                w(f"B resolve: {rb.status} | pattern: {rb.pattern}")
                if rb.hits:
                    w(f"B hits: {', '.join(rb.hits[:10])}" + (" ..." if len(rb.hits) > 10 else ""))
                continue

            # gather compare params
            sheeta = r.get("sheeta", "1")
            sheetb = r.get("sheetb", "1")
            keya = r.get("keya", "A")
            keyb = r.get("keyb", "A")
            colsa = r.get("colsa", "B:K")
            colsb = r.get("colsb", "B:K")
            starta = r.get("starta", "1")
            enda = r.get("enda", "1")
            startb = r.get("startb", "1")
            endb = r.get("endb", "1")

            # run compare (existing core)
            single_report = ""
            try:
                single_report, meta = run_compare_with_meta(
                    ra.path, rb.path,
                    sheeta, keya, colsa, starta, enda,
                    sheetb, keyb, colsb, startb, endb,
                    key_fallback=key_fallback,
                )
                lines = read_report_lines(single_report)
                is_ok = any("Beide Datenbereiche sind identisch." in ln for ln in lines)

                w("=" * 70)
                w(f"RULE {rsec} | Zeitraum: {period_type} {period_val} | Ordner: {sub_label}")
                w(f"A: {os.path.basename(ra.path)} | Blatt: {sheeta} | resolve: {ra.status} | pattern: {ra.pattern}")
                w(f"B: {os.path.basename(rb.path)} | Blatt: {sheetb} | resolve: {rb.status} | pattern: {rb.pattern}")
                w(f"Konfig: A key={keya} cols={colsa} rows={starta}-{enda} | B key={keyb} cols={colsb} rows={startb}-{endb}")
                if meta.get("key_fallback", "none") in ("excel_row", "block_row"):
                    a_fb = meta.get("A_fallback_rows", []) or []
                    b_fb = meta.get("B_fallback_rows", []) or []
                    w(f"Fallback-Key-Zeilen A: {len(a_fb)}, B: {len(b_fb)}")

                if is_ok:
                    ok_cnt += 1
                    w("Ergebnis: OK")
                else:
                    abw_cnt += 1
                    w("Ergebnis: ABWEICHUNG")
                    if meta.get("key_fallback", "none") in ("excel_row", "block_row"):
                        a_abw = meta.get("A_fallback_abw_rows", []) or []
                        b_abw = meta.get("B_fallback_abw_rows", []) or []
                        if a_abw or b_abw:
                            def _fmt(nums):
                                nums = sorted(set(int(x) for x in nums))
                                if len(nums) > 30:
                                    return ", ".join(map(str, nums[:30])) + ", ..."
                                return ", ".join(map(str, nums))
                            w(f"Fallback-Zeilen mit ABW (Excel): A[{_fmt(a_abw)}] | B[{_fmt(b_abw)}]")
                    detail = extract_detail_from_single_report(lines, report_max_detail_lines)
                    for dl in detail:
                        w(dl)

            except Exception as e:
                skip_cnt += 1
                skip_reasons["ERROR"] += 1
                w("=" * 70)
                w(f"SKIP {rsec} | Fehler beim Vergleich | Ordner: {sub_label}")
                w(f"Fehler: {e}")
            finally:
                # aggregate_only: Einzelreport immer löschen (intern genutzt)
                if single_report:
                    try:
                        os.remove(single_report)
                    except Exception:
                        pass

    # footer summary
    dt = time.time() - t0
    w("")
    w("#" * 70)
    w("SUMMARY")
    w(f"OK: {ok_cnt}")
    w(f"ABWEICHUNG: {abw_cnt}")
    w(f"SKIP: {skip_cnt}")
    w(
        "SKIP Gründe: "
        f"folder_missing={skip_reasons['FOLDER_MISSING']}, "
        f"file_missing={skip_reasons['FILE_MISSING']}, "
        f"ambiguous={skip_reasons['AMBIGUOUS']}, "
        f"rule_disabled={skip_reasons['RULE_DISABLED']}, "
        f"rule_skipped={skip_reasons['RULE_SKIPPED']}, "
        f"error={skip_reasons['ERROR']}"
    )
    w(f"Laufzeit: {dt:.1f}s")
    w(f"INI: {ini_path}")
    return out_path

    for sub in left_subs:
        # Ordner-Matching rechts
        if subfolder_mode == "single_folder":
            left_dir = left_abs
            right_dir = right_abs
            sub_label = "(single_folder)"
        else:
            left_dir = os.path.join(left_abs, sub)
            right_dir = os.path.join(right_abs, sub)
            sub_label = sub

        if not os.path.isdir(right_dir):
            skip_cnt += 1
            skip_reasons["FOLDER_MISSING"] += 1
            w("=" * 60)
            w(f"SKIP Unterordner fehlt rechts | {sub_label}")
            w(f"Rechts erwartet: {right_dir}")
            continue

        # Zeitraum bestimmen
        if period_mode == "none":
            period_type, period_val = "", ""
        elif period_mode == "inso_style":
            try:
                period_type, period_val = detect_period_inso_style(sub if sub != "." else "")
            except Exception as e:
                skip_cnt += 1
                skip_reasons["ERROR"] += 1
                w("=" * 60)
                w(f"SKIP Zeitraum konnte nicht erkannt werden | {sub_label}")
                w(f"Fehler: {e}")
                continue
        else:
            raise ValueError(f"period_mode unbekannt: {period_mode}")

        # Rules abarbeiten
        for rsec in selected_rules:
            r = cfg[rsec]
            group, rid = parse_rule_section_name(rsec)

            r_period_type = r.get("period_type", "").strip().upper()
            if period_mode != "none" and r_period_type and r_period_type != period_type:
                continue

            token = r.get("token", "").strip() or group_token(group)

            filea_t = r.get("filea", "").strip()
            fileb_t = r.get("fileb", "").strip()
            if not filea_t or not fileb_t:
                skip_cnt += 1
                skip_reasons["ERROR"] += 1
                w("=" * 60)
                w(f"SKIP Rule unvollständig | {rsec}")
                continue

            apply_a = (glob_apply_to in ("a_only", "both"))
            apply_b = (glob_apply_to in ("b_only", "both"))
            mode_a = file_resolve_mode if apply_a else "exact_only"
            mode_b = file_resolve_mode if apply_b else "exact_only"

            ra = resolve_file(left_dir, filea_t, period_val, token, mode_a, glob_policy, glob_patterns)
            rb = resolve_file(right_dir, fileb_t, period_val, token, mode_b, glob_policy, glob_patterns)

            if ra.status in ("MISSING", "AMBIGUOUS") or rb.status in ("MISSING", "AMBIGUOUS"):
                skip_cnt += 1
                if ra.status == "AMBIGUOUS" or rb.status == "AMBIGUOUS":
                    skip_reasons["AMBIGUOUS"] += 1
                else:
                    skip_reasons["FILE_MISSING"] += 1

                w("=" * 60)
                w(f"SKIP {rsec} | Zeitraum: {period_type} {period_val} | Ordner: {sub_label}")
                w(f"A resolve: {ra.status} | pattern: {ra.pattern}")
                if ra.hits:
                    w(f"A hits: {', '.join(ra.hits[:10])}" + (" ..." if len(ra.hits) > 10 else ""))
                w(f"B resolve: {rb.status} | pattern: {rb.pattern}")
                if rb.hits:
                    w(f"B hits: {', '.join(rb.hits[:10])}" + (" ..." if len(rb.hits) > 10 else ""))
                continue

            # Vergleichsparameter
            sheeta = r.get("sheeta", "1")
            sheetb = r.get("sheetb", "1")
            keya = r.get("keya", "A")
            keyb = r.get("keyb", "A")
            colsa = r.get("colsa", "B:K")
            colsb = r.get("colsb", "B:K")
            starta = r.get("starta", "1")
            enda = r.get("enda", "1")
            startb = r.get("startb", "1")
            endb = r.get("endb", "1")

            try:
                # key-fallback: rule override > profile default
                kf = (r.get("key_fallback", "").strip().lower() or key_fallback_profile).strip().lower()
                if kf in ("row_number", "rownumber", "row"):
                    kf = "excel_row"

                single_report = run_compare(
                    ra.path, rb.path,
                    sheeta, keya, colsa, starta, enda,
                    sheetb, keyb, colsb, startb, endb,
                    key_fallback=kf,
                )
                lines = read_report_lines(single_report)
                is_ok = any("Beide Datenbereiche sind identisch." in ln for ln in lines)

                w("=" * 60)
                w(f"RULE {rsec} | Zeitraum: {period_type} {period_val} | Ordner: {sub_label}")
                w(f"A: {os.path.basename(ra.path)} | Blatt: {sheeta} | resolve: {ra.status} | pattern: {ra.pattern}")
                w(f"B: {os.path.basename(rb.path)} | Blatt: {sheetb} | resolve: {rb.status} | pattern: {rb.pattern}")
                w(f"Konfig: A key={keya} cols={colsa} rows={starta}-{enda} | B key={keyb} cols={colsb} rows={startb}-{endb}")

                # Key-Fallback lines from single report (for transparency)
                kb = [ln for ln in lines if ln.startswith("Key-Fallback:") or ln.startswith("Fallback-Key-Zeilen")]
                for ln in kb:
                    w(ln)

                if is_ok:
                    ok_cnt += 1
                    w("Ergebnis: OK")
                else:
                    abw_cnt += 1
                    w("Ergebnis: ABWEICHUNG")
                    detail = extract_detail_from_single_report(lines, report_max_detail_lines)
                    for dl in detail:
                        w(dl)

                if report_only_single:
                    try:
                        os.remove(single_report)
                    except Exception:
                        pass

            except Exception as e:
                skip_cnt += 1
                skip_reasons["ERROR"] += 1
                w("=" * 60)
                w(f"SKIP {rsec} | Fehler beim Vergleich | Ordner: {sub_label}")
                w(f"Fehler: {e}")

    dt = time.time() - t0
    w("")
    w("#" * 60)
    w("SUMMARY")
    w(f"OK: {ok_cnt}")
    w(f"ABWEICHUNG: {abw_cnt}")
    w(f"SKIP: {skip_cnt}")
    w(
        "SKIP Gründe: "
        f"folder_missing={skip_reasons['FOLDER_MISSING']}, "
        f"file_missing={skip_reasons['FILE_MISSING']}, "
        f"ambiguous={skip_reasons['AMBIGUOUS']}, "
        f"error={skip_reasons['ERROR']}"
    )
    w(f"Laufzeit: {dt:.1f}s")
    w(f"INI: {ini_path}")
    return out_path


def run_compare(
    fileA_path: str, fileB_path: str,
    sheetA_spec: str, keyA: str, colsA_spec: str, startA: str, endA: str,
    sheetB_spec: str, keyB: str, colsB_spec: str, startB: str, endB: str,
    key_fallback: str = "none",  # none | excel_row | block_row
):
    bootlog("START run_compare")

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

    A = read_block(fileA_path, sheetA_spec, keyA, colsA, rsA, reA, key_fallback=key_fallback)
    B = read_block(fileB_path, sheetB_spec, keyB, colsB, rsB, reB, key_fallback=key_fallback)

    nvals = len(colsA)
    m = compare_blocks(A, B, nvals=nvals)

    sheetA_name = A.attrs.get("sheet_name", sheetA_spec)
    sheetB_name = B.attrs.get("sheet_name", sheetB_spec)

    # key-fallback metadata (for reporting)
    mode = (key_fallback or "none").strip().lower()
    if mode in ("row_number", "rownumber", "row"):
        mode = "excel_row"
    if mode not in ("none", "excel_row", "block_row"):
        mode = "none"

    fb_a = A.attrs.get("fallback_rows_excel") or A.attrs.get("fallback_rows") or []
    fb_b = B.attrs.get("fallback_rows_excel") or B.attrs.get("fallback_rows") or []

    fb_abw_a: list[int] = []
    fb_abw_b: list[int] = []
    try:
        if "_key2" in m.columns:
            mask_fb = m["_key2"].astype(str).str.startswith(("ROW:", "R:"))
            mask_abw = m["STATUS"].astype(str) == "ABWEICHUNG"
            sel = m[mask_fb & mask_abw].copy()
            if "_excel_row_A" in sel.columns:
                fb_abw_a = [int(x) for x in sel["_excel_row_A"].dropna().tolist()]
            if "_excel_row_B" in sel.columns:
                fb_abw_b = [int(x) for x in sel["_excel_row_B"].dropna().tolist()]
    except Exception:
        pass

    out_txt = safe_write_path(make_report_filename(sheet_b_name=sheetB_name))

    write_text_report(
        m=m,
        out_txt_path=out_txt,
        fileA=os.path.basename(fileA_path), sheetA=sheetA_name, keyA=keyA, colsA=colsA, rsA=rsA, reA=reA,
        fileB=os.path.basename(fileB_path), sheetB=sheetB_name, keyB=keyB, colsB=colsB, rsB=rsB, reB=reB,
        key_fallback_mode=mode,
        fallback_rows_a_excel=fb_a,
        fallback_rows_b_excel=fb_b,
        fallback_abw_a_excel=fb_abw_a,
        fallback_abw_b_excel=fb_abw_b,
    )

    return out_txt

def run_compare_with_meta(
    fileA_path: str, fileB_path: str,
    sheetA_spec: str, keyA: str, colsA_spec: str, startA: str, endA: str,
    sheetB_spec: str, keyB: str, colsB_spec: str, startB: str, endB: str,
    key_fallback: str = "none",
) -> tuple[str, dict]:
    """
    Wie run_compare, aber liefert zusätzlich Meta-Infos zurück (u.a. Fallback-Key-Zeilen).
    """
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

    rsA, reA = int(startA), int(endA)
    rsB, reB = int(startB), int(endB)

    A = read_block(fileA_path, sheetA_spec, keyA, colsA, rsA, reA, key_fallback=key_fallback)
    B = read_block(fileB_path, sheetB_spec, keyB, colsB, rsB, reB, key_fallback=key_fallback)

    nvals = len(colsA)
    m = compare_blocks(A, B, nvals=nvals)

    sheetA_name = A.attrs.get("sheet_name", sheetA_spec)
    sheetB_name = B.attrs.get("sheet_name", sheetB_spec)

    mode = (key_fallback or "none").strip().lower()
    if mode in ("row_number", "rownumber", "row"):
        mode = "excel_row"
    if mode not in ("none", "excel_row", "block_row"):
        mode = "none"

    fb_a = A.attrs.get("fallback_rows_excel") or A.attrs.get("fallback_rows") or []
    fb_b = B.attrs.get("fallback_rows_excel") or B.attrs.get("fallback_rows") or []

    fb_abw_a: list[int] = []
    fb_abw_b: list[int] = []
    try:
        if "_key2" in m.columns:
            mask_fb = m["_key2"].astype(str).str.startswith(("ROW:", "R:"))
            mask_abw = m["STATUS"].astype(str) == "ABWEICHUNG"
            sel = m[mask_fb & mask_abw].copy()
            if "_excel_row_A" in sel.columns:
                fb_abw_a = [int(x) for x in sel["_excel_row_A"].dropna().tolist()]
            if "_excel_row_B" in sel.columns:
                fb_abw_b = [int(x) for x in sel["_excel_row_B"].dropna().tolist()]
    except Exception:
        pass

    out_txt = safe_write_path(make_report_filename(sheet_b_name=sheetB_name))

    write_text_report(
        m=m,
        out_txt_path=out_txt,
        fileA=os.path.basename(fileA_path), sheetA=sheetA_name, keyA=keyA, colsA=colsA, rsA=rsA, reA=reA,
        fileB=os.path.basename(fileB_path), sheetB=sheetB_name, keyB=keyB, colsB=colsB, rsB=rsB, reB=reB,
        key_fallback_mode=mode,
        fallback_rows_a_excel=fb_a,
        fallback_rows_b_excel=fb_b,
        fallback_abw_a_excel=fb_abw_a,
        fallback_abw_b_excel=fb_abw_b,
    )

    meta = {
        "key_fallback": mode,
        "A_fallback_rows": fb_a,
        "B_fallback_rows": fb_b,
        "A_fallback_abw_rows": fb_abw_a,
        "B_fallback_abw_rows": fb_abw_b,
    }
    return out_txt, meta


# ================== GUI ==================

def main_gui():
    bootlog("START main_gui")

    folder = get_base_dir()
    ini_path = os.path.join(folder, INI_NAME)
    cfg = load_ini(ini_path)

    bootlog("BEFORE tk.Tk")
    root = tk.Tk()
    root.title(f"Excel Blockvergleich (v{__version__})")

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

    status = tk.StringVar(value=f"Ausgabeordner: {get_default_output_dir()} | Presets: {INI_NAME}")
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


    def _rules_for_profile(automation_sec: str) -> list[str]:
        if automation_sec not in cfg:
            return []
        asec = cfg[automation_sec]
        groups = _split_list(asec.get("rule_groups", ""))
        if not groups:
            return []
        out = []
        for rsec in list_rule_sections(cfg):
            g, _id = parse_rule_section_name(rsec)
            if g in groups:
                out.append(rsec)
        return sorted(out)

    def open_rules_editor(automation_sec: str):
        # Rule-Übersicht: Gruppe | Regel | Aktiv | Auto-Skip
        rules = _rules_for_profile(automation_sec)
        if not rules:
            messagebox.showinfo("Regeln", "Keine Regeln für dieses Profil gefunden.")
            return

        win = tk.Toplevel(root)
        win.title("Regeln bearbeiten")
        win.transient(root)

        ttk.Label(win, text=f"Profil: {automation_display_name(cfg, automation_sec)}").grid(row=0, column=0, padx=10, pady=(10,4), sticky="w")

        cols = ("group", "rid", "enabled", "skip")
        tv = ttk.Treeview(win, columns=cols, show="headings", height=18)
        tv.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        tv.heading("group", text="Gruppe")
        tv.heading("rid", text="Regel")
        tv.heading("enabled", text="Regel aktiv")
        tv.heading("skip", text="In Automatik überspringen")

        tv.column("group", width=160, anchor="w")
        tv.column("rid", width=200, anchor="w")
        tv.column("enabled", width=110, anchor="center")
        tv.column("skip", width=170, anchor="center")

        def mark(b: bool) -> str:
            return "☑" if b else "☐"

        def bool_from_ini(v: str, default: bool) -> bool:
            if v is None:
                return default
            s = str(v).strip().lower()
            if s in ("1", "true", "yes", "on"):
                return True
            if s in ("0", "false", "no", "off"):
                return False
            return default

        # map iid -> rule section
        iid_to_rsec: dict[str, str] = {}
        for rsec in rules:
            g, rid = parse_rule_section_name(rsec)
            r = cfg[rsec]
            en = bool_from_ini(r.get("enabled", "1"), True)
            sk = bool_from_ini(r.get("skip_in_automation", "0"), False)
            iid = tv.insert("", "end", values=(g, rid, mark(en), mark(sk)))
            iid_to_rsec[iid] = rsec

        def toggle_cell(iid: str, colname: str):
            if iid not in iid_to_rsec:
                return
            rsec = iid_to_rsec[iid]
            r = cfg[rsec]
            vals = list(tv.item(iid, "values"))

            if colname == "enabled":
                cur = vals[2] == "☑"
                newv = not cur
                vals[2] = mark(newv)
                r["enabled"] = "1" if newv else "0"

            elif colname == "skip":
                cur = vals[3] == "☑"
                newv = not cur
                vals[3] = mark(newv)
                r["skip_in_automation"] = "1" if newv else "0"

            tv.item(iid, values=tuple(vals))
            save_ini(cfg, ini_path)

        def on_click(event):
            iid = tv.identify_row(event.y)
            col = tv.identify_column(event.x)  # '#1'..'#4'
            if not iid or not col:
                return
            if col == "#3":
                toggle_cell(iid, "enabled")
            elif col == "#4":
                toggle_cell(iid, "skip")

        tv.bind("<Button-1>", on_click)

        ttk.Label(
            win,
            text="Hinweis: Klick auf ☑/☐ toggelt. Änderungen werden sofort in die INI geschrieben.",
        ).grid(row=2, column=0, padx=10, pady=(0,10), sticky="w")

        win.columnconfigure(0, weight=1)
        win.rowconfigure(1, weight=1)

    def start_automation_dialog():
        start_root = filedialog.askdirectory(title="Startordner wählen (z.B. ...\\2025-12)")
        if not start_root:
            return

        autos = list_automation_sections(cfg)
        if not autos:
            messagebox.showerror("Automatik", "Keine [AUTOMATION:*] Profile in der INI gefunden.")
            return

        items = [(automation_display_name(cfg, s), s) for s in autos]
        names = [x[0] for x in items]

        win = tk.Toplevel(root)
        win.title("Automatik starten")
        win.transient(root)
        win.grab_set()

        ttk.Label(win, text="Profil wählen:").grid(row=0, column=0, padx=10, pady=10, sticky="w")

        sel = tk.StringVar(value=names[0])
        combo = ttk.Combobox(win, textvariable=sel, values=names, state="readonly", width=40)
        combo.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        def chosen_sec() -> str:
            chosen_name = sel.get()
            for nm, sec in items:
                if nm == chosen_name:
                    return sec
            return ""

        def open_rules():
            sec = chosen_sec()
            if not sec:
                messagebox.showerror("Regeln", "Profil nicht gewählt/gefunden.")
                return
            open_rules_editor(sec)

        def run_now():
            try:
                sec = chosen_sec()
                if not sec:
                    messagebox.showerror("Automatik", "Profil nicht gewählt/gefunden.")
                    return
                out_txt = run_automation(cfg, ini_path, sec, start_root)
                messagebox.showinfo("Fertig", f"Sammelprotokoll:\n{out_txt}")
                status.set(f"AUTOMATIK OK: {out_txt}")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Fehler", str(e))
                status.set(f"AUTOMATIK Fehler: {e}")

        ttk.Button(win, text="Regeln…", command=open_rules).grid(row=1, column=0, padx=10, pady=(0,10), sticky="w")
        ttk.Button(win, text="Start", command=run_now).grid(row=1, column=1, padx=10, pady=(0,10), sticky="e")
        ttk.Button(win, text="Abbrechen", command=win.destroy).grid(row=2, column=1, padx=10, pady=(0,10), sticky="e")
    ttk.Button(frm, text="Start Vergleich", command=on_start).grid(column=0, row=14, sticky="w")
    ttk.Button(frm, text="Automatik starten…", command=start_automation_dialog).grid(column=1, row=14, sticky="w", padx=(8,0))
    ttk.Button(frm, text="Beenden", command=root.destroy).grid(column=2, row=14, sticky="w", padx=(8,0))
    ttk.Button(frm, text="Über…", command=show_about).grid(column=3, row=14, sticky="w", padx=(8,0))

    root.mainloop()


if __name__ == "__main__":
    main_gui()

