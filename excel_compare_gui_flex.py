# automation_engine.py
from __future__ import annotations

import configparser
import fnmatch
import glob
import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

# ------------------------------------------------------------
# WICHTIG:
# Prüfkern NICHT verändern. Nur importieren und aufrufen.
# Passe den Import an dein Projekt an.
# ------------------------------------------------------------
# Beispiel:
# from core_compare import run_compare
# ODER:
# from excel_blockvergleich import run_compare

def run_compare(*args, **kwargs):
    """
    PLACEHOLDER.
    -> ERSETZEN durch: from <dein_modul> import run_compare
    Diese Funktion darf es am Ende NICHT geben.
    """
    raise RuntimeError("run_compare Import nicht gesetzt. Bitte in automation_engine.py anpassen.")


# -----------------------------
# Datenmodelle
# -----------------------------
@dataclass(frozen=True)
class ResolveResult:
    status: str  # EXACT | GLOB_HIT | MISSING | AMBIGUOUS
    path: Optional[str]
    hits: List[str]

@dataclass(frozen=True)
class Rule:
    name: str

    a_exact: str
    a_glob: str
    a_sheet: str
    a_key: str
    a_cols: str
    a_start: str
    a_end: str

    b_exact: str
    b_glob: str
    b_sheet: str
    b_key: str
    b_cols: str
    b_start: str
    b_end: str

@dataclass
class CheckResult:
    rule_name: str
    outcome: str  # OK | ABWEICHUNG | SKIP
    skip_reason: Optional[str]
    a_resolve: ResolveResult
    b_resolve: ResolveResult
    details: List[str]  # Detailzeilen (nur bei ABWEICHUNG, limitiert)


# -----------------------------
# Period parsing
# -----------------------------
def parse_period_INSO(subfolder_name: str) -> Optional[str]:
    """
    Erwartet Unterordnernamen wie:
      '1 Monat_12-2025'   -> 2025-12
      '3 Quartal_4-2025'  -> 2025-Q4
      '4 Halbjahr_2-2025' -> 2025-H2
      '5 Jahr_2025'       -> 2025
    """
    s = subfolder_name.strip()

    m = re.search(r"\b1\s*Monat_(\d{1,2})-(\d{4})\b", s)
    if m:
        month = int(m.group(1))
        year = int(m.group(2))
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}"
        return None

    m = re.search(r"\b3\s*Quartal_(\d{1})-(\d{4})\b", s)
    if m:
        q = int(m.group(1))
        year = int(m.group(2))
        if 1 <= q <= 4:
            return f"{year:04d}-Q{q}"
        return None

    m = re.search(r"\b4\s*Halbjahr_(\d{1})-(\d{4})\b", s)
    if m:
        h = int(m.group(1))
        year = int(m.group(2))
        if h in (1, 2):
            return f"{year:04d}-H{h}"
        return None

    m = re.search(r"\b5\s*Jahr_(\d{4})\b", s)
    if m:
        year = int(m.group(1))
        return f"{year:04d}"

    return None


# -----------------------------
# Resolver
# -----------------------------
def _glob_search(root: str, pattern: str, recursive: bool) -> List[str]:
    # pattern kann bereits ** enthalten; wir joinen sauber.
    # Wir erlauben auch rekursiv: **/<pattern>
    if recursive:
        # Wenn pattern keinen path enthält, suchen wir überall
        if os.sep not in pattern and "/" not in pattern and "\\" not in pattern:
            full_pattern = os.path.join(root, "**", pattern)
        else:
            full_pattern = os.path.join(root, pattern)
        hits = glob.glob(full_pattern, recursive=True)
    else:
        full_pattern = os.path.join(root, pattern)
        hits = glob.glob(full_pattern, recursive=False)

    # Nur Dateien, keine Ordner
    hits = [h for h in hits if os.path.isfile(h)]
    # Normalize
    hits = [os.path.normpath(h) for h in hits]
    return sorted(hits)


def resolve_file(root: str, exact_name: str, glob_pattern: str,
                 mode: str = "exact_then_glob",
                 policy: str = "unique_only",
                 recursive: bool = True) -> ResolveResult:
    """
    Returns ResolveResult with status: EXACT | GLOB_HIT | MISSING | AMBIGUOUS
    """
    exact_path = os.path.normpath(os.path.join(root, exact_name)) if exact_name else ""
    if mode == "exact_then_glob" and exact_name:
        if os.path.isfile(exact_path):
            return ResolveResult(status="EXACT", path=exact_path, hits=[exact_path])

    hits: List[str] = []
    if glob_pattern:
        hits = _glob_search(root, glob_pattern, recursive=recursive)

    if policy == "unique_only":
        if len(hits) == 1:
            return ResolveResult(status="GLOB_HIT", path=hits[0], hits=hits)
        if len(hits) == 0:
            return ResolveResult(status="MISSING", path=None, hits=[])
        return ResolveResult(status="AMBIGUOUS", path=None, hits=hits)

    # fallback (falls du später erweitern willst)
    if len(hits) >= 1:
        return ResolveResult(status="GLOB_HIT", path=hits[0], hits=hits)
    return ResolveResult(status="MISSING", path=None, hits=[])


# -----------------------------
# INI Parsing
# -----------------------------
def load_config(ini_path: str) -> configparser.ConfigParser:
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.optionxform = str  # keys case-sensitive lassen (optional)
    with open(ini_path, "r", encoding="utf-8") as f:
        cfg.read_file(f)
    return cfg


def list_automation_profiles(cfg: configparser.ConfigParser) -> List[str]:
    out = []
    for sec in cfg.sections():
        if sec.startswith("AUTOMATION:"):
            out.append(sec.split("AUTOMATION:", 1)[1])
    return sorted(out)


def _get_required(cfg: configparser.ConfigParser, sec: str, key: str) -> str:
    if not cfg.has_option(sec, key):
        raise KeyError(f"INI fehlt: [{sec}] {key}")
    return cfg.get(sec, key).strip()


def parse_rule_line(line: str) -> Rule:
    parts = [p.strip() for p in line.split("|")]
    if len(parts) != 15:
        raise ValueError(f"Regel hat {len(parts)} Felder, erwartet 15: {line}")

    return Rule(
        name=parts[0],

        a_exact=parts[1],
        a_glob=parts[2],
        a_sheet=parts[3],
        a_key=parts[4],
        a_cols=parts[5],
        a_start=parts[6],
        a_end=parts[7],

        b_exact=parts[8],
        b_glob=parts[9],
        b_sheet=parts[10],
        b_key=parts[11],
        b_cols=parts[12],
        b_start=parts[13],
        b_end=parts[14],
    )


def load_rulegroup(cfg: configparser.ConfigParser, group_name: str) -> List[Rule]:
    sec = f"RULEGROUP:{group_name}"
    if not cfg.has_section(sec):
        raise KeyError(f"INI fehlt Regelgruppe: [{sec}]")

    enabled = cfg.get(sec, "enabled", fallback="true").strip().lower() in ("1", "true", "yes", "on")
    if not enabled:
        return []

    rules: List[Tuple[str, str]] = []
    for k, v in cfg.items(sec):
        if k.startswith("rule."):
            rules.append((k, v))

    rules.sort(key=lambda kv: kv[0])  # rule.001, rule.002 ...
    return [parse_rule_line(v) for _, v in rules]


# -----------------------------
# Report Writer
# -----------------------------
class BatchReport:
    def __init__(self, out_path: str, detail_limit: int = 200) -> None:
        self.out_path = out_path
        self.detail_limit = detail_limit
        self.lines: List[str] = []
        self.count_ok = 0
        self.count_abw = 0
        self.count_skip = 0
        self.skip_reasons: Dict[str, int] = {}

    def add(self, s: str = "") -> None:
        self.lines.append(s)

    def add_skip(self, reason: str) -> None:
        self.count_skip += 1
        self.skip_reasons[reason] = self.skip_reasons.get(reason, 0) + 1

    def add_ok(self) -> None:
        self.count_ok += 1

    def add_abw(self) -> None:
        self.count_abw += 1

    def flush(self) -> None:
        os.makedirs(os.path.dirname(self.out_path), exist_ok=True)
        with open(self.out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(self.lines))


# -----------------------------
# Subfolder pairing
# -----------------------------
def _list_subfolders(path: str) -> List[str]:
    if not os.path.isdir(path):
        return []
    items = []
    for name in os.listdir(path):
        p = os.path.join(path, name)
        if os.path.isdir(p):
            items.append(name)
    return sorted(items)


def _prefix_of_subfolder(name: str) -> Optional[str]:
    m = re.match(r"^\s*([1-9])\s+", name)
    if m:
        return m.group(1)
    return None


def build_pairs(start_dir: str, folder_a: str, folder_b: str,
                subfolder_mode: str, allowed_prefixes: List[str]) -> List[Tuple[str, str, Optional[str]]]:
    """
    Returns list of (a_dir, b_dir, subfolder_name_or_None)
    """
    base_a = os.path.join(start_dir, folder_a)
    base_b = os.path.join(start_dir, folder_b)

    if subfolder_mode == "none":
        return [(base_a, base_b, None)]

    subs_a = _list_subfolders(base_a)
    subs_b = _list_subfolders(base_b)

    if subfolder_mode == "name_equal":
        common = sorted(set(subs_a).intersection(subs_b))
        return [(os.path.join(base_a, s), os.path.join(base_b, s), s) for s in common]

    if subfolder_mode == "prefix_map":
        # Map prefix -> list of folder names (normalerweise je prefix genau 1, aber wir bleiben defensiv)
        map_a: Dict[str, List[str]] = {}
        map_b: Dict[str, List[str]] = {}

        for s in subs_a:
            px = _prefix_of_subfolder(s)
            if px and px in allowed_prefixes:
                map_a.setdefault(px, []).append(s)

        for s in subs_b:
            px = _prefix_of_subfolder(s)
            if px and px in allowed_prefixes:
                map_b.setdefault(px, []).append(s)

        pairs: List[Tuple[str, str, Optional[str]]] = []
        for px in allowed_prefixes:
            la = map_a.get(px, [])
            lb = map_b.get(px, [])
            # Wenn je prefix mehrere -> wir paaren nach Namen, sonst SKIP später (Engine kann das als Missing/Ambiguous behandeln)
            # Hier: simplest pairing by exact same folder name intersection, otherwise pair firsts if both exist.
            inter = sorted(set(la).intersection(lb))
            if inter:
                for s in inter:
                    pairs.append((os.path.join(base_a, s), os.path.join(base_b, s), s))
            else:
                if la and lb:
                    pairs.append((os.path.join(base_a, la[0]), os.path.join(base_b, lb[0]), la[0]))
                # else: fehlt -> wird im Runner als MissingFolder/Skip gezählt, weil Pair gar nicht entsteht
        return pairs

    raise ValueError(f"Unbekannter subfolder_mode: {subfolder_mode}")


# -----------------------------
# Runner
# -----------------------------
def _apply_period(s: str, period: str) -> str:
    return (s or "").replace("{PERIOD}", period)


def run_automation(
    ini_path: str,
    profile_name: str,
    start_dir: str,
    report_out_dir: str,
) -> str:
    """
    Führt ein Automatik-Profil aus.
    Gibt den Pfad zum Sammelreport zurück.
    """
    cfg = load_config(ini_path)
    prof_sec = f"AUTOMATION:{profile_name}"
    if not cfg.has_section(prof_sec):
        raise KeyError(f"Profil nicht gefunden: [{prof_sec}]")

    folder_a = _get_required(cfg, prof_sec, "folder_a")
    folder_b = _get_required(cfg, prof_sec, "folder_b")

    subfolder_mode = cfg.get(prof_sec, "subfolder_mode", fallback="none").strip()
    prefixes_raw = cfg.get(prof_sec, "subfolder_prefixes", fallback="").strip()
    allowed_prefixes = [p.strip() for p in prefixes_raw.split(",") if p.strip()]

    period_mode = cfg.get(prof_sec, "period_mode", fallback="INSO").strip()

    resolve_mode = cfg.get(prof_sec, "resolve_mode", fallback="exact_then_glob").strip()
    resolve_policy = cfg.get(prof_sec, "resolve_policy", fallback="unique_only").strip()
    glob_recursive = cfg.get(prof_sec, "glob_recursive", fallback="true").strip().lower() in ("1", "true", "yes", "on")

    detail_limit = int(cfg.get(prof_sec, "detail_max_lines_per_check", fallback="200"))

    rulegroups_raw = _get_required(cfg, prof_sec, "rulegroups")
    rulegroups = [g.strip() for g in rulegroups_raw.split(",") if g.strip()]

    # Report file name
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    report_name = f"AUTOMATION_{profile_name}_{ts}.txt"
    report_path = os.path.join(report_out_dir, report_name)

    rep = BatchReport(report_path, detail_limit=detail_limit)

    rep.add("Excel Blockvergleich – Sammelreport (Automatik)")
    rep.add(f"Profil: {profile_name}")
    rep.add(f"Startordner: {os.path.normpath(start_dir)}")
    rep.add(f"Folder A: {folder_a}")
    rep.add(f"Folder B: {folder_b}")
    rep.add(f"Subfolder-Mode: {subfolder_mode}")
    rep.add(f"Resolver: {resolve_mode} / {resolve_policy} / recursive={glob_recursive}")
    rep.add(f"Rulegroups: {', '.join(rulegroups)}")
    rep.add("".ljust(80, "-"))

    # Build subfolder pairs
    pairs = build_pairs(start_dir, folder_a, folder_b, subfolder_mode, allowed_prefixes)

    if not pairs:
        rep.add("KEINE passenden Ordnerpaare gefunden -> alles SKIP.")
        rep.add_skip("NoPairsFound")
        rep.flush()
        return report_path

    # For each pair: detect period, run all rulegroups
    for (adir, bdir, subname) in pairs:
        rep.add("")
        rep.add("".ljust(80, "="))
        rep.add(f"Ordnerpaar:")
        rep.add(f"  A: {os.path.normpath(adir)}")
        rep.add(f"  B: {os.path.normpath(bdir)}")
        rep.add(f"  Subfolder: {subname or '(none)'}")

        if not os.path.isdir(adir):
            rep.add("  -> SKIP: Folder A fehlt")
            rep.add_skip("MissingFolderA")
            continue
        if not os.path.isdir(bdir):
            rep.add("  -> SKIP: Folder B fehlt")
            rep.add_skip("MissingFolderB")
            continue

        # Period
        period = ""
        if subname:
            if period_mode.upper() == "INSO":
                p = parse_period_INSO(subname)
                if not p:
                    rep.add("  -> SKIP: PERIOD konnte nicht erkannt werden")
                    rep.add_skip("PeriodParseFailed")
                    continue
                period = p
            else:
                rep.add(f"  -> SKIP: Unbekannter period_mode={period_mode}")
                rep.add_skip("UnknownPeriodMode")
                continue
        else:
            # no subfolder -> optional: period leer lassen
            period = ""

        rep.add(f"  PERIOD: {period or '(empty)'}")
        rep.add("".ljust(80, "-"))

        # Load rules once per pair
        for gname in rulegroups:
            rules = load_rulegroup(cfg, gname)
            if not rules:
                rep.add(f"[{gname}] (disabled/leer) -> übersprungen")
                continue

            rep.add(f"[{gname}] Regeln: {len(rules)}")

            for rule in rules:
                # Apply period placeholder
                a_exact = _apply_period(rule.a_exact, period)
                a_glob = _apply_period(rule.a_glob, period)
                b_exact = _apply_period(rule.b_exact, period)
                b_glob = _apply_period(rule.b_glob, period)

                a_sheet = _apply_period(rule.a_sheet, period)
                b_sheet = _apply_period(rule.b_sheet, period)

                rep.add("")
                rep.add(f"CHECK: {rule.name}")
                rep.add(f"  A: exact='{a_exact}' glob='{a_glob}' sheet='{a_sheet}'")
                rep.add(f"     key='{rule.a_key}' cols='{rule.a_cols}' start='{rule.a_start}' end='{rule.a_end}'")
                rep.add(f"  B: exact='{b_exact}' glob='{b_glob}' sheet='{b_sheet}'")
                rep.add(f"     key='{rule.b_key}' cols='{rule.b_cols}' start='{rule.b_start}' end='{rule.b_end}'")

                # Resolve A/B
                a_res = resolve_file(adir, a_exact, a_glob, mode=resolve_mode, policy=resolve_policy, recursive=glob_recursive)
                b_res = resolve_file(bdir, b_exact, b_glob, mode=resolve_mode, policy=resolve_policy, recursive=glob_recursive)

                rep.add(f"  RESOLVE A: {a_res.status}" + (f" -> {a_res.path}" if a_res.path else ""))
                if a_res.status == "AMBIGUOUS":
                    rep.add(f"           hits={len(a_res.hits)}")
                rep.add(f"  RESOLVE B: {b_res.status}" + (f" -> {b_res.path}" if b_res.path else ""))
                if b_res.status == "AMBIGUOUS":
                    rep.add(f"           hits={len(b_res.hits)}")

                # Policy handling: if missing/ambiguous => SKIP
                if a_res.status in ("MISSING", "AMBIGUOUS"):
                    reason = "MissingA" if a_res.status == "MISSING" else "AmbiguousA"
                    rep.add(f"  RESULT: SKIP ({reason})")
                    rep.add_skip(reason)
                    continue
                if b_res.status in ("MISSING", "AMBIGUOUS"):
                    reason = "MissingB" if b_res.status == "MISSING" else "AmbiguousB"
                    rep.add(f"  RESULT: SKIP ({reason})")
                    rep.add_skip(reason)
                    continue

                # Call core compare (Blackbox)
                # Erwartung: run_compare liefert mindestens:
                #   - status/outcome (OK/ABWEICHUNG/...) und optional Detailzeilen
                # Da wir deinen Prüfkern nicht kennen, kapseln wir defensiv.
                try:
                    core_result = run_compare(
                        file_a=a_res.path,
                        sheet_a=a_sheet,
                        key_a=rule.a_key,
                        cols_a=rule.a_cols,
                        start_a=rule.a_start,
                        end_a=rule.a_end,
                        file_b=b_res.path,
                        sheet_b=b_sheet,
                        key_b=rule.b_key,
                        cols_b=rule.b_cols,
                        start_b=rule.b_start,
                        end_b=rule.b_end,
                        header=None,  # du sagtest header-unabhängig; wenn dein Prüfkern das so bekommt
                        automation=True,  # optional, falls du im Kern ignorierst: ok
                    )
                except TypeError:
                    # Falls run_compare andere Signatur hat: du passt hier die Parameternamen an,
                    # ohne den Kern selbst zu ändern.
                    core_result = run_compare(
                        a_res.path, a_sheet, rule.a_key, rule.a_cols, rule.a_start, rule.a_end,
                        b_res.path, b_sheet, rule.b_key, rule.b_cols, rule.b_start, rule.b_end
                    )
                except Exception as e:
                    rep.add(f"  RESULT: SKIP (Exception)")
                    rep.add(f"  EXCEPTION: {type(e).__name__}: {e}")
                    rep.add_skip("Exception")
                    continue

                # Interpret core_result
                outcome, details = normalize_core_result(core_result)

                if outcome == "OK":
                    rep.add("  RESULT: OK")
                    rep.add_ok()
                elif outcome == "ABWEICHUNG":
                    rep.add("  RESULT: ABWEICHUNG")
                    rep.add_abw()
                    if details:
                        rep.add("  DETAILS:")
                        for line in details[:rep.detail_limit]:
                            rep.add(f"    {line}")
                        if len(details) > rep.detail_limit:
                            rep.add(f"    ... (gekürzt, {len(details)-rep.detail_limit} weitere Zeilen)")
                else:
                    # Alles andere behandeln wir als SKIP, um nicht zu "lügen"
                    rep.add(f"  RESULT: SKIP (UnknownOutcome:{outcome})")
                    rep.add_skip(f"UnknownOutcome:{outcome}")

    # Summary
    rep.add("")
    rep.add("".ljust(80, "="))
    rep.add("SUMMARY")
    rep.add(f"OK: {rep.count_ok}")
    rep.add(f"ABWEICHUNG: {rep.count_abw}")
    rep.add(f"SKIP: {rep.count_skip}")
    if rep.skip_reasons:
        rep.add("")
        rep.add("SKIP-Gründe:")
        for k in sorted(rep.skip_reasons.keys()):
            rep.add(f"  {k}: {rep.skip_reasons[k]}")

    rep.flush()
    return report_path


def normalize_core_result(core_result) -> Tuple[str, List[str]]:
    """
    Versucht, das Ergebnis deines Prüfkerns in (outcome, details[]) zu normalisieren.
    Du passt das ggf. minimal an, je nachdem was run_compare wirklich zurückgibt.
    """
    # Fall 1: dict
    if isinstance(core_result, dict):
        outcome = str(core_result.get("outcome") or core_result.get("status") or "").upper()
        details = core_result.get("details") or core_result.get("lines") or []
        if isinstance(details, str):
            details = [details]
        details = [str(x) for x in details]
        return map_outcome(outcome), details

    # Fall 2: tuple (outcome, details)
    if isinstance(core_result, tuple) and len(core_result) >= 1:
        outcome = str(core_result[0]).upper()
        details: List[str] = []
        if len(core_result) >= 2:
            d = core_result[1]
            if isinstance(d, list):
                details = [str(x) for x in d]
            elif isinstance(d, str):
                details = [d]
        return map_outcome(outcome), details

    # Fall 3: string
    if isinstance(core_result, str):
        return map_outcome(core_result.upper()), []

    # Unknown
    return "UNKNOWN", []


def map_outcome(s: str) -> str:
    s = (s or "").strip().upper()
    if s in ("OK",):
        return "OK"
    if s in ("ABWEICHUNG", "DIFF", "DIFFERENT", "MISMATCH"):
        return "ABWEICHUNG"
    if s in ("FEHLT_IN_A", "FEHLT_IN_B", "MISSING_IN_A", "MISSING_IN_B"):
        # Im Automatikmodus willst du das fachlich eher als ABWEICHUNG behandeln
        return "ABWEICHUNG"
    if s in ("SKIP",):
        return "SKIP"
    return s or "UNKNOWN"
