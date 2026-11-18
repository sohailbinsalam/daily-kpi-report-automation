"""Extract -> Validate -> Transform -> Load for the daily KPI Excel file."""
from __future__ import annotations

import logging
import re
import shutil
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd

from .config import Config, Kpi

log = logging.getLogger(__name__)

KEY_COLUMNS = ["Date", "Zone", "Unit_Count"]
FILE_RE = re.compile(r"daily_kpi_(\d{4}-\d{2}-\d{2})\.xlsx$")
HISTORY_FILE = "kpi_history.csv"


class ExtractError(RuntimeError):
    """No usable input file."""


@dataclass
class ValidationReport:
    rows_read: int = 0
    duplicates_removed: int = 0
    rejected: list[dict] = field(default_factory=list)   # {"row", "zone", "column", "value", "reason"}
    missing_values: list[dict] = field(default_factory=list)

    @property
    def rows_rejected(self) -> int:
        return len({r["row"] for r in self.rejected})

    @property
    def ok(self) -> bool:
        return not self.rejected and not self.missing_values and not self.duplicates_removed

    def summary(self) -> str:
        return (f"{self.rows_read} rows read, {self.duplicates_removed} duplicate(s) removed, "
                f"{len(self.missing_values)} blank value(s), {self.rows_rejected} row(s) rejected")


# --------------------------------------------------------------------------- extract

def find_input(inbox: Path, day: date | None = None) -> tuple[Path, date]:
    """The file for `day`, or the newest file in the inbox."""
    files = sorted((p for p in inbox.glob("daily_kpi_*.xlsx") if FILE_RE.search(p.name)),
                   key=lambda p: p.name)
    if day:
        files = [p for p in files if FILE_RE.search(p.name).group(1) == day.isoformat()]
    if not files:
        raise ExtractError(f"No daily KPI file found in {inbox}" + (f" for {day}" if day else ""))
    path = files[-1]
    return path, date.fromisoformat(FILE_RE.search(path.name).group(1))


def extract(path: Path) -> pd.DataFrame:
    log.info("Reading %s", path.name)
    return pd.read_excel(path, sheet_name="KPI", engine="openpyxl")


# --------------------------------------------------------------------------- validate

def validate(df: pd.DataFrame, kpis: list[Kpi]) -> tuple[pd.DataFrame, ValidationReport]:
    """Check columns, remove duplicates, blank-out impossible values and reject unusable rows."""
    rep = ValidationReport(rows_read=len(df))
    needed = KEY_COLUMNS + [k.column for k in kpis]
    missing_cols = [c for c in needed if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Input file is missing column(s): {', '.join(missing_cols)}")

    df = df[needed].copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.normalize()
    df["Zone"] = df["Zone"].astype(str).str.strip()

    before = len(df)
    df = df.drop_duplicates(subset=["Date", "Zone"], keep="first")
    rep.duplicates_removed = before - len(df)

    bad_rows: set[int] = set()
    for idx, row in df.iterrows():
        if pd.isna(row["Date"]) or not row["Zone"] or row["Zone"] == "nan":
            rep.rejected.append({"row": idx + 2, "zone": row["Zone"], "column": "Date/Zone",
                                 "value": "", "reason": "missing date or zone"})
            bad_rows.add(idx)
        if not (pd.notna(row["Unit_Count"]) and row["Unit_Count"] > 0):
            rep.rejected.append({"row": idx + 2, "zone": row["Zone"], "column": "Unit_Count",
                                 "value": row["Unit_Count"], "reason": "unit count must be above 0"})
            bad_rows.add(idx)

    for k in kpis:
        col = pd.to_numeric(df[k.column], errors="coerce")
        for idx in df.index[col.isna()]:
            rep.missing_values.append({"row": idx + 2, "zone": df.at[idx, "Zone"], "column": k.column})
        if k.unit == "%":
            impossible = (col < 0) | (col > 100)
        else:
            impossible = col < 0
        for idx in df.index[impossible.fillna(False)]:
            rep.rejected.append({"row": idx + 2, "zone": df.at[idx, "Zone"], "column": k.column,
                                 "value": col[idx], "reason": "outside possible range"})
            col[idx] = np.nan   # keep the row, drop only the bad value
        df[k.column] = col

    clean = df.drop(index=list(bad_rows))
    log.info("Validation: %s", rep.summary())
    return clean, rep


# --------------------------------------------------------------------------- transform

def overall_summary(df: pd.DataFrame, kpis: list[Kpi]) -> pd.DataFrame:
    """One row per date: weighted averages (by unit count) or sums across zones. NaNs are skipped."""
    out = []
    for day, g in df.groupby("Date"):
        row = {"Date": day}
        for k in kpis:
            vals = g[k.column]
            if k.agg == "sum":
                row[k.column] = vals.sum(min_count=1)
            else:
                m = vals.notna()
                w = g.loc[m, "Unit_Count"]
                row[k.column] = float(np.average(vals[m], weights=w)) if m.any() else np.nan
        out.append(row)
    return pd.DataFrame(out).sort_values("Date").reset_index(drop=True)


def kpi_status(today: pd.Series, yesterday: pd.Series | None, kpis: list[Kpi]) -> list[dict]:
    """Value, target, change vs yesterday and red/green state for each KPI."""
    rows = []
    for k in kpis:
        v = today[k.column]
        prev = yesterday[k.column] if yesterday is not None else np.nan
        delta = v - prev if pd.notna(prev) else np.nan
        if k.agg == "sum" and pd.notna(prev) and prev:
            change_txt = f"{delta / prev:+.1%}"
        elif pd.notna(delta):
            change_txt = f"{delta:+.{k.decimals}f}"
        else:
            change_txt = "n/a"
        rows.append({
            "kpi": k, "value": v, "delta": delta, "change_txt": change_txt,
            "on_target": k.meets_target(v),
            "improved": k.is_improvement(delta) if pd.notna(delta) else None,
        })
    return rows


def zone_breaches(df_day: pd.DataFrame, kpis: list[Kpi]) -> list[dict]:
    """Every zone/KPI pair that missed its target today, worst gap first."""
    out = []
    for k in (k for k in kpis if k.target is not None):
        for _, r in df_day.iterrows():
            v = r[k.column]
            if pd.notna(v) and k.meets_target(v) is False:
                gap = abs(v - k.target) / abs(k.target)
                out.append({"zone": r["Zone"], "kpi": k, "value": v, "gap": gap})
    return sorted(out, key=lambda x: x["gap"], reverse=True)


# --------------------------------------------------------------------------- load

def load(clean: pd.DataFrame, warehouse: Path) -> pd.DataFrame:
    """Upsert today's rows into the history file (re-running a day replaces it). Returns full history."""
    warehouse.mkdir(parents=True, exist_ok=True)
    hist_path = warehouse / HISTORY_FILE
    if hist_path.exists():
        hist = pd.read_csv(hist_path, parse_dates=["Date"])
        hist = hist[~hist["Date"].isin(clean["Date"].unique())]
        hist = pd.concat([hist, clean], ignore_index=True)
    else:
        hist = clean.copy()
    hist = hist.sort_values(["Date", "Zone"]).reset_index(drop=True)
    hist.to_csv(hist_path, index=False)
    log.info("History now holds %d rows (%s to %s)", len(hist),
             hist["Date"].min().date(), hist["Date"].max().date())
    return hist


def save_rejects(rep: ValidationReport, rejects_dir: Path, day: date) -> Path | None:
    issues = rep.rejected + [{**m, "value": "", "reason": "blank value"} for m in rep.missing_values]
    if not issues:
        return None
    rejects_dir.mkdir(parents=True, exist_ok=True)
    out = rejects_dir / f"issues_{day:%Y-%m-%d}.csv"
    pd.DataFrame(issues).to_csv(out, index=False)
    return out


def archive(path: Path, archive_dir: Path) -> Path:
    archive_dir.mkdir(parents=True, exist_ok=True)
    dest = archive_dir / path.name
    shutil.move(str(path), dest)
    return dest


# --------------------------------------------------------------------------- run

@dataclass
class EtlResult:
    day: date
    source: Path
    validation: ValidationReport
    history: pd.DataFrame
    overall: pd.DataFrame           # overall rows, one per date
    today_zones: pd.DataFrame
    status: list[dict]
    breaches: list[dict]
    rejects_file: Path | None


def run_etl(cfg: Config, day: date | None = None, keep_input: bool = False) -> EtlResult:
    path, file_day = find_input(cfg.paths["inbox"], day)
    raw = extract(path)
    clean, rep = validate(raw, cfg.kpis)
    hist = load(clean, cfg.paths["warehouse"])
    overall = overall_summary(hist, cfg.kpis)

    ts = pd.Timestamp(file_day)
    today = overall[overall["Date"] == ts]
    if today.empty:
        raise ExtractError(f"No valid rows left for {file_day} after validation")
    prev = overall[overall["Date"] < ts].tail(1)
    status = kpi_status(today.iloc[0], prev.iloc[0] if len(prev) else None, cfg.kpis)
    today_zones = hist[hist["Date"] == ts].reset_index(drop=True)
    breaches = zone_breaches(today_zones, cfg.kpis)
    rejects = save_rejects(rep, cfg.paths["rejects"], file_day)
    if not keep_input:
        archive(path, cfg.paths["archive"])
    return EtlResult(file_day, path, rep, hist, overall, today_zones, status, breaches, rejects)
