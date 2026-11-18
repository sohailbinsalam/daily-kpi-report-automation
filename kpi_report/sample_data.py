"""Creates realistic *sample* daily KPI Excel files, the same shape an operations team exports every morning.

No real company data is used. A few data-quality problems are planted on purpose
(a duplicate row, a blank value, an impossible value) so the ETL validation has something to catch.
"""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

ZONES = {  # zone: number of units (used as the weight)
    "Zone A": 1450,
    "Zone B": 1180,
    "Zone C": 1320,
    "Zone D": 640,
    "Zone E": 760,
    "Zone F": 520,
}

# typical level, day-to-day noise
BASE = {
    "Uptime_pct": (99.78, 0.08),
    "Request_Success_pct": (99.05, 0.18),
    "Error_Rate_pct": (0.52, 0.07),
    "Login_Success_pct": (99.42, 0.12),
    "Transfer_Success_pct": (97.9, 0.35),
    "Response_Time_ms": (27.0, 2.2),
    "Speed_Mbps": (32.0, 2.8),
}


def file_name(day: date) -> str:
    return f"daily_kpi_{day:%Y-%m-%d}.xlsx"


def make_day(day: date, rng: np.random.Generator, stress: dict[str, float] | None = None) -> pd.DataFrame:
    """One day's rows, one per zone. `stress` pushes one zone below target for a realistic bad day."""
    weekend = day.weekday() in (4, 5)  # Fri/Sat weekend
    rows = []
    for zone, units in ZONES.items():
        row = {"Date": pd.Timestamp(day), "Zone": zone, "Unit_Count": units}
        for col, (mu, sd) in BASE.items():
            row[col] = rng.normal(mu, sd)
        load = units / 1000
        row["Data_Volume_TB"] = rng.normal(95, 6) * load * (1.12 if weekend else 1.0)
        row["Sessions_k"] = rng.normal(41, 3) * load * (0.9 if weekend else 1.0)
        rows.append(row)
    df = pd.DataFrame(rows)

    if stress:
        zone = stress.get("zone", "Zone B")
        m = df["Zone"] == zone
        df.loc[m, "Error_Rate_pct"] += 0.55
        df.loc[m, "Transfer_Success_pct"] -= 1.9
        df.loc[m, "Response_Time_ms"] += 14
        df.loc[m, "Speed_Mbps"] -= 9
        df.loc[m, "Uptime_pct"] -= 0.6

    pct = [c for c in df.columns if c.endswith("_pct")]
    df[pct] = df[pct].clip(upper=100)
    num = df.select_dtypes("number").columns.drop("Unit_Count")
    df[num] = df[num].round(3)
    return df


def generate(inbox: Path, days: int = 30, end: date | None = None, seed: int = 7) -> list[Path]:
    """Write `days` daily files ending on `end` (default: today) into `inbox`."""
    inbox.mkdir(parents=True, exist_ok=True)
    end = end or date.today()
    rng = np.random.default_rng(seed)
    written = []
    for i in range(days):
        day = end - timedelta(days=days - 1 - i)
        stress = {"zone": "Zone B"} if day == end else None
        df = make_day(day, rng, stress)

        if day == end:  # plant data-quality issues on the newest file
            df = pd.concat([df, df.iloc[[1]]], ignore_index=True)        # duplicate row
            df.loc[2, "Speed_Mbps"] = np.nan                         # blank value
            df.loc[3, "Request_Success_pct"] = 187.0                           # impossible value

        out = inbox / file_name(day)
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            df.to_excel(xw, sheet_name="KPI", index=False)
        written.append(out)
    return written
