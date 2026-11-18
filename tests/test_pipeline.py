from datetime import date
from pathlib import Path
import shutil

import numpy as np
import pandas as pd
import pytest

from kpi_report.config import Kpi, load_config
from kpi_report.etl import find_input, kpi_status, load, overall_summary, run_etl, validate
from kpi_report.report import build_report, nice_axis
from kpi_report.mailer import build_message
from kpi_report.sample_data import generate

ROOT = Path(__file__).resolve().parents[1]
KPIS = [
    Kpi("Success_pct", "Success", "%", "weighted", "higher", 98.5),
    Kpi("Errors_pct", "Errors", "%", "weighted", "lower", 0.8),
    Kpi("Volume_TB", "Volume", "TB", "sum", "higher"),
]


def frame(rows):
    return pd.DataFrame(rows, columns=["Date", "Zone", "Unit_Count", "Success_pct", "Errors_pct", "Volume_TB"])


def test_validate_removes_duplicates_and_blanks_impossible_values():
    df = frame([
        ["2026-11-01", "A", 100, 99.0, 0.5, 10],
        ["2026-11-01", "A", 100, 99.0, 0.5, 10],    # duplicate
        ["2026-11-01", "B", 300, 187.0, 0.6, 20],   # impossible %
        ["2026-11-01", "C", 50, np.nan, 0.4, 5],    # blank
    ])
    clean, rep = validate(df, KPIS)
    assert rep.duplicates_removed == 1
    assert len(clean) == 3
    assert pd.isna(clean.loc[clean.Zone == "B", "Success_pct"]).all()
    assert rep.rejected[0]["column"] == "Success_pct"
    assert len(rep.missing_values) == 1       # the real blank; the impossible value is counted as rejected
    assert not rep.ok


def test_validate_rejects_missing_columns():
    with pytest.raises(ValueError, match="missing column"):
        validate(pd.DataFrame({"Date": [], "Zone": []}), KPIS)


def test_validate_rejects_rows_without_units():
    df = frame([["2026-11-01", "A", 0, 99.0, 0.5, 10], ["2026-11-01", "B", 10, 99.0, 0.5, 10]])
    clean, rep = validate(df, KPIS)
    assert list(clean.Zone) == ["B"]
    assert rep.rows_rejected == 1


def test_overall_summary_weights_by_units_and_sums_volume():
    df = frame([["2026-11-01", "A", 100, 99.0, 1.0, 10], ["2026-11-01", "B", 300, 98.0, 0.2, 30]])
    clean, _ = validate(df, KPIS)
    net = overall_summary(clean, KPIS).iloc[0]
    assert net["Success_pct"] == pytest.approx(98.25)    # (99*100 + 98*300) / 400
    assert net["Errors_pct"] == pytest.approx(0.4)
    assert net["Volume_TB"] == 40


def test_status_direction_lower_is_better():
    today = pd.Series({"Success_pct": 98.0, "Errors_pct": 0.9, "Volume_TB": 50})
    prev = pd.Series({"Success_pct": 99.0, "Errors_pct": 1.0, "Volume_TB": 40})
    st = {r["kpi"].column: r for r in kpi_status(today, prev, KPIS)}
    assert st["Success_pct"]["on_target"] is False and st["Success_pct"]["improved"] is False
    assert st["Errors_pct"]["on_target"] is False and st["Errors_pct"]["improved"] is True   # 1.0 -> 0.9 is better
    assert st["Volume_TB"]["on_target"] is None and st["Volume_TB"]["change_txt"] == "+25.0%"


def test_load_is_idempotent(tmp_path):
    df = frame([["2026-11-01", "A", 100, 99.0, 0.5, 10]])
    clean, _ = validate(df, KPIS)
    load(clean, tmp_path)
    hist = load(clean, tmp_path)          # re-running the same day must not double the rows
    assert len(hist) == 1


def test_nice_axis_gives_clean_steps():
    lo, hi, step = nice_axis(99.46, 99.81)
    assert step in (0.05, 0.1)
    assert lo <= 99.46 and hi >= 99.81
    assert round(lo / step, 6).is_integer()


def test_end_to_end(tmp_path):
    shutil.copy(ROOT / "config.toml", tmp_path / "config.toml")
    cfg = load_config(tmp_path / "config.toml")
    generate(cfg.paths["inbox"], days=5, end=date(2026, 11, 18))
    for d in range(14, 19):
        res = run_etl(cfg, date(2026, 11, d))
    assert res.day == date(2026, 11, 18)
    assert res.validation.duplicates_removed == 1
    assert any(b["zone"] == "Zone B" for b in res.breaches)
    assert not list(cfg.paths["inbox"].glob("*.xlsx"))            # all archived
    pptx = build_report(res, cfg)
    assert pptx.exists() and pptx.stat().st_size > 20_000
    msg = build_message(res, cfg, pptx)
    assert "18 Nov 2026" in msg["Subject"]
    assert any(p.get_filename() == pptx.name for p in msg.iter_attachments())
    with pytest.raises(Exception):
        find_input(cfg.paths["inbox"])                              # inbox now empty
