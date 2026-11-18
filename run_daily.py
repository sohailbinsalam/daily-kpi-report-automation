"""Daily KPI report: Excel file -> ETL -> PowerPoint -> email.

Examples
  python run_daily.py --sample 30        create 30 days of sample files, then run
  python run_daily.py                    process every file waiting in the inbox, report on the newest
  python run_daily.py --date 2026-11-18  process only that day's file
  python run_daily.py --no-email         build the report but skip the email step
"""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import date, datetime
from pathlib import Path

from kpi_report.config import load_config
from kpi_report.etl import ExtractError, FILE_RE, run_etl
from kpi_report.mailer import build_message, send
from kpi_report.report import build_report
from kpi_report.sample_data import generate

log = logging.getLogger("run_daily")


def setup_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"run_{datetime.now():%Y-%m-%d}.log"
    fmt = "%(asctime)s %(levelname)-7s %(name)s: %(message)s"
    logging.basicConfig(level=logging.INFO, format=fmt,
                        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)])
    return log_file


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=str(Path(__file__).with_name("config.toml")))
    ap.add_argument("--date", type=date.fromisoformat, help="process only this day (YYYY-MM-DD)")
    ap.add_argument("--sample", type=int, metavar="DAYS", help="first create this many days of sample files")
    ap.add_argument("--no-email", action="store_true", help="skip the email step")
    ap.add_argument("--keep-input", action="store_true", help="leave input files in the inbox (do not archive)")
    args = ap.parse_args(argv)

    cfg = load_config(args.config)
    log_file = setup_logging(cfg.paths["logs"])
    started = datetime.now()

    try:
        if args.sample:
            files = generate(cfg.paths["inbox"], days=args.sample)
            log.info("Created %d sample file(s) in %s", len(files), cfg.paths["inbox"])

        if args.date:
            days = [args.date]
        else:  # catch up on every waiting file, oldest first
            days = sorted(date.fromisoformat(FILE_RE.search(p.name).group(1))
                          for p in cfg.paths["inbox"].glob("daily_kpi_*.xlsx") if FILE_RE.search(p.name))
            if not days:
                raise ExtractError(f"No daily KPI file waiting in {cfg.paths['inbox']}")

        res = None
        for d in days:
            res = run_etl(cfg, d, keep_input=args.keep_input)
        log.info("ETL done for %d file(s); reporting on %s", len(days), res.day)

        pptx = build_report(res, cfg)
        log.info("Report saved: %s", pptx)

        if not args.no_email:
            eml = send(build_message(res, cfg, pptx), cfg, f"{res.day:%Y-%m-%d}")
            log.info("Email record: %s", eml)

        log.info("Finished in %.1f s. Log: %s", (datetime.now() - started).total_seconds(), log_file)
        return 0
    except ExtractError as e:
        log.warning("%s. Nothing to do.", e)
        return 2
    except Exception:
        log.exception("Run failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
