"""Emails the report. Every email is also saved as an .eml file in the outbox, so there is always a record."""
from __future__ import annotations

import logging
import os
import smtplib
import ssl
from email.message import EmailMessage
from html import escape
from pathlib import Path

from .config import Config
from .etl import EtlResult

log = logging.getLogger(__name__)


def build_message(res: EtlResult, cfg: Config, attachment: Path) -> EmailMessage:
    e = cfg.email
    on = sum(1 for r in res.status if r["on_target"])
    total = sum(1 for r in res.status if r["on_target"] is not None)

    msg = EmailMessage()
    msg["From"] = e["sender"]
    msg["To"] = ", ".join(e["recipients"])
    msg["Subject"] = e["subject"].format(date=f"{res.day:%d %b %Y}")

    lines = [f"Daily KPI report for {res.day:%A %d %B %Y}.", "",
             f"{on} of {total} KPIs on target. {len(res.breaches)} zone issue(s).", ""]
    for r in res.status:
        k = r["kpi"]
        flag = "" if r["on_target"] is None else ("OK " if r["on_target"] else "MISS ")
        lines.append(f"  {flag}{k.label}: {k.fmt(r['value'])} {k.unit} ({r['change_txt']} vs yesterday)")
    if res.breaches:
        lines += ["", "Top issues:"]
        lines += [f"  - {b['zone']}: {b['kpi'].label} {b['kpi'].fmt(b['value'])} {b['kpi'].unit}" for b in res.breaches[:5]]
    lines += ["", f"Input file checks: {res.validation.summary()}.", "", "The full report is attached.",
              "", "This email was generated automatically."]
    msg.set_content("\n".join(lines))

    rows = "".join(
        f"<tr><td style='padding:4px 12px 4px 0'>{escape(r['kpi'].label)}</td>"
        f"<td style='padding:4px 12px;font-weight:600;color:{_color(r)}'>{r['kpi'].fmt(r['value'])} {escape(r['kpi'].unit)}</td>"
        f"<td style='padding:4px 0;color:#667282'>{escape(r['change_txt'])} vs yesterday</td></tr>"
        for r in res.status)
    issues = "".join(f"<li>{escape(b['zone'])}: {escape(b['kpi'].label)} "
                     f"{b['kpi'].fmt(b['value'])} {escape(b['kpi'].unit)}</li>" for b in res.breaches[:5])
    html = f"""<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1b2430">
<p>Daily KPI report for <b>{res.day:%A %d %B %Y}</b>.</p>
<p><b>{on} of {total}</b> KPIs on target · <b>{len(res.breaches)}</b> zone issue(s).</p>
<table style="border-collapse:collapse">{rows}</table>
{'<p><b>Top issues</b></p><ul>' + issues + '</ul>' if issues else ''}
<p style="color:#667282">Input file checks: {escape(res.validation.summary())}.<br>The full report is attached. This email was generated automatically.</p>
</div>"""
    msg.add_alternative(html, subtype="html")

    msg.add_attachment(attachment.read_bytes(), maintype="application",
                       subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
                       filename=attachment.name)
    return msg


def _color(r) -> str:
    if r["on_target"] is None:
        return "#1b2430"
    return "#15803d" if r["on_target"] else "#c0263d"


def send(msg: EmailMessage, cfg: Config, day_tag: str, force_send: bool = False) -> Path:
    """Save to the outbox, then send through SMTP when mode = "smtp" (or force_send)."""
    outbox = cfg.paths["outbox"]
    outbox.mkdir(parents=True, exist_ok=True)
    eml = outbox / f"Daily_KPI_Report_{day_tag}.eml"
    eml.write_bytes(bytes(msg))

    e = cfg.email
    if e.get("mode") != "smtp" and not force_send:
        log.info("Dry run: email saved to %s (not sent)", eml)
        return eml

    password = os.environ.get(e.get("password_env", ""), "")
    if not password:
        raise RuntimeError(f"Set the {e.get('password_env')} environment variable to send email")
    with smtplib.SMTP(e["smtp_host"], int(e["smtp_port"]), timeout=60) as smtp:
        if e.get("use_starttls", True):
            smtp.starttls(context=ssl.create_default_context())
        smtp.login(e["username"], password)
        smtp.send_message(msg)
    log.info("Email sent to %s", ", ".join(e["recipients"]))
    return eml
