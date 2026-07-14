from __future__ import annotations

import re
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage

_XLSX_SUBTYPE = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class MailError(Exception):
    """Raised for any configuration or sending failure (message is user-safe)."""


@dataclass
class SmtpConfig:
    host: str = "smtp.gmail.com"
    port: int = 587
    user: str = ""
    password: str = ""
    use_tls: bool = True          # STARTTLS on a plain connection
    use_ssl: bool = False         # implicit TLS (e.g. port 465)
    from_addr: str = ""
    from_name: str = "RCL MIS Automation"

    @property
    def configured(self) -> bool:
        return bool(self.host and self.user and self.password)

    @classmethod
    def from_dict(cls, d: dict) -> "SmtpConfig":
        d = d or {}
        user = str(d.get("user", "") or "")
        return cls(
            host=str(d.get("host", "smtp.gmail.com") or "smtp.gmail.com"),
            port=int(d.get("port", 587) or 587),
            user=user,
            password=str(d.get("password", "") or ""),
            use_tls=bool(d.get("use_tls", True)),
            use_ssl=bool(d.get("use_ssl", False)),
            from_addr=str(d.get("from_addr", "") or user),
            from_name=str(d.get("from_name", "RCL MIS Automation")
                          or "RCL MIS Automation"),
        )


def parse_addrs(addrs) -> list[str]:
    """Split a comma/semicolon/space-separated string (or list) into addresses."""
    if not addrs:
        return []
    parts = re.split(r"[,;\s]+", addrs) if isinstance(addrs, str) else addrs
    return [a.strip() for a in parts if a and str(a).strip()]


_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def valid_addr(addr: str) -> bool:
    return bool(_EMAIL_RE.match(str(addr).strip()))


def send_email(cfg: SmtpConfig, to_addrs, subject: str, html_body: str,
               attachments=None, cc=None) -> list[str]:
    """Send an HTML email with optional .xlsx attachments.

    ``attachments``: iterable of (filename, bytes). Returns the list of To
    addresses on success; raises MailError on any problem.
    """
    if not cfg.configured:
        raise MailError(
            "Email is not configured yet. Add an [smtp] section to "
            ".streamlit/secrets.toml (see .streamlit/secrets.toml.example).")

    to_list = parse_addrs(to_addrs)
    cc_list = parse_addrs(cc)
    if not to_list:
        raise MailError("Please enter at least one recipient address.")
    bad = [a for a in to_list + cc_list if not valid_addr(a)]
    if bad:
        raise MailError("These addresses look invalid: " + ", ".join(bad))

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = (f"{cfg.from_name} <{cfg.from_addr}>"
                   if cfg.from_name else cfg.from_addr)
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content("This report is best viewed in an HTML-capable client.")
    msg.add_alternative(html_body, subtype="html")

    for name, data in (attachments or []):
        msg.add_attachment(data, maintype="application",
                           subtype=_XLSX_SUBTYPE, filename=name)

    try:
        if cfg.use_ssl:
            server = smtplib.SMTP_SSL(cfg.host, cfg.port, timeout=30)
        else:
            server = smtplib.SMTP(cfg.host, cfg.port, timeout=30)
        with server:
            server.ehlo()
            if cfg.use_tls and not cfg.use_ssl:
                server.starttls()
                server.ehlo()
            server.login(cfg.user, cfg.password)
            server.send_message(msg, from_addr=cfg.from_addr,
                                to_addrs=to_list + cc_list)
    except smtplib.SMTPAuthenticationError as exc:
        raise MailError(
            "SMTP login failed — check the user and app password. For Gmail "
            "use a 16-character App Password, not your normal password.") from exc
    except (smtplib.SMTPException, OSError) as exc:
        raise MailError(f"Could not send the email: {exc}") from exc
    return to_list


def report_email_html(heading: str, intro: str, items: list[str],
                      date_str: str = "") -> str:
    """A simple, self-contained HTML body for a report email."""
    rows = "".join(f"<li style='margin:2px 0'>{i}</li>" for i in items)
    date_line = (f"<p style='margin:0 0 8px'><b>Date:</b> {date_str}</p>"
                 if date_str else "")
    return f"""\
<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a">
  <h2 style="color:#8a6d1a;margin:0 0 6px">{heading}</h2>
  {date_line}
  <p style="margin:0 0 8px">{intro}</p>
  <p style="margin:0 0 4px"><b>Sheets included:</b></p>
  <ul style="margin:0 0 12px 18px;padding:0">{rows}</ul>
  <hr style="border:none;border-top:1px solid #ddd;margin:12px 0">
  <p style="font-size:12px;color:#888;margin:0">
    Sent automatically by RCL MIS Automation.
  </p>
</div>"""
