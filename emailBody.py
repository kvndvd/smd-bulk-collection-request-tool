from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date
from pathlib import Path
import sys

from networkDrop import NETWORK_CNL_PATH, NETWORK_CT_PATH
from smdRequest import OutputPaths


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CONFIG_PATH = get_app_dir() / "email.config"


@dataclass
class EmailPayload:
    subject: str
    to: str
    cc: str
    html_body: str
    counsel_count: int
    court_count: int


def write_default_email_config() -> None:
    CONFIG_PATH.write_text(
        "TO=<insert send to emails>\nCC=<insert send to emails>\n",
        encoding="utf-8",
    )

def load_email_config() -> tuple[str, str]:
    if not CONFIG_PATH.exists():
        write_default_email_config()
    try:
        lines = CONFIG_PATH.read_text(encoding="utf-8").splitlines()
    except Exception:
        write_default_email_config()

    values: dict[str, str] = {}
    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip().upper()] = value.strip()

    to_email = values.get("TO", "").strip()
    cc_email = values.get("CC", "").strip()

    repaired = f"TO={to_email}\nCC={cc_email}\n"
    current = CONFIG_PATH.read_text(encoding="utf-8") if CONFIG_PATH.exists() else ""
    if current != repaired:
        CONFIG_PATH.write_text(repaired, encoding="utf-8")

    return to_email, cc_email


def _count_csv_data_rows(csv_path: Path) -> int:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        next(reader, None)
        return sum(1 for row in reader if any(str(cell).strip() for cell in row))


def build_email_payload(paths: OutputPaths) -> EmailPayload:
    today = date.today()
    counsel_count = _count_csv_data_rows(paths.cnl_path)
    court_count = _count_csv_data_rows(paths.ct_path)
    subject = f"Appeals SMD {today:%Y-%m-%d}"
    to_email, cc_email = load_email_config()

    html_body = f"""
<p>Hi Dena and Erin,</p>
<p>I have placed the updated Results File, Court, and Counsel requests in its appropriate folders.</p>

<div style="margin-top: 8px; margin-bottom: 8px;">
<table cellspacing="0" cellpadding="0" border="1" style="border-collapse: collapse; margin-left: 6.75pt;">
    <tr>
        <td style="border: 1.333px solid rgb(171, 171, 171); padding: 4px 7px; background-color: rgb(146, 208, 80); width: 61.6pt;">&nbsp;</td>
        <td style="border: 1.333px solid rgb(171, 171, 171); padding: 4px 7px; background-color: rgb(146, 208, 80); width: 77.95pt; text-align: center;"><b>NO. OF DOCS</b></td>
        <td style="border: 1.333px solid rgb(171, 171, 171); padding: 4px 7px; background-color: rgb(146, 208, 80); width: 715.15pt; text-align: center;"><b>LINK</b></td>
    </tr>
    <tr>
        <td style="border: 1.333px solid rgb(171, 171, 171); padding: 4px 7px; background-color: rgb(244, 176, 132); text-align: center;"><b>COURT</b></td>
        <td style="border: 1.333px solid rgb(171, 171, 171); padding: 4px 7px; text-align: center;"><b>{court_count}</b></td>
        <td style="border: 1.333px solid rgb(171, 171, 171); padding: 4px 7px; text-align: center;">
            <a href="file:///{NETWORK_CT_PATH.as_posix()}">{NETWORK_CT_PATH}</a>
        </td>
    </tr>
    <tr>
        <td style="border: 1.333px solid rgb(171, 171, 171); padding: 4px 7px; background-color: rgb(255, 229, 153); text-align: center;"><b>COUNSEL</b></td>
        <td style="border: 1.333px solid rgb(171, 171, 171); padding: 4px 7px; text-align: center;"><b>{counsel_count}</b></td>
        <td style="border: 1.333px solid rgb(171, 171, 171); padding: 4px 7px; text-align: center;">
            <a href="file:///{NETWORK_CNL_PATH.as_posix()}">{NETWORK_CNL_PATH}</a>
        </td>
    </tr>
</table>
</div>
"""

    return EmailPayload(
        subject=subject,
        to=to_email,
        cc=cc_email,
        html_body=html_body,
        counsel_count=counsel_count,
        court_count=court_count,
    )


def send_email_via_outlook(payload: EmailPayload, send_now: bool = True) -> None:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for Outlook email sending. "
            "Install it with: pip install pywin32"
        ) from exc

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = payload.to
        mail.CC = payload.cc
        mail.Subject = payload.subject

        # Load the user's Outlook signature first
        mail.Display()
        signature_html = mail.HTMLBody
        mail.HTMLBody = payload.html_body + signature_html

        if send_now:
            mail.Send()

    except Exception as exc:
        raise RuntimeError(
            "Unable to send email through Outlook automation.\n\n"
            "This usually happens when Outlook is unavailable for COM automation, "
            "or when the device is using the new Outlook app instead of classic Outlook.\n\n"
            "Recommended action:\n"
            "- Use classic Outlook for automatic sending, or\n"
            "- Turn off 'Send completion email' and send manually.\n\n"
            f"Details: {exc}"
        ) from exc