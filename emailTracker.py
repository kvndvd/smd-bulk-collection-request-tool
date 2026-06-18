from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

from emailBody import EmailPayload
from smdRequest import OutputPaths


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


HISTORY_PATH = get_app_dir() / "email_history.json"


@dataclass
class EmailFingerprint:
    subject: str
    counsel_count: int
    court_count: int
    cnl_name: str
    ct_name: str
    ct_example_name: str
    xlsm_name: str

    def as_dict(self) -> dict[str, str | int]:
        return {
            "subject": self.subject,
            "counsel_count": self.counsel_count,
            "court_count": self.court_count,
            "cnl_name": self.cnl_name,
            "ct_name": self.ct_name,
            "ct_example_name": self.ct_example_name,
            "xlsm_name": self.xlsm_name,
        }


def build_email_fingerprint(payload: EmailPayload, paths: OutputPaths) -> EmailFingerprint:
    return EmailFingerprint(
        subject=payload.subject.strip(),
        counsel_count=payload.counsel_count,
        court_count=payload.court_count,
        cnl_name=paths.cnl_path.name,
        ct_name=paths.ct_path.name,
        ct_example_name=paths.ct_example_path.name if paths.ct_example_path is not None else "",
        xlsm_name=paths.xlsm_path.name,
    )


def _load_history() -> list[dict]:
    if not HISTORY_PATH.exists():
        return []
    try:
        return json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []


def _save_history(items: list[dict]) -> None:
    HISTORY_PATH.write_text(json.dumps(items, indent=2), encoding="utf-8")


def was_email_sent_locally(fingerprint: EmailFingerprint) -> bool:
    history = _load_history()
    target = fingerprint.as_dict()
    return any(item.get("fingerprint") == target for item in history)


def record_sent_email(fingerprint: EmailFingerprint) -> None:
    history = _load_history()
    history.append(
        {
            "sent_at": datetime.now().isoformat(timespec="seconds"),
            "fingerprint": fingerprint.as_dict(),
        }
    )
    _save_history(history)


def was_email_sent_in_outlook(
    fingerprint: EmailFingerprint,
    lookback_days: int = 7,
) -> bool:
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        sent_folder = namespace.GetDefaultFolder(5)  # olFolderSentMail = 5
        items = sent_folder.Items
        items.Sort("[SentOn]", True)

        cutoff = datetime.now() - timedelta(days=lookback_days)

        for item in items:
            try:
                subject = str(getattr(item, "Subject", "")).strip()
                sent_on = getattr(item, "SentOn", None)

                if not subject:
                    continue

                if sent_on is not None:
                    try:
                        py_sent_on = datetime(
                            sent_on.year,
                            sent_on.month,
                            sent_on.day,
                            sent_on.hour,
                            sent_on.minute,
                            sent_on.second,
                        )
                        if py_sent_on < cutoff:
                            break
                    except Exception:
                        pass

                if subject != fingerprint.subject:
                    continue

                body_text = str(getattr(item, "Body", "") or "")
                if str(fingerprint.counsel_count) in body_text and str(fingerprint.court_count) in body_text:
                    return True

            except Exception:
                continue

    except Exception:
        return False

    return False


def was_email_already_sent(fingerprint: EmailFingerprint) -> bool:
    return was_email_sent_locally(fingerprint) or was_email_sent_in_outlook(fingerprint)