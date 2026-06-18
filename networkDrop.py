from __future__ import annotations

import shutil
from pathlib import Path

from smdRequest import OutputPaths


NETWORK_CNL_PATH = Path(r"\\fabwebd5.net\pclprod\SF-Work\Trial_BPM_project\AUTOMATION_COURTLINK_MODULES\SMD\4_COUNSEL_ASSOCIATION\REQUESTS\APP")
NETWORK_CT_PATH = Path(r"\\fabwebd5.net\pclprod\SF-Work\Trial_BPM_project\AUTOMATION_COURTLINK_MODULES\SMD\3_COURT_COLLECTION\REQUESTS")
NETWORK_CT_EXAMPLE_PATH = Path(r"\\fabwebd5.net\pclprod\SF-Work\Trial_BPM_project\AUTOMATION_COURTLINK_MODULES\SMD\3_COURT_COLLECTION")


def ensure_network_folders() -> None:
    for label, path in (
        ("Counsel path", NETWORK_CNL_PATH),
        ("Court path", NETWORK_CT_PATH),
        ("Court example path", NETWORK_CT_EXAMPLE_PATH),
    ):
        try:
            path.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            raise RuntimeError(
                f"Unable to access {label}: {path}\n"
                f"Please check VPN connection, network access, and folder path.\n\n"
                f"Details: {exc}"
            ) from exc


def _copy_file(source: Path, destination_folder: Path, label: str) -> None:
    try:
        shutil.copy2(source, destination_folder / source.name)
    except Exception as exc:
        raise RuntimeError(
            f"Failed to copy {label} to network path:\n"
            f"Source: {source}\n"
            f"Destination: {destination_folder}\n\n"
            f"Please check VPN connection, path access, and file permissions.\n\n"
            f"Details: {exc}"
        ) from exc


def copy_outputs_to_network(paths: OutputPaths) -> None:
    ensure_network_folders()

    _copy_file(paths.cnl_path, NETWORK_CNL_PATH, "Counsel Request")
    _copy_file(paths.ct_path, NETWORK_CT_PATH, "Court Request")

    if paths.ct_example_path is not None:
        _copy_file(paths.ct_example_path, NETWORK_CT_EXAMPLE_PATH, "Court Request Example")