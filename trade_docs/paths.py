from __future__ import annotations

import os
import sys
from pathlib import Path


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def resource_dir() -> Path:
    bundled = getattr(sys, "_MEIPASS", None)
    return Path(bundled) if bundled else project_root()


def app_data_dir() -> Path:
    if getattr(sys, "frozen", False):
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or str(Path.home())
        path = Path(base) / "PIPLYiDianTeng"
        path.mkdir(parents=True, exist_ok=True)
        return path
    return project_root()


def frontend_dist_dir() -> Path:
    return resource_dir() / "pipl-frontend" / "dist"


def default_fabric_database() -> Path:
    env_path = os.environ.get("PIPL_FABRIC_DB")
    candidates = [
        Path(env_path) if env_path else None,
        Path("D:/codex application/合同自动生成工具/0428 fabric_prices.xlsx"),
        resource_dir() / "fabric_database_en.xlsx",
        resource_dir() / "fabric_master_en.csv",
    ]
    for path in candidates:
        if path and path.exists():
            return path
    return resource_dir() / "fabric_master_en.csv"
