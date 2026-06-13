from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path


SEQUENCE_PATH = Path("order_sequences.json")


def suggest_doc_number(prefix: str = "ES", today: date | None = None, path: Path = SEQUENCE_PATH) -> str:
    today = today or date.today()
    key = today.strftime("%Y%m%d")
    data = _load(path)
    next_no = int(data.get(key, 0)) + 1
    return f"{prefix}{key}{next_no:02d}"


def commit_doc_number(doc_no: str, prefix: str = "ES", path: Path = SEQUENCE_PATH) -> None:
    match = re.fullmatch(rf"{re.escape(prefix)}(\d{{8}})(\d{{2,}})", doc_no.strip())
    if not match:
        return
    key, seq = match.groups()
    data = _load(path)
    data[key] = max(int(data.get(key, 0)), int(seq))
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def _load(path: Path) -> dict[str, int]:
    if not path.exists():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    return {str(k): int(v) for k, v in raw.items()}

