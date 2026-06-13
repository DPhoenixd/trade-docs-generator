from __future__ import annotations

import json
import urllib.request
from datetime import datetime, timezone


def fetch_usd_cny(timeout: float = 5.0) -> tuple[float | None, str]:
    url = "https://open.er-api.com/v6/latest/USD"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
        rate = payload.get("rates", {}).get("CNY")
        if rate:
            timestamp = payload.get("time_last_update_utc") or datetime.now(timezone.utc).isoformat()
            return float(rate), f"open.er-api.com, {timestamp}"
    except Exception as exc:  # noqa: BLE001 - UI should show a human-readable fallback reason.
        return None, f"自动获取失败：{exc}"
    return None, "自动获取失败：接口未返回 CNY"

