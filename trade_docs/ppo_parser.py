from __future__ import annotations

import re
from pathlib import Path

import pandas as pd


PPO_LINE_COLUMNS = [
    "use",
    "gmt_color_code",
    "fabric_combo_name",
    "source_fabric_code",
    "fabric_total_yards",
    "ppo_pur_qty_yards",
    "ppo_qty_yards",
    "source_component",
]


def empty_ppo_lines() -> pd.DataFrame:
    return pd.DataFrame(columns=PPO_LINE_COLUMNS)


def parse_buyer_information(text: str) -> dict[str, str]:
    lines = [_clean_line(line) for line in str(text or "").splitlines()]
    lines = [line for line in lines if line]
    if not lines:
        return {"buyer": "", "buyer_address": ""}
    return {"buyer": lines[0], "buyer_address": "\n".join(lines[1:])}


def parse_ppo_pdf(path: str | Path) -> tuple[dict[str, str], pd.DataFrame]:
    path = Path(path)
    try:
        import pypdf
    except ImportError as exc:  # pragma: no cover - shown in Streamlit UI
        raise RuntimeError("缺少 pypdf，无法解析 PPO PDF。请先安装 requirements.txt。") from exc

    reader = pypdf.PdfReader(path)
    pages = [page.extract_text() or "" for page in reader.pages]
    text = "\n".join(pages)
    summary = {
        "file_name": path.name,
        "ppo_no": _match_value(text, r"PPO#:\s*\n?([A-Z0-9-]+)"),
        "related_go_no": _match_value(text, r"Related GO#:\s*\n?([A-Z0-9-]+)"),
        "supplier": _field_after_label(text, "Supplier"),
        "customer": _field_after_label(text, "Customer"),
        "brand": _field_after_label(text, "Brand"),
        "style_no": _field_after_label(text, "Style No"),
        "season": _field_after_label(text, "Season"),
        "quality_code": _field_after_label(text, "Quality Code:"),
        "width": _field_after_label(text, "Width:"),
        "ship_dest": _field_after_label(text, "Ship \nDest"),
        "ship_mode": _field_after_label(text, "Ship Mode"),
    }
    buyer_info = _extract_buyer_block(text)
    summary.update(parse_buyer_information(buyer_info))
    summary["buyer_info_raw"] = buyer_info
    lines = _extract_ppo_lines(pages)
    return summary, lines


def _extract_ppo_lines(pages: list[str]) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for page in pages:
        lines = [_clean_line(line) for line in page.splitlines()]
        summary_rows = _extract_fabric_total_rows(lines)
        if summary_rows:
            rows.extend(summary_rows)
            continue
        for index, line in enumerate(lines):
            if "@" not in line or "PPO#:" in line:
                continue
            candidate = line
            offset = 1
            while ")" not in candidate and index + offset < len(lines) and offset <= 2:
                candidate += " " + lines[index + offset]
                offset += 1
            parsed = _parse_color_line(candidate)
            if not parsed:
                continue
            numbers = _next_numbers(lines, index + offset, limit=3)
            if len(numbers) < 3:
                continue
            component = _nearest_component(lines, index)
            rows.append(
                {
                    "use": True,
                    "gmt_color_code": parsed["gmt_color_code"],
                    "fabric_combo_name": parsed["fabric_combo_name"],
                    "source_fabric_code": parsed["source_fabric_code"],
                    "fabric_total_yards": numbers[0],
                    "ppo_pur_qty_yards": numbers[1],
                    "ppo_qty_yards": numbers[2],
                    "source_component": component,
                }
            )
    if not rows:
        return empty_ppo_lines()
    return pd.DataFrame(rows, columns=PPO_LINE_COLUMNS)


def _extract_fabric_total_rows(lines: list[str]) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for start, line in enumerate(lines):
        if "Gmt Color Code @ Fabric Combo" not in line:
            continue
        header = lines[start : start + 16]
        if not _is_fabric_total_header(header):
            continue
        end = _next_header_index(lines, start + 1)
        block_end = end if end is not None else min(len(lines), start + 80)
        index = start + 1
        while index < block_end:
            candidate = lines[index]
            if "@" not in candidate or "PPO#:" in candidate:
                index += 1
                continue
            offset = 1
            while ")" not in candidate and index + offset < block_end and offset <= 2:
                candidate += " " + lines[index + offset]
                offset += 1
            parsed = _parse_color_line(candidate)
            if not parsed:
                index += 1
                continue
            numbers = _next_numbers(lines, index + offset, limit=2)
            if len(numbers) < 2:
                index += offset
                continue
            rows.append(
                {
                    "use": True,
                    "gmt_color_code": parsed["gmt_color_code"],
                    "fabric_combo_name": parsed["fabric_combo_name"],
                    "source_fabric_code": parsed["source_fabric_code"],
                    "fabric_total_yards": numbers[1],
                    "ppo_pur_qty_yards": numbers[0],
                    "ppo_qty_yards": None,
                    "source_component": _nearest_component(lines, start),
                }
            )
            index += offset
    return rows


def _is_fabric_total_header(lines: list[str]) -> bool:
    text = " ".join(lines).casefold()
    return "fabric total" in text or ("fabric" in text and "total(yds)" in text)


def _next_header_index(lines: list[str], start: int) -> int | None:
    for index in range(start, len(lines)):
        if "Gmt Color Code @ Fabric Combo" in lines[index]:
            return index
    return None


def _parse_color_line(line: str) -> dict[str, str] | None:
    text = re.sub(r"\s+", " ", line).strip()
    match = re.match(r"^(.+?)@(.+?)\s*\(([A-Z0-9_/-]+)\)$", text)
    if not match:
        return None
    return {
        "gmt_color_code": match.group(1).strip(),
        "fabric_combo_name": match.group(2).strip(),
        "source_fabric_code": match.group(3).strip(),
    }


def _next_numbers(lines: list[str], start: int, limit: int) -> list[float]:
    numbers: list[float] = []
    for line in lines[start : start + 12]:
        value = _to_float(line)
        if value is None:
            continue
        numbers.append(value)
        if len(numbers) >= limit:
            break
    return numbers


def _nearest_component(lines: list[str], index: int) -> str:
    for line in reversed(lines[max(0, index - 35) : index]):
        upper = line.upper()
        if "MAIN BODY" in upper or "TRIM FAB" in upper or "FAB" in upper and " - " in line:
            return line
    return ""


def _extract_buyer_block(text: str) -> str:
    match = re.search(r"Wash Remark\s*\n(.+?)\nPrinting Type", text, flags=re.S)
    if match:
        return _clean_multiline(match.group(1))
    match = re.search(r"FABRIC NEED TO BE SHIP TO.+?\n(.+?)\nfabrication:", text, flags=re.S | re.I)
    if match:
        return _clean_multiline(match.group(1))
    return ""


def _field_after_label(text: str, label: str) -> str:
    lines = text.splitlines()
    normalized_label = _clean_line(label).casefold()
    for index, line in enumerate(lines):
        if _clean_line(line).casefold() != normalized_label:
            continue
        collected: list[str] = []
        for candidate in lines[index + 1 : index + 4]:
            value = _clean_line(candidate)
            if value:
                collected.append(value)
            if collected:
                break
        return " ".join(collected)
    return ""


def _match_value(text: str, pattern: str) -> str:
    match = re.search(pattern, text)
    return _clean_line(match.group(1)) if match else ""


def _clean_multiline(text: str) -> str:
    return "\n".join(line for line in (_clean_line(part) for part in text.splitlines()) if line)


def _clean_line(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _to_float(value: object) -> float | None:
    try:
        text = str(value).replace(",", "").strip()
        if not re.fullmatch(r"\d+(?:\.\d+)?", text):
            return None
        return float(text)
    except (TypeError, ValueError):
        return None
