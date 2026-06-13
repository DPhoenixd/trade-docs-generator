from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd

from .models import FabricRecord


MASTER_COLUMNS = [
    "fabric_code",
    "quality_grade_cn",
    "quality_grade_en",
    "fabric_name_cn",
    "fabric_name_en",
    "composition_cn",
    "composition_en",
    "width_cn",
    "width_en",
    "weight_cn",
    "weight_en",
    "quantification_m_per_kg",
    "tube_plus_allowance_kg_per_roll",
    "reference_roll_weight_kg",
    "reference_roll_weight",
    "remarks_cn",
    "remarks_en",
]


def discover_fabric_database(base_dir: Path) -> Path | None:
    for name in [
        "fabric_database_en.json",
        "fabric_master_en.csv",
        "fabric_database_en.xlsx",
    ]:
        path = base_dir / name
        if path.exists():
            return path
    return None


def load_fabric_master(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if path.suffix.lower() == ".json":
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        if isinstance(data, dict):
            rows: Iterable[dict] = data.get("fabric_master_en") or data.get("rows") or data.values()
            if not isinstance(rows, list):
                rows = list(rows)
        else:
            rows = data
        df = pd.DataFrame(rows)
    elif path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    elif path.suffix.lower() in {".xlsx", ".xlsm"}:
        df = pd.read_excel(path, sheet_name="fabric_master_en", dtype=str)
    else:
        raise ValueError(f"Unsupported fabric database format: {path.suffix}")

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if "fabric_code" not in df.columns:
        raise ValueError("面料数据库缺少 fabric_code 字段")
    df["fabric_code"] = df["fabric_code"].astype(str).str.strip()
    for column in MASTER_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    return df


def load_price_rules(path: str | Path | None) -> pd.DataFrame:
    if path is None:
        return pd.DataFrame()
    path = Path(path)
    if not path.exists():
        return pd.DataFrame()
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, sheet_name="fabric_price_rules", dtype=str)
    return pd.DataFrame()


def _clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _to_float(value: object) -> float | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def fabric_record_from_row(row: pd.Series) -> FabricRecord:
    reference = row.get("reference_roll_weight_kg", "")
    if _clean_text(reference) == "":
        reference = row.get("reference_roll_weight", "")
    return FabricRecord(
        fabric_code=_clean_text(row.get("fabric_code", "")),
        quality_grade_cn=_clean_text(row.get("quality_grade_cn", "")),
        quality_grade_en=_clean_text(row.get("quality_grade_en", "")),
        fabric_name_cn=_clean_text(row.get("fabric_name_cn", "")),
        fabric_name_en=_clean_text(row.get("fabric_name_en", "")),
        composition_cn=_clean_text(row.get("composition_cn", "")),
        composition_en=_clean_text(row.get("composition_en", "")),
        width_cn=_clean_text(row.get("width_cn", "")),
        width_en=_clean_text(row.get("width_en", "")),
        weight_cn=_clean_text(row.get("weight_cn", "")),
        weight_en=_clean_text(row.get("weight_en", "")),
        quantification_m_per_kg=_to_float(row.get("quantification_m_per_kg", "")),
        tube_plus_allowance_kg_per_roll=_to_float(row.get("tube_plus_allowance_kg_per_roll", "")),
        reference_roll_weight_kg=_to_float(reference),
        remarks_cn=_clean_text(row.get("remarks_cn", "")),
        remarks_en=_clean_text(row.get("remarks_en", "")),
        raw={k: _clean_text(v) for k, v in row.to_dict().items()},
    )


def find_fabric(df: pd.DataFrame, fabric_code: str, quality_grade: str | None = None) -> FabricRecord | None:
    code = str(fabric_code).strip()
    matches = df[df["fabric_code"].astype(str).str.strip().str.casefold() == code.casefold()]
    if matches.empty:
        return None
    if quality_grade:
        grade = quality_grade.strip().casefold()
        grade_matches = matches[
            matches["quality_grade_en"].fillna("").astype(str).str.casefold().eq(grade)
            | matches["quality_grade_cn"].fillna("").astype(str).str.casefold().eq(grade)
        ]
        if not grade_matches.empty:
            matches = grade_matches
    return fabric_record_from_row(matches.iloc[0])


def save_fabric_record(path: str | Path, fabric: FabricRecord) -> tuple[str, Path]:
    path = Path(path)
    if path.suffix.lower() != ".csv":
        raise ValueError("当前只支持把面料反写入 CSV 数据库，请选择 fabric_master_en.csv")
    if not path.exists():
        raise FileNotFoundError(f"面料数据库不存在：{path}")

    df = load_fabric_master(path)
    for column in MASTER_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    row = _record_to_row(fabric, df.columns)
    code = str(row.get("fabric_code", "")).strip().casefold()
    grade_en = str(row.get("quality_grade_en", "")).strip().casefold()
    grade_cn = str(row.get("quality_grade_cn", "")).strip().casefold()
    mask = df["fabric_code"].astype(str).str.strip().str.casefold().eq(code)
    if grade_en or grade_cn:
        mask = mask & (
            df["quality_grade_en"].fillna("").astype(str).str.strip().str.casefold().eq(grade_en)
            | df["quality_grade_cn"].fillna("").astype(str).str.strip().str.casefold().eq(grade_cn)
        )

    backup_path = path.with_suffix(path.suffix + f".bak-{datetime.now().strftime('%Y%m%d-%H%M%S')}")
    shutil.copy2(path, backup_path)

    action = "updated" if mask.any() else "added"
    if mask.any():
        idx = df[mask].index[0]
        for column in df.columns:
            df.at[idx, column] = row.get(column, "")
    else:
        df = pd.concat([df, pd.DataFrame([row], columns=df.columns)], ignore_index=True)

    df.to_csv(path, index=False, encoding="utf-8-sig")
    return action, backup_path


def _record_to_row(fabric: FabricRecord, columns: Iterable[str]) -> dict[str, str]:
    row = {column: "" for column in columns}
    row.update({k: _clean_text(v) for k, v in fabric.raw.items() if k in row})
    row.update(
        {
            "fabric_code": fabric.fabric_code,
            "quality_grade_cn": fabric.quality_grade_cn,
            "quality_grade_en": fabric.quality_grade_en,
            "fabric_name_cn": fabric.fabric_name_cn,
            "fabric_name_en": fabric.fabric_name_en,
            "composition_cn": fabric.composition_cn,
            "composition_en": fabric.composition_en,
            "width_cn": fabric.width_cn,
            "width_en": fabric.width_en,
            "weight_cn": fabric.weight_cn,
            "weight_en": fabric.weight_en,
            "quantification_m_per_kg": _float_text(fabric.quantification_m_per_kg),
            "tube_plus_allowance_kg_per_roll": _float_text(fabric.tube_plus_allowance_kg_per_roll),
            "reference_roll_weight_kg": _float_text(fabric.reference_roll_weight_kg),
            "remarks_cn": fabric.remarks_cn,
            "remarks_en": fabric.remarks_en,
        }
    )
    return {k: _clean_text(v) for k, v in row.items()}


def _float_text(value: float | None) -> str:
    if value is None:
        return ""
    return f"{float(value):.6f}".rstrip("0").rstrip(".")
