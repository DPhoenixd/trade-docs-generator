from __future__ import annotations

import json
import re
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
        try:
            df = pd.read_excel(path, sheet_name="fabric_master_en", dtype=str)
        except ValueError:
            df = pd.read_excel(path, sheet_name=0, dtype=str)
            df = _normalise_chinese_price_sheet(df)
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
        try:
            return pd.read_excel(path, sheet_name="fabric_price_rules", dtype=str)
        except ValueError:
            df = pd.read_excel(path, sheet_name=0, dtype=str)
            return _normalise_chinese_price_rules(df)
    return pd.DataFrame()


def _normalise_chinese_price_sheet(df: pd.DataFrame) -> pd.DataFrame:
    df = _clean_chinese_columns(df)
    required = {"面料编号", "品质等级", "成分", "幅宽", "克重"}
    if not required.issubset(set(df.columns)):
        raise ValueError("Excel 面料数据库需要 fabric_master_en 工作表，或包含中文价格库表头：面料编号/品质等级/成分/幅宽/克重")

    rows: list[dict[str, object]] = []
    for _, row in df.iterrows():
        code = _clean_text(row.get("面料编号", ""))
        if not code:
            continue
        grade_cn = _clean_text(row.get("品质等级", ""))
        quantification = _first_number(row.get("量化", ""))
        tube_allowance = _sum_numbers(row.get("纸筒+空差", ""))
        reference_roll = _first_number(row.get("参考条重", ""))
        rows.append(
            {
                "fabric_code": code,
                "quality_grade_cn": grade_cn,
                "quality_grade_en": _grade_cn_to_en(grade_cn),
                "fabric_name_cn": _clean_text(row.get("面料名称", "")),
                "fabric_name_en": "",
                "composition_cn": _clean_text(row.get("成分", "")),
                "composition_en": _composition_cn_to_en(row.get("成分", "")),
                "width_cn": _clean_text(row.get("幅宽", "")),
                "width_en": _width_to_en(row.get("幅宽", "")),
                "weight_cn": _clean_text(row.get("克重", "")),
                "weight_en": _weight_to_en(row.get("克重", "")),
                "quantification_m_per_kg": _float_text(quantification),
                "tube_plus_allowance_kg_per_roll": _float_text(tube_allowance),
                "reference_roll_weight_kg": _float_text(reference_roll),
                "reference_roll_weight": _clean_text(row.get("参考条重", "")),
                "remarks_cn": _clean_text(row.get("特殊标注", "")),
                "remarks_en": _clean_text(row.get("特殊标注", "")),
                "source_series_cn": _clean_text(row.get("系列来源", "")),
                "color_rule_cn": _clean_text(row.get("色号", "")),
                "price_adjustment_cn": _clean_text(row.get("价格调整", "")),
                "bulk_price_rmb_per_kg": _first_number_text(row.get("大货价", "")),
                "net_price_rmb_per_kg": _first_number_text(row.get("净布价", "")),
                "net_price_rmb_per_meter": _first_number_text(row.get("净布米价", "")),
                "sample_price_rmb_per_meter": _first_number_text(row.get("散剪价", "")),
                "default_bulk_price_rmb_per_kg": _first_number_text(row.get("默认大货价", "")),
                "default_net_price_rmb_per_kg": _first_number_text(row.get("默认净布KG价", "")),
                "default_net_price_rmb_per_meter": _first_number_text(row.get("默认参考净布米价", "")),
            }
        )
    return pd.DataFrame(rows)


def _normalise_chinese_price_rules(df: pd.DataFrame) -> pd.DataFrame:
    master = _normalise_chinese_price_sheet(df)
    keep = [
        "fabric_code",
        "quality_grade_cn",
        "quality_grade_en",
        "color_rule_cn",
        "price_adjustment_cn",
        "bulk_price_rmb_per_kg",
        "net_price_rmb_per_kg",
        "net_price_rmb_per_meter",
        "sample_price_rmb_per_meter",
        "default_bulk_price_rmb_per_kg",
        "default_net_price_rmb_per_kg",
        "default_net_price_rmb_per_meter",
        "remarks_cn",
        "remarks_en",
    ]
    return master[[column for column in keep if column in master.columns]].copy()


def _clean_chinese_columns(df: pd.DataFrame) -> pd.DataFrame:
    clean = df.copy()
    clean.columns = [str(column).strip() for column in clean.columns]
    clean = clean.loc[:, ~clean.columns.str.startswith("Unnamed")]
    return clean


def _grade_cn_to_en(value: object) -> str:
    text = _clean_text(value)
    if "一等" in text:
        return "First Grade"
    if "合格" in text:
        return "Qualified Grade"
    return text


def _composition_cn_to_en(value: object) -> str:
    text = _clean_text(value)
    replacements = {
        "粘纤": "Viscose",
        "粘胶": "Viscose",
        "棉": "Cotton",
        "氨纶": "Elastane",
        "聚酯纤维": "Polyester",
        "涤纶": "Polyester",
        "腈纶": "Acrylic",
        "锦纶": "Nylon",
        "羊毛": "Wool",
        "麻": "Linen",
    }
    for cn, en in replacements.items():
        text = text.replace(cn, en)
    return text


def _width_to_en(value: object) -> str:
    text = _clean_text(value).upper()
    match = re.search(r"\d+(?:\.\d+)?", text)
    return f"{match.group(0)}CM" if match else text


def _weight_to_en(value: object) -> str:
    text = _clean_text(value).upper()
    match = re.search(r"\d+(?:\.\d+)?", text)
    return f"{match.group(0)}G" if match else text


def _first_number(value: object) -> float | None:
    numbers = _numbers(value)
    return numbers[0] if numbers else None


def _sum_numbers(value: object) -> float | None:
    numbers = _numbers(value)
    if not numbers:
        return None
    return sum(numbers)


def _numbers(value: object) -> list[float]:
    text = _clean_text(value)
    return [float(match.group(0)) for match in re.finditer(r"\d+(?:\.\d+)?", text)]


def _first_number_text(value: object) -> str:
    number = _first_number(value)
    return _float_text(number)


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
