from __future__ import annotations

import csv
import io
import re
from collections import OrderedDict

import pandas as pd

from .models import FabricRecord, OrderInfo


ROLL_COLUMNS = ["item", "color", "lot", "net_weight_kg", "usd_price_per_kg"]
INVOICE_COLUMNS = [
    "color",
    "art_no",
    "ppo_reference_yards",
    "quantity_input",
    "input_unit",
    "usd_price_per_kg",
    "source_fabric_code",
    "source_note",
]


def sample_invoice_dataframe() -> pd.DataFrame:
    return pd.DataFrame(
        [
            ["MALIBLU", "MALIBLU - 13", 1320.0, 498.4, "KG", 10.3, "E260144_0002", "PPO reference"],
            ["ALMOND", "ALMOND - 6", 1600.0, 619.0, "KG", 10.3, "E260144_0003", "PPO reference"],
        ],
        columns=INVOICE_COLUMNS,
    )


def empty_invoice_dataframe() -> pd.DataFrame:
    return pd.DataFrame(
        [["", "", None, None, "KG", None, "", ""] for _ in range(4)],
        columns=INVOICE_COLUMNS,
    )


def sample_roll_dataframe() -> pd.DataFrame:
    navy = [
        17.6,
        20.6,
        21.4,
        22.9,
        22.9,
        23.1,
        23.2,
        23.3,
        23.4,
        23.4,
        23.5,
        23.5,
        23.6,
        23.6,
        23.6,
        23.6,
        23.6,
        23.7,
        23.7,
        23.7,
        23.8,
        23.8,
        23.8,
        23.8,
        26.1,
        26.1,
    ]
    thyme = [
        18.4,
        19.0,
        19.8,
        20.6,
        21.6,
        21.7,
        22.9,
        23.0,
        23.4,
        23.5,
        23.6,
        23.7,
        23.7,
        23.8,
        23.8,
        23.8,
        23.9,
        26.5,
        27.7,
    ]
    rows = []
    for weight in navy:
        rows.append(["25A109A", "NAVY/#0", "LDA2603-0032", weight, 10.3])
    for weight in thyme:
        rows.append(["25A109A", "THYME/#4", "LDA2603-0034B1", weight, 10.3])
    return pd.DataFrame(rows, columns=ROLL_COLUMNS)


def empty_roll_dataframe() -> pd.DataFrame:
    return pd.DataFrame(
        [["", "", "", None, None] for _ in range(8)],
        columns=ROLL_COLUMNS,
    )


def parse_roll_text(text: str, default_item: str = "", default_price: float | None = None) -> pd.DataFrame:
    text = text.strip()
    if not text:
        return empty_roll_dataframe()

    if "\t" in text:
        rows = _parse_roll_text_delimited(text, default_item=default_item, default_price=default_price)
        if rows:
            return pd.DataFrame(rows, columns=ROLL_COLUMNS)

    rows = _parse_roll_text_loose(text, default_item=default_item, default_price=default_price)
    if rows:
        return pd.DataFrame(rows, columns=ROLL_COLUMNS)

    rows = _parse_roll_text_delimited(text, default_item=default_item, default_price=default_price)
    return pd.DataFrame(rows, columns=ROLL_COLUMNS) if rows else empty_roll_dataframe()


def _parse_roll_text_delimited(text: str, default_item: str = "", default_price: float | None = None) -> list[list[object]]:
    rows = []
    reader = csv.reader(io.StringIO(text), delimiter="\t" if "\t" in text else ",")
    for raw in reader:
        cells = [c.strip() for c in raw if c.strip() != ""]
        if not cells:
            continue
        lower = [c.lower() for c in cells]
        if any(h in lower for h in ["color", "颜色", "net_weight_kg", "细码数量"]):
            continue
        if len(cells) >= 5:
            item, color, lot, weights, price = cells[:5]
        elif len(cells) == 4:
            item = default_item
            color, lot, weights, price = cells
        elif len(cells) == 3:
            item = default_item
            color, lot, weights = cells
            price = default_price
        else:
            continue
        for weight in _split_weights(str(weights)):
            rows.append([item or default_item, color, lot, weight, _to_float(price)])
    return rows


def _parse_roll_text_loose(text: str, default_item: str = "", default_price: float | None = None) -> list[list[object]]:
    rows: list[list[object]] = []
    current_item = default_item
    current_color = ""
    current_lot = ""
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if re.search(r"(合计|总计|本单|货款|扣余|实收|备注)", line, flags=re.I):
            continue
        tokens = line.split()
        lot_index = next((i for i, token in enumerate(tokens) if _looks_like_lot(token)), None)
        number_source = line
        if lot_index is not None:
            current_lot = tokens[lot_index]
            if lot_index >= 1:
                current_item = tokens[0] if default_item == "" else default_item
                current_color = " ".join(tokens[1:lot_index]) or current_color
            after_lot = line.split(current_lot, 1)[1]
            number_source = after_lot
        if not current_lot or not current_color:
            continue
        weights = _numbers_before_unit(number_source)
        for weight in weights:
            rows.append([current_item or default_item, current_color, current_lot, weight, default_price])
    return rows


def _looks_like_lot(value: str) -> bool:
    text = str(value or "").strip().strip(":：,，")
    return bool(re.fullmatch(r"[A-Z]{2,}\d[\w-]*", text, flags=re.I))


def _numbers_before_unit(value: str) -> list[float]:
    segment = re.split(r"(?:公斤|KG|KGS|kg|kgs|公厅|单位|条数|单价|金额)", str(value), maxsplit=1)[0]
    numbers = [_to_float(match.group(0)) for match in re.finditer(r"\d+(?:\.\d+)?", segment)]
    numbers = [number for number in numbers if number is not None and 3 <= float(number) <= 80]
    if len(numbers) >= 4 and float(numbers[-2]).is_integer() and int(numbers[-2]) == len(numbers) - 2:
        numbers = numbers[:-2]
    return [float(number) for number in numbers]


def calculate_rolls(
    rolls_df: pd.DataFrame,
    quantification_m_per_kg: float,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = rolls_df.copy()
    for column in ROLL_COLUMNS:
        if column not in df.columns:
            df[column] = None
    df = df[ROLL_COLUMNS]
    df["item"] = df["item"].fillna("").astype(str).str.strip()
    df["color"] = df["color"].fillna("").astype(str).str.strip()
    df["lot"] = df["lot"].fillna("").astype(str).str.strip()
    df["net_weight_kg"] = pd.to_numeric(df["net_weight_kg"], errors="coerce")
    df["usd_price_per_kg"] = pd.to_numeric(df["usd_price_per_kg"], errors="coerce")
    df = df.dropna(subset=["net_weight_kg"])
    df = df[(df["color"] != "") & (df["lot"] != "")]
    df["roll"] = df.groupby(["item", "color", "lot"], sort=False).cumcount() + 1
    df["meter"] = df["net_weight_kg"] * quantification_m_per_kg
    df["yard"] = df["meter"] / 0.9144
    df["amount_usd"] = df["net_weight_kg"] * df["usd_price_per_kg"]

    group_rows = []
    grouped = df.groupby(["item", "color", "lot"], sort=False, dropna=False)
    for (item, color, lot), group in grouped:
        total_net = float(group["net_weight_kg"].sum())
        total_amount = float(group["amount_usd"].sum())
        price = total_amount / total_net if total_net else 0.0
        group_rows.append(
            {
                "item": item,
                "color": color,
                "lot": lot,
                "rolls": int(len(group)),
                "total_net_weight_kg": total_net,
                "total_meter": float(group["meter"].sum()),
                "total_yard": float(group["yard"].sum()),
                "usd_price_per_kg": price,
                "amount_usd": total_amount,
                "usd_prices": ", ".join(_unique_strings(group["usd_price_per_kg"])),
            }
        )
    summary = pd.DataFrame(group_rows)
    return df, summary


def calculate_invoice_lines(invoice_df: pd.DataFrame, quantification_m_per_kg: float) -> pd.DataFrame:
    df = invoice_df.copy()
    for column in INVOICE_COLUMNS:
        if column not in df.columns:
            df[column] = None
    df = df[INVOICE_COLUMNS]
    df["color"] = df["color"].fillna("").astype(str).str.strip()
    df["art_no"] = df["color"]
    df["source_fabric_code"] = df["source_fabric_code"].fillna("").astype(str).str.strip()
    df["source_note"] = df["source_note"].fillna("").astype(str).str.strip()
    df["input_unit"] = df["input_unit"].fillna("KG").astype(str).str.strip()
    df["input_unit"] = df["input_unit"].replace({"YDS": "Yard", "Yards": "Yard", "Meters": "Meter", "KG": "KG"})
    df.loc[~df["input_unit"].isin(["KG", "Meter", "Yard"]), "input_unit"] = "KG"
    df["ppo_reference_yards"] = pd.to_numeric(df["ppo_reference_yards"], errors="coerce")
    df["quantity_input"] = pd.to_numeric(df["quantity_input"], errors="coerce")
    df["usd_price_per_kg"] = pd.to_numeric(df["usd_price_per_kg"], errors="coerce")
    df = df.dropna(subset=["quantity_input", "usd_price_per_kg"])
    df = df[df["color"] != ""]
    q = float(quantification_m_per_kg or 0)
    df["total_net_weight_kg"] = df.apply(lambda row: _invoice_kg(row, q), axis=1)
    df = df.dropna(subset=["total_net_weight_kg"])
    df["total_meter"] = df.apply(lambda row: _invoice_meter(row, q), axis=1)
    df["total_yard"] = df["total_meter"] / 0.9144
    df["amount_usd"] = df["total_net_weight_kg"] * df["usd_price_per_kg"]
    return df


def validate_order(
    fabric: FabricRecord | None,
    order: OrderInfo,
    rolls_df: pd.DataFrame,
    summary_df: pd.DataFrame,
) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []
    if fabric is None:
        errors.append("面料编号未在数据库中找到，请手动补充或检查编号")
        return errors, warnings

    if not fabric.quantification_m_per_kg or fabric.quantification_m_per_kg <= 0:
        errors.append("面料数据库缺少有效的 quantification_m_per_kg（量化）")
    if fabric.tube_plus_allowance_kg_per_roll is None:
        errors.append("面料数据库缺少 tube_plus_allowance_kg_per_roll（纸筒+空差）")
    if not order.buyer.strip():
        errors.append("Buyer/TO 不能为空")
    if not order.po_no.strip():
        errors.append("PO NO 不能为空")
    if not order.pi_no.strip():
        errors.append("P/I NO 不能为空")
    if rolls_df.empty:
        errors.append("码单明细为空，请录入至少一条 Net Weight")
    if summary_df.empty:
        errors.append("没有可生成的颜色/LOT 汇总")
    if len(summary_df) > 4:
        errors.append("当前 Packing List 模板最多支持 4 个颜色/LOT 分组，超过部分请先拆单或扩展模板")

    for idx, row in rolls_df.iterrows():
        row_no = int(idx) + 1
        if not row.get("color"):
            errors.append(f"第 {row_no} 行颜色为空")
        if not row.get("lot"):
            errors.append(f"第 {row_no} 行 LOT 为空")
        weight = row.get("net_weight_kg")
        if pd.isna(weight) or float(weight) <= 0:
            errors.append(f"第 {row_no} 行 Net Weight 不是有效正数")
        price = row.get("usd_price_per_kg")
        if pd.isna(price) or float(price) <= 0:
            errors.append(f"第 {row_no} 行 USD/KG 单价不是有效正数")

    for _, row in summary_df.iterrows():
        prices = [p for p in str(row.get("usd_prices", "")).split(",") if p.strip()]
        if len(prices) > 1:
            warnings.append(f"{row['color']} / {row['lot']} 同组存在多个 USD/KG 单价，汇总单价按加权平均显示")
    return errors, warnings


def validate_invoice_order(
    fabric: FabricRecord | None,
    order: OrderInfo,
    invoice_df: pd.DataFrame,
    quantity_unit: str,
) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []
    df = invoice_df.copy()
    for column in INVOICE_COLUMNS:
        if column not in df.columns:
            df[column] = None
    df = df[INVOICE_COLUMNS]
    df["color"] = df["color"].fillna("").astype(str).str.strip()
    df["art_no"] = df["color"]
    df["source_fabric_code"] = df["source_fabric_code"].fillna("").astype(str).str.strip()
    df["input_unit"] = df["input_unit"].fillna("Yard").astype(str).str.strip()
    df["input_unit"] = df["input_unit"].replace({"YDS": "Yard", "Yards": "Yard", "Meters": "Meter"})
    df.loc[~df["input_unit"].isin(["KG", "Meter", "Yard"]), "input_unit"] = "Yard"
    df["quantity_input"] = pd.to_numeric(df["quantity_input"], errors="coerce")
    df["usd_price_per_kg"] = pd.to_numeric(df["usd_price_per_kg"], errors="coerce")
    blank_rows = (
        df["color"].eq("")
        & df["art_no"].eq("")
        & df["source_fabric_code"].eq("")
        & df["quantity_input"].isna()
        & df["usd_price_per_kg"].isna()
    )
    df = df.loc[~blank_rows].copy()
    if fabric is None:
        errors.append("面料编号未在数据库中找到，请手动补充或检查编号")
        return errors, warnings
    if not fabric.quantification_m_per_kg or fabric.quantification_m_per_kg <= 0:
        errors.append("当前面料缺少量化 m/kg，无法完成 KG / Meter / Yard 换算。请在第 2 步“面料数据人工修正”里填写量化。")
    if not order.buyer.strip():
        errors.append("Buyer 不能为空")
    if not order.po_no.strip():
        errors.append("PO NO 不能为空")
    if not order.pi_no.strip():
        errors.append("P/I NO 不能为空")
    if not order.ci_no.strip():
        errors.append("C/I NO 不能为空")
    if df.empty:
        errors.append("P.I/C.I 明细为空，请录入至少一个颜色")
    missing_quantity_rows: list[int] = []
    missing_price_rows: list[int] = []
    invalid_unit_rows: list[int] = []
    for idx, row in df.iterrows():
        row_no = int(idx) + 1
        if not str(row.get("color", "")).strip():
            errors.append(f"第 {row_no} 行颜色为空")
        quantity = row.get("quantity_input")
        if pd.isna(quantity) or float(quantity) <= 0:
            missing_quantity_rows.append(row_no)
        if str(row.get("input_unit", "")).strip() not in {"KG", "Meter", "Yard"}:
            invalid_unit_rows.append(row_no)
        price = row.get("usd_price_per_kg")
        if pd.isna(price) or float(price) <= 0:
            missing_price_rows.append(row_no)
    if missing_quantity_rows:
        errors.append(f"请填写第 {_format_row_numbers(missing_quantity_rows)} 行的 Quantity；COLOR 是文本，不需要是数字。")
    if invalid_unit_rows:
        errors.append(f"第 {_format_row_numbers(invalid_unit_rows)} 行 Unit 必须是 Yard / Meter / KG。")
    if missing_price_rows:
        errors.append(f"请填写第 {_format_row_numbers(missing_price_rows)} 行的 USD/KG 单价。")
    return errors, warnings


def _format_row_numbers(rows: list[int]) -> str:
    return "、".join(str(row) for row in rows)


def apply_gross_weight(summary_df: pd.DataFrame, tube_allowance: float) -> pd.DataFrame:
    df = summary_df.copy()
    if df.empty or "total_net_weight_kg" not in df.columns or "rolls" not in df.columns:
        return df
    df["total_gross_weight_kg"] = df["total_net_weight_kg"] + df["rolls"] * tube_allowance
    return df


def _split_weights(value: str) -> list[float]:
    return [_to_float(x) for x in re.split(r"[\s,;，、]+", value.strip()) if _to_float(x) is not None]


def _to_float(value: object) -> float | None:
    if value is None:
        return None
    try:
        text = str(value).strip()
        if text == "":
            return None
        return float(text)
    except ValueError:
        return None


def _invoice_kg(row: pd.Series, quantification_m_per_kg: float) -> float | None:
    quantity = float(row["quantity_input"])
    unit = str(row.get("input_unit") or "Yard")
    if unit == "KG":
        return quantity
    if quantification_m_per_kg <= 0:
        return None
    if unit == "Meter":
        return quantity / quantification_m_per_kg
    return quantity * 0.9144 / quantification_m_per_kg


def _invoice_meter(row: pd.Series, quantification_m_per_kg: float) -> float:
    quantity = float(row["quantity_input"])
    unit = str(row.get("input_unit") or "Yard")
    if unit == "KG":
        return quantity * quantification_m_per_kg if quantification_m_per_kg > 0 else 0.0
    if unit == "Meter":
        return quantity
    return quantity * 0.9144


def _unique_strings(values: pd.Series) -> list[str]:
    seen: OrderedDict[str, None] = OrderedDict()
    for value in values:
        if pd.isna(value):
            continue
        text = f"{float(value):.6f}".rstrip("0").rstrip(".")
        seen[text] = None
    return list(seen.keys())
