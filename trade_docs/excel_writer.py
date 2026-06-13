from __future__ import annotations

import json
import re
from copy import copy
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

from .models import FabricRecord, GeneratedFile, GeneratedInvoiceFiles, OrderInfo
from .numbering import commit_doc_number


BLOCKS = [
    {"color": 1, "lot": 2, "roll": 3, "net": 4, "meter": 5, "yard": 6, "gross": 4},
    {"color": 7, "lot": 8, "roll": 9, "net": 10, "meter": 11, "yard": 12, "gross": 10},
    {"color": 13, "lot": 14, "roll": 15, "net": 16, "meter": 17, "yard": 18, "gross": 16},
    {"color": 19, "lot": 20, "roll": 21, "net": 22, "meter": 23, "yard": 24, "gross": 22},
]
DATA_START_ROW = 10
DATA_END_ROW = 37
TOTAL_ROW = 38
GROSS_ROW = 40


def write_packing_list(
    template_path: str | Path,
    order: OrderInfo,
    fabric: FabricRecord,
    rolls_df: pd.DataFrame,
    summary_df: pd.DataFrame,
) -> GeneratedFile:
    template_path = Path(template_path)
    output_dir = Path(order.output_dir)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = output_dir / stamp
    run_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(template_path)
    ws = wb.active

    _write_header(ws, order, fabric)
    _clear_roll_blocks(ws)
    _write_roll_blocks(ws, rolls_df, summary_df, fabric)
    _write_summary_area(ws, order, summary_df)
    _write_gross_area(ws, order, summary_df, fabric)

    filename = f"Packing List-{_safe(order.po_no)}-{_safe(order.art_no or order.fabric_code)}-{order.order_date.replace('-', '')}.xlsx"
    output_path = run_dir / filename
    wb.save(output_path)

    log = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "template": str(template_path),
        "output": str(output_path),
        "price_source": "USD/KG from PI/CI or user input",
        "order": {
            "po_no": order.po_no,
            "art_no": order.art_no,
            "fabric_code": order.fabric_code,
            "buyer": order.buyer,
            "buyer_address": order.buyer_address,
            "pi_no": order.pi_no,
            "ci_no": order.ci_no,
            "order_date": order.order_date,
            "advance_payment_usd": order.advance_payment_usd,
            "note": order.note,
        },
        "fabric": fabric.raw,
        "rolls": rolls_df.to_dict(orient="records"),
        "summary": summary_df.to_dict(orient="records"),
    }
    log_path = run_dir / f"{output_path.stem}.json"
    log_path.write_text(json.dumps(log, indent=2, ensure_ascii=False), encoding="utf-8")
    commit_doc_number(order.pi_no)
    return GeneratedFile(path=output_path, log_path=log_path)


def write_invoice_pair(
    template_path: str | Path,
    order: OrderInfo,
    fabric: FabricRecord,
    invoice_df: pd.DataFrame,
    exchange_rate: float | None = None,
    exchange_source: str = "",
    quantity_unit: str = "Yard",
) -> GeneratedInvoiceFiles:
    template_path = Path(template_path)
    output_dir = Path(order.output_dir)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = output_dir / stamp
    run_dir.mkdir(parents=True, exist_ok=True)

    pi_path = _write_invoice(
        template_path=template_path,
        run_dir=run_dir,
        order=order,
        fabric=fabric,
        invoice_df=invoice_df,
        title="PROFORMA INVOICE",
        doc_label="P/I NO",
        doc_no=order.pi_no,
        prefix="PI",
        quantity_unit=quantity_unit,
    )
    ci_path = _write_invoice(
        template_path=template_path,
        run_dir=run_dir,
        order=order,
        fabric=fabric,
        invoice_df=invoice_df,
        title="COMMERCIAL INVOICE",
        doc_label="C/I NO",
        doc_no=order.ci_no,
        prefix="CI",
        quantity_unit=quantity_unit,
    )

    log = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "template": str(template_path),
        "pi_output": str(pi_path),
        "ci_output": str(ci_path),
        "exchange_rate_usd_cny": exchange_rate,
        "exchange_source": exchange_source,
        "order": {
            "po_no": order.po_no,
            "art_no": order.art_no,
            "fabric_code": order.fabric_code,
            "buyer": order.buyer,
            "buyer_address": order.buyer_address,
            "pi_no": order.pi_no,
            "ci_no": order.ci_no,
            "order_date": order.order_date,
            "advance_payment_usd": order.advance_payment_usd,
            "note": order.note,
        },
        "fabric": fabric.raw,
        "quantity_unit": quantity_unit,
        "invoice_lines": invoice_df.to_dict(orient="records"),
    }
    log_path = run_dir / f"PI-CI-{_safe(order.po_no)}-{_safe(order.art_no or order.fabric_code)}-{order.order_date.replace('-', '')}.json"
    log_path.write_text(json.dumps(log, indent=2, ensure_ascii=False), encoding="utf-8")
    commit_doc_number(order.pi_no)
    commit_doc_number(order.ci_no)
    return GeneratedInvoiceFiles(pi_path=pi_path, ci_path=ci_path, log_path=log_path)


def _write_invoice(
    template_path: Path,
    run_dir: Path,
    order: OrderInfo,
    fabric: FabricRecord,
    invoice_df: pd.DataFrame,
    title: str,
    doc_label: str,
    doc_no: str,
    prefix: str,
    quantity_unit: str,
) -> Path:
    wb = load_workbook(template_path)
    ws = wb.active
    _set(ws, "A4", title)
    _set(ws, "F5", f"{doc_label}:{doc_no}")
    _set(ws, "F6", f"DATE: {order.order_date}")
    _set(ws, "A11", f"BUYER: {order.buyer}")
    _set(ws, "A12", order.buyer_address)
    price_width = ws.column_dimensions["E"].width
    amount_width = ws.column_dimensions["F"].width
    _set(ws, "D16", f"QUANTITY/{quantity_unit.upper()}")
    _set(ws, "E16", "QUANTITY/KG")
    _set(ws, "F16", "PRICES/KG\n(USD)FOB")
    _set(ws, "G16", "TOTAL AMOUNT:")
    ws.column_dimensions["G"].hidden = False
    ws.column_dimensions["F"].width = price_width
    ws.column_dimensions["G"].width = amount_width
    _copy_cell_format(ws, "F16", "G16")

    for row in range(17, 21):
        for col in range(1, 8):
            _safe_set(ws, row, col, None)
        _copy_cell_format(ws, f"F{row}", f"G{row}")

    description = _invoice_description(order, fabric)
    for offset, item in enumerate(invoice_df.itertuples(index=False), start=17):
        if offset > 20:
            break
        color = str(item.color)
        _safe_set(ws, offset, 1, _invoice_po_line(order, item))
        _safe_set(ws, offset, 2, description)
        _safe_set(ws, offset, 3, color)
        _safe_set(ws, offset, 4, round(_invoice_quantity(item, quantity_unit), 2))
        _safe_set(ws, offset, 5, round(float(item.total_net_weight_kg), 4))
        _safe_set(ws, offset, 6, round(float(item.usd_price_per_kg), 4))
        _safe_set(ws, offset, 7, f"=E{offset}*F{offset}")

    _set(ws, "D21", float(order.advance_payment_usd or 0))
    _set(ws, "C22", "=SUM(G17:G20)-D21")

    filename = f"{prefix}-{_safe(order.po_no)}-{_safe(order.art_no or order.fabric_code)}-{order.order_date.replace('-', '')}.xlsx"
    output_path = run_dir / filename
    wb.save(output_path)
    return output_path


def _invoice_quantity(item, quantity_unit: str) -> float:
    if quantity_unit == "KG":
        return float(item.total_net_weight_kg)
    if quantity_unit == "Meter":
        return float(item.total_meter)
    return float(item.total_yard)


def _invoice_po_line(order: OrderInfo, item) -> str:
    source_code = str(getattr(item, "source_fabric_code", "") or order.fabric_code).strip()
    color = str(getattr(item, "color", "")).strip()
    art_no = str(getattr(item, "art_no", "") or color).strip()
    return (
        f"PO NO.: {order.po_no}\n"
        f"code: {source_code}\n"
        f"ART NO.:{art_no}\n"
        "Knit fabric"
    )


def _invoice_description(order: OrderInfo, fabric: FabricRecord) -> str:
    return (
        f"ITEM: {order.art_no or fabric.fabric_code}\n"
        f"COMP:  {fabric.composition_en or fabric.composition_cn}\n"
        f"WEIGHT :{fabric.weight_en or fabric.weight_cn}        WIDTH :{fabric.width_en or fabric.width_cn}\n"
        "FINISH:\n"
        "COUNTRY OF ORIGIN:CHINA"
    )


def _write_header(ws, order: OrderInfo, fabric: FabricRecord) -> None:
    _set(ws, "G5", order.buyer)
    _set(ws, "P5", f"P/I NO:{order.pi_no}")
    _set(ws, "P6", f"DATE: {order.order_date}")
    _set(ws, "G6", fabric.packing_description(order.art_no))


def _clear_roll_blocks(ws) -> None:
    for block in BLOCKS:
        _safe_set(ws, DATA_START_ROW, block["color"], None)
        for row in range(DATA_START_ROW, DATA_END_ROW + 1):
            for key in ["lot", "roll", "net", "meter", "yard"]:
                _safe_set(ws, row, block[key], None)
        for key in ["roll", "net", "meter", "yard"]:
            _safe_set(ws, TOTAL_ROW, block[key], None)
        _safe_set(ws, GROSS_ROW, block["gross"], None)

    for row in range(6, 10):
        for col in range(26, 31):
            _safe_set(ws, row, col, None)
    for cell in ["AB39", "AC39", "AB40", "AC40", "AD40", "AB41", "AC41", "AD41", "AE40"]:
        _set(ws, cell, None)


def _write_roll_blocks(ws, rolls_df: pd.DataFrame, summary_df: pd.DataFrame, fabric: FabricRecord) -> None:
    quantification = _num(fabric.quantification_m_per_kg)
    yard_factor = 1 / 0.9144
    groups = list(summary_df[["item", "color", "lot"]].itertuples(index=False, name=None))
    for index, (item, color, lot) in enumerate(groups):
        block = BLOCKS[index]
        group = rolls_df[
            (rolls_df["item"].fillna("").astype(str) == str(item))
            & (rolls_df["color"].astype(str) == str(color))
            & (rolls_df["lot"].astype(str) == str(lot))
        ].copy()
        color_text = str(color)
        _safe_set(ws, DATA_START_ROW, block["color"], color_text)

        for offset, row_data in enumerate(group.itertuples(index=False)):
            row = DATA_START_ROW + offset
            _safe_set(ws, row, block["lot"], str(lot))
            _safe_set(ws, row, block["roll"], int(getattr(row_data, "roll")))
            _safe_set(ws, row, block["net"], float(getattr(row_data, "net_weight_kg")))
            net_cell = f"{get_column_letter(block['net'])}{row}"
            meter_cell = f"{get_column_letter(block['meter'])}{row}"
            _safe_set(ws, row, block["meter"], f"={net_cell}*{quantification}")
            _safe_set(ws, row, block["yard"], f"={meter_cell}*{_num(yard_factor)}")

        roll_col = get_column_letter(block["roll"])
        net_col = get_column_letter(block["net"])
        meter_col = get_column_letter(block["meter"])
        yard_col = get_column_letter(block["yard"])
        _safe_set(ws, TOTAL_ROW, block["roll"], f"=COUNT({roll_col}{DATA_START_ROW}:{roll_col}{DATA_END_ROW})")
        _safe_set(ws, TOTAL_ROW, block["net"], f"=SUM({net_col}{DATA_START_ROW}:{net_col}{DATA_END_ROW})")
        _safe_set(ws, TOTAL_ROW, block["meter"], f"=SUM({meter_col}{DATA_START_ROW}:{meter_col}{DATA_END_ROW})")
        _safe_set(ws, TOTAL_ROW, block["yard"], f"=SUM({yard_col}{DATA_START_ROW}:{yard_col}{DATA_END_ROW})")


def _write_summary_area(ws, order: OrderInfo, summary_df: pd.DataFrame) -> None:
    for i, row in enumerate(summary_df.itertuples(index=False), start=6):
        _set(ws, f"Z{i}", str(row.color))
        _set(ws, f"AA{i}", round(float(row.total_net_weight_kg), 2))
        _set(ws, f"AB{i}", round(float(row.usd_price_per_kg), 4))
        _set(ws, f"AC{i}", f"=AA{i}*AB{i}")
    _set(ws, "AD6", float(order.advance_payment_usd or 0))
    _set(ws, "AE6", "=SUM(AC6:AC9)-AD6")


def _write_gross_area(ws, order: OrderInfo, summary_df: pd.DataFrame, fabric: FabricRecord) -> None:
    allowance = _num(fabric.tube_plus_allowance_kg_per_roll)
    _set(ws, "AB39", order.art_no or order.fabric_code)
    _set(ws, "AC39", order.fabric_code)
    for block in BLOCKS[: len(summary_df)]:
        net_cell = f"{get_column_letter(block['net'])}{TOTAL_ROW}"
        roll_cell = f"{get_column_letter(block['roll'])}{TOTAL_ROW}"
        _safe_set(ws, GROSS_ROW, block["gross"], f"={net_cell}+{roll_cell}*{allowance}")
    gross_cells = ",".join(f"{get_column_letter(block['gross'])}{GROSS_ROW}" for block in BLOCKS[: len(summary_df)])
    _set(ws, "AB40", f"=SUM({gross_cells})" if gross_cells else 0)
    _set(ws, "AD40", "=AB40")
    _set(ws, "AE40", order.note or None)


def _set(ws, coordinate: str, value) -> None:
    cell = ws[coordinate]
    if isinstance(cell, MergedCell):
        raise ValueError(f"Cannot write to merged child cell {coordinate}")
    cell.value = value


def _safe_set(ws, row: int, col: int, value) -> None:
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        return
    cell.value = value


def _copy_cell_format(ws, source: str, target: str) -> None:
    source_cell = ws[source]
    target_cell = ws[target]
    if isinstance(source_cell, MergedCell) or isinstance(target_cell, MergedCell):
        return
    if source_cell.has_style:
        target_cell._style = copy(source_cell._style)
    if source_cell.number_format:
        target_cell.number_format = source_cell.number_format
    if source_cell.alignment:
        target_cell.alignment = copy(source_cell.alignment)
    if source_cell.border:
        target_cell.border = copy(source_cell.border)
    if source_cell.fill:
        target_cell.fill = copy(source_cell.fill)
    if source_cell.font:
        target_cell.font = copy(source_cell.font)


def _num(value: float | None) -> str:
    if value is None:
        return "0"
    return f"{float(value):.8f}".rstrip("0").rstrip(".")


def _safe(value: str) -> str:
    text = str(value or "").strip() or "UNKNOWN"
    return re.sub(r'[<>:"/\\\\|?*\\s]+', "_", text).strip("_")
