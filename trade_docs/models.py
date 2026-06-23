from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class FabricRecord:
    fabric_code: str
    quality_grade_cn: str = ""
    quality_grade_en: str = ""
    fabric_name_cn: str = ""
    fabric_name_en: str = ""
    composition_cn: str = ""
    composition_en: str = ""
    width_cn: str = ""
    width_en: str = ""
    weight_cn: str = ""
    weight_en: str = ""
    quantification_m_per_kg: float | None = None
    tube_plus_allowance_kg_per_roll: float | None = None
    reference_roll_weight_kg: float | None = None
    remarks_cn: str = ""
    remarks_en: str = ""
    raw: dict[str, Any] = field(default_factory=dict)

    @property
    def display_name(self) -> str:
        name = self.fabric_name_en or self.fabric_name_cn or "Not specified"
        grade = self.quality_grade_en or self.quality_grade_cn
        return f"{self.fabric_code} - {name}" + (f" ({grade})" if grade else "")

    def packing_description(self, art_no: str | None = None) -> str:
        item = (art_no or self.fabric_code).strip()
        parts = [
            f"ITEM: {item}",
            f"COMP:  {self.composition_en or self.composition_cn}",
            f"WEIGHT :{self.weight_en or self.weight_cn}",
            f"WIDTH :{self.width_en or self.width_cn}",
        ]
        return "\n".join(parts)


@dataclass(frozen=True)
class OrderInfo:
    po_no: str
    art_no: str
    fabric_code: str
    buyer: str
    pi_no: str
    ci_no: str
    order_date: str
    advance_payment_usd: float = 0.0
    output_dir: Path = Path("outputs")
    note: str = ""
    buyer_address: str = ""
    payment_terms: str = ""
    delivery_time: str = ""
    port_destination: str = ""
    port_loading: str = "GUANGZHOU,CHINA"


@dataclass(frozen=True)
class GeneratedFile:
    path: Path
    log_path: Path


@dataclass(frozen=True)
class GeneratedInvoiceFiles:
    pi_path: Path
    ci_path: Path
    log_path: Path
