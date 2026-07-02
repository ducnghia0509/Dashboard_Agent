# -*- coding: utf-8 -*-
"""Pydantic v2: cac model dung xuyen suot luong 1 (ingest)."""
from typing import Optional

from pydantic import BaseModel, Field


class Discovery(BaseModel):
    file_name: str
    fingerprint: str
    sheets: list[str] = Field(default_factory=list)
    columns_per_sheet: dict[str, list[str]] = Field(default_factory=dict)
    detected_report_type: Optional[str] = None
    header_row: Optional[int] = None
    mapping: dict[str, int] = Field(default_factory=dict)
    period: Optional[str] = None
    confidence: float = 0.0
    anomalies: list[str] = Field(default_factory=list)
    created_at: Optional[str] = None


class TemplateSpec(BaseModel):
    file_name: str
    report_type: Optional[str] = None
    header_row: int = 0
    headers: list[str] = Field(default_factory=list)
    mapping: dict[str, int] = Field(default_factory=dict)
    confidence: float = 0.0
    missing_required: list[str] = Field(default_factory=list)
    low_confidence_fields: list[str] = Field(default_factory=list)
    anomalies: list[str] = Field(default_factory=list)
    sample_rows: list[list] = Field(default_factory=list)
    all_sheets: list[str] = Field(default_factory=list)


class ImportPlan(BaseModel):
    file_name: str
    sheet: Optional[str] = None
    report_type: str
    period: Optional[str] = None
    dataset_kind: str = "day"  # "day" (importer.commit) | "month" (importer_month.import_workbook)
    header_row: int = 0
    mapping: dict[str, int] = Field(default_factory=dict)
    idempotency_key: Optional[str] = None
    validations: list[str] = Field(default_factory=list)


class ValidationResult(BaseModel):
    ok: bool
    errors: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ImportResult(BaseModel):
    dry_run: bool
    dataset_id: Optional[str] = None
    by_type: dict[str, int] = Field(default_factory=dict)
    skipped_duplicate: bool = False
    message: str = ""


# ---- Extension 2: generic sheet discovery cho template không có report_type cố định ----

class SheetProfile(BaseModel):
    """Sự thật thô về 1 sheet - KHÔNG suy luận gì, chỉ để analyst (LLM) tự đọc và suy luận."""
    sheet: str
    dimensions: Optional[str] = None
    merged_ranges: list[str] = Field(default_factory=list)
    row_sample: list[list] = Field(default_factory=list)   # N dòng đầu, mọi cột - dò header ngang
    col_sample: dict[str, list] = Field(default_factory=dict)  # N cột đầu, M dòng - dò layout theo cột


class ColumnRole(BaseModel):
    index: int
    role: str   # tên tự do do analyst đặt, vd 'entity_code','entity_name','dau_ky_no','ps_co',...
    label: str = ""


class EntityColumn(BaseModel):
    col_index: int
    entity_name: str


class RowRole(BaseModel):
    row_index: int
    role: str   # vd mã số dòng CĐKT ('270','400'), hoặc tên khoản mục
    label: str = ""


class DisplaySpec(BaseModel):
    """CHỈ để FE hiển thị - không dùng khi ghi raw_rows. qa không cần format riêng,
    có thể dùng field này làm ngữ cảnh khi giải thích số liệu nếu cần."""
    title: str = ""
    unit: str = ""
    columns: list[dict] = Field(default_factory=list)  # [{key,label}]
    note: str = ""


class SheetMapping(BaseModel):
    """Kết luận của analyst sau khi khớp 1 sheet lạ với guideline.xlsx (kpi_glossary.json)."""
    file_name: str
    sheet: str
    orientation: str  # "row_major" | "column_major"
    header_rows: list[int] = Field(default_factory=list)
    data_start_row: int = 0
    columns: list[ColumnRole] = Field(default_factory=list)     # dùng khi orientation=row_major
    entities: list[EntityColumn] = Field(default_factory=list)  # dùng khi orientation=column_major
    row_roles: list[RowRole] = Field(default_factory=list)      # dùng khi orientation=column_major
    matched_kpi_ids: list[int] = Field(default_factory=list)
    target_report_type: str  # PHẢI có tiền tố GEN_
    canonical_kind: Optional[str] = None  # vd 'TK131'/'TSCD' - xem servers/common/canonical.py;
                                           # cho phép tìm lại mapping đã học BẤT KỂ tên sheet/file
                                           # khác nhau giữa các công ty (report_spec_search)
    company: Optional[str] = None  # công ty/đơn vị sở hữu file này (từ discover_files hoặc do
                                    # người dùng xác nhận) - stamp vào raw_rows.cong_ty khi ghi
    confidence: float = 0.0
    notes: list[str] = Field(default_factory=list)
    display_spec: DisplaySpec = Field(default_factory=DisplaySpec)
