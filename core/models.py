from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Literal, Optional


@dataclass
class ImportConfig:
    """Configuration options for batch importing into Phôi DS.docx."""

    doc_date: Optional[date] = field(default_factory=date.today)
    convert_old_provinces: bool = True
    fill_missing_phone: bool = True

    @property
    def formatted_date_vn(self) -> Optional[str]:
        if self.doc_date is None:
            return None
        day_str = f"{self.doc_date.day:02d}"
        month_str = (
            f"{self.doc_date.month:02d}"
            if self.doc_date.month in (1, 2)
            else str(self.doc_date.month)
        )
        return f"Hà Nội, ngày {day_str} tháng {month_str} năm {self.doc_date.year:04d}"


@dataclass
class ValidationWarning:
    """A single data validation issue logged into Bao_cao_tong_hop.xlsx (never modifies Word formatting)."""

    file_name: str
    table_name: str
    row_stt: str
    mssv: str
    full_name: str
    column_name: str
    original_value: str
    reason: str


@dataclass
class ClassSummary:
    """8 Youth Union statistics computed from the 16-col Member List table."""

    file_name: str
    class_name: str = ""
    secretary_name: str = ""
    total_students: int = 0
    total_union_members: int = 0
    total_union_books: int = 0
    female_members: int = 0
    male_members: int = 0
    party_members: int = 0
    party_candidates: int = 0
    youth_non_members: int = 0
    warning_count: int = 0


@dataclass
class ExtractedClassData:
    """Structured cleaned data extracted from one input Chi đoàn .docx file."""

    file_name: str
    class_name_title: str
    class_name_upper: str
    secretary_name: str
    member_rows: list[list[str]]  # each row has 16 columns (index 0..15)
    bch_rows: list[list[str]]     # each row has 14 columns (index 0..13)
    summary: ClassSummary
    warnings: list[ValidationWarning] = field(default_factory=list)
    member_invalid_cells: set[tuple[int, int]] = field(default_factory=set)  # (0-based row idx, col idx)
    bch_invalid_cells: set[tuple[int, int]] = field(default_factory=set)     # (0-based row idx, col idx)


@dataclass
class FileImportResult:
    """Outcome of importing one input .docx into Phôi DS.docx."""

    original_name: str
    output_name: str
    status: Literal["Thành công", "Lỗi cấu trúc", "Lỗi xử lý"]
    message: str = ""
    class_name: str = ""
    secretary_name: str = ""
    output_bytes: Optional[bytes] = None
    summary: Optional[ClassSummary] = None
    warnings: list[ValidationWarning] = field(default_factory=list)


@dataclass
class BatchImportResult:
    """Aggregated outcome of batch importing multiple .docx files."""

    results: list[FileImportResult] = field(default_factory=list)
    excel_bytes: bytes = b""
    zip_bytes: bytes = b""

    @property
    def total_files(self) -> int:
        return len(self.results)

    @property
    def success_count(self) -> int:
        return sum(1 for r in self.results if r.status == "Thành công")

    @property
    def error_count(self) -> int:
        return sum(1 for r in self.results if r.status != "Thành công")

    @property
    def total_warnings(self) -> int:
        return sum(len(r.warnings) for r in self.results)
