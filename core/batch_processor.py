from __future__ import annotations

import io
import zipfile
from pathlib import Path, PurePosixPath
from typing import Callable, Iterable, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .extractor import StructureError, extract_class_data
from .models import (
    BatchImportResult,
    ClassSummary,
    FileImportResult,
    ImportConfig,
    ValidationWarning,
)
from .template_importer import import_into_template

ProgressCallback = Callable[[int, int, FileImportResult], None]


def _make_output_name(original_name: str, used_names: set[str]) -> str:
    stem = Path(original_name).stem
    candidate = f"{stem}_imported.docx"
    counter = 2
    while candidate.lower() in used_names:
        candidate = f"{stem}_{counter}_imported.docx"
        counter += 1
    used_names.add(candidate.lower())
    return candidate


def _decode_zip_member_name(info: zipfile.ZipInfo) -> str:
    name = info.filename
    if not (info.flag_bits & 0x800):
        try:
            return name.encode("cp437").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return name


def extract_docx_inputs(
    uploaded_items: Iterable[tuple[str, bytes]],
) -> tuple[list[tuple[str, bytes]], list[FileImportResult]]:
    """Expand uploaded .docx and .zip items into a flat list of (docx_filename, docx_bytes)."""
    docx_items: list[tuple[str, bytes]] = []
    pre_errors: list[FileImportResult] = []

    for name, data in uploaded_items:
        lower = name.lower()
        if lower.endswith(".docx"):
            base = Path(name).name
            if base.startswith("~$"):
                continue
            docx_items.append((base, data))
        elif lower.endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
                    for info in zf.infolist():
                        if info.is_dir():
                            continue
                        decoded = _decode_zip_member_name(info)
                        posix = PurePosixPath(decoded.replace("\\", "/"))
                        if "__MACOSX" in posix.parts:
                            continue
                        base = posix.name
                        if not base.lower().endswith(".docx"):
                            continue
                        if base.startswith("~$") or base.startswith("._"):
                            continue
                        docx_items.append((base, zf.read(info)))
            except Exception as exc:
                pre_errors.append(
                    FileImportResult(
                        original_name=name,
                        output_name="",
                        status="Lỗi xử lý",
                        message=f"Không thể giải nén file ZIP: {exc}",
                    )
                )
        else:
            pre_errors.append(
                FileImportResult(
                    original_name=name,
                    output_name="",
                    status="Lỗi cấu trúc",
                    message="Định dạng file không được hỗ trợ (chỉ nhận .docx hoặc .zip).",
                )
            )

    return docx_items, pre_errors


def build_summary_excel(
    summaries: list[ClassSummary],
    warnings: list[ValidationWarning],
) -> bytes:
    """Create Bao_cao_tong_hop.xlsx with sheets 'Tong_hop' and 'Canh_bao_loi'."""
    wb = Workbook()

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    warn_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    header_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Times New Roman", size=11, bold=False, color="000000")
    total_font = Font(name="Times New Roman", size=11, bold=True, color="000000")

    thin = Side(border_style="thin", color="BFBFBF")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---------------- Sheet 1: Tong_hop ----------------
    ws_sum = wb.active
    ws_sum.title = "Tong_hop"
    ws_sum.freeze_panes = "A2"

    sum_headers = [
        "Tên file",
        "Tên Chi đoàn",
        "Bí thư Chi đoàn",
        "Tổng SV",
        "Tổng ĐV",
        "Số sổ ĐV",
        "Nữ",
        "Nam",
        "Đảng viên",
        "ĐTĐ",
        "Thanh niên",
    ]
    ws_sum.append(sum_headers)
    for col_idx in range(1, len(sum_headers) + 1):
        c = ws_sum.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = cell_border
    ws_sum.row_dimensions[1].height = 26

    for row_idx, s in enumerate(summaries, start=2):
        ws_sum.append(
            [
                s.file_name,
                s.class_name,
                s.secretary_name,
                s.total_students,
                s.total_union_members,
                s.total_union_books,
                s.female_members,
                s.male_members,
                s.party_members,
                s.party_candidates,
                s.youth_non_members,
            ]
        )
        for col_idx in range(1, len(sum_headers) + 1):
            c = ws_sum.cell(row=row_idx, column=col_idx)
            c.font = body_font
            c.border = cell_border
            c.alignment = Alignment(
                horizontal="left" if col_idx <= 3 else "center",
                vertical="center",
            )

    if summaries:
        tot_row = len(summaries) + 2
        ws_sum.cell(row=tot_row, column=1, value="TỔNG CỘNG")
        for col_idx in range(4, len(sum_headers) + 1):
            col_letter = get_column_letter(col_idx)
            ws_sum.cell(
                row=tot_row,
                column=col_idx,
                value=f"=SUM({col_letter}2:{col_letter}{tot_row - 1})",
            )
        for col_idx in range(1, len(sum_headers) + 1):
            c = ws_sum.cell(row=tot_row, column=col_idx)
            c.fill = total_fill
            c.font = total_font
            c.border = cell_border
            c.alignment = Alignment(
                horizontal="left" if col_idx <= 3 else "center",
                vertical="center",
            )

    sum_widths = [38, 34, 24, 11, 11, 11, 9, 9, 12, 9, 12]
    for idx, w in enumerate(sum_widths, start=1):
        ws_sum.column_dimensions[get_column_letter(idx)].width = w

    # ---------------- Sheet 2: Canh_bao_loi ----------------
    ws_warn = wb.create_sheet(title="Canh_bao_loi")
    ws_warn.freeze_panes = "A2"

    warn_headers = [
        "Tên file",
        "Dòng/STT",
        "MSSV",
        "Họ tên",
        "Cột lỗi",
        "Giá trị gốc",
        "Lý do lỗi",
    ]
    ws_warn.append(warn_headers)
    for col_idx in range(1, len(warn_headers) + 1):
        c = ws_warn.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = cell_border
    ws_warn.row_dimensions[1].height = 26

    for row_idx, w in enumerate(warnings, start=2):
        stt_display = (
            w.row_stt
            if w.table_name == "Danh sách Đoàn viên"
            else f"{w.table_name} - STT {w.row_stt}"
        )
        ws_warn.append(
            [
                w.file_name,
                stt_display,
                w.mssv,
                w.full_name,
                w.column_name,
                w.original_value,
                w.reason,
            ]
        )
        for col_idx in range(1, len(warn_headers) + 1):
            c = ws_warn.cell(row=row_idx, column=col_idx)
            c.font = body_font
            c.border = cell_border
            if col_idx == 6:
                c.fill = warn_fill
            c.alignment = Alignment(
                horizontal="center" if col_idx in (2, 3, 6) else "left",
                vertical="center",
                wrap_text=True,
            )

    warn_widths = [38, 16, 14, 26, 22, 18, 44]
    for idx, w in enumerate(warn_widths, start=1):
        ws_warn.column_dimensions[get_column_letter(idx)].width = w

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def process_batch_import(
    uploaded_items: Iterable[tuple[str, bytes]],
    template_path: Path,
    config: ImportConfig,
    progress_callback: Optional[ProgressCallback] = None,
) -> BatchImportResult:
    """
    Batch-import multiple .docx (or .zip) files into clean copies of Phôi DS.docx.
    Wraps each file in try...except so a single broken file never halts the batch.
    """
    template_bytes = template_path.read_bytes()

    docx_items, pre_errors = extract_docx_inputs(uploaded_items)
    batch_result = BatchImportResult(results=list(pre_errors))

    used_output_names: set[str] = set()
    summaries: list[ClassSummary] = []
    all_warnings: list[ValidationWarning] = []

    total = len(docx_items)
    for idx, (file_name, file_bytes) in enumerate(docx_items, start=1):
        out_name = _make_output_name(file_name, used_output_names)
        try:
            extracted = extract_class_data(
                file_bytes=file_bytes,
                file_name=file_name,
                config=config,
            )
            imported_bytes = import_into_template(
                template_bytes=template_bytes,
                data=extracted,
                doc_date=config.doc_date,
            )
            warn_note = (
                f"{len(extracted.warnings)} cảnh báo dữ liệu"
                if extracted.warnings
                else "dữ liệu hợp lệ"
            )
            res = FileImportResult(
                original_name=file_name,
                output_name=out_name,
                status="Thành công",
                message=f"{extracted.summary.total_students} SV, {len(extracted.bch_rows)} BCH, {warn_note}",
                class_name=extracted.class_name_title,
                secretary_name=extracted.secretary_name,
                output_bytes=imported_bytes,
                summary=extracted.summary,
                warnings=extracted.warnings,
            )
            summaries.append(extracted.summary)
            all_warnings.extend(extracted.warnings)
        except StructureError as exc:
            res = FileImportResult(
                original_name=file_name,
                output_name="",
                status="Lỗi cấu trúc",
                message=str(exc),
            )
        except Exception as exc:
            res = FileImportResult(
                original_name=file_name,
                output_name="",
                status="Lỗi xử lý",
                message=f"Không thể xử lý file: {exc}",
            )

        batch_result.results.append(res)
        if progress_callback is not None:
            progress_callback(idx, total, res)

    excel_bytes = build_summary_excel(summaries, all_warnings)
    batch_result.excel_bytes = excel_bytes

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for res in batch_result.results:
            if res.status == "Thành công" and res.output_bytes and res.output_name:
                zf.writestr(res.output_name, res.output_bytes)
        zf.writestr("Bao_cao_tong_hop.xlsx", excel_bytes)

    batch_result.zip_bytes = zip_buf.getvalue()
    return batch_result
