from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Iterable, Optional

from docx import Document
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

from .cleaner import (
    clean_and_validate_date,
    clean_and_validate_phone,
    clean_text,
    generate_vn_phone,
    normalize_ethnicity,
    normalize_gender,
    normalize_position,
    normalize_religion,
    parse_truthy_flag,
    vn_title_case,
)
from .models import (
    ClassSummary,
    ExtractedClassData,
    ImportConfig,
    ValidationWarning,
)
from .province_mapping import normalize_birthplace

_DOT_ONLY_RE = re.compile(r"^[\.\u2026\-\s_*]+$")
_BODY_CLASS_TITLE_RE = re.compile(
    r"(?:đoàn\s+viên\s+thanh\s+niên\s+chi\s+đoàn|ban\s+chấp\s+hành\s+chi\s+đoàn)\s+(.+)",
    re.IGNORECASE,
)
_HEADER_CHI_DOAN_RE = re.compile(r"chi\s+đoàn\s+(.+)", re.IGNORECASE)
_FILENAME_PREFIX_RE = re.compile(
    r"^(?:ds\s+|danh\s+sách\s+)?chi\s+đoàn[\s_\-–]*", re.IGNORECASE
)
_KEEP_UPPER_TOKEN_RE = re.compile(r"^(?:k\d{2}|cnvl|cnc|hk|cntt|&|-|–|\d+)$", re.IGNORECASE)


class StructureError(ValueError):
    """Raised when an input .docx lacks the required 16-column Youth Union member table."""


def _iter_cell_paragraphs_recursive(cell: _Cell) -> Iterable[Paragraph]:
    for p in cell.paragraphs:
        yield p
    for nested in cell.tables:
        for row in nested.rows:
            seen: set[int] = set()
            for ncell in row.cells:
                cid = id(ncell._tc)
                if cid in seen:
                    continue
                seen.add(cid)
                yield from _iter_cell_paragraphs_recursive(ncell)


def _is_placeholder_text(text: str) -> bool:
    s = clean_text(text)
    if not s:
        return True
    if _DOT_ONLY_RE.match(s):
        return True
    # Also reject templates like "……………..."
    if len(re.sub(r"[\.\u2026\s_\-*]", "", s)) < 2:
        return True
    return False


def _clean_class_name_candidate(raw: str) -> str:
    s = clean_text(raw)
    s = re.sub(r"\*{2,}\s*$", "", s).strip()
    s = re.sub(r"^[\s:\-–_]+", "", s).strip()
    s = re.sub(r"[\s:\-_]+$", "", s).strip()
    return s


def _smart_title_from_upper(text: str) -> str:
    """Convert an ALL-CAPS class name into readable Vietnamese mixed/title case."""
    tokens = clean_text(text).split(" ")
    out: list[str] = []
    for idx, tok in enumerate(tokens):
        if _KEEP_UPPER_TOKEN_RE.match(tok):
            out.append(tok.upper())
        elif idx == 0:
            out.append(tok[0].upper() + tok[1:].lower())
        else:
            out.append(tok[0].upper() + tok[1:].lower())
    return " ".join(out)


_STOP_TITLE_LINE_RE = re.compile(
    r"^(?:năm\s+học|nhiệm\s+kỳ|\-{3,}|_{3,}|\*{2,}|\()", re.IGNORECASE
)


def _extract_class_names(doc: Document, file_name: str) -> tuple[str, str]:
    """
    Extract (class_name_title, class_name_upper) from the input document.
    Supports multi-line class names split across consecutive paragraphs before
    'Năm học' / 'Nhiệm kỳ' / '***'.
    """
    body_candidate = ""
    body_ps = list(doc.paragraphs)
    for idx, p in enumerate(body_ps):
        txt = clean_text(p.text)
        m = _BODY_CLASS_TITLE_RE.search(txt)
        if m:
            parts = [_clean_class_name_candidate(m.group(1))]
            # Collect any continuation line(s) before 'Năm học' / 'Nhiệm kỳ' / '---'
            for nxt in body_ps[idx + 1 : idx + 3]:
                ntxt = clean_text(nxt.text)
                if not ntxt or _STOP_TITLE_LINE_RE.match(ntxt):
                    break
                parts.append(_clean_class_name_candidate(ntxt))
            cand = clean_text(" ".join(pt for pt in parts if pt))
            if not _is_placeholder_text(cand):
                body_candidate = cand
                break

    header_candidate = ""
    for table in doc.tables:
        if len(table.rows) > 4 or (table.rows and len(table.rows[0].cells) > 4):
            continue
        for row in table.rows:
            for cell in row.cells:
                cell_ps = list(_iter_cell_paragraphs_recursive(cell))
                for idx, p in enumerate(cell_ps):
                    txt = clean_text(p.text)
                    low = txt.lower()
                    if "tm." in low or "bch" in low or "liên chi đoàn" in low:
                        continue
                    m = _HEADER_CHI_DOAN_RE.search(txt)
                    if m:
                        parts = [_clean_class_name_candidate(m.group(1))]
                        for nxt in cell_ps[idx + 1 : idx + 3]:
                            ntxt = clean_text(nxt.text)
                            if not ntxt or _STOP_TITLE_LINE_RE.match(ntxt):
                                break
                            parts.append(_clean_class_name_candidate(ntxt))
                        cand = clean_text(" ".join(pt for pt in parts if pt))
                        if not _is_placeholder_text(cand):
                            header_candidate = cand
                            break
                if header_candidate:
                    break
            if header_candidate:
                break
        if header_candidate:
            break

    if body_candidate:
        title_ver = (
            _smart_title_from_upper(body_candidate)
            if body_candidate.isupper()
            else body_candidate
        )
        # Prefer authentic uppercase header from Table 0 if available and at least as complete
        if header_candidate and len(header_candidate) >= len(body_candidate):
            upper_ver = header_candidate.upper()
        else:
            upper_ver = body_candidate.upper()
        return title_ver, upper_ver

    if header_candidate:
        title_ver = (
            _smart_title_from_upper(header_candidate)
            if header_candidate.isupper()
            else header_candidate
        )
        return title_ver, header_candidate.upper()

    # Fallback: derive from filename
    stem = clean_text(Path(file_name).stem)
    stem = _FILENAME_PREFIX_RE.sub("", stem).strip(" -–_")
    if not stem:
        stem = clean_text(Path(file_name).stem)
    title_ver = _smart_title_from_upper(stem) if stem.isupper() else stem
    return title_ver, stem.upper()


def _extract_secretary_from_signatures(doc: Document) -> str:
    """Find Chi đoàn Bí thư name from TM. BCH CHI ĐOÀN signature block(s)."""
    for table in doc.tables:
        if len(table.rows) > 4 or (table.rows and len(table.rows[0].cells) > 4):
            continue
        seen_top: set[int] = set()
        for row in table.rows:
            for top_cell in row.cells:
                cid = id(top_cell._tc)
                if cid in seen_top:
                    continue
                seen_top.add(cid)
                paragraphs = list(_iter_cell_paragraphs_recursive(top_cell))
                block_text = " ".join(clean_text(p.text) for p in paragraphs).lower()
                if (
                    "chi đoàn" in block_text
                    and "bí thư" in block_text
                    and "đoàn trường" not in block_text
                ):
                    bi_thu_idx: Optional[int] = None
                    for idx, p in enumerate(paragraphs):
                        if "bí thư" in clean_text(p.text).lower():
                            bi_thu_idx = idx
                            break
                    if bi_thu_idx is not None:
                        for p in paragraphs[bi_thu_idx + 1 :]:
                            cand = clean_text(p.text)
                            if not _is_placeholder_text(cand):
                                return vn_title_case(cand)
    return ""


def _detect_data_tables(doc: Document) -> tuple[Optional[Table], Optional[Table]]:
    """
    Dynamically detect:
      - Main 16-col Member List table (Row 0 has 'MSSV' and 'Ghi chú', >= 15 cols)
      - 14-col BCH table (Row 0 has 'MSSV' and 'Chức vụ', >= 13 cols)
    """
    members_table: Optional[Table] = None
    bch_table: Optional[Table] = None

    for table in doc.tables:
        if not table.rows:
            continue
        row0 = [clean_text(c.text).lower() for c in table.rows[0].cells]
        ncols = len(table.rows[0].cells)
        has_mssv = any("mssv" in t or "mã sv" in t or "mã sinh viên" in t for t in row0)
        has_ghi_chu = any("ghi chú" in t or "ghi chu" in t for t in row0)
        has_chuc_vu = any("chức vụ" in t or "chuc vu" in t for t in row0)

        if has_mssv and has_ghi_chu and ncols >= 15 and members_table is None:
            members_table = table
        elif has_mssv and has_chuc_vu and ncols >= 13 and bch_table is None:
            bch_table = table

    return members_table, bch_table


def _is_empty_row(raw_cells: list[str]) -> bool:
    """Ignore col 0 (STT); row is empty if all remaining cells are blank."""
    return all(not val for val in raw_cells[1:])


def _get_or_generate_phone(phone_cache: dict[str, str], mssv: str, full_name: str) -> str:
    key = mssv or full_name.lower()
    if key and key in phone_cache:
        return phone_cache[key]
    num = generate_vn_phone()
    if key:
        phone_cache[key] = num
    return num


def extract_class_data(
    file_bytes: bytes,
    file_name: str,
    config: ImportConfig,
) -> ExtractedClassData:
    """
    Read an input Chi đoàn .docx file, detect tables dynamically, clean all fields,
    compute the 8 statistics, and collect validation warnings (without modifying input).
    """
    doc = Document(io.BytesIO(file_bytes))
    members_table, bch_table = _detect_data_tables(doc)
    if members_table is None:
        raise StructureError(
            "Không tìm thấy bảng Danh sách Đoàn viên (16 cột chứa tiêu đề 'MSSV' và 'Ghi chú')."
        )
    if len(members_table.rows[0].cells) < 16:
        raise StructureError(
            f"Bảng Danh sách Đoàn viên chỉ có {len(members_table.rows[0].cells)} cột (yêu cầu 16 cột)."
        )

    class_name_title, class_name_upper = _extract_class_names(doc, file_name)
    secretary_name = _extract_secretary_from_signatures(doc)

    warnings: list[ValidationWarning] = []
    member_invalid_cells: set[tuple[int, int]] = set()
    bch_invalid_cells: set[tuple[int, int]] = set()
    phone_cache: dict[str, str] = {}
    summary = ClassSummary(file_name=file_name, class_name=class_name_title)

    # 1. Extract & clean 16-col Member List rows
    member_rows: list[list[str]] = []
    stt = 0
    for row in members_table.rows[1:]:
        raw = [clean_text(row.cells[i].text) for i in range(16)]
        if _is_empty_row(raw):
            continue
        stt += 1

        mssv = raw[1].replace(" ", "")
        full_name = vn_title_case(raw[2])
        gender = normalize_gender(raw[3])
        ethnicity = normalize_ethnicity(raw[4])
        religion = normalize_religion(raw[5])
        birthplace, place_err = normalize_birthplace(
            raw[6], apply_merger=config.convert_old_provinces
        )

        dob_val, dob_err = clean_and_validate_date(raw[7], is_dob=True)
        social_group = raw[8]
        join_val, join_err = clean_and_validate_date(raw[9], is_dob=False)

        has_book = parse_truthy_flag(raw[10])
        book_val = "x" if has_book else ""

        is_party = parse_truthy_flag(raw[11])
        party_val = "x" if is_party else ""

        is_dtd = parse_truthy_flag(raw[12])
        dtd_val = "x" if is_dtd else ""

        is_youth = parse_truthy_flag(raw[13])
        youth_val = "x" if is_youth else ""

        phone_val, phone_err = clean_and_validate_phone(
            raw[14],
            fill_missing=config.fill_missing_phone,
            phone_generator=lambda _m=mssv, _n=full_name: _get_or_generate_phone(phone_cache, _m, _n),
        )
        if not phone_err and phone_val and (mssv or full_name):
            phone_cache.setdefault(mssv or full_name.lower(), phone_val)

        note_val = raw[15]

        row_idx_0 = stt - 1
        if place_err:
            member_invalid_cells.add((row_idx_0, 6))
            warnings.append(
                ValidationWarning(
                    file_name=file_name,
                    table_name="Danh sách Đoàn viên",
                    row_stt=str(stt),
                    mssv=mssv,
                    full_name=full_name,
                    column_name="Nơi sinh",
                    original_value=raw[6],
                    reason=place_err,
                )
            )
        if dob_err:
            member_invalid_cells.add((row_idx_0, 7))
            warnings.append(
                ValidationWarning(
                    file_name=file_name,
                    table_name="Danh sách Đoàn viên",
                    row_stt=str(stt),
                    mssv=mssv,
                    full_name=full_name,
                    column_name="Ngày/tháng/năm sinh",
                    original_value=raw[7],
                    reason=dob_err,
                )
            )
        if join_err:
            member_invalid_cells.add((row_idx_0, 9))
            warnings.append(
                ValidationWarning(
                    file_name=file_name,
                    table_name="Danh sách Đoàn viên",
                    row_stt=str(stt),
                    mssv=mssv,
                    full_name=full_name,
                    column_name="Ngày vào Đảng, Đoàn",
                    original_value=raw[9],
                    reason=join_err,
                )
            )
        if phone_err:
            member_invalid_cells.add((row_idx_0, 14))
            warnings.append(
                ValidationWarning(
                    file_name=file_name,
                    table_name="Danh sách Đoàn viên",
                    row_stt=str(stt),
                    mssv=mssv,
                    full_name=full_name,
                    column_name="Số điện thoại",
                    original_value=raw[14],
                    reason=phone_err,
                )
            )

        member_rows.append(
            [
                str(stt),
                mssv,
                full_name,
                gender,
                ethnicity,
                religion,
                birthplace,
                dob_val,
                social_group,
                join_val,
                book_val,
                party_val,
                dtd_val,
                youth_val,
                phone_val,
                note_val,
            ]
        )

        # Accumulate 8 statistics
        # A student is a Đoàn viên if they have a join date, a union book, or aren't exclusively Thanh niên.
        is_union_member = bool(raw[9]) or has_book or (not is_youth)
        summary.total_students += 1
        if is_union_member:
            summary.total_union_members += 1
            if gender == "Nữ":
                summary.female_members += 1
            elif gender == "Nam":
                summary.male_members += 1

        if is_youth:
            summary.youth_non_members += 1
        if has_book:
            summary.total_union_books += 1
        if is_party:
            summary.party_members += 1
        if is_dtd:
            summary.party_candidates += 1

    # 2. Extract & clean 14-col BCH Executive Committee rows
    bch_rows: list[list[str]] = []
    if bch_table is not None and len(bch_table.rows[0].cells) >= 14:
        bch_stt = 0
        for row in bch_table.rows[1:]:
            raw = [clean_text(row.cells[i].text) for i in range(14)]
            if _is_empty_row(raw):
                continue
            bch_stt += 1

            mssv = raw[1].replace(" ", "")
            full_name = vn_title_case(raw[2])
            gender = normalize_gender(raw[3])
            ethnicity = normalize_ethnicity(raw[4])
            religion = normalize_religion(raw[5])
            birthplace, place_err = normalize_birthplace(
                raw[6], apply_merger=config.convert_old_provinces
            )

            dob_val, dob_err = clean_and_validate_date(raw[7], is_dob=True)
            social_group = raw[8]
            join_val, join_err = clean_and_validate_date(raw[9], is_dob=False)
            party_val = "x" if parse_truthy_flag(raw[10]) else ""
            dtd_val = "x" if parse_truthy_flag(raw[11]) else ""
            phone_val, phone_err = clean_and_validate_phone(
                raw[12],
                fill_missing=config.fill_missing_phone,
                phone_generator=lambda _m=mssv, _n=full_name: _get_or_generate_phone(phone_cache, _m, _n),
            )
            position_val = normalize_position(raw[13])

            bch_row_idx_0 = bch_stt - 1
            if place_err:
                bch_invalid_cells.add((bch_row_idx_0, 6))
                warnings.append(
                    ValidationWarning(
                        file_name=file_name,
                        table_name="Lý lịch BCH",
                        row_stt=str(bch_stt),
                        mssv=mssv,
                        full_name=full_name,
                        column_name="Nơi sinh",
                        original_value=raw[6],
                        reason=place_err,
                    )
                )
            if dob_err:
                bch_invalid_cells.add((bch_row_idx_0, 7))
                warnings.append(
                    ValidationWarning(
                        file_name=file_name,
                        table_name="Lý lịch BCH",
                        row_stt=str(bch_stt),
                        mssv=mssv,
                        full_name=full_name,
                        column_name="Ngày/tháng/năm sinh",
                        original_value=raw[7],
                        reason=dob_err,
                    )
                )
            if join_err:
                bch_invalid_cells.add((bch_row_idx_0, 9))
                warnings.append(
                    ValidationWarning(
                        file_name=file_name,
                        table_name="Lý lịch BCH",
                        row_stt=str(bch_stt),
                        mssv=mssv,
                        full_name=full_name,
                        column_name="Ngày vào Đảng, Đoàn",
                        original_value=raw[9],
                        reason=join_err,
                    )
                )
            if phone_err:
                bch_invalid_cells.add((bch_row_idx_0, 12))
                warnings.append(
                    ValidationWarning(
                        file_name=file_name,
                        table_name="Lý lịch BCH",
                        row_stt=str(bch_stt),
                        mssv=mssv,
                        full_name=full_name,
                        column_name="Số điện thoại",
                        original_value=raw[12],
                        reason=phone_err,
                    )
                )

            bch_rows.append(
                [
                    str(bch_stt),
                    mssv,
                    full_name,
                    gender,
                    ethnicity,
                    religion,
                    birthplace,
                    dob_val,
                    social_group,
                    join_val,
                    party_val,
                    dtd_val,
                    phone_val,
                    position_val,
                ]
            )

            # Fallback for secretary_name from BCH table if signature was blank
            if not secretary_name:
                pos_low = position_val.lower()
                if "bí thư" in pos_low and "phó" not in pos_low:
                    secretary_name = full_name

    summary.secretary_name = secretary_name
    summary.warning_count = len(warnings)

    return ExtractedClassData(
        file_name=file_name,
        class_name_title=class_name_title,
        class_name_upper=class_name_upper,
        secretary_name=secretary_name,
        member_rows=member_rows,
        bch_rows=bch_rows,
        summary=summary,
        warnings=warnings,
        member_invalid_cells=member_invalid_cells,
        bch_invalid_cells=bch_invalid_cells,
    )
