from __future__ import annotations

import copy
import io
import random
import re
from datetime import date
from typing import Iterable, Optional

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

from .cleaner import clean_text, format_stat_number
from .models import ExtractedClassData

WARNING_FILL_HEX = "FFF2CC"  # Light yellow highlight for invalid data cells
_DOTS_RE = re.compile(r"\.{3,}|…+")
_DATE_LINE_RE = re.compile(r"hà\s*nội\s*,?\s*ngày.*?tháng.*?năm.*", re.IGNORECASE)
_W14_PARA_ID = "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _set_cell_shading(cell: _Cell, fill_hex: str) -> None:
    """Apply background shading to an invalid cell without modifying any other cell property."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old_shd in tcPr.findall(qn("w:shd")):
        tcPr.remove(old_shd)
    shd = parse_xml(
        f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>'
    )
    tcPr.append(shd)


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


def _iter_table_paragraphs_recursive(table: Table) -> Iterable[Paragraph]:
    seen: set[int] = set()
    for row in table.rows:
        for cell in row.cells:
            cid = id(cell._tc)
            if cid in seen:
                continue
            seen.add(cid)
            yield from _iter_cell_paragraphs_recursive(cell)


def _replace_dots_in_paragraph(paragraph: Paragraph, replacement: str) -> bool:
    """
    Replace '...' or '....' (or '…') directly on the existing run(s) of a paragraph,
    preserving 100% of the template run's <w:rPr> formatting.
    Returns True if a replacement was made.
    """
    # Fast path: placeholder sits inside a single run (true for all 14 placeholders in Phôi DS.docx)
    for run in paragraph.runs:
        if _DOTS_RE.search(run.text):
            run.text = _DOTS_RE.sub(replacement, run.text, count=1)
            return True

    # Fallback: placeholder spans across multiple adjacent runs
    full_text = "".join(r.text for r in paragraph.runs)
    m = _DOTS_RE.search(full_text)
    if not m:
        return False

    start, end = m.span()
    pos = 0
    replaced = False
    for run in paragraph.runs:
        r_len = len(run.text)
        r_start, r_end = pos, pos + r_len
        pos = r_end

        if r_end <= start or r_start >= end:
            continue

        local_s = max(0, start - r_start)
        local_e = min(r_len, end - r_start)
        if not replaced:
            run.text = run.text[:local_s] + replacement + run.text[local_e:]
            replaced = True
        else:
            run.text = run.text[:local_s] + run.text[local_e:]
    return replaced


def _write_cell_preserving_template_style(cell: _Cell, value: str) -> None:
    """
    Write text into an empty/template cell while keeping <w:tcPr> and <w:pPr>
    100% untouched and copying the template's <w:rPr> onto the new run.
    """
    p = cell.paragraphs[0]
    p_el = p._p

    # Capture template <w:rPr> from an existing run or from <w:pPr><w:rPr>
    template_rPr = None
    existing_runs = p_el.findall(qn("w:r"))
    for r_el in existing_runs:
        rPr = r_el.find(qn("w:rPr"))
        if rPr is not None:
            template_rPr = copy.deepcopy(rPr)
            break

    if template_rPr is None:
        pPr = p_el.find(qn("w:pPr"))
        if pPr is not None:
            p_rPr = pPr.find(qn("w:rPr"))
            if p_rPr is not None:
                template_rPr = copy.deepcopy(p_rPr)

    # Clear existing run elements while leaving <w:pPr> completely intact
    for r_el in existing_runs:
        p_el.remove(r_el)

    if not value:
        return

    new_r = OxmlElement("w:r")
    if template_rPr is not None:
        new_r.append(template_rPr)

    new_t = OxmlElement("w:t")
    if value.startswith(" ") or value.endswith(" "):
        new_t.set(_XML_SPACE, "preserve")
    new_t.text = value
    new_r.append(new_t)
    p_el.append(new_r)


def _refresh_para_ids(tr_el) -> None:
    """Assign fresh unique w14:paraId values on cloned <w:p> elements."""
    for p_el in tr_el.findall(f".//{qn('w:p')}"):
        if _W14_PARA_ID in p_el.attrib:
            p_el.set(_W14_PARA_ID, f"{random.randint(0x10000000, 0x7FFFFFFF):08X}")


def _populate_template_data_table(
    table: Table,
    rows_data: list[list[str]],
    expected_cols: int,
    invalid_cells: Optional[set[tuple[int, int]]] = None,
) -> None:
    """
    Import rows_data into a template data table (Row 0 header untouched).
    Clones Row 1 XML if more rows are needed; removes trailing excess rows if fewer.
    Skips column 0 (STT) when the cell uses Word auto-numbering (<w:numPr>).
    Applies light yellow shading (#FFF2CC) only to cells listed in invalid_cells.
    """
    if len(table.rows) < 2:
        raise ValueError("Bảng dữ liệu trong file phôi không có dòng dữ liệu mẫu (Row 1).")

    # Take an untouched snapshot of Row 1 before any mutation
    template_tr = copy.deepcopy(table.rows[1]._tr)

    needed = len(rows_data)
    existing = len(table.rows) - 1  # excluding Row 0 header

    if needed > existing:
        tbl_el = table._tbl
        for _ in range(needed - existing):
            cloned = copy.deepcopy(template_tr)
            _refresh_para_ids(cloned)
            tbl_el.append(cloned)
    elif needed < existing:
        # Remove excess blank rows from the end of the table
        tbl_el = table._tbl
        all_trs = list(tbl_el.findall(qn("w:tr")))
        for tr_to_remove in all_trs[1 + needed :]:
            tbl_el.remove(tr_to_remove)

    bad_set = invalid_cells or set()

    # Populate data into each row
    for row_idx, values in enumerate(rows_data, start=1):
        row = table.rows[row_idx]
        data_row_idx = row_idx - 1
        for col_idx in range(min(expected_cols, len(row.cells), len(values))):
            cell = row.cells[col_idx]
            if col_idx == 0:
                # Preserve Word auto-numbering if template uses <w:numPr>
                if cell._tc.find(f".//{qn('w:numPr')}") is not None:
                    continue
            _write_cell_preserving_template_style(cell, values[col_idx])
            if (data_row_idx, col_idx) in bad_set:
                _set_cell_shading(cell, WARNING_FILL_HEX)


def _detect_template_tables(doc: Document) -> tuple[Optional[Table], Optional[Table], Optional[Table]]:
    """Locate the 16-col member table, 14-col BCH table, and 8-col stats table in Phôi DS.docx."""
    main_tbl: Optional[Table] = None
    bch_tbl: Optional[Table] = None
    stats_tbl: Optional[Table] = None

    for table in doc.tables:
        if not table.rows:
            continue
        row0 = [clean_text(c.text).lower() for c in table.rows[0].cells]
        ncols = len(table.rows[0].cells)
        has_mssv = any("mssv" in t for t in row0)
        has_ghi_chu = any("ghi chú" in t or "ghi chu" in t for t in row0)
        has_chuc_vu = any("chức vụ" in t or "chuc vu" in t for t in row0)

        if has_mssv and has_ghi_chu and ncols >= 16 and main_tbl is None:
            main_tbl = table
            continue
        if has_mssv and has_chuc_vu and ncols >= 14 and bch_tbl is None:
            bch_tbl = table
            continue
        full_low = " ".join(clean_text(c.text).lower() for r in table.rows for c in r.cells)
        if "tổng số" in full_low and not has_mssv and stats_tbl is None:
            stats_tbl = table

    return main_tbl, bch_tbl, stats_tbl


def _replace_stats_placeholders(stats_tbl: Table, data: ExtractedClassData) -> None:
    """Replace the 8 '...' placeholders inside Table 2 (Thống kê) in-place."""
    s = data.summary
    for row in stats_tbl.rows:
        for cell in row.cells:
            cell_low = clean_text(cell.text).lower()
            if not _DOTS_RE.search(cell.text):
                continue

            if "sinh viên" in cell_low:
                val = format_stat_number(s.total_students)
            elif "sổ" in cell_low and "đoàn viên" in cell_low:
                val = format_stat_number(s.total_union_books)
            elif "nữ" in cell_low:
                val = format_stat_number(s.female_members)
            elif "nam" in cell_low:
                val = format_stat_number(s.male_members)
            elif "đảng viên" in cell_low:
                val = format_stat_number(s.party_members)
            elif "đối tượng đảng" in cell_low or "đtđ" in cell_low:
                val = format_stat_number(s.party_candidates)
            elif "thanh niên" in cell_low:
                val = format_stat_number(s.youth_non_members)
            elif "đoàn viên" in cell_low:
                val = format_stat_number(s.total_union_members)
            else:
                continue

            for p in cell.paragraphs:
                if _DOTS_RE.search(p.text):
                    _replace_dots_in_paragraph(p, val)
                    break


def _replace_date_in_paragraph(paragraph: Paragraph, doc_date: date) -> bool:
    """
    Update 'Hà Nội, ngày ... tháng ... năm ...' in-place on the template's original runs,
    preserving 100% of every run's <w:rPr> and paragraph's <w:pPr>.
    Uses Vietnamese administrative convention: day 01..09 has '0'; month 01..02 has '0',
    months 3..12 have no leading '0'.
    """
    if not _DATE_LINE_RE.search(clean_text(paragraph.text)):
        return False

    day_str = f"{doc_date.day:02d}"
    month_str = f"{doc_date.month:02d}" if doc_date.month in (1, 2) else str(doc_date.month)
    year_str = f"{doc_date.year:04d}"

    runs = paragraph.runs
    if not runs:
        return False

    # Fast path for Phôi DS.docx: ['Hà Nội, ngày ', '05', ' tháng ', '11', ' năm ', '2026']
    if (
        len(runs) >= 6
        and "ngày" in runs[0].text.lower()
        and "tháng" in runs[2].text.lower()
        and "năm" in runs[4].text.lower()
    ):
        runs[1].text = day_str
        runs[3].text = month_str
        runs[5].text = year_str
        for extra_r in runs[6:]:
            extra_r.text = ""
        return True

    # Fallback: preserve Run 0's exact <w:rPr> and write full formatted date
    runs[0].text = f"Hà Nội, ngày {day_str} tháng {month_str} năm {year_str}"
    for r in runs[1:]:
        r.text = ""
    return True


def _replace_all_placeholders(
    doc: Document,
    data: ExtractedClassData,
    main_tbl: Optional[Table],
    bch_tbl: Optional[Table],
    stats_tbl: Optional[Table],
    doc_date: Optional[date] = None,
) -> None:
    """
    Replace all 14 '...' / '....' placeholders (and optionally the header date line)
    in Phôi DS.docx without altering any formatting of the template:
      - Body paragraphs: 'Đoàn viên Thanh niên Chi đoàn ...', 'Ban Chấp hành Chi đoàn ...'
      - Table 0 & Table 4 nested tables: 'CHI ĐOÀN ...' & 'CHI ĐOÀN ....' -> UPPERCASE class name
      - Table 0 & Table 4 date lines (when doc_date is provided): 'Hà Nội, ngày ... tháng ... năm ...'
      - Table 2 (8 stats cells): '...' -> formatted stat counts (01..09 for 1..9)
      - Table 3 & Table 6 nested tables: '...' -> Chi đoàn Bí thư name
    """
    # 1. Body title paragraphs
    for p in doc.paragraphs:
        txt = clean_text(p.text)
        if not _DOTS_RE.search(txt):
            continue
        low = txt.lower()
        if "chi đoàn" in low:
            _replace_dots_in_paragraph(p, data.class_name_title)

    # 2. Stats table (8 cells)
    if stats_tbl is not None:
        _replace_stats_placeholders(stats_tbl, data)

    # 3. Remaining layout tables (Table 0, Table 3, Table 4, Table 6 — recursively through nested tables)
    skip_ids = {
        id(t._tbl) for t in (main_tbl, bch_tbl, stats_tbl) if t is not None
    }
    for table in doc.tables:
        if id(table._tbl) in skip_ids:
            continue

        # Check each top-level cell block so we know whether a standalone '...' belongs to Bí thư Chi đoàn
        seen_top: set[int] = set()
        for row in table.rows:
            for top_cell in row.cells:
                cid = id(top_cell._tc)
                if cid in seen_top:
                    continue
                seen_top.add(cid)

                block_paragraphs = list(_iter_cell_paragraphs_recursive(top_cell))
                block_low = " ".join(clean_text(p.text) for p in block_paragraphs).lower()
                is_chi_doan_sig = (
                    "chi đoàn" in block_low
                    and "bí thư" in block_low
                    and "đoàn trường" not in block_low
                )

                for p in block_paragraphs:
                    raw = clean_text(p.text)
                    if doc_date is not None and _DATE_LINE_RE.search(raw):
                        _replace_date_in_paragraph(p, doc_date)
                        continue
                    if not _DOTS_RE.search(raw):
                        continue
                    low = raw.lower()
                    if "chi đoàn" in low and "tm." not in low and "bch" not in low:
                        _replace_dots_in_paragraph(p, data.class_name_upper)
                    elif is_chi_doan_sig:
                        _replace_dots_in_paragraph(p, data.secretary_name)


def import_into_template(
    template_bytes: bytes,
    data: ExtractedClassData,
    doc_date: Optional[date] = None,
) -> bytes:
    """
    Load a clean copy of Phôi DS.docx, replace the 14 placeholders (and optional date)
    in-place, and import member + BCH rows into Table 1 & Table 5 without altering any
    template layout, margins, fonts, sizes, column widths, or borders.
    """
    doc = Document(io.BytesIO(template_bytes))
    main_tbl, bch_tbl, stats_tbl = _detect_template_tables(doc)

    if main_tbl is None:
        raise ValueError("File phôi không có bảng Danh sách Đoàn viên (16 cột).")

    # A. Replace all 14 '...' / '....' placeholders (and optional header date) in-place
    _replace_all_placeholders(doc, data, main_tbl, bch_tbl, stats_tbl, doc_date=doc_date)

    # B. Import data rows into Table 1 (16 cols) & Table 5 (14 cols)
    _populate_template_data_table(
        main_tbl,
        data.member_rows,
        expected_cols=16,
        invalid_cells=data.member_invalid_cells,
    )
    if bch_tbl is not None and data.bch_rows:
        _populate_template_data_table(
            bch_tbl,
            data.bch_rows,
            expected_cols=14,
            invalid_cells=data.bch_invalid_cells,
        )
    elif bch_tbl is not None and not data.bch_rows:
        _populate_template_data_table(bch_tbl, [], expected_cols=14)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
