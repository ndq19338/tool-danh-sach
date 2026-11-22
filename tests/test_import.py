from __future__ import annotations

import io
import re
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree
from openpyxl import load_workbook

PROJECT_CON_DIR = Path(__file__).resolve().parent.parent
WORKSPACE_ROOT = PROJECT_CON_DIR.parent
TEMPLATE_PATH = PROJECT_CON_DIR / "Phôi DS.docx"

sys.path.insert(0, str(PROJECT_CON_DIR))

from core.batch_processor import process_batch_import
from core.models import ImportConfig
from core.template_importer import WARNING_FILL_HEX


def _cell_fill_hex(cell) -> str | None:
    tcPr = cell._tc.find(qn("w:tcPr"))
    if tcPr is None:
        return None
    shd = tcPr.find(qn("w:shd"))
    return shd.get(qn("w:fill")) if shd is not None else None

_DOTS_RE = re.compile(r"\.{3,}|…+")


def _all_paragraphs_text(doc: Document) -> list[str]:
    texts = [p.text for p in doc.paragraphs]

    def walk_tbl(tbl):
        for r in tbl.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    texts.append(p.text)
                for nt in c.tables:
                    walk_tbl(nt)

    for t in doc.tables:
        walk_tbl(t)
    return texts


def _verify_phoi_untouched(
    template_bytes: bytes,
    output_bytes: bytes,
    expected_members: int,
    expected_bch: int,
    expected_main_yellow: set[tuple[int, int]] | None = None,
) -> Document:
    tpl = Document(io.BytesIO(template_bytes))
    out = Document(io.BytesIO(output_bytes))

    # 1. ZERO remaining '...' or '....' anywhere in the document!
    all_texts = _all_paragraphs_text(out)
    leftover = [t for t in all_texts if _DOTS_RE.search(t)]
    assert not leftover, f"Found leftover dot placeholders in output: {leftover}"

    # 2. Exact sections & page setup identical to Phôi DS.docx
    assert len(out.sections) == len(tpl.sections)
    for i, (s_tpl, s_out) in enumerate(zip(tpl.sections, out.sections)):
        assert s_out.orientation == s_tpl.orientation, f"Section {i} orientation changed"
        assert s_out.page_width == s_tpl.page_width, f"Section {i} width changed"
        assert s_out.page_height == s_tpl.page_height, f"Section {i} height changed"
        assert s_out.left_margin == s_tpl.left_margin, f"Section {i} left margin changed"
        assert s_out.right_margin == s_tpl.right_margin, f"Section {i} right margin changed"
        assert s_out.top_margin == s_tpl.top_margin, f"Section {i} top margin changed"
        assert s_out.bottom_margin == s_tpl.bottom_margin, f"Section {i} bottom margin changed"

    # 3. Exact same table count, tblPr, and tblGrid on ALL tables
    assert len(out.tables) == len(tpl.tables)
    for ti, (t_tpl, t_out) in enumerate(zip(tpl.tables, out.tables)):
        assert etree.tostring(t_out._tbl.tblPr) == etree.tostring(t_tpl._tbl.tblPr), f"Table {ti} tblPr modified!"
        g_tpl = t_tpl._tbl.find(qn("w:tblGrid"))
        g_out = t_out._tbl.find(qn("w:tblGrid"))
        assert etree.tostring(g_out) == etree.tostring(g_tpl), f"Table {ti} tblGrid modified!"

    # 4. Exact Row 0 header untouched on Table 1 & Table 5
    assert etree.tostring(out.tables[1].rows[0]._tr) == etree.tostring(tpl.tables[1].rows[0]._tr), "Table 1 header modified!"
    assert etree.tostring(out.tables[5].rows[0]._tr) == etree.tostring(tpl.tables[5].rows[0]._tr), "Table 5 header modified!"

    # 5. Exact row counts & preserved w:numPr on STT col + exact yellow shading ONLY on invalid cells
    yellow_expected = expected_main_yellow or set()
    t_main = out.tables[1]
    assert len(t_main.rows) == 1 + expected_members, f"Main rows={len(t_main.rows)}, expected={1 + expected_members}"
    for ri, row in enumerate(t_main.rows[1:], start=1):
        assert row.cells[0]._tc.find(f".//{qn('w:numPr')}") is not None, f"Main row {ri} lost w:numPr"
        for ci, cell in enumerate(row.cells):
            fill = _cell_fill_hex(cell)
            if (ri, ci) in yellow_expected:
                assert fill == WARNING_FILL_HEX, f"Main R{ri}C{ci} expected yellow fill {WARNING_FILL_HEX}, got {fill}"
            else:
                assert fill != WARNING_FILL_HEX, f"Main R{ri}C{ci} should NOT have yellow fill"
            # Check non-empty runs keep 10pt Times New Roman from template
            for p in cell.paragraphs:
                for r in p.runs:
                    if r.text.strip():
                        assert r.font.name == "Times New Roman"
                        assert r.font.size.pt == 10.0

    t_bch = out.tables[5]
    assert len(t_bch.rows) == 1 + expected_bch, f"BCH rows={len(t_bch.rows)}, expected={1 + expected_bch}"
    for ri, row in enumerate(t_bch.rows[1:], start=1):
        assert row.cells[0]._tc.find(f".//{qn('w:numPr')}") is not None, f"BCH row {ri} lost w:numPr"

    return out


def test_samples_and_shrink_import() -> None:
    template_bytes = TEMPLATE_PATH.read_bytes()
    samples_dir = WORKSPACE_ROOT / "samples"
    sample_files = sorted(p for p in samples_dir.glob("*.docx") if not p.name.startswith("~$"))
    assert len(sample_files) == 2

    # Include a file with < 43 students (31 students) to test shrinking rows too!
    polyme01 = WORKSPACE_ROOT / "CNVL Polyme vaf Compozit 01.docx"
    test_files = list(sample_files)
    if polyme01.exists():
        test_files.append(polyme01)

    inputs = [(p.name, p.read_bytes()) for p in test_files]
    batch = process_batch_import(inputs, TEMPLATE_PATH, ImportConfig(fill_missing_phone=True))
    print(f"[Import Batch] total={batch.total_files}, success={batch.success_count}, errors={batch.error_count}, warnings={batch.total_warnings}")
    assert batch.success_count == len(test_files)
    assert batch.error_count == 0

    with zipfile.ZipFile(io.BytesIO(batch.zip_bytes), "r") as zf:
        names = set(zf.namelist())
        print("[ZIP entries]", sorted(names))
        assert "Bao_cao_tong_hop.xlsx" in names
        assert "Danh sách chi đoàn_imported.docx" in names
        assert "CN thép và HK - K69_imported.docx" in names

        # 1. Verify Danh sách chi đoàn_imported.docx (46 students -> cloned 3 rows, R41C7 yellow!)
        doc_raw = _verify_phoi_untouched(
            template_bytes,
            zf.read("Danh sách chi đoàn_imported.docx"),
            expected_members=46,
            expected_bch=3,
            expected_main_yellow={(41, 7)},
        )
        full_raw = "\n".join(_all_paragraphs_text(doc_raw))
        assert "CHI ĐOÀN KỸ THUẬT IN 01 – K68" in full_raw
        assert "Đoàn viên Thanh niên Chi đoàn Kỹ Thuật In 01 – K68" in full_raw
        assert "Ban Chấp hành Chi đoàn Kỹ Thuật In 01 – K68" in full_raw
        assert "Phạm Anh Tuấn" in full_raw
        # Verify Table 1 cleaned data & 'x'/'' book format & date format
        t1 = doc_raw.tables[1]
        assert t1.rows[1].cells[10].text.strip() == ""   # 'Không' -> ''
        assert t1.rows[2].cells[10].text.strip() == "x"  # 'Có' -> 'x'
        assert t1.rows[6].cells[5].text.strip() == "Không"
        assert t1.rows[9].cells[5].text.strip() == "Không"
        assert t1.rows[11].cells[3].text.strip() == "Nam"
        assert t1.rows[11].cells[4].text.strip() == "Kinh"
        assert t1.rows[41].cells[7].text.strip() == "28/02/2025"  # invalid DOB preserved
        # Verify Table 2 stats populated
        st_raw = " | ".join(c.text.replace("\n", " ") for c in doc_raw.tables[2].rows[0].cells)
        print("[Stats Danh sách chi đoàn_imported]", st_raw)
        assert "sinh viên: 46" in st_raw
        assert "Đoàn viên: 38" in st_raw
        assert "Viên nữ: 15" in st_raw
        assert "viên nam: 31" in st_raw

        # 2. Verify CN thép và HK - K69_imported.docx (47 students -> 4 yellow invalid cells!)
        doc_k69 = _verify_phoi_untouched(
            template_bytes,
            zf.read("CN thép và HK - K69_imported.docx"),
            expected_members=47,
            expected_bch=3,
            expected_main_yellow={(7, 7), (17, 14), (21, 14), (22, 7)},
        )
        full_k69 = "\n".join(_all_paragraphs_text(doc_k69))
        assert "CHI ĐOÀN CÔNG NGHỆ THÉP VÀ HỢP KIM – K69" in full_k69
        assert "Đào Minh Duy" in full_k69

        # 3. Verify CNVL Polyme vaf Compozit 01_imported.docx (31 students < 43 -> deleted 12 excess rows, female='01'!)
        if polyme01.exists():
            doc_p = _verify_phoi_untouched(
                template_bytes,
                zf.read("CNVL Polyme vaf Compozit 01_imported.docx"),
                expected_members=31,
                expected_bch=3,
            )
            st_p = " | ".join(c.text.replace("\n", " ") for c in doc_p.tables[2].rows[0].cells)
            print("[Stats Polyme 01_imported]", st_p)
            assert "sinh viên: 31" in st_p
            assert "Viên nữ: 01" in st_p  # leading '0' for 1..9!
            assert "Trịnh Quốc Hưng" in "\n".join(_all_paragraphs_text(doc_p))

        # 4. Verify Excel report
        wb = load_workbook(io.BytesIO(zf.read("Bao_cao_tong_hop.xlsx")))
        assert wb.sheetnames == ["Tong_hop", "Canh_bao_loi"]
        ws_sum = list(wb["Tong_hop"].iter_rows(values_only=True))
        ws_warn = list(wb["Canh_bao_loi"].iter_rows(values_only=True))
        print("\n[Excel Tong_hop]")
        for r in ws_sum:
            print(" ", r)
        print("\n[Excel Canh_bao_loi]")
        for r in ws_warn:
            print(" ", r)
        assert len(ws_warn) - 1 == 5  # 4 from K69 + 1 from Danh sách chi đoàn


def test_zip_and_all_workspace_files() -> None:
    """Test ZIP input + error isolation across ALL 10 real workspace files."""
    template_bytes = TEMPLATE_PATH.read_bytes()
    real_files = sorted(p for p in WORKSPACE_ROOT.glob("*.docx") if not p.name.startswith("~$"))
    if not real_files:
        return

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w") as zf:
        for p in real_files:
            zf.write(p, arcname=f"lop/{p.name}")

    batch = process_batch_import(
        [("tat_ca_chi_doan.zip", zip_buf.getvalue())],
        TEMPLATE_PATH,
        ImportConfig(fill_missing_phone=True),
    )
    print(f"\n[All Workspace Files via ZIP] total={batch.total_files}, success={batch.success_count}, errors={batch.error_count}, warnings={batch.total_warnings}")
    for r in batch.results:
        s = r.summary
        stat_str = (
            f"SV={s.total_students} ĐV={s.total_union_members} Nữ={s.female_members} Nam={s.male_members} TN={s.youth_non_members}"
            if s
            else ""
        )
        print(f"  - {r.original_name}: {r.status} | {stat_str} | Chi đoàn={r.class_name!r} | Bí thư={r.secretary_name!r}")
        if r.status == "Thành công" and r.output_bytes and s:
            assert s.total_union_members > 0, f"{r.original_name} counted 0 Đoàn viên!"
            assert (s.male_members + s.female_members) > 0, f"{r.original_name} counted 0 Nam/Nữ!"
            doc = Document(io.BytesIO(r.output_bytes))
            leftover = [t for t in _all_paragraphs_text(doc) if _DOTS_RE.search(t)]
            assert not leftover, f"{r.original_name} still has dots: {leftover}"
            if "TẠO HÌNH" in r.original_name.upper():
                assert s.total_union_members == 48 and s.female_members == 3 and s.male_members == 45
                st_txt = " | ".join(c.text.replace("\n", " ") for c in doc.tables[2].rows[0].cells)
                assert "Viên nữ: 03" in st_txt and "viên nam: 45" in st_txt, f"Unexpected stats: {st_txt}"
    assert batch.success_count >= len(real_files) - 1


if __name__ == "__main__":
    test_samples_and_shrink_import()
    test_zip_and_all_workspace_files()
    print("\nALL PROJECT CON IMPORT TESTS PASSED SUCCESSFULLY!")
