from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from core.batch_processor import process_batch_import
from core.models import BatchImportResult, FileImportResult, ImportConfig

APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "Phôi DS.docx"

st.set_page_config(
    page_title="Import hàng loạt Danh sách Chi đoàn vào Phôi chuẩn",
    page_icon="📥",
    layout="wide",
)


def _render_sidebar() -> ImportConfig:
    st.sidebar.header("⚙️ Cấu hình Import vào Phôi")

    if TEMPLATE_PATH.exists():
        st.sidebar.success(f"✅ File phôi sẵn sàng:\n`{TEMPLATE_PATH.name}`")
    else:
        st.sidebar.error(f"❌ Không tìm thấy file phôi:\n`{TEMPLATE_PATH}`")

    doc_date = st.sidebar.date_input(
        "Ngày/tháng/năm điền vào văn bản",
        value=date.today(),
        format="DD/MM/YYYY",
        help="Thay thế trực tiếp dòng 'Hà Nội, ngày ... tháng ... năm ...' trong Phôi DS.docx (giữ nguyên font/cỡ chữ gốc).",
    )

    convert_old_provinces = st.sidebar.checkbox(
        "Tự động quy đổi tỉnh/thành cũ sang tên mới sau sáp nhập",
        value=True,
        help="Quy đổi cột Nơi sinh (bảng 16 cột & bảng BCH 14 cột) về danh sách 34 tỉnh/thành phố mới sau sáp nhập (Nghị quyết 202/2025/QH15) và các tỉnh cũ như Hà Tây -> Hà Nội.",
    )

    fill_missing_phone = st.sidebar.checkbox(
        "Tự động điền SĐT ngẫu nhiên chuẩn VN khi bị trống",
        value=True,
        help="Nếu ô SĐT trong file đầu vào bị trống, tự động sinh số 10 chữ số chuẩn đầu số nhà mạng VN.",
    )

    cfg = ImportConfig(
        doc_date=doc_date,
        convert_old_provinces=convert_old_provinces,
        fill_missing_phone=fill_missing_phone,
    )

    st.sidebar.divider()
    if cfg.formatted_date_vn:
        st.sidebar.caption(f"📅 Dòng ngày xuất ra:\n**{cfg.formatted_date_vn}**")
    st.sidebar.caption(
        "🔒 **Nguyên tắc bảo toàn Phôi:**\n"
        "- Giữ nguyên 100% khổ giấy, lề trang, section, font chữ, cỡ chữ, độ rộng cột, viền bảng, chiều cao dòng và auto-numbering `STT` của `Phôi DS.docx`.\n"
        "- Chỉ cập nhật dòng ngày tháng năm, thay thế đúng 14 vị trí dấu `...` / `....` và điền dữ liệu vào Bảng 16 cột & Bảng BCH 14 cột.\n"
        "- Các ô bị lỗi dữ liệu (ngày sinh, SĐT hoặc Nơi sinh không nhận diện được) sẽ được **bôi vàng nhạt** trong file Word và ghi chi tiết vào `Bao_cao_tong_hop.xlsx`."
    )
    return cfg


def _render_results(batch: BatchImportResult) -> None:
    st.divider()
    st.subheader("📦 Kết quả Import hàng loạt")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng số file", batch.total_files)
    c2.metric("Thành công", batch.success_count)
    c3.metric("Lỗi cấu trúc / xử lý", batch.error_count)
    c4.metric("Cảnh báo dữ liệu (Excel)", batch.total_warnings)

    zip_filename = f"Ket_qua_import_phoi_{date.today().strftime('%Y%m%d')}.zip"
    col_dl1, col_dl2 = st.columns([2, 1])
    with col_dl1:
        st.download_button(
            label="⬇️ Tải xuống trọn bộ (.zip gồm các file _imported.docx & Bao_cao_tong_hop.xlsx)",
            data=batch.zip_bytes,
            file_name=zip_filename,
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )
    with col_dl2:
        st.download_button(
            label="📊 Chỉ tải Bao_cao_tong_hop.xlsx",
            data=batch.excel_bytes,
            file_name="Bao_cao_tong_hop.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    tab_status, tab_sum, tab_warn = st.tabs(
        [
            f"📋 Trạng thái từng file ({batch.total_files})",
            f"📊 Bảng Tổng hợp ({batch.success_count})",
            f"⚠️ Cảnh báo lỗi dữ liệu ({batch.total_warnings})",
        ]
    )

    with tab_status:
        rows = [
            {
                "Tên file gốc": r.original_name,
                "Tên Chi đoàn trích xuất": r.class_name or "—",
                "Bí thư Chi đoàn": r.secretary_name or "—",
                "File đầu ra": r.output_name or "—",
                "Trạng thái": "✅ Thành công" if r.status == "Thành công" else f"❌ {r.status}",
                "Chi tiết": r.message,
            }
            for r in batch.results
        ]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    with tab_sum:
        sum_rows = [
            {
                "Tên file": r.summary.file_name,
                "Tên Chi đoàn": r.summary.class_name,
                "Bí thư Chi đoàn": r.summary.secretary_name,
                "Tổng SV": r.summary.total_students,
                "Tổng ĐV": r.summary.total_union_members,
                "Số sổ ĐV": r.summary.total_union_books,
                "Nữ": r.summary.female_members,
                "Nam": r.summary.male_members,
                "Đảng viên": r.summary.party_members,
                "ĐTĐ": r.summary.party_candidates,
                "Thanh niên": r.summary.youth_non_members,
            }
            for r in batch.results
            if r.summary is not None
        ]
        if sum_rows:
            df_sum = pd.DataFrame(sum_rows)
            totals = {
                "Tên file": "TỔNG CỘNG",
                "Tên Chi đoàn": "",
                "Bí thư Chi đoàn": "",
                "Tổng SV": int(df_sum["Tổng SV"].sum()),
                "Tổng ĐV": int(df_sum["Tổng ĐV"].sum()),
                "Số sổ ĐV": int(df_sum["Số sổ ĐV"].sum()),
                "Nữ": int(df_sum["Nữ"].sum()),
                "Nam": int(df_sum["Nam"].sum()),
                "Đảng viên": int(df_sum["Đảng viên"].sum()),
                "ĐTĐ": int(df_sum["ĐTĐ"].sum()),
                "Thanh niên": int(df_sum["Thanh niên"].sum()),
            }
            st.dataframe(
                pd.concat([df_sum, pd.DataFrame([totals])], ignore_index=True),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("Chưa có file nào import thành công.")

    with tab_warn:
        warn_rows = [
            {
                "Tên file": w.file_name,
                "Dòng/STT": (
                    w.row_stt
                    if w.table_name == "Danh sách Đoàn viên"
                    else f"{w.table_name} - STT {w.row_stt}"
                ),
                "MSSV": w.mssv,
                "Họ tên": w.full_name,
                "Cột lỗi": w.column_name,
                "Giá trị gốc": w.original_value,
                "Lý do lỗi": w.reason,
            }
            for r in batch.results
            for w in r.warnings
        ]
        if warn_rows:
            st.dataframe(pd.DataFrame(warn_rows), use_container_width=True, hide_index=True)
        else:
            st.success("🎉 Không phát hiện lỗi dữ liệu nào (ngày sinh, SĐT, Nơi sinh đều hợp lệ)!")


def main() -> None:
    config = _render_sidebar()

    st.title("📥 Import hàng loạt Danh sách Chi đoàn vào `Phôi DS.docx`")
    st.markdown(
        "Tải lên **nhiều file `.docx`** hoặc **1 file `.zip`** chứa danh sách các Chi đoàn. "
        "Ứng dụng sẽ đọc & làm sạch dữ liệu rồi **import trực tiếp vào bản sao của `Phôi DS.docx`**, "
        "bảo toàn tuyệt đối 100% định dạng gốc của phôi."
    )

    uploaded_files = st.file_uploader(
        "Chọn các file Word đầu vào (.docx) hoặc file nén (.zip)",
        type=["docx", "zip"],
        accept_multiple_files=True,
    )

    run_btn = st.button(
        "🚀 Bắt đầu Import vào Phôi",
        type="primary",
        disabled=(not uploaded_files or not TEMPLATE_PATH.exists()),
    )

    if run_btn and uploaded_files and TEMPLATE_PATH.exists():
        raw_items = [(uf.name, uf.getvalue()) for uf in uploaded_files]
        progress_bar = st.progress(0.0, text="Đang chuẩn bị danh sách file...")
        status_log = st.container()

        def on_progress(current: int, total: int, res: FileImportResult) -> None:
            ratio = current / max(total, 1)
            badge = "✅ Thành công" if res.status == "Thành công" else f"❌ {res.status}"
            progress_bar.progress(
                ratio,
                text=f"Đang import ({current}/{total}): {res.original_name} — {badge}",
            )
            with status_log:
                if res.status == "Thành công":
                    st.write(
                        f"✅ **{res.original_name}** → `{res.output_name}` "
                        f"*(Chi đoàn: {res.class_name} | Bí thư: {res.secretary_name or 'trống'} — {res.message})*"
                    )
                else:
                    st.error(f"❌ **{res.original_name}** — {res.status}: {res.message}")

        batch = process_batch_import(
            uploaded_items=raw_items,
            template_path=TEMPLATE_PATH,
            config=config,
            progress_callback=on_progress,
        )
        progress_bar.progress(1.0, text=f"Hoàn tất import {batch.total_files} file!")
        st.session_state["last_import_batch"] = batch

    if "last_import_batch" in st.session_state:
        _render_results(st.session_state["last_import_batch"])


if __name__ == "__main__":
    main()
