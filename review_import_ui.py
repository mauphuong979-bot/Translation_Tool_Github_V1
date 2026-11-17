import hashlib

import streamlit as st

from review_roundtrip import apply_ai_translations, preview_ai_translations
from review_filenames import review_output_filename


STATUS_LABELS = {
    "ready": "Sẵn sàng", "blank": "Ô trống", "blocked": "Cần xử lý thủ công",
    "missing": "Thiếu bảng", "applied": "Đã áp dụng", "unchanged": "Không đổi",
}
IMPORT_PREVIEW_VERSION = 5


def _render_preview_rows(rows, fingerprint):
    st.dataframe([
        {"Mã": row["review_id"],
         "Trạng thái": "Track Changes có sẵn" if row["status"] == "blocked" and row.get("source_has_revisions") else STATUS_LABELS[row["status"]],
         "Nội dung gốc": row["original"], "Bản dịch AI": row["translation"], "Ghi chú": row["reason"]}
        for row in rows
    ], hide_index=True, width="stretch", height=min(404, 44 + 72 * len(rows)), row_height=72,
        column_config={
            "Mã": st.column_config.TextColumn(width=140, pinned=True),
            "Trạng thái": st.column_config.TextColumn(width=155),
            "Nội dung gốc": st.column_config.TextColumn(width=210),
            "Bản dịch AI": st.column_config.TextColumn(width=230),
            "Ghi chú": st.column_config.TextColumn(width=380, help="Kết quả kiểm tra và lý do. Ghi chú đầy đủ có ở bên dưới bảng."),
        })
    existing = [row["review_id"] for row in rows if row.get("source_has_revisions") and row["status"] != "applied"]
    if existing:
        st.info("Báo cáo nguồn đã có sẵn Track Changes ở các mã: " + ", ".join(dict.fromkeys(existing))
                + ". Các đoạn này có thể đã hiện bản dịch dù không được chọn để nhập từ AI. Lần nhập này giữ nguyên các thay đổi đó.")
    with st.container(border=True):
        st.markdown("**Chi tiết từng mã và ghi chú**")
        default_index = next((index for index, row in enumerate(rows) if row.get("source_has_revisions") and row["status"] == "blocked"),
                             next((index for index, row in enumerate(rows) if row["status"] == "blocked"), 0))
        identity_column, note_column = st.columns([1, 2], gap="small")
        with identity_column:
            selected_index = st.selectbox("Chọn mã để xem", range(len(rows)), index=default_index,
                                          format_func=lambda index: f"{rows[index]['review_id']} · {STATUS_LABELS[rows[index]['status']]}",
                                          key=f"ai_detail_{fingerprint}")
        row = rows[selected_index]
        with note_column:
            st.markdown("**Ghi chú (đầy đủ)**")
            st.info(row["reason"] or "Sẵn sàng nhập. Chỉ các mã được chọn và xác nhận mới được áp dụng.")
        with st.expander("So sánh nội dung gốc và bản dịch AI", expanded=False):
            source_column, translation_column = st.columns(2, gap="small")
            with source_column:
                st.markdown("**Nội dung gốc**")
                st.text(row["original"])
            with translation_column:
                st.markdown("**Bản dịch AI (trong bảng AI_Form)**")
                st.text(row["translation"] or "AI chưa điền bản dịch.")
            if row.get("source_has_revisions"):
                st.markdown("**Nội dung đang có trong báo cáo nguồn (trước lần nhập này)**")
                st.text(row.get("source_current_text", ""))


def _render_counts(counts, total):
    attention = counts.get("blocked", 0) + counts.get("missing", 0)
    skipped = counts.get("blank", 0) + counts.get("unchanged", 0)
    cards = (("Tổng số", total), ("Sẵn sàng", counts.get("ready", 0)),
             ("Cần chú ý", attention), ("Trống / không đổi", skipped))
    st.markdown('<div class="ai-review-summary">' + ''.join(
        f'<div class="ai-review-stat"><span>{label}</span><strong>{int(value)}</strong></div>'
        for label, value in cards) + '</div>', unsafe_allow_html=True)


def render_ai_import():
    with st.expander("📤 Nhập bản dịch từ AI", expanded=False), st.container(key="ai_import_workspace", gap="small"):
        st.markdown('<div class="ai-review-steps"><b>01 · Tải file</b><span>02 · Rà soát mã</span><span>03 · Áp dụng & tải về</span></div>', unsafe_allow_html=True)
        st.caption("Dùng đúng cặp báo cáo và bảng AI_Form đã được AI điền. Chỉ các mã hợp lệ và được chọn mới được áp dụng; bạn vẫn cần rà soát chất lượng bản dịch.")
        current_report = st.session_state.get("processed_outputs", {}).get("review")
        use_current = False
        report_column, worksheet_column = st.columns(2, gap="small")
        with report_column:
            if current_report is not None:
                use_current = st.checkbox("Dùng báo cáo Comments + Track Changes vừa xử lý trong phiên này", value=True, key="ai_use_current_report")
            if use_current:
                report_data = current_report.getvalue() if hasattr(current_report, "getvalue") else current_report
                report_name = review_output_filename(st.session_state.get("processed_filename") or "Report.docx", "review")
                st.caption("Báo cáo đi kèm: " + report_name)
            else:
                uploaded_report = st.file_uploader("Báo cáo Comments + Track Changes đi kèm", type=["docx"], key="ai_import_report")
                report_data = uploaded_report.getvalue() if uploaded_report is not None else None
                report_name = uploaded_report.name if uploaded_report is not None else "Report.docx"
        with worksheet_column:
            worksheet = st.file_uploader("Bảng AI_Form đã được AI điền", type=["docx"], key="ai_import_worksheet")
        if report_data is None or worksheet is None:
            st.info("Chọn đủ hai file để kiểm tra bản dịch. Bảng AI cũ chưa có mã cần được tạo lại.")
            return
        worksheet_data = worksheet.getvalue()
        fingerprint = hashlib.sha256(report_data + b"\0" + worksheet_data).hexdigest()
        if st.button("Kiểm tra bản dịch AI", key="ai_check_translations", width="stretch"):
            st.session_state.pop("ai_import_preview", None)
            st.session_state.pop("ai_import_result", None)
            try:
                with st.spinner("Đang kiểm tra mã, nội dung nguồn và vị trí comment..."):
                    preview = preview_ai_translations(report_data, worksheet_data)
                st.session_state.ai_import_preview = {"fingerprint": fingerprint, "version": IMPORT_PREVIEW_VERSION, "data": preview}
            except Exception as error:
                st.error(f"Không nhập được các file này: {error}")
        saved_preview = st.session_state.get("ai_import_preview", {})
        if saved_preview.get("fingerprint") != fingerprint:
            return
        if saved_preview.get("version") != IMPORT_PREVIEW_VERSION:
            st.info("Bấm Kiểm tra bản dịch AI lần nữa để làm mới kết quả kiểm tra.")
            return
        preview = saved_preview["data"]
        _render_counts(preview["counts"], len(preview["rows"]))
        _render_preview_rows(preview["rows"], fingerprint)
        ready_ids = [row["review_id"] for row in preview["rows"] if row["status"] == "ready"]
        if not ready_ids:
            st.info("Chưa có bản dịch nào sẵn sàng để áp dụng. Các mục trống, bị chặn hoặc thiếu bảng giữ nguyên nội dung và comment gốc.")
            return
        selection_column, apply_column = st.columns([3, 2], gap="small")
        with selection_column:
            selected_ids = st.multiselect("Các mã sẽ áp dụng", ready_ids, default=ready_ids, key=f"ai_selection_{fingerprint}")
            st.caption("Ứng dụng tạo một bản báo cáo mới có Track Changes và chỉ xóa comment của các mã đã áp dụng. Reject khôi phục nội dung nhưng không khôi phục comment đã xóa; hãy giữ báo cáo gốc để dự phòng.")
        selection_key = hashlib.sha256("\n".join(sorted(selected_ids)).encode("utf-8")).hexdigest()
        with apply_column:
            confirmed = st.checkbox("Tôi đã rà soát các bản dịch được chọn và đồng ý xóa comment của chúng trong bản mới", key=f"ai_confirm_{fingerprint}_{selection_key}")
            apply_clicked = st.button("Áp dụng bản dịch AI", key="ai_apply_translations", disabled=not confirmed or not selected_ids, width="stretch")
        if apply_clicked:
            st.session_state.pop("ai_import_result", None)
            try:
                with st.spinner("Đang tạo báo cáo mới có Track Changes..."):
                    output, summary = apply_ai_translations(report_data, worksheet_data, selected_ids)
                if set(summary["applied_ids"]) != set(selected_ids):
                    raise ValueError("Các mã đã áp dụng không khớp với lựa chọn. Không tạo file tải về.")
                st.session_state.ai_import_result = {
                    "fingerprint": fingerprint, "data": output.getvalue(), "count": summary["applied"],
                    "selection": sorted(selected_ids),
                    "applied_ids": summary["applied_ids"], "preserved_ids": summary["preserved_ids"],
                    "preexisting_revision_ids": summary["preexisting_revision_ids"],
                    "name": review_output_filename(report_name, "ai_result"),
                }
            except Exception as error:
                st.error(f"Không tạo được file kết quả: {error}")
        result = st.session_state.get("ai_import_result", {})
        if result.get("fingerprint") == fingerprint and result.get("selection") == sorted(selected_ids):
            st.success(f"Đã áp dụng {result['count']} bản dịch vào báo cáo mới. Comment của các mã không chọn được giữ lại.")
            st.caption("Bản dịch AI CHỈ được áp dụng cho: " + ", ".join(result["applied_ids"]))
            if result["preserved_ids"]:
                st.caption("Các mã giữ nguyên trong lần nhập này: " + ", ".join(result["preserved_ids"]))
            if result["preexisting_revision_ids"]:
                st.info("Track Changes có sẵn trong báo cáo nguồn được giữ nguyên, KHÔNG phải do lần nhập AI này: " + ", ".join(result["preexisting_revision_ids"]))
            st.download_button("📥 Tải báo cáo đã áp dụng bản dịch AI", result["data"], file_name=result["name"],
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               key="ai_download_result", width="stretch")
