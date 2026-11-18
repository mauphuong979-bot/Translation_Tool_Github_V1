import hashlib
import html
import json

import streamlit as st

from review_roundtrip import (apply_ai_results, apply_ai_translations, preview_ai_results, preview_ai_translations,
                              read_review_manifest)
from review_filenames import review_output_filename


STATUS_LABELS = {
    "ready": "Sẵn sàng", "blank": "Ô trống", "blocked": "Cần xử lý thủ công",
    "missing": "Thiếu bảng", "applied": "Đã áp dụng", "unchanged": "Không đổi",
}
IMPORT_PREVIEW_VERSION = 5
SOURCE_API = "🤖 Dịch tự động bằng AI"
SOURCE_FILE = "📄 Nhập file AI_Form"


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


# ---------------------------------------------------------------- API source

def _api_config():
    """AI API settings from st.secrets["anthropic"]; None when not configured or the user is not allowed."""
    from ai_translate_api import load_config

    try:
        settings = dict(st.secrets["anthropic"])
    except Exception:
        return None
    allowed = settings.pop("enabled_users", None) or []
    if allowed and st.session_state.get("username") not in allowed:
        return None
    return load_config(settings)


def log_event(*args):
    from usage_logger import log_event as write

    write(*args)


def _with_notes(preview, notes):
    """Show the AI's own notes next to the checker's reason for each code."""
    for row in preview["rows"]:
        note = notes.get(row["review_id"])
        if note:
            row["reason"] = (row["reason"] + " " if row["reason"] else "") + "AI ghi chú: " + note
    return preview


def _saved_names():
    """Confirmed company names {Vietnamese: international}; an unavailable store never blocks translation."""
    try:
        from company_names_store import company_name_map

        return company_name_map()
    except Exception as error:
        st.warning(f"Không đọc được bảng tên công ty đã xác nhận: {error}")
        return {}


def _render_company_names(outcome):
    """Found, saved and not-found company names; users can confirm or correct them and save to the shared table."""
    found = outcome.get("company_names", {})
    missing = [name for name in outcome.get("searched_names", []) if name not in found]
    if not found and not missing:
        return
    with st.expander(f"Tên công ty ({len(found)} có tên quốc tế, {len(missing)} chưa tìm thấy)", expanded=False):
        rows = [{"Lưu": False, "Tên tiếng Việt": name, "Tên quốc tế": entry["international"],
                 "Trạng thái": "Đã lưu trước đó" if entry.get("saved") else "Tra được", "Nguồn": entry.get("source_url", "")}
                for name, entry in found.items()]
        rows += [{"Lưu": False, "Tên tiếng Việt": name, "Tên quốc tế": "", "Trạng thái": "Chưa tìm thấy", "Nguồn": ""} for name in missing]
        edited = st.data_editor(rows, hide_index=True, width="stretch", key="ai_company_names_editor",
                                disabled=["Tên tiếng Việt", "Trạng thái", "Nguồn"],
                                column_config={"Lưu": st.column_config.CheckboxColumn(width="small"),
                                               "Nguồn": st.column_config.LinkColumn()})
        st.caption("Kiểm tra tên quốc tế (có thể sửa hoặc tự nhập), đánh dấu cột Lưu rồi bấm lưu. Tên đã lưu được dùng chung cho mọi người: "
                   "tự thay khi xử lý báo cáo, gửi kèm khi dịch bằng AI và không phải tra web lại. Tên chưa lưu vẫn giữ tiếng Việt và tô vàng.")
        chosen = [row for row in edited if row["Lưu"]]
        if st.button(f"💾 Lưu {len(chosen)} tên đã xác nhận", key="ai_company_names_save", disabled=not chosen, width="stretch"):
            blank = [row["Tên tiếng Việt"] for row in chosen if not str(row["Tên quốc tế"] or "").strip()]
            if blank:
                st.error("Hãy nhập tên quốc tế trước khi lưu: " + ", ".join(blank))
                return
            try:
                from company_names_store import save_company_names

                count = save_company_names([{"vietnamese": row["Tên tiếng Việt"], "international": row["Tên quốc tế"], "source": row["Nguồn"]}
                                            for row in chosen], st.session_state.get("username", "unknown"))
            except Exception as error:
                st.error(f"Không lưu được tên công ty: {error}")
                return
            log_event(st.session_state.get("username", "unknown"), "Company Names",
                      f"Lưu {count} tên: " + "; ".join(f"{row['Tên tiếng Việt']} = {row['Tên quốc tế']}" for row in chosen))
            st.success(f"Đã lưu {count} tên công ty. Các lần xử lý và dịch sau sẽ dùng tên này.")


def _run_translation(items, target_col, config, previous=None, saved_names=None):
    from ai_translate_api import translate_items, usage_cost

    bar = st.progress(0.0, text=f"Đang dịch 0/{len(items)} mục...")
    outcome = translate_items(items, target_col, config, saved_names=saved_names,
                              progress=lambda done, total: bar.progress(done / total, text=f"Đang dịch {done}/{total} mục..."))
    bar.empty()
    cost = usage_cost(outcome["usage"], outcome["model"])
    usage = outcome["usage"]
    log_event(st.session_state.get("username", "unknown"), "AI API",
              f"{len(items)} mục {target_col}, model {outcome['model']}, vào {usage.get('input_tokens', 0)}, "
              f"ra {usage.get('output_tokens', 0)}, cache đọc {usage.get('cache_read_input_tokens', 0)}, "
              f"tra web {usage.get('web_search_requests', 0)} lần, chi phí ~${cost or 0:.4f}, lỗi {len(outcome['failed_ids'])} mục")
    if previous:
        replaced = {result["review_id"] for result in outcome["results"]}
        outcome["results"] = [result for result in previous["results"] if result["review_id"] not in replaced] + outcome["results"]
        outcome["notes"] = {**previous["notes"], **outcome["notes"]}
        outcome["company_names"] = {**previous.get("company_names", {}), **outcome.get("company_names", {})}
        outcome["searched_names"] = list(dict.fromkeys(previous.get("searched_names", []) + outcome.get("searched_names", [])))
        outcome["usage"] = {key: previous["usage"].get(key, 0) + usage.get(key, 0) for key in set(previous["usage"]) | set(usage)}
    return outcome


def _show_api_error(prefix, error):
    from ai_translate_api import describe_api_error, technical_detail

    st.error(f"{prefix}: {describe_api_error(error)}")
    with st.expander("Chi tiết kỹ thuật (gửi cho quản trị viên khi cần)", expanded=False):
        st.code(technical_detail(error), language=None, wrap_lines=True)


def _cell(text):
    return html.escape(str(text)).replace("\n", "<br>")


def _render_payload(payload, names):
    """What will be sent, as a readable table built from the payload itself; the exact JSON is one toggle away."""
    if names:
        st.markdown(f"**Bước 1 – Tra cứu {len(names)} tên công ty** (chỉ gửi các tên này)")
        st.markdown("<div class='ai-payload-names'>" + "".join(f"<span>{_cell(name)}</span>" for name in names) + "</div>",
                    unsafe_allow_html=True)
        st.markdown(f"**Bước 2 – Dịch {len(payload)} mục**")
    with_names = any(entry.get("company_names") for entry in payload)
    header = "<th>Mã</th><th>Nội dung cần dịch</th><th>Câu tham chiếu (tiếng Việt → bản dịch)</th>"
    header += "<th>Tên công ty gửi kèm</th>" if with_names else ""
    rows = []
    for entry in payload:
        references = "".join(
            f"<div class='ai-payload-ref'><b>{_cell(reference['label'])}</b><br>{_cell(reference['vietnamese'])}"
            f"<div class='ai-payload-arrow'>→ {_cell(reference['translation'])}</div></div>"
            for reference in entry["references"]) or "<span class='ai-payload-empty'>Không có</span>"
        row = f"<td class='ai-payload-id'>{_cell(entry['review_id'])}</td><td>{_cell(entry['source'])}</td><td>{references}</td>"
        if with_names:
            row += "<td>" + "<br>".join(f"{_cell(name['vietnamese'])} → {_cell(name['international'])}"
                                        for name in entry.get("company_names", [])) + "</td>"
        rows.append(f"<tr>{row}</tr>")
    with st.container(height=480 if len(payload) > 3 else "content"):
        st.markdown(f"<table class='ai-payload'><thead><tr>{header}</tr></thead><tbody>{''.join(rows)}</tbody></table>",
                    unsafe_allow_html=True)
    if st.toggle("Hiện dạng JSON (đúng nội dung gửi đi)", key="ai_api_payload_json"):
        if names:
            st.caption("Bước 1 – tra cứu tên công ty")
            st.code(json.dumps(names, ensure_ascii=False, indent=1), language="json", wrap_lines=True)
            st.caption("Bước 2 – dịch")
        st.code(json.dumps(payload, ensure_ascii=False, indent=1), language="json", wrap_lines=True)


def _render_api_source(report_data, config):
    """Translate through the API; returns (results, notes) once a run exists for this report, else None."""
    from ai_translate_api import estimate_cost

    report_key = hashlib.sha256(report_data).hexdigest()
    try:
        manifest = read_review_manifest(report_data)
    except Exception as error:
        st.error(f"Không đọc được dữ liệu AI trong báo cáo: {error}")
        return None
    items = manifest["items"]
    run = st.session_state.get("ai_api_run", {})
    if run.get("report") != report_key:
        run = {}
    if not items and not run:
        st.info("Báo cáo này không còn mục nào cần AI dịch.")
        return None

    st.caption(f"{len(items)} mục cần dịch sang {manifest['target_col']}. Chỉ gửi đúng nội dung của bảng AI_Form: mã, nội dung cần dịch "
               "và câu tham chiếu. Không gửi file, các đoạn khác của báo cáo, comment hay tên file.")
    from ai_translate_api import LOOKUP_DOMAINS, build_payload, names_to_look_up

    saved = _saved_names()
    names = names_to_look_up(items, saved) if config.web_search else []
    used = sorted(name for name in saved if any(name in item["original"] for item in items))
    if used:
        st.caption(f"Dùng {len(used)} tên công ty đã xác nhận trước đó (không tra lại).")
    if names:
        st.caption(f"Tra tên quốc tế của {len(names)} tên công ty trên {' và '.join(LOOKUP_DOMAINS)}. Bước tra cứu chỉ gửi tên công ty; "
                   "tên tìm được sẽ được gửi kèm mục có tên đó ở bước dịch.")
    with st.expander("Xem dữ liệu sẽ gửi", expanded=False):
        _render_payload(build_payload(items, saved), names)
    estimate_column, confirm_column = st.columns([2, 3], gap="small")
    with estimate_column:
        if st.button("Ước tính chi phí", key="ai_api_estimate", width="stretch", disabled=not items):
            try:
                with st.spinner("Đang đếm token..."):
                    st.session_state.ai_api_estimate_data = {"report": report_key, "data": estimate_cost(items, manifest["target_col"], config, saved_names=saved)}
            except Exception as error:
                _show_api_error("Không ước tính được chi phí", error)
        estimate = st.session_state.get("ai_api_estimate_data", {})
        if estimate.get("report") == report_key:
            data = estimate["data"]
            cost = f"khoảng ${data['usd_estimate']:.2f}" if data["usd_estimate"] is not None else "chưa có bảng giá cho model này"
            searches = data.get("web_searches_estimate", 0)
            web = f" Gồm tối đa ~{searches} lần tra web; chưa tính token của bước tra cứu." if searches else ""
            st.caption(f"{data['input_tokens']:,} token vào, ~{data['output_tokens_estimate']:,} token ra: {cost}.{web}")
    with confirm_column:
        consent = st.checkbox("Tôi đồng ý gửi nội dung AI_Form của báo cáo này tới Claude API để dịch", key=f"ai_api_consent_{report_key}")
    if st.button("Dịch bằng AI", key="ai_translate_api", disabled=not consent or not items, width="stretch"):
        try:
            outcome = _run_translation(items, manifest["target_col"], config, saved_names=saved)
        except Exception as error:
            _show_api_error("Không gọi được AI", error)
            return None
        run = {"report": report_key, "outcome": outcome}
        st.session_state.ai_api_run = run
        st.session_state.pop("ai_import_preview", None)
        st.session_state.pop("ai_import_result", None)
    if not run:
        return None

    outcome = run["outcome"]
    for error in outcome["errors"]:
        st.error(error)
    for warning in outcome["warnings"]:
        st.warning(warning)
    if outcome.get("error_details"):
        with st.expander("Chi tiết kỹ thuật (gửi cho quản trị viên khi cần)", expanded=False):
            st.code("\n".join(outcome["error_details"]), language=None, wrap_lines=True)
    _render_company_names(outcome)
    if not outcome["results"]:
        st.info("AI chưa trả về bản dịch nào. Kiểm tra lỗi ở trên rồi thử lại.")
        return None
    retry_items = [item for item in items if item["review_id"] not in {result["review_id"] for result in outcome["results"]}]
    retry_column, download_column = st.columns(2, gap="small")
    with retry_column:
        if retry_items and st.button(f"Dịch lại {len(retry_items)} mục còn thiếu", key="ai_api_retry", disabled=not consent, width="stretch"):
            try:
                run = {"report": report_key, "outcome": _run_translation(retry_items, manifest["target_col"], config, outcome, saved)}
            except Exception as error:
                _show_api_error("Không gọi được AI", error)
            else:
                st.session_state.ai_api_run = run
                st.session_state.pop("ai_import_preview", None)
                st.session_state.pop("ai_import_result", None)
                st.rerun()
    with download_column:
        st.download_button("📥 Tải kết quả AI (.json)", json.dumps(outcome["results"], ensure_ascii=False, indent=1),
                           file_name="AI_results.json", mime="application/json", key="ai_api_download_json", width="stretch")
    return outcome["results"], outcome["notes"]


# ---------------------------------------------------------------- review and apply (shared)

def _render_review_and_apply(fingerprint, report_name, apply):
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
                output, summary = apply(selected_ids)
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


def render_ai_import():
    with st.expander("📤 Nhập bản dịch từ AI", expanded=False), st.container(key="ai_import_workspace", gap="small"):
        st.markdown('<div class="ai-review-steps"><b>01 · Tải file</b><span>02 · Rà soát mã</span><span>03 · Áp dụng & tải về</span></div>', unsafe_allow_html=True)
        config = _api_config()
        source = st.radio("Nguồn bản dịch", (SOURCE_API, SOURCE_FILE), horizontal=True, key="ai_source_mode") if config else SOURCE_FILE
        if source == SOURCE_FILE:
            st.caption("Dùng đúng cặp báo cáo và bảng AI_Form đã được AI điền. Chỉ các mã hợp lệ và được chọn mới được áp dụng; bạn vẫn cần rà soát chất lượng bản dịch.")
        else:
            st.caption("Ứng dụng gửi các mục của bảng AI_Form tới AI, rồi kiểm tra kết quả như khi nhập file. Bạn vẫn cần rà soát chất lượng bản dịch trước khi áp dụng.")
        current_report = st.session_state.get("processed_outputs", {}).get("review")
        use_current = False
        report_column, worksheet_column = st.columns(2, gap="small") if source == SOURCE_FILE else (st.container(), None)
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

        if source == SOURCE_API:
            if report_data is None:
                st.info("Chọn báo cáo Comments + Track Changes để dịch các mục còn lại bằng AI.")
                return
            run = _render_api_source(report_data, config)
            if run is None:
                return
            results, notes = run
            fingerprint = hashlib.sha256(report_data + b"\0api\0" + json.dumps(results, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()
            if st.session_state.get("ai_import_preview", {}).get("fingerprint") != fingerprint:
                try:
                    preview = _with_notes(preview_ai_results(report_data, results), notes)
                    st.session_state.ai_import_preview = {"fingerprint": fingerprint, "version": IMPORT_PREVIEW_VERSION, "data": preview}
                except Exception as error:
                    st.error(f"Không kiểm tra được kết quả AI: {error}")
                    return
            _render_review_and_apply(fingerprint, report_name, lambda selected: apply_ai_results(report_data, results, selected))
            return

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
        _render_review_and_apply(fingerprint, report_name, lambda selected: apply_ai_translations(report_data, worksheet_data, selected))
