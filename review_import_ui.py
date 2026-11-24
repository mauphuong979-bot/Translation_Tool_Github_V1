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


COLUMN_APPLY = "Áp dụng"
COLUMN_TRANSLATION = "Bản dịch AI (sửa được)"


def _review_state(scope):
    """Per report (and AI_Form) review state kept across reruns: user edits, unticked codes and the table version."""
    state = st.session_state.get("ai_review_state")
    if not state or state.get("scope") != scope:
        state = {"scope": scope, "edits": {}, "excluded": [], "version": 0}
        st.session_state.ai_review_state = state
    return state


def _on_table_edit(editor_key, scope, row_ids):
    """Move the table's pending edits into the review state, then rebuild the table so statuses are re-checked."""
    state = st.session_state.get("ai_review_state")
    if not state or state.get("scope") != scope:
        return
    excluded = set(state["excluded"])
    for index, change in st.session_state.get(editor_key, {}).get("edited_rows", {}).items():
        review_id = row_ids[int(index)]
        if COLUMN_TRANSLATION in change:
            state["edits"][review_id] = str(change[COLUMN_TRANSLATION] or "")
        if COLUMN_APPLY in change:
            if change[COLUMN_APPLY]:
                excluded.discard(review_id)
            else:
                excluded.add(review_id)
    state["excluded"] = sorted(excluded)
    state["version"] += 1


def _set_selection(scope, excluded):
    state = st.session_state.get("ai_review_state")
    if state and state.get("scope") == scope:
        state["excluded"] = sorted(excluded)
        state["version"] += 1


def _clear_edits(scope):
    state = st.session_state.get("ai_review_state")
    if state and state.get("scope") == scope:
        state["edits"] = {}
        state["version"] += 1


def _status_label(row):
    label = "Track Changes có sẵn" if row["status"] == "blocked" and row.get("source_has_revisions") else STATUS_LABELS[row["status"]]
    if row.get("edited"):
        return f"✏️ Đã sửa · {label}"
    return f"📝 Sửa trong Word · {label}" if row.get("word_edited") else label


def _full_note(row):
    note = row["reason"] or "Sẵn sàng. Chỉ các mã được tích và xác nhận mới được áp dụng."
    if row.get("source_has_revisions") and row["status"] != "applied" and row.get("source_current_text"):
        note += "\nNội dung đang có trong báo cáo nguồn: " + row["source_current_text"]
    return note


def _render_review_table(rows, state, scope, ready_ids, selected_ids):
    """One table for reviewing, editing and choosing codes: tick Áp dụng, double-click a cell to read or edit the full text."""
    existing = [row["review_id"] for row in rows if row.get("source_has_revisions") and row["status"] != "applied"]
    if existing:
        st.info("Báo cáo nguồn đã có sẵn Track Changes ở các mã: " + ", ".join(dict.fromkeys(existing))
                + ". Các đoạn này có thể đã hiện bản dịch dù không được chọn để nhập từ AI. Lần nhập này giữ nguyên các thay đổi đó.")
    edits = state["edits"]
    all_column, none_column, undo_column = st.columns(3, gap="small")
    with all_column:
        st.button(f"Chọn tất cả {len(ready_ids)} mã sẵn sàng", key="ai_select_all", width="stretch",
                  disabled=len(selected_ids) == len(ready_ids), on_click=_set_selection, args=(scope, []))
    with none_column:
        st.button("Bỏ chọn tất cả", key="ai_select_none", width="stretch", disabled=not selected_ids,
                  on_click=_set_selection, args=(scope, ready_ids))
    with undo_column:
        st.button(f"↩️ Bỏ {len(edits)} chỉnh sửa, dùng lại bản AI", key="ai_clear_edits", width="stretch", disabled=not edits,
                  on_click=_clear_edits, args=(scope,))
    selected = set(selected_ids)
    table = [{COLUMN_APPLY: row["review_id"] in selected, "Mã": row["review_id"], "Trạng thái": _status_label(row),
              "Nội dung gốc": row["original"], COLUMN_TRANSLATION: row["translation"], "Ghi chú": _full_note(row)} for row in rows]
    editor_key = f"ai_review_editor_{scope[:16]}_{state['version']}"
    st.data_editor(
        table, key=editor_key, hide_index=True, width="stretch", num_rows="fixed",
        height=min(620, 40 + 104 * len(rows)), row_height=104,
        disabled=["Mã", "Trạng thái", "Nội dung gốc", "Ghi chú"],
        on_change=_on_table_edit, args=(editor_key, scope, [row["review_id"] for row in rows]),
        column_config={
            COLUMN_APPLY: st.column_config.CheckboxColumn(width=92, pinned=True, help="Chỉ các mã Sẵn sàng được tích mới được áp dụng."),
            "Mã": st.column_config.TextColumn(width=130, pinned=True),
            "Trạng thái": st.column_config.TextColumn(width=150),
            "Nội dung gốc": st.column_config.TextColumn(width=280),
            COLUMN_TRANSLATION: st.column_config.TextColumn(
                width=320, help="Bấm đúp vào ô để xem đầy đủ hoặc sửa bản dịch. Sửa xong bấm ra ngoài ô; trạng thái được kiểm tra lại ngay."),
            "Ghi chú": st.column_config.TextColumn(width=320, help="Kết quả kiểm tra và lý do. Bấm đúp vào ô để xem đầy đủ."),
        })
    st.caption("Bấm đúp vào ô để xem đầy đủ nội dung. Sửa trực tiếp ở cột Bản dịch AI rồi bấm ra ngoài ô: ứng dụng kiểm tra lại mã đó "
               "(để trống = không áp dụng). Tích cột Áp dụng để chọn mã; chỉ mã Sẵn sàng mới được áp dụng.")


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
            st.success(f"Đã lưu {count} tên công ty (trạng thái Chờ duyệt trên Google Sheet, sheet CompanyNames). "
                       "Các lần xử lý và dịch sau sẽ dùng tên này ngay.")


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


def _api_scope(report_data):
    return "api:" + hashlib.sha256(report_data).hexdigest()


def _effective_results(results, edits):
    """AI results with the user's table edits, in the same JSON shape, so a downloaded file can be imported again."""
    effective = [dict(result, translation=edits.get(result["review_id"], result["translation"])) for result in results]
    known = {result["review_id"] for result in results}
    effective += [{"review_id": review_id, "translation": text, "highlight": []} for review_id, text in edits.items() if review_id not in known]
    return effective


def _plain(text):
    return " ".join(str(text or "").split())


def _mark_word_edits(preview, results):
    """Flag codes whose translation in the uploaded Word file differs from the API translation."""
    api = {result["review_id"]: _plain(result["translation"]) for result in results}
    for row in preview["rows"]:
        row["word_edited"] = bool(row.get("translation")) and _plain(row["translation"]) != api.get(row["review_id"], "")
    return preview


def _prefilled_worksheet(manifest, outcome, edits):
    """The AI_Form Word worksheet with the API translations (and in-app edits) already in the result column."""
    from review_export import build_ai_review_document

    results = {result["review_id"]: dict(result, note=outcome["notes"].get(result["review_id"], ""))
               for result in _effective_results(outcome["results"], edits)}
    items = [dict(item, batch=manifest["batch"]) for item in manifest["items"]]
    document = build_ai_review_document(items, manifest["target_col"], results)
    return document.getvalue() if document is not None else None


def _clear_word(scope):
    st.session_state.pop("ai_api_word", None)
    st.session_state.ai_word_uploader = st.session_state.get("ai_word_uploader", 0) + 1
    _clear_edits(scope)


def _render_word_round_trip(manifest, outcome, report_key, scope):
    """Download the translations as an AI_Form Word file for translators, and take the edited file back.

    Returns {"data", "name", "hash"} of the uploaded Word file for this report, or None to use the API results."""
    with st.container(border=True):
        st.caption("Tải file Word có sẵn bản dịch AI (ghi chú AI ở comment, tên cần kiểm tra tô vàng) → chỉ sửa cột Bản dịch hoàn thiện, "
                   "được bật Track Changes → lưu và tải lên ô bên phải. Không sửa 4 cột đầu, không gộp/tách bảng.")
        download_column, upload_column = st.columns(2, gap="small")
        with download_column:
            edits = _review_state(scope)["edits"]
            cache_key = hashlib.sha256(json.dumps([report_key, outcome["results"], outcome["notes"], edits], ensure_ascii=False,
                                                  sort_keys=True).encode("utf-8")).hexdigest()
            cached = st.session_state.get("ai_api_word_export", {})
            if cached.get("key") != cache_key:
                try:
                    cached = {"key": cache_key, "data": _prefilled_worksheet(manifest, outcome, edits)}
                except Exception as error:
                    st.error(f"Không tạo được file Word: {error}")
                    cached = {"key": cache_key, "data": None}
                st.session_state.ai_api_word_export = cached
            if cached["data"]:
                st.download_button("📄 Tải AI_Form (Word) đã điền bản dịch AI", cached["data"], file_name="AI_Form_API.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   key="ai_api_download_word", width="stretch")
        with upload_column:
            uploaded = st.file_uploader("Tải lên AI_Form đã sửa trong Word", type=["docx"],
                                        key=f"ai_api_word_upload_{st.session_state.get('ai_word_uploader', 0)}")
        word = st.session_state.get("ai_api_word")
        if uploaded is None:
            if word and word.get("report") == report_key:
                st.session_state.pop("ai_api_word", None)
                _clear_edits(scope)
            return None
        if uploaded is not None:
            data = uploaded.getvalue()
            digest = hashlib.sha256(data).hexdigest()
            if not word or word.get("hash") != digest or word.get("report") != report_key:
                word = {"report": report_key, "data": data, "name": uploaded.name, "hash": digest}
                st.session_state.ai_api_word = word
                _clear_edits(scope)
                st.session_state.pop("ai_import_result", None)
        if not word or word.get("report") != report_key:
            return None
        info_column, reset_column = st.columns([3, 1], gap="small")
        with info_column:
            st.success(f"Bảng rà soát bên dưới đang dùng bản dịch từ file Word: {word['name']}. "
                       "Mã có nhãn 📝 là mã biên dịch viên đã sửa so với bản AI.")
        with reset_column:
            st.button("Bỏ file Word, dùng lại bản AI", key="ai_api_word_clear", width="stretch", on_click=_clear_word, args=(scope,))
        return word


def _section(number, title, hint=""):
    st.markdown(f'<div class="ai-section"><b>{number}</b><span>{html.escape(title)}</span>'
                + (f'<em>{html.escape(hint)}</em>' if hint else "") + "</div>", unsafe_allow_html=True)


def _render_translation_setup(items, manifest, config, report_key, saved):
    """What will be sent, cost estimate and consent; returns (consent, translate clicked)."""
    from ai_translate_api import LOOKUP_DOMAINS, build_payload, estimate_cost, names_to_look_up

    names = names_to_look_up(items, saved) if config.web_search else []
    notes = [f"{len(items)} mục cần dịch sang {manifest['target_col']}. Chỉ gửi nội dung bảng AI_Form (mã, nội dung cần dịch, câu tham chiếu); "
             "không gửi file, các đoạn khác của báo cáo, comment hay tên file."]
    used = sorted(name for name in saved if any(name in item["original"] for item in items))
    if used:
        notes.append(f"Dùng {len(used)} tên công ty đã xác nhận trước đó (không tra lại).")
    if names:
        notes.append(f"Tra tên quốc tế của {len(names)} tên công ty trên {' và '.join(LOOKUP_DOMAINS)}; bước tra cứu chỉ gửi tên công ty.")
    st.caption(" ".join(notes))
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
    clicked = st.button("Dịch bằng AI", key="ai_translate_api", type="primary", disabled=not consent or not items, width="stretch")
    return consent, clicked


def _render_api_source(report_data, config, scope):
    """Step 1 (translate) and step 2 (optional Word round trip); returns (results, notes, word) once a run exists, else None."""
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

    saved = _saved_names()
    _section("1", "Dịch bằng AI")
    if run:
        translated = len(run["outcome"]["results"])
        st.markdown(f"✅ Đã dịch **{translated}/{len(items)}** mục. Bản dịch nằm trong bảng rà soát ở bước 3.")
        with st.expander("Xem dữ liệu đã gửi / dịch lại toàn bộ", expanded=False):
            consent, clicked = _render_translation_setup(items, manifest, config, report_key, saved)
    else:
        consent, clicked = _render_translation_setup(items, manifest, config, report_key, saved)
    if clicked:
        try:
            outcome = _run_translation(items, manifest["target_col"], config, saved_names=saved)
        except Exception as error:
            _show_api_error("Không gọi được AI", error)
            return None
        st.session_state.ai_api_run = {"report": report_key, "outcome": outcome}
        _review_state(scope)
        _clear_word(scope)
        st.session_state.pop("ai_import_preview", None)
        st.session_state.pop("ai_import_result", None)
        st.rerun()
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
    retry_items = [item for item in items if item["review_id"] not in {result["review_id"] for result in outcome["results"]}]
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
    _render_company_names(outcome)
    if not outcome["results"]:
        st.info("AI chưa trả về bản dịch nào. Kiểm tra lỗi ở trên rồi thử lại.")
        return None
    _section("2", "Sửa bằng Word", "Tùy chọn – dành cho biên dịch viên")
    word = _render_word_round_trip(manifest, outcome, report_key, scope)
    _section("3", "Rà soát & áp dụng")
    return outcome["results"], outcome["notes"], word


# ---------------------------------------------------------------- review and apply (shared)

def _render_review_and_apply(scope, base_key, report_name, preview, apply, error_prefix="Không kiểm tra được bản dịch AI"):
    """Shared review table and apply step. preview(overrides) and apply(selected_ids, overrides) always read the report
    as it was before any AI import, so translations can be edited and applied again as many times as needed."""
    state = _review_state(scope)
    edits = dict(state["edits"])
    fingerprint = hashlib.sha256((base_key + "\0" + json.dumps(edits, ensure_ascii=False, sort_keys=True)).encode("utf-8")).hexdigest()
    saved_preview = st.session_state.get("ai_import_preview", {})
    if saved_preview.get("fingerprint") != fingerprint or saved_preview.get("version") != IMPORT_PREVIEW_VERSION:
        try:
            data = preview(edits)
        except Exception as error:
            st.error(f"{error_prefix}: {error}")
            return
        saved_preview = {"fingerprint": fingerprint, "version": IMPORT_PREVIEW_VERSION, "data": data}
        st.session_state.ai_import_preview = saved_preview
    data = saved_preview["data"]
    rows = data["rows"]
    _render_counts(data["counts"], len(rows))
    ready_ids = [row["review_id"] for row in rows if row["status"] == "ready"]
    excluded = set(state["excluded"])
    selected_ids = [review_id for review_id in ready_ids if review_id not in excluded]
    _render_review_table(rows, state, scope, ready_ids, selected_ids)
    if not ready_ids:
        st.info("Chưa có bản dịch nào sẵn sàng để áp dụng. Các mục trống, bị chặn hoặc thiếu bảng giữ nguyên nội dung và comment gốc.")
        return
    edited_ready = [review_id for review_id in selected_ids if review_id in edits]
    summary_column, apply_column = st.columns([3, 2], gap="small")
    with summary_column:
        st.markdown(f"**Sẽ áp dụng {len(selected_ids)}/{len(ready_ids)} mã sẵn sàng**"
                    + (f" (trong đó {len(edited_ready)} mã đã sửa tay)" if edited_ready else ""))
        st.caption("Mỗi lần áp dụng tạo một bản báo cáo MỚI từ báo cáo gốc (trước khi nhập AI), có Track Changes và chỉ xóa comment của các mã đã áp dụng. "
                   "Muốn sửa sau khi đã tải về: sửa trong bảng, bấm Áp dụng lại và dùng file mới nhất. "
                   "Reject khôi phục nội dung nhưng không khôi phục comment đã xóa; hãy giữ báo cáo gốc để dự phòng.")
    selection_key = hashlib.sha256("\n".join(sorted(selected_ids)).encode("utf-8")).hexdigest()
    with apply_column:
        confirmed = st.checkbox("Tôi đã rà soát các bản dịch được chọn và đồng ý xóa comment của chúng trong bản mới",
                                key=f"ai_confirm_{fingerprint}_{selection_key}")
        apply_clicked = st.button("Áp dụng bản dịch AI", key="ai_apply_translations", disabled=not confirmed or not selected_ids, width="stretch")
    if apply_clicked:
        st.session_state.pop("ai_import_result", None)
        try:
            with st.spinner("Đang tạo báo cáo mới có Track Changes..."):
                output, summary = apply(selected_ids, edits)
            if set(summary["applied_ids"]) != set(selected_ids):
                raise ValueError("Các mã đã áp dụng không khớp với lựa chọn. Không tạo file tải về.")
            st.session_state.ai_import_result = {
                "scope": scope, "fingerprint": fingerprint, "data": output.getvalue(), "count": summary["applied"],
                "selection": sorted(selected_ids), "edited_ids": edited_ready,
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
        if result.get("edited_ids"):
            st.caption("Dùng bản dịch đã sửa tay cho: " + ", ".join(result["edited_ids"]))
        if result["preserved_ids"]:
            st.caption("Các mã giữ nguyên trong lần nhập này: " + ", ".join(result["preserved_ids"]))
        if result["preexisting_revision_ids"]:
            st.info("Track Changes có sẵn trong báo cáo nguồn được giữ nguyên, KHÔNG phải do lần nhập AI này: " + ", ".join(result["preexisting_revision_ids"]))
        st.download_button("📥 Tải báo cáo đã áp dụng bản dịch AI", result["data"], file_name=result["name"],
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           key="ai_download_result", width="stretch")
    elif result.get("scope") == scope:
        st.info("Bản dịch hoặc lựa chọn đã thay đổi sau lần áp dụng trước. Xác nhận lại và bấm Áp dụng bản dịch AI để tạo file mới; file đã tải trước đó không còn khớp.")


def render_ai_import():
    with st.expander("📤 Nhập bản dịch từ AI", expanded=False), st.container(key="ai_import_workspace", gap="small"):
        config = _api_config()
        source = st.radio("Nguồn bản dịch", (SOURCE_API, SOURCE_FILE), horizontal=True, key="ai_source_mode") if config else SOURCE_FILE
        if source == SOURCE_FILE:
            st.markdown('<div class="ai-review-steps"><b>01 · Tải file</b><span>02 · Rà soát mã</span><span>03 · Áp dụng & tải về</span></div>', unsafe_allow_html=True)
            st.caption("Dùng đúng cặp báo cáo và bảng AI_Form đã được AI điền. Chỉ các mã hợp lệ và được chọn mới được áp dụng; bạn vẫn cần rà soát chất lượng bản dịch.")
        else:
            st.caption("Các bước: ① AI dịch → ② (tùy chọn) biên dịch viên sửa bằng Word → ③ rà soát, chọn mã và áp dụng vào báo cáo.")
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
            scope = _api_scope(report_data)
            run = _render_api_source(report_data, config, scope)
            if run is None:
                return
            results, notes, word = run
            base_key = hashlib.sha256(report_data + b"\0api\0" + json.dumps(results, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()
            if word is not None:
                _render_review_and_apply(scope, base_key + "\0word\0" + word["hash"], report_name,
                                         lambda edits: _mark_word_edits(_with_notes(preview_ai_translations(report_data, word["data"], edits), notes), results),
                                         lambda selected, edits: apply_ai_translations(report_data, word["data"], selected, edits),
                                         "Không nhập được file Word này")
                return
            _render_review_and_apply(scope, base_key, report_name,
                                     lambda edits: _with_notes(preview_ai_results(report_data, results, edits), notes),
                                     lambda selected, edits: apply_ai_results(report_data, results, selected, edits),
                                     "Không kiểm tra được kết quả AI")
            return

        with worksheet_column:
            worksheet = st.file_uploader("Bảng AI_Form đã được AI điền", type=["docx"], key="ai_import_worksheet")
        if report_data is None or worksheet is None:
            st.info("Chọn đủ hai file để kiểm tra bản dịch. Bảng AI cũ chưa có mã cần được tạo lại.")
            return
        worksheet_data = worksheet.getvalue()
        scope = "file:" + hashlib.sha256(report_data + b"\0" + worksheet_data).hexdigest()
        if st.button("Kiểm tra bản dịch AI", key="ai_check_translations", width="stretch"):
            st.session_state.pop("ai_import_preview", None)
            st.session_state.pop("ai_import_result", None)
            st.session_state.ai_file_checked = scope
        if st.session_state.get("ai_file_checked") != scope:
            return
        _render_review_and_apply(scope, scope, report_name,
                                 lambda edits: preview_ai_translations(report_data, worksheet_data, edits),
                                 lambda selected, edits: apply_ai_translations(report_data, worksheet_data, selected, edits),
                                 "Không nhập được các file này")
