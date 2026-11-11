import hashlib

import streamlit as st

from review_roundtrip import apply_ai_translations, preview_ai_translations
from review_filenames import review_output_filename


STATUS_LABELS = {
    "ready": "Ready", "blank": "Blank", "blocked": "Needs manual review",
    "missing": "Missing table", "applied": "Already applied", "unchanged": "Unchanged",
}
IMPORT_PREVIEW_VERSION = 5


def _render_preview_rows(rows, fingerprint):
    st.dataframe([
        {"ID": row["review_id"],
         "Status": "Source Track Changes" if row["status"] == "blocked" and row.get("source_has_revisions") else STATUS_LABELS[row["status"]],
         "Source": row["original"], "AI translation": row["translation"], "Note": row["reason"]}
        for row in rows
    ], hide_index=True, width="stretch", height=min(404, 44 + 72 * len(rows)), row_height=72,
        column_config={
            "ID": st.column_config.TextColumn(width=140, pinned=True),
            "Status": st.column_config.TextColumn(width=155),
            "Source": st.column_config.TextColumn(width=210),
            "AI translation": st.column_config.TextColumn(width=230),
            "Note": st.column_config.TextColumn(width=380, help="Import checks and reasons. The full note is also available below the table."),
        })
    existing = [row["review_id"] for row in rows if row.get("source_has_revisions") and row["status"] != "applied"]
    if existing:
        st.info("Track Changes already present in the source report: " + ", ".join(dict.fromkeys(existing))
                + ". They may already display translated text even though they are not selected for AI import. This import leaves those changes untouched.")
    with st.container(border=True):
        st.markdown("**Full row details and note**")
        default_index = next((index for index, row in enumerate(rows) if row.get("source_has_revisions") and row["status"] == "blocked"),
                             next((index for index, row in enumerate(rows) if row["status"] == "blocked"), 0))
        identity_column, note_column = st.columns([1, 2], gap="small")
        with identity_column:
            selected_index = st.selectbox("Review ID to inspect", range(len(rows)), index=default_index,
                                          format_func=lambda index: f"{rows[index]['review_id']} · {STATUS_LABELS[rows[index]['status']]}",
                                          key=f"ai_detail_{fingerprint}")
        row = rows[selected_index]
        with note_column:
            st.markdown("**Note (full text)**")
            st.info(row["reason"] or "Ready for import. Only selected, confirmed IDs will be applied.")
        with st.expander("Compare source and AI translation", expanded=False):
            source_column, translation_column = st.columns(2, gap="small")
            with source_column:
                st.markdown("**Source**")
                st.text(row["original"])
            with translation_column:
                st.markdown("**AI translation (worksheet candidate)**")
                st.text(row["translation"] or "No completed AI translation.")
            if row.get("source_has_revisions"):
                st.markdown("**Text already present in the source report (before this import)**")
                st.text(row.get("source_current_text", ""))


def _render_counts(counts, total):
    attention = counts.get("blocked", 0) + counts.get("missing", 0)
    skipped = counts.get("blank", 0) + counts.get("unchanged", 0)
    cards = (("Total", total), ("Ready", counts.get("ready", 0)),
             ("Needs attention", attention), ("Blank / unchanged", skipped))
    st.markdown('<div class="ai-review-summary">' + ''.join(
        f'<div class="ai-review-stat"><span>{label}</span><strong>{int(value)}</strong></div>'
        for label, value in cards) + '</div>', unsafe_allow_html=True)


def render_ai_import():
    with st.expander("📤 Import AI translations", expanded=False), st.container(key="ai_import_workspace", gap="small"):
        st.markdown('<div class="ai-review-steps"><b>01 · Upload</b><span>02 · Review IDs</span><span>03 · Apply & download</span></div>', unsafe_allow_html=True)
        st.caption("Use the matching report and AI-completed worksheet. Only validated, selected IDs are applied; translation quality still needs your review.")
        current_report = st.session_state.get("processed_outputs", {}).get("review")
        use_current = False
        report_column, worksheet_column = st.columns(2, gap="small")
        with report_column:
            if current_report is not None:
                use_current = st.checkbox("Use the Comments + Track Changes report from this session", value=True, key="ai_use_current_report")
            if use_current:
                report_data = current_report.getvalue() if hasattr(current_report, "getvalue") else current_report
                report_name = review_output_filename(st.session_state.get("processed_filename") or "Report.docx", "review")
                st.caption("Matching report: " + report_name)
            else:
                uploaded_report = st.file_uploader("Matching Comments + Track Changes report", type=["docx"], key="ai_import_report")
                report_data = uploaded_report.getvalue() if uploaded_report is not None else None
                report_name = uploaded_report.name if uploaded_report is not None else "Report.docx"
        with worksheet_column:
            worksheet = st.file_uploader("AI-completed translation worksheet", type=["docx"], key="ai_import_worksheet")
        if report_data is None or worksheet is None:
            st.info("Choose both files to check the translations. Older worksheets without IDs must be regenerated.")
            return
        worksheet_data = worksheet.getvalue()
        fingerprint = hashlib.sha256(report_data + b"\0" + worksheet_data).hexdigest()
        if st.button("Check AI translations", key="ai_check_translations", width="stretch"):
            st.session_state.pop("ai_import_preview", None)
            st.session_state.pop("ai_import_result", None)
            try:
                with st.spinner("Checking IDs, source content and comment anchors..."):
                    preview = preview_ai_translations(report_data, worksheet_data)
                st.session_state.ai_import_preview = {"fingerprint": fingerprint, "version": IMPORT_PREVIEW_VERSION, "data": preview}
            except Exception as error:
                st.error(f"Cannot import these files: {error}")
        saved_preview = st.session_state.get("ai_import_preview", {})
        if saved_preview.get("fingerprint") != fingerprint:
            return
        if saved_preview.get("version") != IMPORT_PREVIEW_VERSION:
            st.info("Click Check AI translations again to refresh the preview and source-change details.")
            return
        preview = saved_preview["data"]
        _render_counts(preview["counts"], len(preview["rows"]))
        _render_preview_rows(preview["rows"], fingerprint)
        ready_ids = [row["review_id"] for row in preview["rows"] if row["status"] == "ready"]
        if not ready_ids:
            st.info("No translations are ready to apply. Blank, blocked and missing entries keep their original text and comments.")
            return
        selection_column, apply_column = st.columns([3, 2], gap="small")
        with selection_column:
            selected_ids = st.multiselect("Translations to apply", ready_ids, default=ready_ids, key=f"ai_selection_{fingerprint}")
            st.caption("A new copy records Track Changes and removes only applied comments. Reject restores text, not deleted comments; keep the source report as a backup.")
        selection_key = hashlib.sha256("\n".join(sorted(selected_ids)).encode("utf-8")).hexdigest()
        with apply_column:
            confirmed = st.checkbox("I have reviewed the selected translations and agree to remove their comments in the new copy", key=f"ai_confirm_{fingerprint}_{selection_key}")
            apply_clicked = st.button("Apply AI translations", key="ai_apply_translations", disabled=not confirmed or not selected_ids, width="stretch")
        if apply_clicked:
            st.session_state.pop("ai_import_result", None)
            try:
                with st.spinner("Creating a new report with Track Changes..."):
                    output, summary = apply_ai_translations(report_data, worksheet_data, selected_ids)
                if set(summary["applied_ids"]) != set(selected_ids):
                    raise ValueError("The applied IDs do not match the selection. No download will be provided.")
                st.session_state.ai_import_result = {
                    "fingerprint": fingerprint, "data": output.getvalue(), "count": summary["applied"],
                    "selection": sorted(selected_ids),
                    "applied_ids": summary["applied_ids"], "preserved_ids": summary["preserved_ids"],
                    "preexisting_revision_ids": summary["preexisting_revision_ids"],
                    "name": review_output_filename(report_name, "ai_result"),
                }
            except Exception as error:
                st.error(f"No output was created: {error}")
        result = st.session_state.get("ai_import_result", {})
        if result.get("fingerprint") == fingerprint and result.get("selection") == sorted(selected_ids):
            st.success(f"Applied {result['count']} translation(s) to a new report. Unselected comments are retained.")
            st.caption("New AI changes applied ONLY to: " + ", ".join(result["applied_ids"]))
            if result["preserved_ids"]:
                st.caption("IDs kept unchanged by this import: " + ", ".join(result["preserved_ids"]))
            if result["preexisting_revision_ids"]:
                st.info("Pre-existing source Track Changes preserved, NOT newly applied by AI import: " + ", ".join(result["preexisting_revision_ids"]))
            st.download_button("📥 Download report with AI Track Changes", result["data"], file_name=result["name"],
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               key="ai_download_result", width="stretch")
