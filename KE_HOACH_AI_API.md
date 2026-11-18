# Kế hoạch: Tự động xử lý AI_Form qua API

## Mục tiêu

Thay 3 bước thủ công (tải AI_Form → gửi AI → tải file về nhập lại) bằng một nút **"🤖 Dịch tự động bằng AI"**. Kết quả đi qua đúng bộ kiểm tra và màn hình rà soát/áp dụng hiện có, rồi ghi vào báo cáo bằng Track Changes như trước. Luồng AI_Form thủ công được giữ nguyên làm phương án dự phòng.

## Phạm vi dữ liệu gửi ra API (đã chốt)

Được phép gửi nội dung báo cáo ra API, **nhưng chỉ đúng phần đang có trong AI_Form**, không gửi toàn bộ báo cáo.

- **Được gửi:** với mỗi mục trong AI_Form, chỉ gồm bốn cột đầu: `review_id` (mã), nội dung cần dịch, tiếng Việt tham chiếu và bản dịch tham chiếu (kèm nhãn nguồn và tỷ lệ). Ngoài ra có ngôn ngữ đích và prompt quy tắc dịch.
- **Không gửi:**
  - file `.docx`, các đoạn khác của báo cáo, ngữ cảnh xung quanh (đoạn trước, đoạn sau, tiêu đề, bảng);
  - `report_text`, nội dung comment, tên tác giả, tên file, metadata;
  - toàn bộ từ điển hoặc báo cáo kỳ trước. Chỉ gửi đúng các câu tham chiếu đã chọn cho từng mục, giống AI_Form.
- **Cách bảo đảm:**
  - `build_payload` dùng danh sách trường cho phép (allowlist), không truyền nguyên `item` của manifest;
  - có kiểm thử khẳng định payload chỉ chứa các trường trên, và tập mục gửi đi trùng khớp với tập mục của AI_Form tương ứng;
  - giao diện có nút **"Xem dữ liệu sẽ gửi"** hiện đúng JSON trước khi gửi;
  - nhật ký chỉ ghi số mục và số token, không ghi nội dung.
- **Tra tên công ty qua web (giai đoạn 5):** làm thành một bước riêng trước bước dịch. Request tra cứu **chỉ chứa danh sách tên công ty** (do ứng dụng tự tách từ nội dung cần dịch), công cụ `web_search` chỉ được tìm trên hai tên miền đã cho phép, và kết quả chỉ được nhận khi đường dẫn nguồn thuộc hai tên miền đó. Bước dịch không dùng công cụ web; tên tìm được được gửi kèm đúng mục có tên đó. Lý do tách bước: tài liệu API chưa nói rõ web search (luôn kèm citation) có dùng chung được với structured output hay không.

## Nguyên tắc thiết kế

- **Không gửi file Word cho AI.** Gửi JSON (mã, nội dung cần dịch, tham chiếu) và nhận JSON (mã, bản dịch, phần cần tô vàng). AI không thể làm hỏng cấu trúc bảng, cột gốc hay mã.
- **Một bộ kiểm tra cho cả hai luồng.** Kết quả API được chuyển thành cùng dạng "dòng" mà `_worksheet_rows` đang tạo, rồi đi qua cùng logic kiểm tra neo, comment, bookmark và đoạn không an toàn.
- **Người dùng vẫn duyệt trước khi áp dụng.** Không tự động ghi vào báo cáo.
- **Tắt được hoàn toàn.** Không có khóa API thì nút không hiện, ứng dụng chạy như hiện nay.

## Thay đổi theo file

### 1. `review_roundtrip.py`: tách phần kiểm tra khỏi phần đọc worksheet

- Tách `_inspect(report_source, worksheet_source)` thành:
  - `_validate(report, manifest, supplied_rows, identifier=None)`: toàn bộ logic kiểm tra hiện có (dòng 318 trở đi), không đổi hành vi.
  - `_inspect`: đọc worksheet bằng `_worksheet_rows`, rồi gọi `_validate`.
- Thêm:
  - `read_review_manifest(report_source)`: trả về `target_col`, `batch` và danh sách mục chưa áp dụng, để gửi cho API.
  - `rows_from_ai_results(manifest, results)`: chuyển JSON từ API thành dòng chuẩn, trong đó `source_values` là giá trị kỳ vọng, `has_revisions` và `has_objects` là False, `highlights` là vị trí ký tự tính từ các đoạn AI đánh dấu.
  - `preview_ai_results(report, results)` và `apply_ai_results(report, results, selected_ids)`: phiên bản dùng JSON của hai hàm hiện có. Có thể gộp bằng tham số `rows=` vào `apply_ai_translations` để tránh lặp code.
- Cần kiểm thử hồi quy: luồng nhập file AI_Form phải cho kết quả giống hệt trước khi tách.

### 2. `ai_translate_api.py` (mới): gọi Claude API

- `build_payload(items, target_col)`: dùng allowlist, mỗi mục chỉ gồm `review_id`, `source`, `references` (nguồn, tỷ lệ, câu Việt, bản dịch), đúng bốn cột đầu của AI_Form. Xem mục "Phạm vi dữ liệu gửi ra API".
- `translate_items(items, target_col, progress_cb)`:
  - Chia lô khoảng 25 mục mỗi lần gọi; chạy song song 3–4 lô; SDK tự thử lại khi gặp lỗi 429/5xx.
  - Model mặc định `claude-opus-5-5`, effort `medium` (đổi được trong secrets); bật `fallbacks: "default"` để có model dự phòng khi bị từ chối.
  - System prompt: phần quy tắc dịch rút từ `get_ai_review_prompt`, bỏ các câu về sửa và xuất file Word, gắn prompt caching.
  - Structured output, schema: `{"items": [{"review_id", "translation", "highlight": [chuỗi con cần tô vàng], "note"}]}`.
  - Hậu kiểm: mã lạ hoặc trùng thì loại; mã bị thiếu thì đánh dấu `missing`; `note` khác rỗng thì đưa vào cột Ghi chú.
  - Trả về kết quả và thống kê token (input, output, cache) để ghi nhật ký.
- Tra tên công ty (tùy chọn, giai đoạn 2): bật công cụ `web_search` với `allowed_domains` là masothue.com và thuvienphapluat.vn. Nếu tắt, quy tắc là giữ tên tiếng Việt và tô vàng.

### 3. `review_import_ui.py`: thêm luồng API vào giao diện

- Trong expander "📤 Nhập bản dịch từ AI", thêm lựa chọn nguồn: **Dịch tự động bằng API** hoặc **Nhập file AI_Form**.
- Luồng API:
  1. Dùng báo cáo `_TC` của phiên hiện tại hoặc tải lên.
  2. Hiện số mục và chi phí ước tính (dùng `count_tokens`), kèm ô xác nhận về việc gửi dữ liệu ra ngoài.
  3. Nút "Dịch bằng AI", có thanh tiến trình theo lô.
  4. Dùng lại `_render_counts`, `_render_preview_rows`, ô chọn mã, xác nhận và nút áp dụng; fingerprint tính trên báo cáo cộng với JSON kết quả.
  5. Thêm nút tải **AI_Form đã điền** (dùng `build_ai_review_document` với cột cuối đã có bản dịch) để lưu hồ sơ và đối chiếu.
- Các mục bị chặn hoặc trống: cho phép "Dịch lại các mục lỗi" đối với mục `missing` hoặc lỗi API.

### 4. Cấu hình và nhật ký

- `.streamlit/secrets.toml` (không commit) và secrets trên Streamlit Cloud:
  ```toml
  [anthropic]
  api_key = "..."
  model = "claude-opus-5-5"
  effort = "medium"
  enabled_users = ["..."]   # để trống = mọi người dùng
  web_search = false
  ```
- `requirements.txt`: thêm `anthropic`.
- `usage_logger.py`: thêm cột số mục, token vào, token ra, token đọc cache, model, thời gian.
- `user_guide.py`: thêm mục hướng dẫn luồng API.

### 5. Kiểm thử

- Kiểm thử phạm vi dữ liệu: payload chỉ chứa các trường được phép, và tập mục gửi đi trùng khớp với AI_Form của cùng báo cáo.
- Kiểm thử đơn vị với client giả lập (không tốn phí): chia lô, hậu kiểm JSON, chuyển chuỗi tô vàng thành vị trí ký tự, mã lạ/trùng/thiếu.
- Kiểm thử hồi quy luồng AI_Form thủ công với các file QA sẵn có trong `scratch/review_*_qa`.
- Chạy thật trên 2–3 báo cáo mẫu (E, Hs, Ht), so chất lượng với bản dịch qua chat AI hiện nay, ghi lại chi phí mỗi báo cáo. Bước này tốn phí API nên cần bạn đồng ý trước.

## Giai đoạn

| GĐ | Nội dung | Kết quả |
|---|---|---|
| 0 | ~~Chốt chính sách dữ liệu~~ (đã chốt: chỉ phần AI_Form); có khóa API | Sẵn sàng gọi API |
| 1 ✅ | Tách `_validate` và kiểm thử hồi quy | Luồng cũ chạy y như trước (đã xong) |
| 2 ✅ | `ai_translate_api.py` cùng kiểm thử giả lập | Dịch được JSON, có thống kê token (đã xong, chưa gọi API thật) |
| 3 ✅ | Giao diện, cấu hình, nhật ký, hướng dẫn | Dùng được từ đầu đến cuối (đã xong, kiểm bằng client giả lập) |
| 4 | Chạy thử trên báo cáo thật, tinh chỉnh prompt và effort | Số liệu chất lượng và chi phí |
| 5 ✅ | Tra tên công ty bằng web search; lưu tên đã xác nhận | Đã xong (kiểm bằng client giả lập). Tra cứu bật bằng `web_search = true`. Tên đã xác nhận lưu ở bảng riêng `company_names_store.py` (CSV khi chạy máy, worksheet "CompanyNames" trên Google Sheet khi chạy cloud), dùng khi xử lý báo cáo và khi dịch bằng AI; admin tải về theo mẫu Dictionary_v3 |

## Cần bạn quyết định

1. ~~Có cho phép gửi nội dung báo cáo ra API không?~~ Đã chốt: cho phép, giới hạn đúng phần AI_Form. Còn mở: có cần thỏa thuận không lưu dữ liệu (ZDR) với nhà cung cấp không?
2. Có bật tra tên công ty qua web ngay từ đầu không, hay để giai đoạn 5?
3. Ai được dùng tính năng: mọi người dùng hay một danh sách cụ thể?
