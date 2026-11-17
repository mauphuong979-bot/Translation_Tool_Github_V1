import streamlit as st


QUICK_START = """
**Ứng dụng dịch báo cáo tài chính Word (.docx) từ tiếng Việt sang tiếng Anh (E), tiếng Hoa giản thể (Hs) hoặc phồn thể (Ht).**
Chỉ hỗ trợ file **.docx**; không hỗ trợ .doc, .pdf, .xlsx.

| Bước | Việc cần làm | Ở đâu |
|---|---|---|
| 1 | Tải báo cáo tiếng Việt cần dịch | **📂 Báo cáo cần dịch** |
| 2 | (Tùy chọn) Tải cặp báo cáo năm trước: bản tiếng Việt + bản đã dịch | **📅 Báo cáo năm trước** |
| 3 | Kiểm tra thông tin tự điền, chọn ngôn ngữ đích | **📝 Thông tin báo cáo** |
| 4 | Chọn **Chế độ xuất file**, bấm **🚀 Xử lý báo cáo** | Hàng nút xử lý |
| 5 | Tải file kết quả; nếu còn đoạn chưa dịch, tải **AI_Form.docx** gửi AI | Các nút **📥 Tải…** |
| 6 | Nhập file AI đã điền để ghi bản dịch bằng Track Changes | **📤 Nhập bản dịch từ AI** |

**Các nút ở góc phải trên cùng**
* **❓** – mở hướng dẫn này.
* **⚙️** – bật/tắt **Các bước xử lý** và tùy chọn **Tự mở/tải kết quả**.
* **👤** – thông tin tài khoản và nút **Đăng xuất**.
"""

PROCESS = """
**1. Báo cáo cần dịch**
* Tải file .docx tiếng Việt. Ứng dụng tự trích tên công ty, ngày kết thúc năm tài chính, ngày lập báo cáo, kỳ báo cáo và tên người ký vào **Thông tin báo cáo**.

**2. Báo cáo năm trước (tùy chọn, nên dùng)**
* Bật **Dùng báo cáo năm trước để gợi ý**, tải **Bản tiếng Việt năm trước** (trái) và **Bản đã dịch năm trước** (phải).
* Dưới mỗi ô hiện một dòng tóm tắt: tên file đầy đủ, tên công ty, mã số thuế, kỳ báo cáo – dùng để kiểm tra đã chọn đúng cặp file.
* Nếu ô **Tên công ty đã dịch** còn trống hoặc là giá trị mẫu, ứng dụng tự điền tên công ty lấy từ bản dịch năm trước. Hãy kiểm tra lại.

**3. Thông tin báo cáo**
* **Dịch sang:** `E` tiếng Anh, `Hs` tiếng Hoa giản thể, `Ht` tiếng Hoa phồn thể.
* **Tên công ty đã dịch** (chữ đỏ) là tên công ty bằng ngôn ngữ đích – luôn kiểm tra kỹ; chữ in hoa toàn bộ được tự chuyển sang dạng viết hoa chữ cái đầu.
* **Tên công ty tiếng Việt, ngày, kỳ báo cáo, người ký:** sửa nếu trích xuất chưa đúng.

**4. Chế độ xuất file và xử lý**
* **Comments + Track Changes (khuyên dùng):** báo cáo có Track Changes cho câu khớp 100% với báo cáo năm trước, comment gợi ý cho câu chưa khớp, kèm bảng **AI_Form.docx**.
* **Bản sạch:** báo cáo đã dịch, không chèn gợi ý trong đoạn.
* **Bản gốc:** kiểu cũ, gợi ý `[Original] / [Prior VN] / [Suggest]` chèn ngay trong văn bản.
* **Tất cả các bản:** tạo tất cả các file trên.
* Bấm **🚀 Xử lý báo cáo**. Sau khi xong, ô tóm tắt cho biết số Track Changes, số comment và số mục cần AI.

**5. Tên file kết quả**
* `…_TC.docx` – báo cáo Comments + Track Changes.
* `AI_Form.docx` – bảng/prompt gửi AI (tên cố định, ngắn gọn).
* `…_AI_TC.docx` – báo cáo sau khi nhập bản dịch AI.
* Việc ghép dữ liệu dựa trên mã ID bên trong file, không dựa vào tên file.

**6. Các bước xử lý (nút ⚙️)**
* Gồm 14 bước tự động: chuẩn hóa Unicode, CleanV, mẫu đoạn văn, dịch theo từ điển, font kép tiếng Hoa, đổi định dạng số, căn chỉnh bảng, định dạng ngày, textbox, dấu tên người ký, tô vàng tiếng Việt chưa dịch, gợi ý từ điển…
* Mặc định bật tất cả. Chỉ tắt khi biết rõ mục đích.
* **Tự mở/tải kết quả:** tự mở file kết quả bằng Word (máy Windows) hoặc tự tải xuống (bản cloud).
"""

AI_FLOW = """
**Khi nào dùng:** sau khi xử lý ở chế độ **Comments + Track Changes** hoặc **Tất cả các bản**, các đoạn còn tô vàng hoặc chỉ có gợi ý dưới 100% được đưa vào **AI_Form.docx**.

**Bước 1 – Tải AI_Form.docx**
* File trang ngang, mỗi mục là một bảng 5 cột: **Mã**, **Nội dung cần dịch**, **Tiếng Việt tham chiếu**, **Bản dịch tham chiếu**, **Bản dịch hoàn thiện**.
* Đầu file có **PROMPT chung cho AI** quy định ngôn ngữ đích, font, cách viết ngày, số.

**Bước 2 – Gửi cho AI**
* Dùng AI có khả năng **sửa và xuất file Word** (và truy cập web để tra tên công ty). Yêu cầu AI làm theo prompt ở đầu file, chỉ điền cột cuối.
* Quy tắc tên riêng trong prompt:
  * **Tên công ty** (ví dụ *Công ty TNHH ABC*): AI tra tên quốc tế trên **masothue.com** hoặc **thuvienphapluat.vn**. Không tìm thấy thì giữ nguyên tên tiếng Việt và **tô vàng** để bạn kiểm tra.
  * **Tên người** (ví dụ *Nguyễn Văn A*): chuyển thành không dấu – *Nguyen Van A*.
  * Thẻ mẫu như **【自年月日至年月日止階段】** được giữ nguyên, không làm trống ô.

**Bước 3 – Nhập lại (📤 Nhập bản dịch từ AI)**
* Chọn báo cáo `_TC` của phiên hiện tại (mặc định) hoặc tải đúng báo cáo đi kèm, rồi tải file AI đã điền.
* Bấm **Kiểm tra bản dịch AI**, xem trạng thái từng mã: *Sẵn sàng*, *Ô trống*, *Cần xử lý thủ công*, *Thiếu bảng*, *Đã áp dụng*, *Không đổi*, *Track Changes có sẵn*.
* Chọn các mã hợp lệ, đánh dấu xác nhận đã rà soát, bấm **Áp dụng bản dịch AI**, rồi tải **…_AI_TC.docx**.
* Bản dịch được ghi bằng Track Changes, giữ định dạng đoạn và cỡ chữ; font DFKai-SB cho chữ Hoa, Times New Roman cho chữ Latin/số. Phần AI **tô vàng** (tên công ty chưa tra được) vẫn được **giữ tô vàng** trong báo cáo.
* Comment của các mã đã áp dụng được xóa; Reject sẽ khôi phục câu gốc (không khôi phục comment).

**Các trường hợp không tự thay (cần xử lý tay)**
* Ô trống, mã trùng, AI sửa 4 cột đầu, nguồn/comment đã bị sửa trong báo cáo.
* Đoạn có liên kết Excel, field Word, hyperlink, định dạng hỗn hợp trong câu, bookmark bao một phần câu hoặc nhiều đoạn.
* Đoạn đã có Track Changes từ trước (ví dụ khớp 100% với báo cáo năm trước) – được giữ nguyên.
"""

READ_RESULT = """
**Trong báo cáo kết quả**
* **Tô vàng:** cụm tiếng Việt chưa dịch được.
* **Track Changes:** câu được thay tự động (khớp 100% báo cáo năm trước, hoặc bản dịch AI đã nhập). Accept để giữ, Reject để quay về câu gốc.
* **Comment:** gợi ý dịch dưới 100% kèm nguồn (báo cáo năm trước / từ điển) và tỷ lệ tương đồng.

**Chế độ Bản gốc (gợi ý trong văn bản)**
* **Màu đỏ – gợi ý từ báo cáo năm trước** (ưu tiên cao): `[Original]`, `[Prior VN]`, `[Suggest]`; khác biệt giữa `[Original]` và `[Prior VN]` được tô xanh lá.
* **Màu xanh dương – gợi ý từ từ điển:** `[Original]`, `[Dict VN]`, `[Suggest]`.
* **`[Original]` đứng riêng:** câu tô vàng không có gợi ý (khi bật bước 13).

**Tab 📖 Tra từ điển**
* Tra cứu từ điển theo tiếng Việt hoặc bản dịch; chọn ngôn ngữ hiển thị E / Hs / Ht hoặc tất cả.
"""

TROUBLESHOOTING = """
* **Không tải được file:** kiểm tra đúng đuôi .docx, file không bị khóa (Protected/Read-only), không đang mở ở chế độ chỉ đọc.
* **Xử lý lâu:** báo cáo lớn cần vài chục giây; không tải lại trang khi đang chạy.
* **Tải lại trang (F5) sẽ phải đăng nhập lại** và mất các file đã tải lên trong phiên.
* **Thông tin tự điền sai:** sửa trực tiếp trong **Thông tin báo cáo** trước khi bấm **Xử lý báo cáo**.
* **Nhập AI báo lỗi “thuộc báo cáo khác hoặc ngôn ngữ dịch khác”:** file AI không cùng báo cáo/ngôn ngữ với báo cáo `_TC` đang chọn – tải đúng cặp file.
* **Worksheet cũ không có mã ID:** xử lý lại báo cáo nguồn để tạo AI_Form.docx mới.
* **Bản cloud ngủ đông:** lần truy cập đầu sau thời gian dài cần đợi 1–2 phút.
* **Quyền Admin:** tab **📈 Quản trị** có nhật ký sử dụng, quản lý người dùng, đồng bộ mẫu Excel và tải từ điển đã giải mã.
"""


@st.dialog("📘 Hướng dẫn sử dụng", width="large")
def show_user_guide():
    sections = st.tabs(["🚀 Bắt đầu nhanh", "📂 Xử lý báo cáo", "🤖 Dịch bằng AI", "🔍 Đọc kết quả", "🛠️ Sự cố thường gặp"])
    for tab, content in zip(sections, (QUICK_START, PROCESS, AI_FLOW, READ_RESULT, TROUBLESHOOTING)):
        with tab:
            st.markdown(content)
