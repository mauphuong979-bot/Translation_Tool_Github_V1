import streamlit as st

def render_user_guide_sidebar():
    """
    Renders an interactive, step-by-step User Guide inside the Streamlit sidebar.
    Optimized for new users, using English UI headers and Vietnamese content.
    The navigation buttons are rendered compactly at the top of the guide.
    """
    # 1. Initialize session state for user guide step if not existing
    if "user_guide_step" not in st.session_state:
        st.session_state.user_guide_step = 1

    # 2. Render the outer expander to keep it compact
    with st.sidebar.expander("📘 Quick User Guide", expanded=False):
        # Inject custom CSS to make these specific expander buttons very compact
        st.markdown("""
            <style>
            [data-testid="stSidebar"] .stExpander div.stButton > button {
                padding: 0.3rem 0.5rem !important;
                font-size: 0.8rem !important;
                min-height: 28px !important;
                height: 30px !important;
                line-height: 1.2 !important;
            }
            </style>
        """, unsafe_allow_html=True)

        total_steps = 9
        step = st.session_state.user_guide_step

        # Render navigation buttons at the TOP of the expander
        col_prev, col_next = st.columns(2)
        
        with col_prev:
            if step > 1:
                if st.button("← Previous", key="guide_prev_btn", use_container_width=True):
                    st.session_state.user_guide_step -= 1
                    st.rerun()
            else:
                st.button("← Previous", key="guide_prev_btn_disabled", disabled=True, use_container_width=True)

        with col_next:
            if step < total_steps:
                if st.button("Next →", key="guide_next_btn", use_container_width=True):
                    st.session_state.user_guide_step += 1
                    st.rerun()
            else:
                st.button("Next →", key="guide_next_btn_disabled", disabled=True, use_container_width=True)

        # Divider between navigation and content
        st.markdown("<hr style='margin: 8px 0; border: none; border-top: 1px solid #e9ecef;' />", unsafe_allow_html=True)

        # Define headers and content for each step
        if step == 1:
            st.markdown("### **Step 1/9: Overview & Supported Formats**")
            st.markdown("""
            **Tổng quan:**
            Công cụ chuyên dùng để tự động hóa quá trình phân tích, căn chỉnh định dạng bảng biểu và dịch thuật báo cáo tài chính chuyên nghiệp từ Tiếng Việt sang các ngôn ngữ khác (Tiếng Anh, Tiếng Trung Giản thể, Tiếng Trung Phồn thể).
            
            **Supported File Formats (Định dạng hỗ trợ):**
            * Chỉ hỗ trợ duy nhất tài liệu Word có đuôi mở rộng **`.docx`**.
            * Các định dạng khác như `.doc` (Word cũ), `.pdf`, `.xlsx` (Excel) hoặc `.txt` đều **không được hỗ trợ** bởi hệ thống.
            """)

        elif step == 2:
            st.markdown("### **Step 2/9: Upload Document**")
            st.markdown("""
            **Tải tệp tin đầu vào:**
            1. Tại giao diện chính, di chuyển tới tab **🚀 Process**.
            2. Nhấp vào vùng kéo thả hoặc bấm nút duyệt tại mục **Upload Financial Statements** để chọn tệp báo cáo tài chính `.docx` cần xử lý từ máy tính (ví dụ: báo cáo của năm nay).
            3. Hệ thống sẽ tự động quét tệp tin ngay khi tải lên để trích xuất các thông tin siêu dữ liệu cần thiết.
            """)

        elif step == 3:
            st.markdown("### **Step 3/9: Prior Year Suggestion**")
            st.markdown("""
            **Sử dụng báo cáo năm trước để gợi ý dịch:**
            1. Tích chọn **Suggest based on prior year report** (mặc định có chọn). Khi bỏ chọn, các giao diện liên quan sẽ tự động ẩn đi.
            2. **Tải tệp đối chiếu song song:**
               * **Bên trái (Vietnamese):** Tải báo cáo tiếng Việt năm trước (ví dụ: `V 2025.docx`).
               * **Bên phải (Translated):** Tải báo cáo dịch tương ứng của năm trước (Tiếng Anh hoặc Tiếng Hoa, ví dụ: `Dich 2025.docx`).
            3. **Trích xuất Metadata:** Hệ thống sẽ tự động trích xuất bảng đầu tiên của 2 báo cáo năm trước (Tên doanh nghiệp, mã số thuế, niên độ...) và hiển thị trong mục thông tin để bạn dễ đối chiếu.
            """)

        elif step == 4:
            st.markdown("### **Step 4/9: Metadata Selection**")
            st.markdown("""
            **Tùy chọn dịch & Chỉnh sửa:**
            1. Rà soát thông tin được trích xuất tự động tại phần **Report Metadata** (Tên doanh nghiệp, ngày lập báo cáo, niên độ, danh sách người ký).
            2. **Translate into:** Chọn ngôn ngữ dịch mục tiêu (`E` cho Tiếng Anh, `Hs` cho Tiếng Trung Giản thể, `Ht` cho Tiếng Trung Phồn thể).
            3. **Translated Name:** Nhập tên dịch của công ty. Trường này sẽ tự động viết hoa chữ cái đầu (Proper Case) và được in đỏ nổi bật trên giao diện để nhắc nhở kiểm tra kỹ.
            """)

        elif step == 5:
            st.markdown("### **Step 5/9: Pipeline Configuration**")
            st.markdown("""
            **Cấu hình quy trình xử lý:**
            * Phía dưới phần hướng dẫn này trên thanh Sidebar là mục **PROCESSING PIPELINE**.
            * Tập hợp các bước xử lý tự động (chuẩn hóa Unicode, căn chỉnh bảng, phông chữ kép CJK...).
            * **Lưu ý các tính năng thông minh mới:**
              * **11. Highlight VN:** Tự động **xóa toàn bộ highlight cũ** trước khi đánh dấu chữ tiếng Việt mới để tránh nhầm lẫn.
              * **13. Always Show [Original]:** Luôn hiển thị khối gốc tiếng Việt cho các câu chưa được dịch để dễ so sánh (mặc định được bật).
            """)

        elif step == 6:
            st.markdown("### **Step 6/9: Execute Process**")
            st.markdown("""
            **Kích hoạt xử lý báo cáo:**
            1. Sau khi hoàn thành rà soát các trường thông tin, bấm nút **`🚀 Process Report`** màu tím/xanh nổi bật dưới phần Metadata ở giao diện chính để thực thi.
            2. Hệ thống sẽ tiến hành chạy ngầm quy trình biên dịch toàn bộ văn bản và định dạng lại các bảng biểu XML phức tạp.
            3. Thời gian xử lý dao động từ vài giây tới một phút tùy thuộc vào độ lớn của tài liệu.
            """)

        elif step == 7:
            st.markdown("### **Step 7/9: Download Output Files**")
            st.markdown("""
            **Nhận tệp tin kết quả:**
            * **Môi trường cục bộ (Windows):** Tệp Word kết quả đã xử lý sẽ tự động được mở trực tiếp bằng ứng dụng Microsoft Word của bạn (nếu chọn *Auto-open/download result*).
            * **Môi trường Cloud hoặc tải thủ công:** Bấm nút **`📥 Download Report (.docx)`** màu xanh nổi bật tại màn hình chính để tải xuống tệp tin kết quả.
            * **Quyền Admin:** Bạn có thể tải bản từ điển đối chiếu đã giải mã thẻ động qua nút **`Download Resolved Dictionary (.xlsx)`** ở tab **Admin**.
            * **Khớp báo cáo kỳ trước 100%:** Ứng dụng tạo Track Changes ngay ở bước xử lý, áp dụng font theo ngôn ngữ đích và không đưa đoạn đã khớp này vào worksheet AI, kể cả khi từ điển có thêm gợi ý dưới 100%. Đoạn đã thay thành công không được tạo thêm comment, nên sau Accept không có comment thừa cần xóa. Nếu cấu trúc không thể thay an toàn, comment ghi rõ cần kiểm tra thủ công; comment có sẵn của người dùng luôn được giữ lại.
            * **Tên file ngắn:** **_TC.docx** là báo cáo Comments + Track Changes; **_AI_Form.docx** là bảng/prompt gửi AI; **_AI_TC.docx** là báo cáo đã áp dụng bản dịch AI bằng Track Changes. Ứng dụng giữ tên báo cáo, ngôn ngữ và thời điểm xử lý, bỏ phần **tool** thừa và không cộng dồn hậu tố **review_AI_reviewed**. Việc ghép nội dung vẫn dựa trên mã ID và dữ liệu liên kết, không dựa vào tên file.
            * **Bảng hoàn thiện bản dịch bằng AI:** Chọn **`Comments + Track Changes`** hoặc **`All outputs`**, rồi tải **`Download AI Translation Worksheet`**. File gồm gợi ý chưa khớp và các đoạn tiếng Việt tô vàng chưa tìm thấy tham chiếu; mỗi đoạn có comment và mã riêng. Hai cột tham chiếu có thể trống; AI vẫn phải dịch đầy đủ nguồn hiện tại.
            * File này có bố cục trang ngang, mỗi comment một bảng năm cột: **Mã**, **Nội dung cần dịch**, **Tiếng Việt tham chiếu**, **Bản dịch tham chiếu**, **Bản dịch hoàn thiện**. Mã đối chiếu giúp nhập đúng comment; không có vị trí tài liệu hay thông tin thừa. Một prompt chung quy định ngôn ngữ đích và định dạng tiếng Anh hoặc tiếng Hoa.
            * **Tên riêng:** Prompt yêu cầu giữ nguyên tên pháp lý công ty (ví dụ **Công ty TNHH ABC**) và tên người (ví dụ **Nguyễn Văn A**), không tự dịch, phiên âm hoặc bỏ dấu. Chỉ dùng tên đích chính thức khi được người dùng cung cấp/xác nhận. Cách gọi chung **Công ty** trong câu vẫn được dịch theo ngữ cảnh; nếu cả ô chỉ là tên riêng, AI điền lại đúng tên nguồn, không để trống.
            * **Thẻ mẫu và bookmark:** Thẻ hiển thị như **【自年月日至年月日止階段】** không chặn nhập bản dịch hoặc tạo Track Changes. Nếu nguồn chưa cung cấp giá trị thay thế, AI giữ nguyên thẻ và dịch phần còn lại, không chèn cảnh báo vào ô kết quả. Bookmark Word rỗng ở đầu/cuối hoặc bao toàn bộ đoạn được bảo toàn và không còn bị chặn chỉ vì có bookmark; bookmark lỗi, bao một phần câu hoặc đi qua nhiều đoạn vẫn cần xử lý thủ công để tránh làm hỏng liên kết.
            * Tải file này lên AI có khả năng chỉnh sửa và xuất Word, yêu cầu thực hiện prompt ở đầu file, chỉ điền cột cuối và giữ nguyên bốn cột đầu. Tải lại file Word đã hoàn thiện; có thể copy cột cuối vào báo cáo hoặc nhập tự động như sau.
            * Mở **Import AI translations** trong tab xử lý. Dùng bản **Comments + Track Changes** của phiên hiện tại hoặc tải lên đúng bản báo cáo đi kèm worksheet; sau đó đính kèm worksheet đã được AI điền.
            * Bấm **Check AI translations**, xem trạng thái từng mã và chọn các bản dịch hợp lệ. Ô trống, mã trùng, nguồn/comment bị sửa, liên kết Excel hay cấu trúc phức tạp (kể cả định dạng nhấn mạnh hỗn hợp không thể ánh xạ an toàn) sẽ không được tự thay thế. Không ghép gần đúng theo nội dung; kiểm tra này không thay cho rà soát chất lượng dịch.
            * Giao diện nhập AI có hai cột tải file, các thẻ thống kê và bảng rà soát gọn. Cột **Note** được ưu tiên độ rộng; chọn mã tại **Full row details and note** để đọc ghi chú đầy đủ, mở **Compare source and AI translation** để so sánh nội dung mà không làm bảng quá cao.
            * **Source Track Changes** là thay đổi đã có trong báo cáo trước khi nhập AI, chẳng hạn bản dịch khớp 100% từ báo cáo kỳ trước ở bước xử lý ban đầu. Mã này không được áp dụng thêm từ AI; bỏ chọn không hoàn tác các thay đổi có sẵn. Sau khi áp dụng, giao diện liệt kê riêng các mã AI thực sự đã thay và các mã giữ nguyên.
            * Xác nhận đã rà soát rồi bấm **Apply AI translations** và tải báo cáo mới qua **Download report with AI Track Changes**. Ứng dụng ghi xóa đoạn cũ và chèn bản dịch bằng Track Changes, giữ định dạng đoạn, cỡ chữ nguồn và quy tắc font DFKai-SB cho tiếng Hoa, Times New Roman cho chữ Latin/số; chỉ xóa comment của những dòng đã áp dụng thành công.
            * **Lưu ý:** Giữ bản báo cáo gốc để dự phòng. Reject khôi phục nội dung cũ nhưng không khôi phục comment đã xóa. Các file cũ chưa có mã liên kết phải được tạo lại bằng phiên bản mới; không thể nhập ngược an toàn.
            * Bản dịch chèn mới không giữ màu tô vàng đánh dấu cần dịch, kể cả màu kế thừa từ style; bản gốc trong phần xóa vẫn được giữ để Reject. Sau khi cập nhật ứng dụng, hãy xử lý lại báo cáo nguồn để tạo đủ comment, mã và worksheet mới; worksheet cũ không có các đoạn mới được phát hiện.
            """)

        elif step == 8:
            st.markdown("### **Step 8/9: Smart Output Verification**")
            st.markdown("""
            **Cách đọc và rà soát kết quả dịch thông minh:**
            * **Vùng bôi màu vàng (Yellow Highlights):** Đây là các cụm từ Tiếng Việt chưa được dịch.
            * **Định dạng gợi ý trực quan (Bold & Color):** Các nhãn được tách dòng riêng và in đậm để phân biệt với nội dung.
            * **Gợi ý Báo cáo năm trước (Màu Đỏ - Có độ ưu tiên cao hơn):**
              * Hiển thị dạng: `[Original]`, `[Prior VN]`, `[Suggest]`.
              * **Tự động đối chiếu:** Hệ thống sẽ tự động so sánh và **tô highlight xanh lá sáng (bright green)** cho những từ/cụm từ khác biệt ở phần nội dung **`[Original]`** so với **`[Prior VN]`** để bạn chú ý và hiệu chỉnh chính xác.
            * **Gợi ý Từ điển (Màu Xanh Dương):**
              * Hiển thị dạng: `[Original]`, `[Dict VN]`, `[Suggest]`. Không tô highlight xanh lá.
            * **Standalone `[Original]` (Màu Xanh Dương):**
              * Hiển thị cho các câu tô vàng còn lại không có so khớp (khi bật bước 13).
            """)

        elif step == 9:
            st.markdown("### **Step 9/9: Common Issues & Fixes**")
            st.markdown("""
            **Common Issues & Fixes (Sự cố thường gặp):**
            * **Không tải được tệp/Lỗi định dạng:** Đảm bảo tệp tin đúng định dạng `.docx`, không bị lỗi mã hóa XML và không ở trạng thái khóa bảo vệ (Protected/Read-only).
            * **Xử lý lâu với tệp lớn:** Việc phân tích và dịch từng phần tử XML bảng biểu báo cáo tài chính lớn cần thời gian xử lý. Hãy kiên nhẫn và không tải lại trang khi hệ thống đang chạy.
            * **Không tải được tệp kết quả:** Thử tải lại bằng cách bấm nút tải thủ công `Download Report (.docx)` hiển thị bền vững ở giao diện chính.
            * **Streamlit Cloud ngủ đông:** Nếu ứng dụng không được sử dụng trong thời gian dài, Streamlit Cloud sẽ tự động chuyển sang trạng thái ngủ (Sleep). Khi truy cập lại, vui lòng đợi 1-2 phút để máy chủ ảo khởi chạy lại hệ thống trước khi bắt đầu tải tệp.
            """)
