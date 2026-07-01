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
               * **Bên phải (English):** Tải báo cáo tiếng Anh dịch tương ứng năm trước (ví dụ: `Dich 2025.docx`).
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
