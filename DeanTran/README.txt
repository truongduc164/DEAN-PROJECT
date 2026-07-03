=======================================================
               DEANTRANS - V1.0.0
=======================================================

1. GIỚI THIỆU
DEANTRANS là ứng dụng hỗ trợ dịch thuật tài liệu tự động qua API của Google Gemini, chuyên dùng để dịch file Excel, PowerPoint và Word sang các ngôn ngữ khác nhau mà vẫn giữ nguyên format định dạng.

2. CÁCH SỬ DỤNG
- Bước 1: Mở ứng dụng bằng cách click đúp vào file "DEANTRANS.exe".
- Bước 2: Kéo thả file Excel hoặc PowerPoint vào ô danh sách file.
- Bước 3: Đảm bảo bạn đã cấu hình API Key trong mục "API Settings" (vần đăng nhập Admin để thấy). Nút Login nằm ở Menu -> Admin. Mật khẩu mặc định là: admin
- Bước 4: Nhấn nút [▶ Translate] để bắt đầu.

3. KẾT QUẢ DỊCH
- Các file sau khi dịch sẽ có thêm hậu tố (ví dụ: _Vi, _En) và nằm cùng thư mục với file gốc, hoặc nằm trong thư mục cài đặt nếu bạn đã thiết lập khác đi.
- Phần mềm KHÔNG BAO GIỜ ghi đè lên file đã dịch trước đó. Nếu đã có file _Vi, nó sẽ tự tạo file _Vi(1).

4. CÁC THƯ MỤC CẦN LƯU Ý
Trong thư mục chứa DEANTRANS.exe có các thư mục:
- configs/ : Chứa cấu hình ứng dụng, prompts, và API keys (mã hóa). KHÔNG xóa thư mục này. Bạn có thể sửa trực tiếp file .json bên trong nếu hiểu rõ cấu trúc.
- logs/    : Ghi lại các bản báo cáo mỗi khi dịch lỗi để bạn gửi kỹ thuật viên.
- outputs/ : Thư mục lưu trữ mặc định.

5. NÂNG CẤP VÀ XỬ LÝ SỰ CỐ
- Nếu app không dịch được, hãy mở configs/app_settings.json kiểm tra.
- Hãy backup lại thư mục configs/ trước khi chuyển sang phiên bản mới.
