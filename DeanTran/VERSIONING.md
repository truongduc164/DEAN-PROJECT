# CHIẾN LƯỢC QUẢN LÝ PHIÊN BẢN - DEANTRANS

Tài liệu này hướng dẫn bạn cách bảo trì và nâng cấp phiên bản cho DEANTRANS mà không làm vỡ các bản cũ đang chạy ổn định.

## 1. QUY TẮC ĐẶT TÊN PHIÊN BẢN (Semantic Versioning)
Tên phiên bản có dạng: **V{Major}.{Minor}.{Patch}** (Ví dụ: V1.0.0, V1.2.0, V2.0.0)

### Khi nào tăng Patch (Số cuối cùng - ví dụ V1.0.1)
- Khi bạn **SỬA LỖI** (bug fix) mà không thay đổi tính năng hiện tại.
- Ví dụ: Sửa lỗi dịch file PPT bị crash giữa chừng, đổi màu chữ từ xanh sang đỏ, sửa lỗi chính tả trên UI.
- Cách làm: Code thẳng trên nhánh hiện tại, sửa ở `version.py` thành `"1.0.1"`, chạy lại file `build_release.bat`.

### Khi nào tăng Minor (Số ở giữa - ví dụ V1.1.0)
- Khi bạn **THÊM TÍNH NĂNG MỚI** nhưng **VẪN GIỮ ĐƯỢC PHẦN CŨ**.
- Ví dụ: Ngày mai bạn thêm menu "Dịch PDF", hoặc "Dịch SRT". Tính năng Excel/PPT/Word cũ vẫn còn đó và hoạt động bình thường.
- Lưu ý: Chắc chắn tính năng cũ không bị mất hoặc lỗi. Khi tăng Minor, số Patch sẽ quay về 0 (Ví dụ từ V1.0.5 -> V1.1.0).

### Khi nào tăng Major (Số đầu tiên - ví dụ V2.0.0)
- Khi bạn **THAY ĐỔI LỚN** hoặc **ĐẬP ĐI XÂY LẠI** làm mất sự tương thích với bản cũ.
- Ví dụ: Xóa tính năng dịch PPT đi thay bằng thứ khác, viết lại toàn bộ giao diện, hoặc đổi ngôn ngữ code từ Python sang C#.
- Lưu ý: Rất ít khi phải tăng Major. Khi tăng Major, số Minor và Patch quay về 0.

---

## 2. QUY TRÌNH NÂNG CẤP AN TOÀN TRONG CODE (Không phá bản cũ)

Vì bạn là người mới code, nguyên tắc số 1 là **BACKUP (Sao lưu)**.

### Bước 1: Sao lưu mã nguồn (Source Code) hiện tại
Sau khi bản V1.0.0 chạy tốt, hãy COPY toàn bộ thư mục thư mục chứa code `DEANTRANS/` sang một chỗ khác, đổi tên thành `DEANTRANS_Source_V1.0.0`. 
Điều này đảm bảo nếu bạn code sai, bạn vẫn có bản gốc để phục hồi hoặc gửi cho tôi (để sửa).

*(Nếu sau này bạn biết dùng Git, hãy dùng `git branch` và `git tag v1.0.0`, không cần copy thủ công)*

### Bước 2: Nâng cấp ở bản code làm việc
- Bắt đầu sửa code thêm tính năng mới ở thư mục `DEANTRANS/` chính của bạn.
- Vào file `app/version.py`, sửa dòng `APP_VERSION = "1.1.0"`.

### Bước 3: Đóng gói lại
- Chạy lại file `build_release.bat`. Script này sẽ tự tạo ra thư mục `DEANTRANS_V1.1.0` hoàn chỉnh cho bạn.
- Bản `DEANTRANS_V1.0.0` cũ nằm ở ngoài máy bạn hoặc máy người dùng **sẽ không bị ảnh hưởng** (vì nó là một folder riêng, chạy độc lập).

---

## 3. LÀM SAO ĐỂ GIỮ CẤU HÌNH (PROMPT) CỦA NGƯỜI DÙNG?
Khi bạn đưa bản `V1.1.0` cho đồng nghiệp, họ sẽ kêu: *"Thế file configs tôi cấu hình mòn mỏi ở bản V1.0.0 đi đâu hết rồi???"*

**Cách xử lý vô cùng đơn giản (vì ta đã dùng mô hình ONEDIR):**
- Bảo họ: Tải thư mục `DEANTRANS_V1.1.0` mới về.
- Xóa thư mục `configs/` TRONG bản `V1.1.0` đi.
- **Copy thư mục `configs/` từ bản `V1.0.0` sang bản `V1.1.0`.**
- Vậy là toàn bộ API Key (đã lưu máy), App Settings, Custom Prompt được giữ nguyên. Ứng dụng code của bản V1.1.0 sẽ ngoan ngoãn đọc cấu hình từ thư mục `configs/` cũ. 
*(Lưu ý: Code settings_manager.py tôi viết đã lo vụ này, nếu config thiếu key mới của bản V1.1.0, nó sẽ tự nhét thêm vào mà không đè mất cấu hình cũ).*
