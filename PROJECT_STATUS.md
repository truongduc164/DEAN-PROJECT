# Tình Trạng Dự Án DeanTran & Lộ Trình Phát Triển

## 1. Tình Trạng Hiện Tại (Current Status)
**Giai đoạn: Phase 0 - Skeleton (Khung Sườn)**
- **Mức độ hoàn thiện:** ~5%
- **Đã làm được:**
    - [x] Cấu trúc thư mục dự án chuẩn.
    - [x] File chạy chính `run.py` (App launch ok).
    - [x] Cấu hình môi trường (venv, requirements).
    - [x] RUNBOOK hướng dẫn cài đặt.
- **Chưa có:**
    - [ ] Hầu hết logic xử lý file Excel, Word, PPT.
    - [ ] Giao diện người dùng (UI) chi tiết (mới chỉ có nút kiểm tra môi trường).
    - [ ] Các module dịch thuật thực tế (Google, OpenAI, v.v.).
    - [ ] Unit Test (thư mục `tests` đang trống).

> **Kết luận:** App hiện tại chỉ là "bộ khung". Nó chạy lên được nhưng **chưa thể dùng để dịch thuật**.

---

## 2. Lộ Trình Phát Triển (Roadmap)

Dưới đây là các bước tiếp theo để đưa app vào hoạt động thực tế.

### Phase 1: Core Logic (Logic Cốt Lõi)
*Mục tiêu: Xử lý được file và gọi API dịch thuật (chưa cần giao diện đẹp).*
- [ ] **Reader Module**: Viết hàm đọc file Excel (`openpyxl`), Word (`python-docx`), PPT (`python-pptx`).
- [ ] **Translator Module**: Tích hợp Google Translate (free/key) hoặc AI (Gemini/OpenAI).
- [ ] **Writer Module**: Viết hàm ghi kết quả dịch ra file mới.
- [ ] **Unit Tests**: Viết test cho các hàm đọc/ghi/dịch này.

### Phase 2: User Interface (Giao Diện)
*Mục tiêu: Người dùng thao tác dễ dàng, không cần dính tới code.*
- [ ] Thiết kế Layout chính: Chọn file nguồn, chọn ngôn ngữ đích, nút "Bắt đầu".
- [ ] Hiển thị tiến trình (Progress Bar).
- [ ] Log window: Hiện thông báo lỗi/thành công ngay trên app.

### Phase 3: Integration & Polish
*Mục tiêu: Kết nối UI vào Core và hoàn thiện.*
- [ ] Nối nút "Bắt đầu" ở UI vào luồng xử lý của Core.
- [ ] Xử lý luồng (Threading) để app không bị đơ khi đang dịch.
- [ ] Đóng gói (Build .exe) để chạy máy khác không cần cài Python.

---

## 3. Việc Cần Làm Ngay (Next Actions)
Để "tiếp tục test app" như yêu cầu, chúng ta cần **viết code cho Phase 1**.

Bạn muốn bắt đầu module nào trước?
1.  **Reader (Đọc file):** Upload thử 1 file Excel/Word xem code đọc được chưa.
2.  **Translator (Dịch):** Test thử kết nối tới Google Translate/Gemini.
