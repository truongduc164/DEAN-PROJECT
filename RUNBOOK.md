# Hướng Dẫn Cài Đặt & Chạy Project DeanTran (RUNBOOK)

Làm theo checklist này để thiết lập lại môi trường chạy từ ổ D.

---

## 1. Chuẩn Bị & Dọn Dẹp
Mở **Command Prompt (CMD)** (không cần Run as Admin trừ khi có lỗi quyền).

Cho phép chạy script nếu dùng PowerShell (Optional):
```cmd
powershell Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Di chuyển vào thư mục dự án:
```cmd
D:
cd "D:\0. Lập trình\1.DEANTRANS\DeanTran"
```
> **Output Mong Đợi:** cmd hiển thị `D:\0. Lập trình\1.DEANTRANS\DeanTran>`

Xóa venv cũ (nếu có):
```cmd
if exist venv rmdir /s /q venv
if exist .venv rmdir /s /q .venv
```


---

## 2. Tạo Môi Trường Ảo (Virtual Environment)
Tạo venv mới ngay tại thư mục hiện tại:
```cmd
python -m venv venv
```
> **Output Mong Đợi:** Folder `venv` xuất hiện trong thư mục `DeanTran`.

Kích hoạt venv:
```cmd
venv\Scripts\activate
```
> **Output Mong Đợi:** Dòng lệnh hiện `(venv) D:\0. Lập trình\1.DEANTRANS\DeanTran>`

---

## 3. Cài Đặt Thư Viện
Cập nhật pip và cài dependency:
```cmd
python -m pip install --upgrade pip
pip install -r requirements.txt
```
> **Output Mong Đợi:** Pip setup chạy, sau đó cài các gói như PySide6, openpyxl, google-generativeai... Cuối cùng báo `Successfully installed ...`

---

## 4. Chạy Ứng Dụng
Chạy lệnh kiểm tra cuối cùng:
```cmd
py run.py
```
*(Nếu `py` không nhận venv, dùng lệnh: `python run.py`)*

> **Output Mong Đợi:**
> - Cửa sổ ứng dụng "DeanTran Translation App" hiện lên.
> - Console in ra:
>   `[INFO] Project Root: D:\0. Lập trình\1.DEANTRANS\DeanTran`
>   `[INFO] Output Dir:   D:\0. Lập trình\1.DEANTRANS\outputs`

---

## 5. Khắc Phục Lỗi Thường Gặp (Troubleshooting)

| Lỗi | Cách Sửa |
| :--- | :--- |
| **'python' is not recognized...** | Cài lại Python, nhớ tích chọn **"Add Python to PATH"**. Re-open CMD. |
| **'pip' is not recognized...** | Dùng `python -m pip` thay vì `pip` để gọi module pip trực tiếp. |
| **Lỗi Security (PSSecurityException)** | Chạy lệnh ở Bước 1 để cấp quyền cho Script. |
| **Không tìm thấy module** | Quên `activate` venv? Chạy lại bước kích hoạt và cài requirements. |
| **Code vẫn trỏ về ổ C** | Code đã được kiểm tra (`run.py`). Nó tự động cấu hình PATH theo vị trí file hiện tại. Output mặc định là `D:\...\outputs`. |
