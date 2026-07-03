import sys
from pathlib import Path

# Thêm đường dẫn thư mục gốc vào module search path để có thể import
project_root = Path(__file__).resolve().parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from app.core.auth_manager import AuthManager

def main():
    print("----------------------------------------")
    print("   Đổi mật khẩu Admin DeanTrans App     ")
    print("----------------------------------------")
    
    new_password = input("Nhập mật khẩu mới cho admin: ").strip()
    if not new_password:
        print("Mật khẩu không được để trống!")
        return

    # Khởi tạo AuthManager sẽ tự động load file configs/users.json
    am = AuthManager()
    
    # Tài khoản admin luôn được tạo theo mặc định từ AuthManager_load()
    success = am.change_password("admin", new_password)
    if success:
        print(f"Đã cập nhật mật khẩu cho admin thành công! (Lưu tại: {am._file})")
    else:
        print("Có lỗi xảy ra, không thể đổi mật khẩu.")

if __name__ == "__main__":
    main()
