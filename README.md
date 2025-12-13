# XLS to XLSX Converter

Công cụ tự động chuyển đổi hàng loạt file Excel định dạng cũ (`.xls`) sang định dạng mới (`.xlsx`) trên Windows. Tool sử dụng bộ máy Excel (Excel Engine) để đảm bảo tính toàn vẹn dữ liệu, hỗ trợ quét thư mục nhiều cấp và xuất báo cáo chi tiết.

## 🚀 Tính Năng Chính

*   **Giao diện đồ họa (GUI):** Dễ sử dụng, không cần gõ lệnh.
*   **Quét đệ quy:** Tự động tìm file `.xls` trong thư mục được chọn và tất cả các thư mục con.
*   **Chuyển đổi chuẩn xác:** Sử dụng thư viện `win32com` điều khiển trực tiếp Microsoft Excel để chuyển đổi (Save As), giảm thiểu lỗi định dạng so với các công cụ convert dòng lệnh.
*   **Tùy chọn dọn dẹp:** Cho phép xóa file `.xls` gốc sau khi chuyển đổi thành công.
*   **Hệ thống Log chi tiết:**
    *   Hiển thị trạng thái thời gian thực trên giao diện.
    *   Tự động xuất file `conversion_log.txt` tổng hợp danh sách file Thành công/Thất bại tại thư mục làm việc.

## 📋 Yêu Cầu Hệ Thống

*   **Hệ điều hành:** Windows 10, Windows 11.
*   **Phần mềm bắt buộc:** Máy tính **phải cài đặt Microsoft Excel** (2010, 2013, 2016, 2019, 365...).
    *   *Lưu ý:* Excel cần được kích hoạt bản quyền (Activated) để tránh các hộp thoại pop-up làm gián đoạn quá trình chạy tự động.

## 📖 Hướng Dẫn Sử Dụng

1.  **Chạy ứng dụng:** Mở file `.exe` hoặc chạy script Python.
2.  **Chọn thư mục:** Nhấn nút "Chọn Thư Mục" để trỏ đến folder chứa các file Excel cũ.
3.  **Cấu hình:**
    *   Tick vào ô *"Xóa file .xls cũ..."* nếu bạn muốn dọn dẹp ổ cứng sau khi convert.
    *   Bỏ tick nếu muốn giữ lại bản gốc để backup.
4.  **Bắt đầu:** Nhấn nút **BẮT ĐẦU CHUYỂN ĐỔI**.
5.  **Kết quả:**
    *   Sau khi chạy xong, file log chi tiết sẽ được lưu tại đường dẫn: `[Thư mục của bạn]\conversion_log.txt`.

## 🛠️ Dành Cho Nhà Phát Triển (Developer)

Nếu bạn muốn chạy từ mã nguồn hoặc chỉnh sửa code:

### 1. Cài đặt môi trường
Yêu cầu Python 3.x. Cài đặt các thư viện cần thiết:

```bash
pip install pywin32
```

### 2. Chạy mã nguồn
```bash
python converter_log.py
```

### 3. Đóng gói thành file EXE
Sử dụng **PyInstaller** để build file chạy độc lập. Cần lưu ý thêm `hidden-import` để thư viện `win32com` hoạt động ổn định.

Cài đặt PyInstaller:
```bash
pip install pyinstaller
```

Lệnh Build (chạy trong Terminal/CMD):
```bash
pyinstaller --noconsole --onefile --hidden-import="win32com.client" --hidden-import="pythoncom" converter_log.py
```
*File `.exe` sẽ nằm trong thư mục `dist/`.*

## ⚠️ Các Trường Hợp Cần Lưu Ý

1.  **File có Mật khẩu:** Tool sẽ **bỏ qua** và ghi vào mục THẤT BẠI các file yêu cầu mật khẩu để mở (Password to Open).
2.  **File Macro (.xlsm):** File `.xls` chứa Macro khi chuyển sang `.xlsx` sẽ bị **mất Macro** (do định dạng xlsx không hỗ trợ code VBA).
3.  **Lỗi treo Excel:** Nếu đang chạy mà bạn mở một file Excel khác lên can thiệp, tiến trình có thể bị gián đoạn. Nên để máy rảnh khi đang convert số lượng lớn.

## 📝 Định dạng File Log
File `conversion_log.txt` sẽ có cấu trúc như sau:

```text
BÁO CÁO CHUYỂN ĐỔI EXCEL
Thời gian: 2025-12-13 10:00:00
Thư mục quét: D:\Data\OldExcel
==================================================

THÀNH CÔNG: 50
- D:\Data\OldExcel\Sub1\file_01.xls -> D:\Data\OldExcel\Sub1\file_01.xlsx
...

------------------------------

THẤT BẠI: 02
- D:\Data\OldExcel\Sub2\error.xls | Lỗi: Password required
...
```

---
*   **Author:** @danhcp
*   **Version:** 1.0.0
