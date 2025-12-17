# MSO Convert

<p align="center">
  <img src="MSOConvert.png" alt="Giao diện MSO Convert" max-width="700">
</p>

## 🚀 Tính Năng Chính

*   **Giao diện đồ họa (GUI):** Dễ sử dụng, trực quan, không cần gõ lệnh.
*   **Hỗ trợ đa định dạng:** Chuyển đổi các định dạng phổ biến nhất của Excel, Word và PowerPoint.
*   **Tùy chọn định dạng:** Cho phép người dùng chọn loại tệp muốn chuyển đổi (Excel, Word, PowerPoint).
*   **Quét thư mục linh hoạt:**
    *   Tùy chọn quét thư mục được chọn và tất cả các thư mục con của nó (mặc định).
    *   Tùy chọn chỉ quét các tệp trong thư mục cấp cao nhất.
*   **Xử lý xung đột thông minh:** Nếu tệp đích đã tồn tại, công cụ sẽ tự động tạo một phiên bản mới với hậu tố {name}_ thay vì ghi đè.
*   **Tùy chọn dọn dẹp:** Cho phép xóa tệp gốc sau khi chuyển đổi thành công.
*   **Hệ thống Log chi tiết:**
    *   Hiển thị trạng thái thời gian thực trên giao diện.
    *   Tự động xuất file `conversion_log.txt` tổng hợp danh sách tệp Thành công/Thất bại tại thư mục đã quét.

## 📋 Yêu Cầu Hệ Thống

*   **Hệ điều hành:** Windows 10, Windows 11.
*   **Phần mềm bắt buộc:** Máy tính **phải cài đặt bộ Microsoft Office** (2010, 2013, 2016, 2019, 365...).
    *   *Lưu ý:* Bộ Office cần được kích hoạt bản quyền (Activated) để tránh các hộp thoại pop-up có thể làm gián đoạn quá trình chạy tự động.

## 📖 Hướng Dẫn Sử Dụng

1.  **Chạy ứng dụng:** Mở file `MSO Convert.exe` (nếu đã build) hoặc chạy script Python.
2.  **Chọn thư mục:** Nhấn nút "Chọn Thư Mục" để trỏ đến folder chứa các tệp Office cũ.
3.  **Cấu hình:**
    *   **Bao gồm các thư mục con:** Chọn nếu bạn muốn quét tất cả các thư mục bên trong.
    *   **Xóa file cũ...:** Chọn nếu bạn muốn dọn dẹp ổ cứng sau khi convert.
    *   **Chọn định dạng:** Chọn các loại tệp bạn muốn chuyển đổi (Excel, Word, PowerPoint).
4.  **Bắt đầu:** Nhấn nút **BẮT ĐẦU CHUYỂN ĐỔI**.
5.  **Kết quả:**
    *   Theo dõi tiến trình trên cửa sổ ứng dụng.
    *   Sau khi chạy xong, file log chi tiết sẽ được lưu tại: `[Thư mục của bạn]\conversion_log.txt`.

## 🛠️ Dành Cho Nhà Phát Triển (Developer)

Nếu bạn muốn chạy từ mã nguồn hoặc chỉnh sửa code.

### 1. Cài đặt môi trường
Yêu cầu Python 3.x. Cài đặt các thư viện cần thiết bằng tệp `requirements.txt`:

```bash
pip install -r requirements.txt
```

### 2. Chạy mã nguồn
```bash
python "Office Converter.py"
```

### 3. Đóng gói thành file EXE
Tệp `build.bat` đã được cấu hình sẵn để đóng gói ứng dụng một cách dễ dàng. Chỉ cần chạy tệp `build.bat`.

Nó sẽ tự động:
- Cài đặt các thư viện cần thiết.
- Chạy PyInstaller với các tham số tối ưu (nhúng icon, thêm data, ẩn console).
- Dọn dẹp các tệp tạm sau khi build xong.

*File `MSO Convert.exe` hoàn chỉnh sẽ nằm trong thư mục `dist/`.*

---
*   **Author:** @danhcp
*   **Version:** 2.0.0
