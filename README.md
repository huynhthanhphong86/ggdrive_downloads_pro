# Google Drive Downloads Pro

Ứng dụng chuyên nghiệp giúp tải tài liệu, bài thuyết trình và **toàn bộ thư mục Google Drive** bị khóa quyền xem (*View-Only / No-Download*), hỗ trợ xuất ra **PowerPoint (.pptx)** chuẩn 100% slide, **Word (.docx)** chuẩn 100% bản gốc và **PDF (.pdf)** siêu nét (đầy đủ 100% trang cho cả Google Slides, Google Docs và tệp PDF Google Drive).

---

## 🚀 Cách 1: Sử dụng Giao diện Web Trực quan (GUI) - Khuyên Dùng

Giao diện Web hiện đại, dễ thao tác, hỗ trợ chọn thư mục lưu qua Windows Dialog hoặc các nút Preset nhanh, chọn định dạng tải, theo dõi tiến trình thời gian thực.

### 1. Khởi chạy ứng dụng:
- **Bật máy chủ & mở giao diện (Windows):** Nhấp đúp chuột vào tệp [`run_gui.bat`](file:///f:/PM/GGDrive_Downloads_Pro/run_gui.bat).
- **Tắt / Dừng máy chủ (Windows):** Nhấp đúp chuột vào tệp [`stop_server.bat`](file:///f:/PM/GGDrive_Downloads_Pro/stop_server.bat).
- **Hoặc chạy từ Terminal:**
  ```powershell
  python f:\PM\GGDrive_Downloads_Pro\app.py
  ```
- Ứng dụng sẽ tự động mở trình duyệt tại địa chỉ: `http://127.0.0.1:5000`

### 2. Các tính năng nổi bật:
- 🔗 **Hỗ trợ 4 loại liên kết Google Drive & Docs:**
  - 📊 **Bài thuyết trình Google Slides / PPTX View-Only:** Bắt trọn vẹn 100% các slide (ngay cả các slide có hoạt họa / animations) và xuất ra PowerPoint (.pptx) hoặc PDF (.pdf).
  - 📁 **Thư mục Google Drive:** Quét và tải hàng loạt tất cả các tệp con bên trong.
  - 📕 **Tệp PDF Google Drive View-Only:** Bắt trọn vẹn toàn bộ các trang PDF gốc (không giới hạn số trang) bằng công nghệ Adaptive Synchronized Scroller & Network Interception.
  - 📄 **Tài liệu Google Docs View-Only:** Xuất sang Word (.docx) chuẩn 100% bản gốc hoặc PDF Canvas Retina.
- 📁 **Chọn vị trí lưu trữ linh hoạt & tiện lợi:**
  - Nút **"📁 Duyệt..."** mở hộp thoại chọn thư mục Windows chuẩn (luôn hiển thị nổi bật trên cùng).
  - Các nút chọn vị trí nhanh 1 chạm: **📥 Downloads**, **🖥️ Desktop**, **💾 Workspace**.
  - Nút **"📂 Mở thư mục"** để xem các tệp đã tải ngay trong File Explorer.
- 📦 **Tùy chọn định dạng linh hoạt:**
  - **PowerPoint (.pptx):** Chuẩn tỷ lệ 16:9 / 4:3, tương thích 100% với Microsoft PowerPoint, WPS Office, Google Slides.
  - **Word (.docx):** Giữ nguyên bảng biểu nét viền đen, font Times New Roman 13pt, căn lề chuẩn 100% bản gốc.
  - **PDF (.pdf):** Đầy đủ 100% các trang siêu nét.
  - **Tất cả định dạng:** Tải trọn vẹn cả bộ định dạng cùng lúc.

---

## 🐍 Cách 2: Chạy từ Dòng lệnh (CLI)

```powershell
# 1. Tải Google Slides sang PowerPoint (.pptx):
python download_gdoc.py "https://docs.google.com/presentation/d/1rmO-DqqLrK86lyS4GZ37cLTb_gNlYqaf/edit" -f pptx

# 2. Tải Google Slides sang PDF:
python download_gdoc.py "https://docs.google.com/presentation/d/1rmO-DqqLrK86lyS4GZ37cLTb_gNlYqaf/edit" -f pdf

# 3. Tải tệp PDF từ Google Drive View-Only:
python download_gdoc.py "https://drive.google.com/file/d/1s-6MsJbKV_i3MQOa3mGbY03YUUKU5Ad_/view"

# 4. Tải 1 tệp Google Docs sang DOCX chuẩn bản gốc:
python download_gdoc.py "https://docs.google.com/document/d/1KN5IuYT_D3wzsx1tJ0rdOQxrmpSs3uCX/edit" -f docx
```

