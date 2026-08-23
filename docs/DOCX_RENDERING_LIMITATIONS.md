# Giới Hạn Kỹ Thuật Rendering & Giải Pháp Khắc Phục (Rendering Limitations) - Docx-Agent V2.1

## 1. Giới Hạn & Độ Bao Phủ Hiện Tại (Current Known Limitations)

Dưới đây là các giới hạn kỹ thuật được phân loại theo mức độ ưu tiên và hiện trạng xử lý trong Docx-Agent V2.1:

| Hạng Mục | Hành Vi Hiện Tại | Mức Độ Ảnh Hưởng | Giải Pháp / Cơ Chế Khắc Phục (Mitigation) |
| :--- | :--- | :---: | :--- |
| **Vẽ Đồ Họa VML Cổ Điển (`v:shape`)** | Bỏ qua hoặc lưu vào `UnsupportedBlock` | Thấp (Chủ yếu ở tài liệu Word 97-2003) | Bảo toàn mã XML thô nguyên bản (`raw_xml`) trong Canonical Model, khôi phục 100% khi ghi lại ra tệp `.docx`. |
| **Trường Động Phức Tạp (`w:fldSimple` / `w:instrText`)** | Đọc nội dung tĩnh hoặc cập nhật trường khi mở Word | Trung bình | Thêm thuộc tính `<w:updateFields w:val="true"/>` trong `settings.xml` để Microsoft Word tự động tính toán lại mục lục, số trang và công thức khi mở tệp. |
| **SmartArt & Khối 3D** | Chuyển đổi thành khối bảo toàn `UnsupportedBlock` | Thấp | Không làm vỡ định dạng xung quanh; bảo toàn nguyên vẹn quan hệ (relationships) trong gói OPC. |
| **Hiệu Ứng Chữ Nghệ Thuật (WordArt)** | Chuyển thành văn bản thô kèm cảnh báo trong log | Rất thấp | Giữ trọn vẹn nội dung chữ và thuộc tính kích thước để người dùng dễ dàng chỉnh sửa. |
| **Bảng Lồng Nhau Đa Tầng (Nested Tables)** | Hiển thị phẳng hóa (Flattened) trong Webview | Trung bình | Bảo toàn cấu trúc cây OOXML trong Python engine để xuất file Word chuẩn chỉ. |

---

## 2. Chiến Lược Mở Rộng Tiếp Theo (Future Roadmap)

1. **True Font Metrics Rendering**: Tích hợp module tính toán chiều rộng ký tự thực tế (Glyph Advance Widths) dựa trên FreeType/HarfBuzz để đạt độ chính xác ngắt dòng 99.9% so với Microsoft Word DirectWrite Engine.
2. **Interactive Table Border Resizer**: Cho phép kéo giãn kích thước cột và hàng trực tiếp trên giao diện trực quan của Webview.
3. **Multi-Column Layout (`w:cols`)**: Hỗ trợ chia cột báo và tài liệu nghiên cứu khoa học IEEE/ACM 2 cột.
