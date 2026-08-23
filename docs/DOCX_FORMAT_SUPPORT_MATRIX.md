# Ma Trận Hỗ Trợ Định Dạng DOCX (Format Support Matrix) - Docx-Agent V2.1

## 1. Bảng Tổng Hợp Khả Năng Hỗ Trợ (Coverage Matrix)

| Danh Mục OOXML | Phần Tử / Thuộc Tính | Mức Độ Hỗ Trợ | Cơ Chế Xử Lý Trong Pipeline |
| :--- | :--- | :---: | :--- |
| **Cấu trúc Tài liệu** | Sections (`w:sectPr`) | **Đầy đủ (100%)** | `SectionNode`, cách ly độc lập theo thứ tự tuần tự trong body, không gây nhân đôi khối |
| | Khổ giấy A4 (`w:pgSz`) | **Đầy đủ (100%)** | 210mm x 297mm (595.28pt x 841.89pt), chuyển đổi toạ độ chính xác sang CSS và OpenXML |
| | Căn lề (`w:pgMar`) | **Đầy đủ (100%)** | Lề chuẩn TCVN (Trái 3.0cm, Phải 2.0cm, Trên 2.0cm, Dưới 2.0cm) |
| | Định hướng trang (`w:orient`) | **Đầy đủ (100%)** | Portrait / Landscape độc lập trên từng Section |
| | Header / Footer | **Đầy đủ (100%)** | Hỗ trợ `differentFirstPage`, trang bìa ẩn Header/Footer, các trang sau đánh số tự động |
| **Đoạn văn (Paragraph)** | Căn lề (`w:jc`) | **Đầy đủ (100%)** | `left`, `center`, `right`, `both` (justify) |
| | Giãn dòng (`w:spacing/@w:line`) | **Đầy đủ (100%)** | Hỗ trợ `line_spacing` (1.15, 1.4, 1.5, 2.0, exact pt) |
| | Khoảng cách trước/sau (`w:before`, `w:after`) | **Đầy đủ (100%)** | Chuyển đổi chính xác dxa ↔ pt (1 pt = 20 dxa) |
| | Thụt lề đầu dòng (`w:ind/@w:firstLine`) | **Đầy đủ (100%)** | Thụt lề 1.27cm cho văn bản thường, 0.0cm cho tiêu đề, mã nguồn, bảng |
| | Thụt lề treo & lề trái (`w:hanging`, `w:left`) | **Đầy đủ (100%)** | Áp dụng cho danh sách phân cấp và trích dẫn |
| | Ngắt trang trước (`w:pageBreakBefore`) | **Đầy đủ (100%)** | Kích hoạt chuyển trang tự động trong `LayoutEngine` |
| | Giữ với đoạn tiếp (`w:keepNext`) | **Đầy đủ (100%)** | Chống mồ côi tiêu đề (Orphan Heading Prevention) |
| **Ký tự (Run & Font)** | Kiểu chữ (`w:rFonts`) | **Đầy đủ (100%)** | `ascii`, `hAnsi`, `cs` (Times New Roman, Consolas, JetBrains Mono, Arial) |
| | Cỡ chữ (`w:sz`, `w:szCs`) | **Đầy đủ (100%)** | Chuyển đổi chính xác half-points ↔ pt |
| | Đậm, Nghiêng, Gạch chân | **Đầy đủ (100%)** | `w:b`, `w:i`, `w:u` (single, double, dotted) |
| | Màu sắc (`w:color`) | **Đầy đủ (100%)** | Mã màu Hex `#RRGGBB` và Auto |
| | Màu nền / Highlight (`w:highlight`) | **Đầy đủ (100%)** | Ánh xạ bảng màu chuẩn Word (`yellow`, `green`, `cyan`, `magenta`, v.v.) |
| | Chỉ số trên/dưới (`w:vertAlign`) | **Đầy đủ (100%)** | `superscript`, `subscript` |
| | Ký tự ngắt trang (`w:br[@w:type='page']`) | **Đầy đủ (100%)** | Chuyển trang thực tế trong LayoutEngine |
| **Kế Thừa Kiểu (Styles)** | `docDefaults` | **Đầy đủ (100%)** | Kế thừa gốc từ `styles.xml` |
| | Kế thừa `w:basedOn` | **Đầy đủ (100%)** | Duyệt đệ quy chuỗi cha-con tránh lặp vô hạn |
| | Đè kiểu trực tiếp (Direct Formatting) | **Đầy đủ (100%)** | Ưu tiên cao nhất theo chuẩn OOXML ECMA-376 |
| **Danh Sách & Numbering** | `abstractNum` & `num` | **Đầy đủ (100%)** | Trích xuất từ `numbering.xml` |
| | Cấp độ phân cấp (`w:ilvl`) | **Đầy đủ (100%)** | 0 đến 8 cấp |
| | Định dạng số (`w:numFmt`) | **Đầy đủ (100%)** | `decimal`, `lowerLetter`, `upperLetter`, `lowerRoman`, `upperRoman`, `bullet` |
| | Mẫu văn bản (`w:lvlText`) | **Đầy đủ (100%)** | `%1.`, `%1.%2.`, `(%1)`, `•`, `-` |
| **Bảng Biểu (Tables)** | Lưới cột (`w:tblGrid/w:gridCol`) | **Đầy đủ (100%)** | Tính toán chiều rộng cột chính xác theo cm |
| | Nền ô (`w:shd`) | **Đầy đủ (100%)** | Giữ nguyên màu nền bảng theo bảng mã Hex |
| | Đường viền ô (`w:tcBorders`) | **Đầy đủ (100%)** | Đường viền đơn, viền kép, viền màu riêng cho hộp code & nhận xét |
| | Gộp cột (`w:gridSpan`) | **Đầy đủ (100%)** | Ánh xạ trực tiếp sang HTML `colSpan` |
| | Gộp hàng (`w:vMerge`) | **Đầy đủ (100%)** | Ánh xạ trực tiếp sang HTML `rowSpan` |
| | Lặp lại tiêu đề trang (`w:tblHeader`) | **Đầy đủ (100%)** | Lặp lại khi bảng vượt qua ranh giới trang |
| **Hình Ảnh & Sơ Đồ** | DrawingML (`w:drawing/wp:inline`) | **Đầy đủ (100%)** | Trích xuất kích thước `cx`, `cy` (EMUs → cm) và liên kết tệp |
| | Caption hình ảnh | **Đầy đủ (100%)** | Căn giữa, in nghiêng dưới ảnh |
| | Sơ đồ Mermaid / SVG | **Đầy đủ (100%)** | Render động trực tiếp sang SVG |
| **Phần tử mở rộng** | Thẻ OOXML chưa hỗ trợ | **Bảo toàn (Lossless)** | Lưu trong `UnsupportedBlock` (`raw_xml`) và chèn lại khi xuất file |
