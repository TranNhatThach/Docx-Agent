"""
Example script demonstrating complete DocumentAgent lifecycle:
Creation, formatting presets, text editing, table styling, images, TOC, verification, and diff.
"""

from pathlib import Path
from PIL import Image
from docx_agent import DocumentAgent


def run_demo():
    examples_dir = Path(__file__).parent
    output_docx = examples_dir / "demo_report.docx"
    chart_img = examples_dir / "demo_chart.png"

    print("1. Creating fresh document...")
    agent = DocumentAgent()

    print("2. Applying Vietnamese Academic Preset (academic-vn)...")
    agent.apply_preset("academic-vn")

    print("3. Adding title and headings...")
    t_id = agent.append("BÁO CÁO NGHIÊN CỨU HỆ THỐNG AGENT THÔNG MINH", style="Title")
    agent.append("Đề tài: Tối ưu hóa xử lý văn bản tự động với AI", style="Subtitle")

    agent.append("1. Giới thiệu tổng quan", style="Heading 1")
    agent.append(
        "Nghiên cứu này trình bày kiến trúc và phương pháp tiếp cận mới trong việc xây dựng "
        "các hệ thống agent tự động thao tác tài liệu Microsoft Word một cách an toàn và chính xác."
    )

    agent.append("2. Phương pháp thực hiện", style="Heading 1")
    p_id = agent.append(
        "Chúng tôi sử dụng thuật toán phẫu thuật Run (Run Surgery Engine) "
        "để thay thế văn bản mà không làm mất định dạng in đậm, in nghiêng hoặc màu sắc."
    )
    agent.format_text(p_id, font_name="Times New Roman", font_size_pt=13.0)

    agent.append("2.1 Phân tích dữ liệu thực nghiệm", style="Heading 2")
    
    print("4. Creating data table with repeating header...")
    tid = agent.create_table(
        rows=4,
        cols=3,
        data=[
            ["Phương Pháp", "Độ Chính Xác", "Thời Gian (ms)"],
            ["Naive Text Replacement", "62.4%", "120"],
            ["DOM Overwrite", "78.1%", "95"],
            ["Docx-Agent Run Surgery", "99.8%", "45"],
        ],
        alignment="center",
    )
    agent.edit_cell(f"{tid}_r04_c02", text="99.9%", bold=True, bg_color_hex="D4EDDA")

    print("5. Inserting chart image with caption...")
    if chart_img.exists():
        agent.insert_image(
            image_path=str(chart_img),
            width_cm=13.0,
            alignment="center",
            caption="Hình 1: So sánh hiệu năng giữa các phương pháp thao tác DOCX.",
        )

    print("6. Setting page numbering...")
    agent.set_page_numbers(format_str="Trang {PAGE} / {NUMPAGES}", alignment="center")

    print("7. Saving document transactionally with independent verification...")
    saved_path = agent.save(output_docx, verify=True)
    print(f"Document successfully generated and verified at: {saved_path}")


if __name__ == "__main__":
    run_demo()
