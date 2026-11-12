import argparse
from pathlib import Path

from app import inspect_pdf


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Đọc và in các trường nhận diện được từ một file PDF."
    )
    parser.add_argument("pdf", type=Path, help="Đường dẫn file PDF cần kiểm tra")
    args = parser.parse_args()

    if not args.pdf.is_file():
        parser.error(f"Không tìm thấy file: {args.pdf}")
    if args.pdf.suffix.lower() != ".pdf":
        parser.error("File đầu vào phải có đuôi .pdf")

    info = inspect_pdf(args.pdf)
    print(f"File: {args.pdf}")
    print(f"Đã dùng OCR: {'Có' if info.used_ocr else 'Không'}")
    confidence = f"{info.confidence:.1f}" if info.confidence is not None else "Không có"
    print(f"Độ tin cậy OCR: {confidence}")
    print(f"Họ tên gợi ý từ tên file: {info.name_hint or '(không có)'}")
    print(f"CCCD/CMND: {', '.join(sorted(info.ids)) or '(không tìm thấy)'}")
    print(f"Các ngày nhận diện: {', '.join(sorted(info.dates)) or '(không tìm thấy)'}")
    print(f"Ngày sinh: {', '.join(sorted(info.birth_dates)) or '(không tìm thấy)'}")
    if info.error:
        print(f"Lỗi/cảnh báo: {info.error}")
    print("\n--- Toàn bộ nội dung đọc được ---")
    print(info.text or "(không đọc được nội dung)")


if __name__ == "__main__":
    main()
