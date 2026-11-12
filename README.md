# Tool đối soát hồ sơ - MVP

Ứng dụng desktop chạy offline, đọc Excel và thư mục PDF, đối soát họ tên, CCCD,
ngày sinh rồi xuất Excel có trạng thái, chi tiết sai lệch và hyperlink mở hồ sơ.

## Dùng trên Windows

1. Cài **Python 3.12 64-bit**, nhớ chọn **Add Python to PATH**.
2. Double-click `chay_source_windows.bat` để chạy ngay.
3. Đảm bảo có hai model DeepDoc tại:
   `deepdoc_vietocr\onnx\det.onnx` và
   `deepdoc_vietocr\vietocr\weight\vgg_seq2seq.pth`.
4. Double-click `build_windows.bat`; bản portable nằm tại `dist\DoiSoatHoSo.exe`.
5. Chỉ cần gửi file `.exe` này. Máy nhận không cần cài Python hay bộ OCR riêng.

Các file model bị loại khỏi Git theo `.gitignore`; cần đặt chúng ở đúng các đường dẫn trên
trước khi chạy OCR hoặc build bản portable.

Nếu trước đó đã chạy bằng phiên bản Python khác và thư mục `.venv` đã tồn tại,
hãy xóa riêng thư mục `.venv` trong thư mục app rồi chạy lại file `.bat`.

## OCR cho PDF scan

- App dùng DeepDoc để phát hiện dòng chữ và VietOCR để nhận dạng tiếng Việt trên CPU.
- PDF có lớp text vẫn được đọc trực tiếp; PDF scan được chuyển thành ảnh và xử lý bằng DeepDoc.
- Cột **Độ tin cậy OCR** hiển thị trung bình xác suất ký tự VietOCR theo thang 0–100.
  Đây là điểm model, không phải xác suất đã hiệu chuẩn rằng toàn bộ hồ sơ chính xác.
- Hồ sơ scan có điểm OCR dưới 50 sẽ được đánh dấu **Cần kiểm tra**.
- Nếu thiếu model hoặc thư viện OCR, lỗi được ghi vào kết quả đối soát để hồ sơ scan cần kiểm tra.

Bản `.exe` portable sẽ lớn hơn và lần mở đầu có thể chậm vài giây vì Windows cần
giải nén engine OCR và model vào thư mục tạm. Windows SmartScreen cũng có thể cảnh báo vì
file chưa được ký số; chỉ chọn **More info - Run anyway** khi nhận file từ nguồn tin cậy.

## Định dạng Excel đầu vào

Hàng đầu tiên là tiêu đề. App tự nhận các tên cột thường gặp, nhưng ổn định nhất là:

- `Họ và tên`
- `Số CCCD`
- `Ngày sinh`

Tiêu đề `Số chứng minh` trong file xuất từ hệ thống sát hạch cũng được app tự nhận.

CCCD nên để dạng Text để không mất số 0 ở đầu. App đọc cả `.xls` và `.xlsx`, nhưng
luôn xuất kết quả dạng `.xlsx`.

## Nguyên tắc ghép hồ sơ

1. Ưu tiên CCCD tìm thấy trong tên file hoặc nội dung PDF.
2. Nếu chưa tìm được, so khớp họ tên không dấu trong tên file.
3. Kiểm tra CCCD và ngày sinh nếu PDF đọc được các trường đó.
4. File PDF không ghép với dòng nào được liệt kê tại sheet `PDF chưa ghép`.

Nên đặt tên PDF theo mẫu `Ho Ten - CCCD - ngay.thang.nam.pdf`; cách này giúp app
ghép đúng ngay cả khi bản scan mờ.
