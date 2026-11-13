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
- PDF có lớp text vẫn được đọc trực tiếp; trang scan được render bằng PDFium ở 216 DPI,
  cùng cách với `deepdoc_vietocr/t_ocr.py`, rồi xử lý bằng DeepDoc.
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

1. Ưu tiên ID 12 chữ số đọc được từ nội dung PDF.
2. Chỉ khi nội dung PDF không có ID 12 chữ số hợp lệ, app mới dùng ID trong tên file
   làm phương án dự phòng; ID dự phòng có thể dài 9 hoặc 12 chữ số.
3. Nếu ID đọc từ PDF không khớp dòng Excel, app có thể thử ID từ tên file; kết quả ghép
   bằng ID tên file luôn là **Cần kiểm tra**. Không ghép theo tên file nếu không có ID
   khớp.
4. Một PDF đã được ghép với độ tin cậy cao (điểm từ 0.8 trở lên, bao gồm ghép ID chính
   xác) không được ghép với dòng Excel khác.
5. File PDF không ghép với dòng nào được liệt kê tại sheet `PDF chưa ghép`.

Nếu không có dòng Excel khớp chính xác ID trong tên file, app có thể ghép gần đúng khi
cả độ tương đồng tên và ID trong tên file đều đạt ít nhất 85%. App chọn cặp có điểm trung
bình cao nhất, chỉ ghép mỗi PDF với một dòng, và đánh dấu kết quả **Sai thông tin**;
quy tắc này dựa trên tên file, không phụ thuộc điểm OCR của PDF.

Ngày sinh khớp sau khi đảo ngày/tháng sẽ được đánh dấu **Cần kiểm tra**, không phải
**Sai thông tin**. Sai lệch chỉ được đánh dấu **Sai thông tin** khi độ tin cậy PDF từ 80
trở lên; nếu điểm thấp hơn 80 hoặc không có điểm, kết quả là **Cần kiểm tra**.

Khi có nhiều số ID 12 chữ số trong PDF, app chọn số có điểm OCR dòng cao nhất; ba chữ
số đầu phải từ `001` đến `096` (bao gồm cả hai đầu). Chỉ các ID OCR có điểm từ 60 trở
lên mới được xét để tránh nhận nhầm chuỗi số nhiễu. Nếu số bị tách qua hai dòng, điểm
dùng là trung bình hai dòng. ID 9 chữ số không bị loại khi được dùng làm dự phòng từ tên
file. Bảng quét trực tiếp hiển thị điểm của ID đã chọn. PDF có lớp text và tên file không
có điểm OCR.

Ngày sinh được tìm trong cùng trang, ưu tiên ngày cách nhãn không quá hai dòng. Chỉ khi
không có ngày hợp lệ gần nhãn, app mở rộng tìm kiếm tối đa 12 dòng và đánh dấu kết quả
đó cần kiểm tra thủ công. Trong từng nhóm, chọn khoảng cách gần nhất trước, rồi mới dùng
điểm OCR để phân định. Ngày phải trong khoảng 18–100 tuổi tính đến ngày chạy app. Ngày
bị tách qua hai dòng dùng điểm trung bình. Khoảng tuổi này là bộ lọc hợp lý do người dùng
chọn cho hồ sơ bằng lái, không phải kết luận về điều kiện pháp lý cho mọi hạng giấy phép.
Điểm ngày sinh đã chọn cũng hiển thị trong bảng quét. PDF có lớp text không có điểm OCR;
nếu nhiều ngày cùng khoảng cách và điểm, app không tự chọn và yêu cầu kiểm tra thủ công.

Nên đặt tên PDF theo mẫu `Ho Ten - CCCD - ngay.thang.nam.pdf`; cách này giúp app
ghép đúng ngay cả khi bản scan mờ.
