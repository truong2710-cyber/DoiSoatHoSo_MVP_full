# Tool đối soát hồ sơ - MVP

Ứng dụng desktop chạy offline, đọc Excel và thư mục PDF, đối soát họ tên, CCCD,
ngày sinh rồi xuất Excel có trạng thái, chi tiết sai lệch và hyperlink mở hồ sơ.

## Dùng trên Windows

1. Cài **Python 3.12 64-bit**, nhớ chọn **Add Python to PATH**.
2. Double-click `chay_source_windows.bat` để chạy ngay.
3. Đảm bảo có hai model DeepDoc tại:
   `deepdoc_vietocr\onnx\det.onnx` và
   `deepdoc_vietocr\vietocr\weight\vgg_seq2seq.pth`.
4. Double-click `build_windows.bat`; bản portable nằm trong `dist\DoiSoatHoSo\`.
5. Gửi toàn bộ thư mục `DoiSoatHoSo` cho máy nhận, không gửi riêng file `.exe`.
   Chạy `DoiSoatHoSo.exe` bên trong thư mục đó; máy nhận không cần cài Python
   hay bộ OCR riêng.

### Build bằng GitHub Actions

Hai model OCR cần được commit tại các đường dẫn ở trên; workflow sẽ kiểm tra hash model,
build Windows app khi push lên nhánh `main`, rồi đăng toàn bộ thư mục portable thành artifact trong tab
**Actions**. Có thể chạy thủ công bằng cách mở **Actions**, chọn **Build Windows EXE**
và bấm **Run workflow**. Tải rồi giải nén artifact
`DoiSoatHoSo-windows-x64-onedir` của lần chạy thành công; chạy EXE bên trong
thư mục đã giải nén, giữ các file đi kèm cùng thư mục.

Để cài thư viện chạy app thủ công, dùng `python -m pip install -r requirements.txt`.
File này bao gồm dependency của app và DeepDoc; `requirements-build.txt` bổ sung
PyInstaller dành riêng cho bước đóng gói.
Các file model bị loại khỏi Git theo `.gitignore`; cần đặt chúng ở đúng các đường dẫn trên
trước khi chạy OCR hoặc build bản portable.

Nếu trước đó đã chạy bằng phiên bản Python khác và thư mục `.venv` đã tồn tại,
hãy xóa riêng thư mục `.venv` trong thư mục app rồi chạy lại file `.bat`.

Trên cửa sổ chính, bật/tắt **Hiển thị kết quả trực tiếp trong khi xử lý** bất cứ lúc nào
để ẩn hoặc mở lại bảng tiến trình và kết quả. Khi mở lại, các dòng đã quét/đối soát vẫn
được giữ lại. Tùy chọn này mặc định được bật.

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

## Bản Tesseract thử nghiệm

`app_tesseract.py` giữ nguyên giao diện, nhận dạng thông tin và quy tắc đối soát của
`app.py`, nhưng thay OCR DeepDoc + VietOCR bằng Tesseract. Cài Tesseract OCR có dữ liệu
tiếng Việt (`vie.traineddata`), sau đó chạy `python app_tesseract.py`. Nếu lệnh
`tesseract` không có trong `PATH`, đặt biến môi trường `TESSERACT_CMD` thành đường dẫn
đến chương trình Tesseract. Bản này tự dùng thêm tiếng Anh nếu có `eng.traineddata`.
Các điểm OCR của Tesseract không được hiệu chuẩn như điểm VietOCR; cần đánh giá kết quả
thực tế trước khi dùng thay cho bản hiện tại.

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
dùng là trung bình hai dòng. Nếu OCR dính thêm chữ số làm thành chuỗi 13–15 số, app thử
các cửa sổ 12 số hợp lệ và ưu tiên ID trùng với tên file. ID 9 chữ số không bị loại khi
được dùng làm dự phòng từ tên file. Bảng quét trực tiếp hiển thị điểm của ID đã chọn.
PDF có lớp text và tên file không có điểm OCR.

Ngày sinh được ghép với nhãn gần nhất theo vị trí trên trang; khoảng cách chuẩn hóa là
giá trị nhỏ hơn giữa khoảng hở ngang và khoảng hở dọc, tối đa `0.04`. Vì vậy ngày cùng
hàng hoặc cùng cột với nhãn vẫn có thể ghép dù cách xa theo trục còn lại. Mỗi cặp
nhãn/ngày là một lượt bầu độc lập, kể cả khi có nhiều lượt trên cùng một trang. Mỗi lượt
có trọng số `0.5 + 0.5 × (điểm OCR / 100)`, vì vậy ngày được
nhận diện lặp lại có thể thắng một lần đọc có điểm cao hơn. Nếu tổng trọng số của hai
ngày đứng đầu chênh nhau không quá `0.004`, kết quả được để trống và yêu cầu kiểm tra thủ
công. Khi OCR năm sinh bị lỗi định dạng hoặc thiếu chữ số, app có thể dùng năm suy ra từ
CCCD 12 số đọc trực tiếp từ PDF với điểm OCR từ 80 trở lên; tên file không được dùng để
suy ra năm. Cách đọc ưu tiên ngày có dấu `/` trước dấu `.`. Khi OCR không đọc được ngày
sinh đáng tin cậy, app thử OCR lại trang khi xoay `+90°` và `-90°` nếu điểm OCR dưới
65%, nhãn ngày sinh chưa có ngày đọc được, hoặc ID yếu. Khi tài liệu đã có nhãn ngày sinh
nhưng chưa tìm được ngày, các trang OCR tiếp theo cũng được thử xoay. Parser chấp nhận
các dấu phân cách ngày hỗn hợp như `11.07/2003` và một số lỗi OCR thường gặp trong nhãn
`Ngày sinh`, rồi chọn kết quả có bằng chứng ngày sinh/CCCD tốt hơn. Khi thiếu tọa độ OCR, app dùng cách tìm theo dòng: ưu
tiên ngày cách nhãn không quá hai dòng, sau đó mới mở rộng tối đa 12 dòng và đánh dấu
kết quả cần kiểm tra thủ công. Ngày phải trong khoảng 18–100 tuổi tính đến ngày chạy app.
Khoảng tuổi này là bộ lọc hợp lý do người dùng chọn cho hồ sơ bằng lái, không phải kết
luận về điều kiện pháp lý cho mọi hạng giấy phép. Điểm ngày sinh đã chọn cũng hiển thị
trong bảng quét; PDF có lớp text không có điểm OCR.

Nên đặt tên PDF theo mẫu `Ho Ten - CCCD - ngay.thang.nam.pdf`; cách này giúp app
ghép đúng ngay cả khi bản scan mờ.
