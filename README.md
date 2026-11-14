# Đối soát hồ sơ

Ứng dụng desktop Windows giúp đối chiếu danh sách Excel với hồ sơ PDF theo họ tên,
CCCD và ngày sinh. Ứng dụng đọc trực tiếp PDF có lớp văn bản; với PDF scan, ứng dụng
dùng DeepDoc để phát hiện văn bản và VietOCR để nhận dạng tiếng Việt.

## Tính năng

- Đọc danh sách `.xlsx` và `.xls`; quét PDF trong thư mục đã chọn, bao gồm thư mục con.
- Trích xuất thông tin từ lớp văn bản PDF hoặc dùng OCR cho trang scan.
- Hiển thị tiến độ, kết quả đối chiếu và kết quả quét PDF; hỗ trợ tìm kiếm, lọc và sắp
  xếp trong giao diện.
- Xuất workbook Excel có trạng thái, giải thích chênh lệch, điểm OCR, đường dẫn hồ sơ
  và sheet riêng cho PDF chưa ghép.
- Lưu cache kết quả quét trên máy để lần chạy sau có thể bỏ qua các PDF không thay đổi.

Ứng dụng xử lý dữ liệu trên máy đang chạy. Bộ cài portable bao gồm Python và các thư viện
cần thiết; máy nhận không cần cài Python hoặc DeepDoc riêng.

## Chạy trên Windows từ mã nguồn

1. Cài Python **3.12 64-bit**.
2. Đảm bảo hai model OCR có trong repository tại đúng các đường dẫn:
   - `deepdoc_vietocr/onnx/det.onnx`
   - `deepdoc_vietocr/vietocr/weight/vgg_seq2seq.pth`
3. Chạy `chay_source_windows.bat`. Script tạo `.venv` và cài các gói từ
   `requirements.txt` nếu môi trường `.venv` chưa tồn tại, sau đó chạy `app.py`.

Nếu đã có `.venv` từ môi trường cũ hoặc dependency đã thay đổi, hãy xóa thư mục `.venv`
trong thư mục dự án rồi chạy lại script để tạo môi trường mới.

## Tạo bản portable

### GitHub Actions (Windows)

Workflow `.github/workflows/build-windows.yml` chạy tự động khi có push lên nhánh `main`;
nó cũng có thể chạy thủ công:

1. Mở repository trên GitHub và vào **Actions**.
2. Chọn **Build Windows EXE**, bấm **Run workflow**, rồi xác nhận nhánh cần build.
3. Chờ workflow hoàn tất thành công và tải artifact
   `DoiSoatHoSo-windows-x64-onedir`.
4. Giải nén artifact và chạy `DoiSoatHoSo.exe` ngay trong thư mục vừa giải nén;
   giữ tất cả các file đi kèm trong cùng thư mục.

Artifact được giữ trong 14 ngày. Cần giải nén và giữ nguyên **toàn bộ thư mục ứng dụng**;
không di chuyển riêng file `.exe` ra khỏi các file đi kèm. Workflow dùng Windows x64 và
Python 3.12, xác minh hash của hai model, kiểm tra import và khởi tạo OCR, rồi đóng gói
ứng dụng.

Hai model nêu trên được theo dõi trong Git tại các đường dẫn cố định. Workflow sẽ dừng
nếu thiếu model hoặc nội dung không khớp hash mong đợi; không thay model riêng lẻ mà
không cập nhật bước xác minh trong workflow.

### Build thủ công trên Windows

Trên Windows có Python 3.12 64-bit và đủ model tại chỗ, chạy `build_windows.bat`.
Script cài dependency từ `requirements-build.txt`, kiểm tra Torch/torchvision và DeepDoc,
rồi tạo thư mục portable tại `dist\DoiSoatHoSo\`. Gửi toàn bộ thư mục này sang máy đích.

`requirements.txt` chứa dependency chạy ứng dụng và tham chiếu
`deepdoc_vietocr/requirements.txt`; `requirements-build.txt` bổ sung PyInstaller. Không
dùng script Windows để tạo `.exe` trên macOS.

## Sử dụng ứng dụng

Trong cửa sổ chính:

1. Chọn file Excel danh sách.
2. Chọn thư mục chứa PDF.
3. Chọn đường dẫn file kết quả `.xlsx` (mặc định nằm cạnh file Excel nguồn).
4. Chọn **Bắt đầu kiểm tra**. Có thể hủy tác vụ đang chạy; cache quét PDF có thể được
   dùng lại ở lần chạy sau.

Nên giữ CCCD ở dạng văn bản trong Excel để bảo toàn số `0` ở đầu. Hàng đầu tiên phải là
hàng tiêu đề. Tên cột được nhận diện linh hoạt; nên dùng:

| Trường | Tiêu đề khuyến nghị |
|---|---|
| Họ tên | `Họ và tên` |
| CCCD | `Số CCCD` |
| Ngày sinh | `Ngày sinh` |

Cột ngày sinh có thể để trống hoặc không có. Cột họ tên và CCCD là bắt buộc. File `.xls`
được hỗ trợ làm đầu vào; kết quả luôn xuất thành `.xlsx`.

## Kết quả đối chiếu

Workbook đầu ra có:

- **Kết quả đối soát** — dữ liệu Excel ban đầu cùng trạng thái, chi tiết chênh lệch, điểm
  OCR, đường dẫn hồ sơ và ghi chú cần kiểm tra.
- **PDF chưa ghép** — các PDF không được ghép với dòng Excel nào.

Các trạng thái thường gặp là **OK**, **Cần kiểm tra**, **Sai thông tin** và
**Không tìm thấy hồ sơ**. Kết quả không thay thế việc kiểm tra hồ sơ gốc: OCR có thể đọc
sai; điểm OCR là điểm của mô hình, không phải xác suất đã hiệu chuẩn rằng toàn bộ thông
tin chính xác. Chênh lệch ở điểm OCR thấp được đưa vào diện kiểm tra thay vì mặc định kết
luận là sai.

Ứng dụng ưu tiên CCCD đọc từ nội dung PDF. Trong trường hợp cần thiết, thông tin trong
tên file có thể hỗ trợ ghép hồ sơ; những trường hợp ghép dự phòng hoặc gần đúng được
đánh dấu để người dùng xem lại. Có thể đặt tên PDF theo dạng
`Ho Ten - CCCD.pdf` để hỗ trợ nhận diện hồ sơ khi bản scan mờ.

OCR xử lý tối đa 8 trang mỗi PDF. PDF nhiều trang hơn có thể còn nội dung chưa được quét;
hãy kiểm tra thủ công nếu thông tin cần đối chiếu nằm ngoài các trang được xử lý.

## Dữ liệu cục bộ và xử lý sự cố

- Cài đặt, cache quét và log lỗi nằm trong `%USERPROFILE%\.doi_soat_cache\`.
- Nếu lỗi xảy ra trong lúc xử lý, xem `errors.log` trong thư mục trên để lấy traceback.
  Thông báo lỗi trong ứng dụng cũng hiển thị đường dẫn log nếu ghi log thành công.
- Nếu ứng dụng báo thiếu model, xác nhận cả hai file model còn ở đúng các đường dẫn nêu
  trên và tạo lại bản build.
- Nếu file `.exe` không khởi động hoặc thiếu thành phần, tải artifact mới nhất và giải nén
  lại nguyên thư mục `DoiSoatHoSo`.
- Bản build chưa được ký số. Windows SmartScreen hoặc phần mềm bảo mật có thể cảnh báo
  với ứng dụng mới hoặc chưa có danh tiếng phát hành. Chỉ chạy bản build từ repository
  và workflow mà bạn tin cậy; không bỏ qua cảnh báo khi chưa xác minh nguồn file.
