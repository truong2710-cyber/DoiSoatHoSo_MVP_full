# Đối soát hồ sơ

Ứng dụng desktop Windows giúp đối chiếu danh sách Excel với hồ sơ PDF theo họ tên,
CCCD và ngày sinh. Có thể chọn một sheet hoặc đối chiếu tất cả sheet trong workbook.
PDF có lớp văn bản được đọc trực tiếp; PDF scan được xử lý bằng DeepDoc để phát hiện
văn bản và VietOCR để nhận dạng tiếng Việt.

## Tính năng

- Đọc file `.xlsx` và `.xls`; tìm PDF trong thư mục đã chọn, bao gồm cả thư mục con.
- Chọn riêng một sheet Excel hoặc chọn **Tất cả các sheet**.
- Theo dõi tiến độ, kết quả đối chiếu và nội dung đọc từ PDF ngay trong giao diện; hỗ trợ
  tìm kiếm, lọc trạng thái và sắp xếp theo cột.
- Xuất workbook kết quả có trạng thái, giải thích chênh lệch, điểm OCR, đường dẫn hồ sơ
  và danh sách PDF chưa ghép.
- Lưu cache quét PDF trên máy để có thể dùng lại kết quả nếu file PDF không thay đổi.

Ứng dụng xử lý dữ liệu trên máy đang chạy. Bản portable chứa runtime và thư viện cần
thiết; máy nhận không cần cài Python hoặc DeepDoc riêng.

## Chạy trên Windows từ mã nguồn

1. Cài Python **3.12 64-bit**.
2. Đảm bảo hai model OCR có trong repository tại:
   - `deepdoc_vietocr/onnx/det.onnx`
   - `deepdoc_vietocr/vietocr/weight/vgg_seq2seq.pth`
3. Chạy `chay_source_windows.bat`. Nếu chưa có môi trường `.venv`, script tạo môi trường,
   cài các gói từ `requirements.txt`, rồi chạy `app.py`.

Nếu dependency đã thay đổi hoặc `.venv` được tạo bằng Python khác, hãy xóa thư mục
`.venv` trong dự án rồi chạy lại script.

## Tạo bản portable

### GitHub Actions

Workflow `.github/workflows/build-windows.yml` chạy khi push lên nhánh `main`; có thể
chạy thủ công từ trang **Actions**:

1. Mở repository trên GitHub và chọn **Actions**.
2. Chọn workflow **Build Windows EXE**, bấm **Run workflow** và chọn nhánh cần build.
3. Khi workflow thành công, tải artifact `DoiSoatHoSo-windows-x64-onedir`.
4. Giải nén artifact và chạy `DoiSoatHoSo.exe` trong thư mục đã giải nén.

Giữ nguyên toàn bộ thư mục ứng dụng; không tách riêng file `.exe` khỏi các file bên cạnh.
Artifact được giữ trong 14 ngày. Workflow dùng Windows x64 và Python 3.12, xác minh hash
model, kiểm tra Torch/torchvision và khởi tạo DeepDoc trước khi đóng gói.

Hai model cần được commit tại các đường dẫn ở trên. Workflow dừng nếu thiếu model hoặc
nội dung file không khớp hash đã cấu hình; thay model cũng cần cập nhật bước xác minh
trong workflow.

### Build thủ công trên Windows

Trên Windows có Python 3.12 64-bit và đủ model, chạy `build_windows.bat`. Script cài các
gói từ `requirements-build.txt`, kiểm tra import và khởi tạo OCR, rồi tạo bản portable ở
`dist\DoiSoatHoSo\`. Gửi toàn bộ thư mục này cho máy đích.

`requirements.txt` chứa dependency chạy ứng dụng và tham chiếu
`deepdoc_vietocr/requirements.txt`; `requirements-build.txt` bổ sung PyInstaller. Script
build là dành cho Windows, không dùng để tạo `.exe` trên macOS.

## Chuẩn bị file Excel

Hàng đầu tiên của mỗi sheet cần đối chiếu phải là tiêu đề. Tên cột khuyến nghị:

| Trường | Tiêu đề khuyến nghị |
|---|---|
| Họ tên | `Họ và tên` |
| CCCD | `Số CCCD` |
| Ngày sinh | `Ngày sinh` |

Ứng dụng nhận diện một số biến thể tiêu đề phổ biến. Cột họ tên và CCCD là bắt buộc;
cột ngày sinh có thể để trống hoặc không có. Nên lưu CCCD ở dạng văn bản để giữ số `0`
ở đầu. File `.xls` được đọc bằng `xlrd`, file `.xlsx` bằng `openpyxl`; `.xlsb` và các
định dạng khác không được hỗ trợ. Kết quả luôn xuất thành `.xlsx`.

## Chọn sheet và cách đối chiếu

Sau khi chọn workbook, danh sách **Sheet cần đối soát** hiển thị các sheet có trong file.
Nếu có nhiều sheet, có thể:

- Chọn một sheet để đối chiếu riêng sheet đó.
- Chọn **Tất cả các sheet** để xử lý mọi sheet dùng được. Đây là lựa chọn tường minh;
  ứng dụng không tự động gộp tất cả sheet khi mở file.

Khi chọn tất cả, sheet thiếu cột họ tên hoặc CCCD, hoặc không có dòng dữ liệu, sẽ bị bỏ
qua và được ghi lý do vào kết quả. Nếu không còn sheet nào dùng được, ứng dụng báo lỗi
thay vì xuất kết quả rỗng.

PDF được quét một lần, sau đó đối chiếu độc lập với từng sheet được chọn. Vì mỗi sheet
có lượt ghép riêng, cùng một PDF có thể được ghép với một dòng trong mỗi sheet khác
nhau. Nếu workbook chứa các danh sách chồng lặp, cần lưu ý điều này khi chọn **Tất cả
các sheet**.

## Kết quả xuất ra

Khi chọn một sheet, workbook kết quả gồm:

- **Kết quả đối soát** — dữ liệu sheet nguồn cùng trạng thái, chi tiết, điểm OCR, đường
  dẫn hồ sơ và ghi chú cần kiểm tra.
- **PDF chưa ghép** — PDF không được ghép với dòng nào trong sheet đang xử lý.

Khi chọn **Tất cả các sheet**, workbook kết quả gồm:

- **Tổng hợp sheet** — trạng thái đối soát, số dòng, số lượng theo từng kết quả và ghi
  chú cho mỗi sheet được xử lý hoặc bị bỏ qua.
- Một sheet kết quả riêng cho mỗi sheet Excel được xử lý; mỗi sheet giữ các cột dữ liệu
  ban đầu và được bổ sung thông tin đối soát.
- **PDF chưa ghép** — PDF không được sử dụng bởi các sheet đã xử lý. Nếu có sheet bị bỏ
  qua, các PDF chỉ thuộc những sheet đó cũng có thể xuất hiện ở đây.

Tên sheet kết quả được tạo từ tên sheet nguồn; Excel giới hạn độ dài tên sheet nên tên
dài hoặc trùng lặp có thể được rút gọn/điều chỉnh.

Các trạng thái gồm **OK**, **Cần kiểm tra**, **Sai thông tin** và **Không tìm thấy hồ sơ**.
Ứng dụng ưu tiên CCCD đọc từ nội dung PDF; trong một số trường hợp, CCCD trong tên file
có thể hỗ trợ ghép và kết quả sẽ được đánh dấu để kiểm tra. Có thể đặt tên PDF theo dạng
`Ho Ten - CCCD.pdf` để hỗ trợ nhận diện khi bản scan mờ.

Điểm OCR là điểm của mô hình, không phải xác suất đã hiệu chuẩn rằng toàn bộ thông tin
chính xác. OCR có thể đọc sai; hãy kiểm tra PDF gốc, nhất là khi ứng dụng đánh dấu
**Cần kiểm tra** hoặc **Sai thông tin**. Mỗi PDF được OCR tối đa 8 trang; thông tin nằm
ở trang sau có thể không được phát hiện.

## Dữ liệu cục bộ và xử lý sự cố

- Cài đặt, cache quét và log lỗi nằm trong `%USERPROFILE%\.doi_soat_cache\`.
- Nếu lỗi xảy ra trong khi đối chiếu, xem `errors.log` trong thư mục trên để lấy traceback.
  Hộp thoại lỗi cũng hiển thị đường dẫn log nếu ghi log thành công.
- Nếu ứng dụng báo thiếu model, xác nhận cả hai model ở đúng đường dẫn rồi tạo lại bản
  build.
- Nếu ứng dụng portable không khởi động hoặc báo thiếu thành phần, tải artifact mới nhất
  và giải nén lại toàn bộ thư mục ứng dụng.
- Bản build chưa được ký số. Windows SmartScreen hoặc phần mềm bảo mật có thể cảnh báo
  với ứng dụng mới/chưa có danh tiếng phát hành. Chỉ chạy bản build từ repository và
  workflow mà bạn tin cậy; hãy xác minh nguồn file trước khi cho phép chạy.
