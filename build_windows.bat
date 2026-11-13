@echo off
setlocal
cd /d "%~dp0"
py -3.12 --version >nul 2>&1
if errorlevel 1 (
  echo LOI: Khong tim thay Python 3.12.
  echo Hay cai Python 3.12 64-bit tu python.org, sau do chay lai file nay.
  pause
  exit /b 1
)

if not exist "deepdoc_vietocr\onnx\det.onnx" (
  echo LOI: Thieu model deepdoc_vietocr\onnx\det.onnx.
  pause
  exit /b 1
)
if not exist "deepdoc_vietocr\vietocr\weight\vgg_seq2seq.pth" (
  echo LOI: Thieu model deepdoc_vietocr\vietocr\weight\vgg_seq2seq.pth.
  pause
  exit /b 1
)

py -3.12 -m venv .venv
if errorlevel 1 goto :error
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install --only-binary=:all: -r requirements.txt
if errorlevel 1 goto :error
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "DoiSoatHoSo" --paths "deepdoc_vietocr" --collect-all pymupdf --collect-all pdfplumber --collect-all pypdfium2 --collect-all torch --collect-all onnxruntime --collect-all cv2 --collect-all vietocr --add-data "deepdoc_vietocr\conf\service_conf.yaml;conf" --add-data "deepdoc_vietocr\onnx\det.onnx;onnx" --add-data "deepdoc_vietocr\vietocr\weight\vgg_seq2seq.pth;vietocr\weight" app.py
if errorlevel 1 goto :error
echo.
echo Da tao ban portable: dist\DoiSoatHoSo.exe
echo May nhan KHONG can cai Python hoac bo OCR rieng.
pause
exit /b 0

:error
echo.
echo KHONG THE BUILD. Xem thong bao loi ben tren.
pause
exit /b 1
