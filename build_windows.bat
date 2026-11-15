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
  if not defined CI pause
  exit /b 1
)
if not exist "deepdoc_vietocr\vietocr\weight\vgg_seq2seq.pth" (
  echo LOI: Thieu model deepdoc_vietocr\vietocr\weight\vgg_seq2seq.pth.
  if not defined CI pause
  exit /b 1
)

py -3.12 -m venv .venv
if errorlevel 1 goto :error
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install -r requirements-build.txt
if errorlevel 1 goto :error
python -c "import xlrd; print('xlrd', xlrd.__version__)"
if errorlevel 1 goto :error
python -c "import torch, torchvision; print('Torch', torch.__version__, '- torchvision', torchvision.__version__)"
if errorlevel 1 goto :error
python -c "import sys; sys.path.insert(0, 'deepdoc_vietocr'); from module.ocr import OCR; OCR(); print('DeepDoc OCR import and model initialization: OK')"
if errorlevel 1 goto :error
python -m PyInstaller --noconfirm --clean --onedir --windowed --name "DoiSoatHoSo" --paths "deepdoc_vietocr" --hidden-import xlrd --hidden-import module.ocr --hidden-import module.recognizer --hidden-import module.layout_recognizer --hidden-import module.table_structure_recognizer --collect-all pymupdf --collect-all pdfplumber --collect-all pypdfium2 --collect-all torch --collect-all torchvision --collect-all onnxruntime --collect-all cv2 --collect-all vietocr --add-data "deepdoc_vietocr\conf\service_conf.yaml;conf" --add-data "deepdoc_vietocr\onnx\det.onnx;onnx" --add-data "deepdoc_vietocr\vietocr\weight\vgg_seq2seq.pth;vietocr\weight" app.py
if errorlevel 1 goto :error
echo.
echo Da tao ban portable: dist\DoiSoatHoSo\DoiSoatHoSo.exe
echo Gui toan bo thu muc DoiSoatHoSo cho may nhan.
echo May nhan KHONG can cai Python hoac bo OCR rieng.
if not defined CI pause
exit /b 0

:error
echo.
echo KHONG THE BUILD. Xem thong bao loi ben tren.
if not defined CI pause
exit /b 1
