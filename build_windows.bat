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

set "TESS_SOURCE=C:\Program Files\Tesseract-OCR"
if not exist "%TESS_SOURCE%\tesseract.exe" (
  echo LOI: Chua tim thay Tesseract tai %TESS_SOURCE%.
  echo Hay cai Tesseract OCR 5 kem ngon ngu Vietnamese truoc khi build.
  pause
  exit /b 1
)
if not exist "%TESS_SOURCE%\tessdata\vie.traineddata" (
  echo LOI: Tesseract chua co ngon ngu Vietnamese - vie.traineddata.
  pause
  exit /b 1
)

if exist bundled_tesseract rmdir /s /q bundled_tesseract
mkdir bundled_tesseract\tessdata
copy /y "%TESS_SOURCE%\tesseract.exe" bundled_tesseract\ >nul
copy /y "%TESS_SOURCE%\*.dll" bundled_tesseract\ >nul
copy /y "%TESS_SOURCE%\tessdata\eng.traineddata" bundled_tesseract\tessdata\ >nul
copy /y "%TESS_SOURCE%\tessdata\vie.traineddata" bundled_tesseract\tessdata\ >nul
copy /y "%TESS_SOURCE%\tessdata\osd.traineddata" bundled_tesseract\tessdata\ >nul
if exist "%TESS_SOURCE%\tessdata\pdf.ttf" copy /y "%TESS_SOURCE%\tessdata\pdf.ttf" bundled_tesseract\tessdata\ >nul
if exist "%TESS_SOURCE%\tessdata\configs" xcopy /e /i /y "%TESS_SOURCE%\tessdata\configs" bundled_tesseract\tessdata\configs >nul
if exist "%TESS_SOURCE%\tessdata\tessconfigs" xcopy /e /i /y "%TESS_SOURCE%\tessdata\tessconfigs" bundled_tesseract\tessdata\tessconfigs >nul

py -3.12 -m venv .venv
if errorlevel 1 goto :error
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install --only-binary=:all: -r requirements.txt
if errorlevel 1 goto :error
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "DoiSoatHoSo" --collect-all pymupdf --add-data "bundled_tesseract;tesseract" app.py
if errorlevel 1 goto :error
echo.
echo Da tao ban portable: dist\DoiSoatHoSo.exe
echo May nhan KHONG can cai Python hoac Tesseract.
pause
exit /b 0

:error
echo.
echo KHONG THE BUILD. Xem thong bao loi ben tren.
pause
exit /b 1
