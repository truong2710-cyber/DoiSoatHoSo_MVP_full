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
if not exist .venv\Scripts\python.exe (
  py -3.12 -m venv .venv
  if errorlevel 1 goto :error
  call .venv\Scripts\activate.bat
  python -m pip install --upgrade pip
  if errorlevel 1 goto :error
  python -m pip install -r requirements.txt
  if errorlevel 1 goto :error
) else (
  call .venv\Scripts\activate.bat
)
python app.py
if errorlevel 1 goto :error
exit /b 0

:error
echo.
echo KHONG THE CHAY APP. Xem thong bao loi ben tren.
pause
exit /b 1
