@echo off
REM Office Zoom Reset - exe 빌드 스크립트
REM 필요한 패키지가 없으면 먼저: python -m pip install pywin32 pyinstaller pystray pillow

cd /d "%~dp0"

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --noconsole ^
  --name ConsultingReady ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import win32com.gen_py ^
  --hidden-import win32timezone ^
  --hidden-import pystray._win32 ^
  --hidden-import PIL.Image ^
  --hidden-import PIL.ImageDraw ^
  --hidden-import PIL.ImageFont ^
  --hidden-import tkinter ^
  consulting_ready.py

echo.
echo === 빌드 완료: dist\ConsultingReady.exe ===
pause
