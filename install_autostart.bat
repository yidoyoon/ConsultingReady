@echo off
REM Windows 로그인 시 ConsultingReady.exe 를 자동 실행하도록
REM 시작프로그램(Startup) 폴더에 바로가기를 만든다.

set "EXE=%~dp0dist\ConsultingReady.exe"
if not exist "%EXE%" (
  echo [오류] dist\ConsultingReady.exe 가 없습니다. 먼저 build.bat 로 빌드하세요.
  pause
  exit /b 1
)

set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
powershell -NoProfile -Command ^
  "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%STARTUP%\ConsultingReady.lnk');" ^
  "$s.TargetPath='%EXE%';" ^
  "$s.WorkingDirectory='%~dp0dist';" ^
  "$s.Description='Office 문서 저장 시 자동 정리';" ^
  "$s.Save()"

echo.
echo 자동 시작 등록 완료: %STARTUP%\ConsultingReady.lnk
echo (지금 바로 실행하려면 dist\ConsultingReady.exe 를 더블클릭하세요.)
pause
