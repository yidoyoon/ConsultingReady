@echo off
REM 자동 시작 등록을 해제한다.
set "LNK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ConsultingReady.lnk"
if exist "%LNK%" (
  del "%LNK%"
  echo 자동 시작 등록을 해제했습니다.
) else (
  echo 등록된 자동 시작 바로가기가 없습니다.
)
echo (실행 중인 프로그램은 트레이 아이콘 - [종료] 또는 작업 관리자에서 종료하세요.)
pause
