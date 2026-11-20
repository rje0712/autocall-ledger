@echo off
chcp 65001 > nul
title 오토콜 회비 장부 웹사이트 업데이트
cls

echo ========================================================
echo   [오토콜 회비 장부] GitHub Pages 웹사이트 동기화
echo ========================================================
echo.
echo [1/3] 최신 데이터 동기화 및 웹페이지 빌드 중...
python build_ledger.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo ❌ 데이터 빌드 중 오류가 발생했습니다.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [2/3] GitHub로 최신 장부 전송 중 (Git Push)...
git add .
git commit -m "update: 최신 회비 장부 업데이트"
git push origin main
if %ERRORLEVEL% neq 0 (
    echo.
    echo ⚠️ Git 푸시 중 오류가 발생했습니다. 네트워크 또는 깃 설정을 확인해주세요.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [3/3] 공개 웹사이트 주소를 클립보드에 복사 중...
set "PAGE_URL=https://rje0712.github.io/autocall-ledger"
echo|set /p="%PAGE_URL%"| clip

echo.
echo ========================================================
echo   🎉 성공! 실시간 웹사이트에 최신 장부가 반영되었습니다.
echo ========================================================
echo.
echo   📱 계원들에게 보낼 주소 (클립보드에 자동 복사 완료!):
echo      %PAGE_URL%
echo.
echo   💡 안내:
echo      - 카카오톡 채팅방에서 [Ctrl + V] 로 바로 붙여넣기 하세요!
echo      - 'github.com/...'(코드 저장소)이 아닌 위 'github.io/...'
echo        주소를 보내셔야 계원들의 폰에서 예쁜 장부 화면이 열립니다.
echo.
echo ========================================================
echo.
echo 잠시 후 웹 브라우저에서 최신 장부 사이트가 자동으로 열립니다...
timeout /t 3 > nul
start "" "%PAGE_URL%"
