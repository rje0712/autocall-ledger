@echo off
chcp 65001 > nul
title 오토콜 회비 장부 웹사이트 업데이트
cls
echo ========================================================
echo   [오토콜 회비 장부] GitHub Pages 웹사이트 동기화
echo ========================================================
echo.
echo 1. 최신 데이터 HTML 빌드 중...
python build_ledger.py

echo.
echo 2. GitHub 웹사이트로 최신 장부 전송 중...
git add .
git commit -m "update: 최신 회비 장부 업데이트"
git push origin main

echo.
echo ========================================================
echo   성공! 24시간 실시간 웹사이트에 최신 장부가 반영되었습니다.
echo   접속 주소: https://rje0712.github.io/autocall-ledger
echo ========================================================
echo.
pause
