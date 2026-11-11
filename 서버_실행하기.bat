@echo off
chcp 65001 > nul
title 오토콜 회비 장부 로컬 웹서버
cls
echo ========================================================
echo   [오토콜 회비 장부] 로컬 웹 서버를 시작합니다.
echo ========================================================
echo.
echo PC 접속 주소:     http://localhost:8080/오토콜_회비장부.html
echo 모바일 접속 안내: 같은 와이파이(Wi-Fi)에 연결된 스마트폰에서
echo                  아래 IP 주소로 접속할 수 있습니다.
echo.
ipconfig | findstr /i "IPv4"
echo.
echo 서버를 종료하려면 이 창에서 Ctrl + C 를 누르세요.
echo ========================================================
echo.
start "" "http://localhost:8080/오토콜_회비장부.html"
python -m http.server 8080
pause
