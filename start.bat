@echo off
chcp 65001 > nul
title PDF & Image Master Web Server
cls
echo ========================================================
echo    🚀 กำลังเริ่มระบบ PDF & Image Master...
echo ========================================================
echo.
echo กำลังเปิดเว็บเบราว์เซอร์ไปที่ http://localhost:3000
start http://localhost:3000
echo.
echo กด Ctrl+C เพื่อหยุดการทำงานของเซิร์ฟเวอร์
echo ========================================================
echo.
node server.js
pause
