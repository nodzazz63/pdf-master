@echo off
chcp 65001 > nul
title PDF & Image Master Web Server
cls
echo ========================================================
echo    🚀 กำลังเริ่มระบบ PDF & Image Master...
echo ========================================================
echo กำลังเปิดเว็บเบราว์เซอร์ไปที่ http://localhost:3000 (กำลังรอเซิร์ฟเวอร์พร้อม)...
start "" /b cmd /c "timeout /t 2 /nobreak > nul & start http://localhost:3000"
echo.
echo กด Ctrl+C เพื่อหยุดการทำงานของเซิร์ฟเวอร์
echo ========================================================
echo.
node server.js
pause
