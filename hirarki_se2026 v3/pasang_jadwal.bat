@echo off
chcp 65001 >nul
REM Mendaftarkan pembaruan harian pukul 08.15. Klik kanan -> Run as administrator.

schtasks /Create /F /TN "SE2026 Hirarki - Update Harian" ^
 /TR "\"C:\Users\supoy\hirarki_se2026\update.bat\"" ^
 /SC DAILY /ST 08:15 /RL HIGHEST

echo.
echo Jadwal terpasang. Untuk menguji sekarang:
echo    schtasks /Run /TN "SE2026 Hirarki - Update Harian"
echo Untuk membatalkan:
echo    schtasks /Delete /TN "SE2026 Hirarki - Update Harian" /F
pause
