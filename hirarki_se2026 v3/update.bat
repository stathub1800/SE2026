@echo off
chcp 65001 >nul
setlocal

set REPO=C:\Users\supoy\hirarki_se2026
set LOG=%REPO%\log_update.txt

cd /d "%REPO%" || (echo Folder %REPO% tidak ditemukan & exit /b 1)

echo. >> "%LOG%"
echo ================================================== >> "%LOG%"
echo [%date% %time%] mulai >> "%LOG%"

REM --- 1. bangun ulang index.html dari berkas di folder data\ ---
python build.py >> "%LOG%" 2>&1
if errorlevel 1 (
  echo [%date% %time%] GAGAL membangun index.html >> "%LOG%"
  exit /b 1
)

REM --- 2. samakan dulu dengan GitHub agar tidak bentrok ---
git pull --rebase --autostash origin main >> "%LOG%" 2>&1

REM --- 3. kirim kalau memang ada perubahan ---
git add -A >> "%LOG%" 2>&1
git diff --cached --quiet
if not errorlevel 1 (
  echo [%date% %time%] tidak ada perubahan, tidak perlu commit >> "%LOG%"
  exit /b 0
)

git commit -m "Perbarui data SE2026 %date% %time%" >> "%LOG%" 2>&1
git push origin main >> "%LOG%" 2>&1
if errorlevel 1 (
  echo [%date% %time%] GAGAL push ke GitHub >> "%LOG%"
  exit /b 1
)

echo [%date% %time%] selesai, halaman sudah diperbarui >> "%LOG%"
exit /b 0
