@echo off
echo ==================================================
echo   SIMULASI SERVER NETLIFY (OFFLINE MODE)
echo ==================================================
echo.
echo Script ini akan menjalankan server lokal untuk 
echo folder "Netlify_Lengkap" agar Anda bisa mencoba
echo simulasi tampilan persis seperti di Netlify.
echo.
echo Akses:
echo - Karyawan: http://localhost:8888/
echo - Admin:    http://localhost:8888/admin/
echo.
echo JANGAN TUTUP JENDELA INI.
echo.

cd /d "%~dp0"
start http://localhost:8888/
start http://localhost:8888/admin/

python -m http.server 8888
pause