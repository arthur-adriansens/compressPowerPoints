@echo off
for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"
<nul set /p ="%ESC%[30;102m🚀 Start compressie%ESC%[0m"
echo.
echo "hi"
pause