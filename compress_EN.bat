:: Copyright (c) 2025 Arthur
:: Licensed under CC BY-NC 4.0 — Non-commercial use only, with attribution
:: https://creativecommons.org/licenses/by-nc/4.0/

@echo off
setlocal enabledelayedexpansion
:: Enable UTF-8 for emojis and special characters
chcp 65001

set count=0
set done=0

for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"

:: Path to 7-Zip and compression tools
set "TOOLS_DIR=%~dp0tools"
set "SEVENZIP=%TOOLS_DIR%\7-Zip\7z.exe"
set "OPTIPNG=%TOOLS_DIR%\optipng.exe"
set "JPEGTRAN=%TOOLS_DIR%\jpegtran.exe"

:: Check for .pptx-files, ">nul 2>&1" prevents echo
dir /b *.pptx >nul 2>&1 || (
    echo X No PowerPoints found
    pause
    exit /b
)

<nul set /p ="%ESC%[30;102m🟢🚀 Step 1: Unpack all PowerPoints...%ESC%[0m"

:: Go through all .pptx-files
for /R %%f in (*.pptx) do (
    :: Create temporary work folder in current folder
    set "BESTANDSNAAM=%%~nf"
    set "WERKMAP=%cd%\pptx-temporary_!BESTANDSNAAM!"

    echo ---------------------------------------------
    echo 🔄 Unpacking: %%f...
    
    :: Delete old work folders if they excist
    if exist "!WERKMAP!" rd /s /q "!WERKMAP!" >nul 2>&1
    md "!WERKMAP!"

    :: Pak .pptx uit
    "%SEVENZIP%" x "%%f" -o"!WERKMAP!" >nul

    echo 📂 Compression of images in !WERKMAP!\ppt\media...

    :: Compress PNG's lossless with progres and filtered output
    if exist "!WERKMAP!\ppt\media\*.png" (
        for %%i in ("!WERKMAP!\ppt\media\*.png") do (
            "%OPTIPNG%" -o2 "%%i" >nul
        )
    )

    :: Compress JPG's lossless
    if exist "!WERKMAP!\ppt\media\*.jpg" (
        for %%j in ("!WERKMAP!\ppt\media\*.jpg") do (
            "%JPEGTRAN%" -optimize -copy none -outfile "%%j" "%%j" >nul
        )
    )

    echo ✅ Compression finished for !BESTANDSNAAM!
    echo ---------------------------------------------
)

echo.
<nul set /p ="%ESC%[30;102m🟠 Press ENTER to pack the PowerPoints%ESC%[0m"
pause

<nul set /p ="%ESC%[30;102m🟢 Step 2: Pack all the PowerPoints...%ESC%[0m"

for /R %%f in (*.pptx) do (
    set "BESTANDSNAAM=%%~nf"
	set "BESTANDSPAD=%%~dpf"
    set "WERKMAP=%cd%\pptx-tijdelijk_!BESTANDSNAAM!"
    
    echo 📦 Packing: !BESTANDSNAAM!...

    pushd "!WERKMAP!"
    "!SEVENZIP!" a -tzip "!BESTANDSPAD!comprim_!BESTANDSNAAM!.pptx" * >nul 2>&1
    popd

    echo 🚀 Finished: !BESTANDSPAD!comprim_!BESTANDSNAAM!.pptx

    if exist "!WERKMAP!" rd /s /q "!WERKMAP!" >nul 2>&1
)

echo.
echo 🎉 Everything is processed. Press ENTER to close the script.
pause
