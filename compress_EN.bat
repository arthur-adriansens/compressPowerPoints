:: Copyright (c) 2025 Arthur
:: Licensed under CC BY-NC 4.0 — Non-commercial use only, with attribution
:: https://creativecommons.org/licenses/by-nc/4.0/

@echo off
setlocal enabledelayedexpansion
:: Enable UTF-8 for emojis and special characters
chcp 65001

set "count=0"
set "done=0"

for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"

:: Path to 7-Zip and compression tools
set "TOOLS_DIR=%~dp0tools"
set "SEVENZIP=%TOOLS_DIR%\7-Zip\7z.exe"
set "OPTIPNG=%TOOLS_DIR%\optipng.exe"
set "JPEGTRAN=%TOOLS_DIR%\jpegtran.exe"

:: Check for .pptx-files, ">nul 2>&1" prevents echo
dir /b *.pptx >nul 2>&1 || (
    echo X Geen PowerPoints gevonden
    pause
    exit /b
)

<nul set /p ="%ESC%[30;102m🟢 Step 1: Unpack all PowerPoints...%ESC%[0m"
echo.

:: Go through all .pptx-files
for /R %%f in (*.pptx) do (
    :: Create temporary work folder in current folder
    set "BESTANDSNAAM=%%~nf"
    set "WERKMAP=%cd%\pptx-tijdelijk_!BESTANDSNAAM!"

    echo ---------------------------------------------
    echo 🔄 Unpacking: %%f...
    
    :: Delete old work folders if they excist
    if exist "!WERKMAP!" rd /s /q "!WERKMAP!" >nul 2>&1
    md "!WERKMAP!"

    :: Unpack .pptx
    "%SEVENZIP%" x "%%f" -o"!WERKMAP!" >nul

    <nul set /p ="%ESC%[30;102m📂 Compression of images in !WERKMAP!\ppt\media...%ESC%[0m"
    echo.

    :: Compress PNG's lossless with progres and filtered output
    if exist "!WERKMAP!\ppt\media\*.png" (
        set /a count=0
        for /f %%i in ('dir /b /a:-d "!WERKMAP!\ppt\media\*.png" 2^>nul') do (
            set /a count+=1
        )

        set /a done=0
        for /f %%i in ('dir /b /a:-d "!WERKMAP!\ppt\media\*.png" 2^>nul') do (
            set /a done+=1
            set /a progress="(!done!*30)/!count!"
            set "bar=["
            for /L %%p in (1,1,!progress!) do set "bar=!bar!#"
            for /L %%s in (!progress!,1,29) do set "bar=!bar!."
            set "bar=!bar!]"
            <nul set /p="🔧 !bar! !done!/!count! `%%~nxi` "

            Rem "%OPTIPNG%" -o2 -quiet "!WERKMAP!\ppt\media\%%i"
            set "output="
            for /f "delims=" %%r in ('cmd /c "%OPTIPNG% -o2 "!WERKMAP!\ppt\media\%%i" 2>&1 | findstr /C:"Output file size =" "') do set "output=%%r"
            
            if defined output (
                echo 👉 !output!
            ) else (
                echo ✅ Reeds geoptimaliseerd: %%~nxi
            )
        )
    )

    :: Compress JPG's lossless
    if exist "!WERKMAP!\ppt\media\*.jpeg" (
        for /f %%j in ('dir /b /a:-d "!WERKMAP!\ppt\media\*.jpeg" 2^>nul') do (
            echo "🤖🪄 JPEG: %%j"

            "%JPEGTRAN%" -optimize -copy none -outfile "!WERKMAP!\ppt\media\%%j" "!WERKMAP!\ppt\media\%%j"
        )
    )

    echo ✅ Compressie afgerond voor !BESTANDSNAAM!
    echo ---------------------------------------------
    echo.
)

echo 🟠 Press ENTER to pack the PowerPoints
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
