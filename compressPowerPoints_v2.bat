@echo off
setlocal enabledelayedexpansion

set count=0
set done=0

for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"

:: Pad naar 7-Zip en compressie tools
set "TOOLS_DIR=%~dp0tools"
set "SEVENZIP=%TOOLS_DIR%\7-Zip\7z.exe"
set "OPTIPNG=%TOOLS_DIR%\optipng.exe"
set "JPEGTRAN=%TOOLS_DIR%\jpegtran.exe"

:: Controleer op .pptx-bestanden, ">nul 2>&1" onderdrukt echo
dir /b *.pptx >nul 2>&1 || (
    echo X Geen PowerPoints gevonden
    pause
    exit /b
)

<nul set /p ="%ESC%[30;102m🟢 Stap 1: Uitpakken van alle PowerPoints...%ESC%[0m"

:: Doorloop alle .pptx-bestanden
for /R %%f in (*.pptx) do (
    :: Maak tijdelijke werkmap aan in de huidige map
    set "BESTANDSNAAM=%%~nf"
    set "WERKMAP=%cd%\pptx-tijdelijk_!BESTANDSNAAM!"

    echo ---------------------------------------------
    echo 🔄 Uitpakken: %%f...
    
    :: Verwijder oude werkmap als die bestaat
    if exist "!WERKMAP!" rd /s /q "!WERKMAP!" >nul 2>&1
    md "!WERKMAP!"

    :: Pak .pptx uit
    "%SEVENZIP%" x "%%f" -o"!WERKMAP!" >nul

    echo 📂 Compressie van afbeeldingen in !WERKMAP!\ppt\media...

    :: Compress PNG's lossless met voortgang en gefilterde output
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

    echo ✅ Compressie afgerond voor !BESTANDSNAAM!
    echo ---------------------------------------------
)

echo.
<nul set /p ="%ESC%[30;102m🟠 Druk op ENTER om de PowerPoints in te pakken%ESC%[0m"
pause

<nul set /p ="%ESC%[30;102m🟢 Stap 2: Inpakken van alle PowerPoints...%ESC%[0m"

for /R %%f in (*.pptx) do (
    set "BESTANDSNAAM=%%~nf"
	set "BESTANDSPAD=%%~dpf"
    set "WERKMAP=%cd%\pptx-tijdelijk_!BESTANDSNAAM!"
    
    echo 📦 Inpakken: !BESTANDSNAAM!...

    pushd "!WERKMAP!"
    "!SEVENZIP!" a -tzip "!BESTANDSPAD!comprim_!BESTANDSNAAM!.pptx" * >nul 2>&1
    popd

    echo ✅ Klaar: !BESTANDSPAD!comprim_!BESTANDSNAAM!.pptx

    if exist "!WERKMAP!" rd /s /q "!WERKMAP!" >nul 2>&1
)

echo.
echo 🎉 Alles is verwerkt. Druk op ENTER om af te sluiten.
pause
