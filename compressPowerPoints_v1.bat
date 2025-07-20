@echo off
setlocal enabledelayedexpansion

:: Pad naar 7-Zip (pas aan indien nodig)
set "SEVENZIP=C:\Program Files\7-Zip\7z.exe"

:: Pad naar compressie tools
set "TOOLS_DIR=%~dp0tools"
set "OPTIPNG=%TOOLS_DIR%\optipng.exe"
set "JPEGTRAN=%TOOLS_DIR%\jpegtran.exe"

:: Controleer op .pptx-bestanden, ">nul 2>&1" onderdrukt echo
dir /b *.pptx >nul 2>&1 || (echo X Geen PowerPoints gevonden & pause & exit /b)

echo 🟢 Stap 1: Uitpakken van alle PowerPoints...

:: Doorloop alle .pptx-bestanden
for /R %%f in (*.pptx) do (
	:: Maak tijdelijke werkmap aan in de huidige map
	set "BESTANDSNAAM=%%~nf"
	set "WERKMAP=%cd%\pptx-tijdelijk_!BESTANDSNAAM!"

	echo ---------------------------------------------
    echo 🔄 Uitpakken: %%f...
	
    :: Verwijder oude werkmap en maak nieuwe, ">nul 2>&1" onderdrukt echo
    rd /s /q "!WERKMAP!" >nul 2>&1
    md "!WERKMAP!"

    :: Pak .pptx uit
    "%SEVENZIP%" x "%%f" -o"!WERKMAP!" >nul

    ::echo 📂 Afbeeldingen voor !BESTANDSNAAM! staan in:
    ::echo !WERKMAP!\ppt\media

	echo 📂 Compressie van afbeeldingen in !WERKMAP!\ppt\media...

	:: Compress PNG's lossless
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
echo 🟠 Druk op ENTER om de PowerPoints in te pakken
pause

echo 🟢 Stap 2: Inpakken van alle PowerPoints...

for /R %%f in (*.pptx) do (
	:: Maak tijdelijke werkmap aan in de huidige map
	set "BESTANDSNAAM=%%~nf"
	set "WERKMAP=%cd%\pptx-tijdelijk_!BESTANDSNAAM!"
	
    echo 📦 Inpakken: !BESTANDSNAAM!...

    pushd "!WERKMAP!"
    "!SEVENZIP!" a -tzip "..\comprim_!BESTANDSNAAM!.pptx" * >nul 2>&1
    popd

    echo ✅ Klaar: comprim_!BESTANDSNAAM!.pptx

	:: Opruimen,">nul 2>&1" onderdrukt echo
	rd /s /q "!WERKMAP!" >nul 2>&1 ">nul	
)

echo.
echo 🎉 Alles is verwerkt. Druk op ENTER om af te sluiten.
pause