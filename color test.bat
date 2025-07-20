@echo off

powershell write-host -fore Cyan This is Cyan text
powershell write-host -back Red This is Red background

call :color 4
call :echo Red foreground
call :color 7 " and "
call :color 4f
echo Red background

call :color 
echo Back to normal

call :color 70 "Black "
call :color 1 "Blue "
call :color 2 "Green "
call :color 3 "Aqua "
call :color 4 "Red "
call :color 5 "Purple "
call :color 6 "Yellow "
call :color 7 "White "
call :color 8 "Gray "
call :color 9 "LightBlue" $
call :color a "LightGreen "
call :color b "LightAqua "
call :color c "LightRed "
call :color d "LightPurple "
call :color e "LightYellow "
call :color f "BrightWhite " $

call :color 1f Blue back
call :color 2f Green back
call :color 3f Aqua back
call :color 4f Red back
call :color 5f Purple back
call :color 6f Yellow back
call :color 7f White back
call :color 8f Gray back
call :color 9f "LightBlue back" $
call :color a0 LightGreen back
call :color b0 LightAqua back
call :color c0 LightRed back
call :color d0 LightPurple back
call :color e0 LightYellow back
call :color f0 LightWhite back $

call :color
echo %ESC%[4mUnderline%ESC%[0m.
pause 


goto :eof

:: Displays a text without new line at the end (unlike echo)
:echo
@<nul set /p ="%*"
@goto :eof

:: Change color to the first parameter (same codes as for the color command) 
:: And display the other parameters (write $ at the end for new line)
:color
@echo off
IF [%ESC%] == [] for /F %%a in ('echo prompt $E ^| cmd') do set "ESC=%%a"
SET color=0%1
IF [%color%] == [0] SET color=07
SET fore=%color:~-1%
SET back=%color:~-2,1% 
SET color=%ESC%[
if %fore% LEQ 7 (
  if %fore% == 0 SET color=%ESC%[30
  if %fore% == 1 SET color=%ESC%[34
  if %fore% == 2 SET color=%ESC%[32
  if %fore% == 3 SET color=%ESC%[36
  if %fore% == 4 SET color=%ESC%[31
  if %fore% == 5 SET color=%ESC%[35
  if %fore% == 6 SET color=%ESC%[33
  if %fore% == 7 SET color=%ESC%[37
) ELSE (
  if %fore% == 8 SET color=%ESC%[90
  if %fore% == 9  SET color=%ESC%[94
  if /i %fore% == a SET color=%ESC%[92
  if /i %fore% == b SET color=%ESC%[96
  if /i %fore% == c SET color=%ESC%[91
  if /i %fore% == d SET color=%ESC%[95
  if /i %fore% == e SET color=%ESC%[93
  if /i %fore% == f SET color=%ESC%[97
)
if %back% == 0 (SET color=%color%;40) ELSE (
  if %back% == 1 SET color=%color%;44
  if %back% == 2 SET color=%color%;42
  if %back% == 3 SET color=%color%;46
  if %back% == 4 SET color=%color%;41
  if %back% == 5 SET color=%color%;45
  if %back% == 6 SET color=%color%;43
  if %back% == 7 SET color=%color%;47
  if %back% == 8 SET color=%color%;100
  if %back% == 9  SET color=%color%;104
  if /i %back% == a SET color=%color%;102
  if /i %back% == b SET color=%color%;106
  if /i %back% == c SET color=%color%;101
  if /i %back% == d SET color=%color%;105
  if /i %back% == e SET color=%color%;103
  if /i %back% == f SET color=%color%;107
)
SET color=%color%m
:repeatcolor
if [%2] NEQ [$] SET color=%color%%~2
shift
if [%2] NEQ [] if [%2] NEQ [$] SET color=%color% & goto :repeatcolor
if [%2] EQU [$] (echo %color%) else (<nul set /p ="%color%")
goto :eof

pause