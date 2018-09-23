@echo off

title DeskBand Debug Installer

SET path_dll=%1
SET path_regasm=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

net session >nul 2>&1
if NOT %errorLevel% == 0 (
	echo You need administrator privileges.
	pause >nul
	exit
)

if NOT exist %path_dll% (
    echo The file %path_dll% could not be found.
	pause >nul
	exit
)

if NOT exist "%path_regasm%" (
    echo The file RegAsm could not be found.
	pause >nul
	exit
)

"%path_regasm%" /codebase %path_dll%

echo.
echo Successfully installed.
echo.