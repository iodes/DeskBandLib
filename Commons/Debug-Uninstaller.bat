@echo off

title DeskBand Debug Uninstaller

SET path_dll=%1
SET path_gacutil=%PROGRAMFILES(X86)%\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.6.1 Tools\x64\gacutil.exe
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

if NOT exist "%path_gacutil%" (
    echo The file gacutil could not be found.
	pause >nul
	exit
)

if NOT exist "%path_regasm%" (
    echo The file RegAsm could not be found.
	pause >nul
	exit
)

"%path_gacutil%" /u DeskBandLib
"%path_gacutil%" /u %~n1
"%path_regasm%" /u %path_dll%

taskkill /im explorer.exe /f
start explorer.exe

echo.
echo Successfully uninstalled.
echo.