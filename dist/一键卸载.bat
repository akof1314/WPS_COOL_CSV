@echo off
@set baseDir=%~dp0

@SET GACUTIL="%baseDir%\pkg\gacutil.exe"

rd /s /Q C:\Windows\Microsoft.NET\assembly\GAC_MSIL\wps-cool-csv

rem rd /s /Q C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Excel

rem rd /s /Q C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Office


C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm /u %baseDir%\pkg\wps-cool-csv.dll /tlb:%baseDir%\pkg\wps-cool-csv.tlb

Echo.


regedit /s %baseDir%\pkg\uninstall.reg

pause
