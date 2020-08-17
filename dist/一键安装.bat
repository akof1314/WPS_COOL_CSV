@echo off
@set baseDir=%~dp0

regedit /s %baseDir%\pkg\install.reg

C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm %baseDir%\pkg\wps-cool-csv.dll /tlb:%baseDir%\pkg\wps-cool-csv.tlb

@SET GACUTIL="%baseDir%\pkg\gacutil.exe"

%GACUTIL% -i %baseDir%\pkg\wps-cool-csv.dll

if not exist C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Excel (
	%GACUTIL% -i %baseDir%\pkg\Excel.dll
)

if not exist C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Office (
	%GACUTIL% -i %baseDir%\pkg\Office.dll
)

pause
