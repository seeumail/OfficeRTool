@cls
@echo off
setLocal EnableExtensions EnableDelayedExpansion

set "installfolder=%~dp0"
set "installfolder=%installfolder:~0,-1%"
set "installername=%~n0.cmd"

set "SingleNul=>nul"
set "SingleNulV1=1>nul"
set "SingleNulV2=2>nul"
set "SingleNulV3=3>nul"
set "MultiNul=1>nul 2>&1"
set "TripleNul=1>nul 2>&1 3>&1"

%MultiNul% fltmc || (
	set "_=call "%~dpfx0" %*" & powershell -nop -c start cmd -args '/d/x/r',$env:_ -verb runas || (
	>"%temp%\Elevate.vbs" echo CreateObject^("Shell.Application"^).ShellExecute "%~dpfx0", "%*" , "", "runas", 1
	%SingleNul% "%temp%\Elevate.vbs" & del /q "%temp%\Elevate.vbs" )
	exit)

cls
cd /D "%installfolder%"
for /F "tokens=*" %%a in (package.info) do (
	SET /a countx=!countx! + 1
	set var!countx!=%%a
)
if %countx% LSS 5 ((echo:)&&(echo Download incomplete - Package unusable - Redo download)&&(echo:)&&(pause)&&(exit))
set "instversion=%var2%"
set "instlang=%var3%"
set "instarch1=%var4%"
set "instupdlocid=%var5%"

if /i "%instarch1%" equ "x86" set "instarch2=32"
if /i "%instarch1%" equ "x64" set "instarch2=64"
if /i "%instarch1%" equ "x64" if not exist "%systemroot%\SysWOW64\cmd.exe" ((echo.)&&(echo ERROR: You can't install x64/64bit Office on x86/32bit Windows)&&(echo.)&&(pause)&&(exit))

if /i "%instarch1%" equ "multi" (

	if not exist configure32.xml (
		pause
		exit /b
	)
	
	if not exist configure64.xml (
		pause
		exit /b
	)
	
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set instarch2=32)
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set instarch2=64)
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set instarch2=64
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set instarch2=64
)

