@cls
@echo off
1>nul 2>nul fltmc || (
	set "_=call "%~dpfx0" %*" & powershell -nop -c start cmd -args '/d/x/r',$env:_ -verb runas || (
	>"%temp%\Elevate.vbs" echo CreateObject^("Shell.Application"^).ShellExecute "%~dpfx0", "%*" , "", "runas", 1
	>nul "%temp%\Elevate.vbs" & del /q "%temp%\Elevate.vbs" )
	exit)
	
if /i "%1" EQU "-Execute" goto :work
start "" "cmd" /k "%~dpfx0" -Execute
exit /b

:work
cd /d "%~dp0"
echo.
echo Script will Start Office(R)Tool in Debug Mode.
echo you Will see Blank Screen, and than 2 Beeps.
echo It's the Choice Menu, from here you can:
echo (1) Close Window / (2) press key for your selection
echo Than wait for 2 beeps, close script.
echo if it take too long, close script.!
echo post log file later in thread.!
echo.
echo.
pause
"OfficeRTool.cmd" -debug >log 2>&1 3>&1
exit /b