@cls
@echo off
cd /d  "%~dp0"
cd..
echo:
echo --- Compress OfficeRTool Folder
>nul call %windir%\Compress.cmd * "%USERPROFILE%\desktop\OfficeRTool"
echo --- Backup OfficeRTool Folder
>nul 2>&1 copy /y "%USERPROFILE%\desktop\OfficeRTool.*" "D:\Software\MS Tools Pack\Office" || (echo:&echo ERROR ### FAIL to copy&echo:&pause)
explorer "D:\Software\MS Tools Pack\Office"
timeout 4
exit /b
