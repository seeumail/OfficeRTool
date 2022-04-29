# OfficeRTool
**OfficeRTool Offical GitHub Page**

**Tool To Install, Activate, Config - Office 2016 - 2021**

**Original author : ratzlefatz - Current maintainer : Mr Dino.**

**Official MDL Thread For help & support**

https://forums.mydigitallife.net/threads/84450/

**Video - how to use [made by BredzPro]**

https://www.youtube.com/watch?v=DpK5R_IOqgk

# Main Changes

- User friendly Interface
- Up to date Activation files
- Auto Create Package Info file
- Auto Detect system Arch. & Lang
- Visual Refresh for Current & LTSC Channels
- Support Multi Language / Architecture ISO Disk
- Support Online / Offline Install Include Create ISO
- Support install from ISO / Offline folder [ ¥ NEW FEATURE ¥ ]
- Support Activation & Convert for Office Products, Include 365 & Home
- Support Downloading Offline Image / Offline Package / Online setup [ ¥ NEW FEATURE ¥ ]
- Special Thanks to abbodi1406 for script advice Inc. VBS core file & Activation script / DLL

# How to get the latest release

**PowerShell Version**

**Option A - Copy / Paste to PowerShell Console**

**Option B - Save as [.PS1] file. And run this command :: [ powershell -noprofile -executionpolicy bypass -file "YOUR_FILE_HERE" ]**

````
<# Based on -- Using Powershell To Retrieve Latest Package Url From Github Releases #>
<# https://copdips.com/2019/12/Using-Powershell-to-retrieve-latest-package-url-from-github-releases.html #>
$url = 'https://github.com/DarkDinosaurEx/OfficeRTool/releases/latest'
$request = [System.Net.WebRequest]::Create($url)
$response = $request.GetResponse()
$realTagUrl = $response.ResponseUri.OriginalString
$version=$realTagUrl.split('/')[-1].Trim('v')
$fileName = "OfficeRTool.rar"
$realDownloadUrl = $realTagUrl.Replace('tag', 'download') + '/' + $fileName
$OutputFile = $env:USERPROFILE+'\desktop\'+$fileName
Invoke-WebRequest -Uri $realDownloadUrl -OutFile $OutputFile
[Environment]::Exit(1)
[Environment]::Exit(1)
````

**Wget Version**

Save as [.cmd] file. Run it later.

````
@cls
@echo off
>nul chcp 437
setlocal enabledelayedexpansion
title Office(R)Tool download tool

>nul fltmc || ( set "_=call "%~dpfx0" %*"
	powershell -nop -c start cmd -args '/d/x/r',$env:_ -verb runas || (
	mshta vbscript:execute^("createobject(""shell.application"").shellexecute(""cmd"",""/d/x/r "" &createobject(""WScript.Shell"").Environment(""PROCESS"")(""_""),,""runas"",1)(window.close)"^))|| (
	cls & echo:& echo Script elavation failed& pause)
	exit )

Set TAG=
set URI=
set OfficeRToolLink=
set Latest="%temp%\latest"
set wget="%windir%\wget.exe"
set "FileName=OfficeRTool.RAR"
set "GitHub=https://github.com/DarkDinosaurEx/OfficeRTool/releases"
set wget_url="https://raw.githubusercontent.com/DarkDinosaurEx/OfficeRTool/main/OfficeFixes/win_x32/wget.exe"
set "output_file=%USERPROFILE%\DESKTOP\%FileName%"
set URL="%GitHub%/latest"

if not exist %wget% >nul bitsadmin /transfer debjob /download /priority normal %wget_url% %wget%
if not exist %wget% goto :theEnd

if exist %Latest% del /q %Latest%
powershell -noprofile -executionpolicy bypass -command start '%wget:~1,-1%' -Wait -WindowStyle hidden -Args '--max-redirect=0 %url% --output-file=\"%Latest:~1,-1%\"'
if exist %Latest% for /f "tokens=2 delims= " %%$ in ('"type %Latest% | find /i "tag""') do set "URI=%%$"
if defined URI echo "%URI:~59%" | >nul findstr /r [0-9].[0-9] 			&& set "TAG=%URI:~59%"
if defined URI echo "%URI:~59%" | >nul findstr /r [0-9][0-9].[0-9][0-9] 	&& set "TAG=%URI:~59%"
if defined URI echo "%URI:~59%" | >nul findstr /r [0-9][0-9].[0-9] 		&& set "TAG=%URI:~59%"
if defined URI echo "%URI:~59%" | >nul findstr /r [0-9].[0-9][0-9] 		&& set "TAG=%URI:~59%"
if defined TAG set "OfficeRToolLink=%GitHub%/download/%tag%/%FileName%"
if defined OfficeRToolLink %wget% --quiet --no-check-certificate --content-disposition --output-document="%output_file%" "%OfficeRToolLink%"

:theEnd
echo:
if exist "%output_file%" (echo the download was successful.) else (echo the downloads have failed.)
echo:
pause
exit /b
````
