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

**PS Version**

**Option A - Copy / Paste to PowerShell Console**

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

**Option B - Save as [.PS1] file. And run this command**

````
powershell -noprofile -executionpolicy bypass -file "YOUR_FILE_HERE"
````

**Wget Version**

Save as [.cmd] file. Run it later.

Must change WGET path, to your path.
````
@cls
@echo off
	
Set TAG=
set URI=
set OfficeRToolLink=
set Wget="c:\windows\wget.exe"  || WGET TOOL FULL ADDRESS
set "FileName=OfficeRTool.RAR"
set Latest="%temp%\latest"
set "GitHub=https://github.com/DarkDinosaurEx/OfficeRTool/releases"
set URL="%GitHub%/latest"
if exist %Latest% del /q %Latest%
REM start "" /min /wait %wget% --max-redirect=0 %url% --output-file=%Latest%
powershell -noprofile -executionpolicy bypass -command start '%wget:~1,-1%' -Wait -WindowStyle hidden -Args '--max-redirect=0 %url% --output-file=\"%Latest:~1,-1%\"'
if exist %Latest% for /f "tokens=2 delims= " %%$ in ('"type %Latest% | find /i "tag""') do set "URI=%%$"
if defined URI echo "%URI:~59%" | >nul findstr /r [0-9].[0-9] 				&& set "TAG=%URI:~59%"
if defined URI echo "%URI:~59%" | >nul findstr /r [0-9][0-9].[0-9][0-9] 	&& set "TAG=%URI:~59%"
if defined URI echo "%URI:~59%" | >nul findstr /r [0-9][0-9].[0-9] 			&& set "TAG=%URI:~59%"
if defined URI echo "%URI:~59%" | >nul findstr /r [0-9].[0-9][0-9] 			&& set "TAG=%URI:~59%"
if defined TAG set "OfficeRToolLink=%GitHub%/download/%tag%/%FileName%"
if defined TAG echo:&echo Download Latest Release --- v%TAG%
if defined OfficeRToolLink 2>nul %wget% --quiet --no-check-certificate --content-disposition --output-document="%USERPROFILE%\DESKTOP\%FileName%" "%OfficeRToolLink%"
````
