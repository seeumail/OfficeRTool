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

**Copy / Paste to PowerShell Console**

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

**OR Save to [.PS1] file. And run this command**

````
powershell -noprofile -executionpolicy bypass -file "YOUR_FILE_HERE"
````
