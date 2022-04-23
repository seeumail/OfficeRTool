
	@cls
	@echo off
	>nul chcp 437
	
	set "SingleNul=>nul"
	set "SingleNulV1=1>nul"
	set "SingleNulV2=2>nul"
	set "SingleNulV3=3>nul"
	set "MultiNul=1>nul 2>&1"
	set "TripleNul=1>nul 2>&1 3>&1"
	setLocal EnableExtensions EnableDelayedExpansion
	
	set "latestversion="
	set "verifiedversion="
	set "Currentversion=1.1"
	for /f "tokens=*" %%$ in ('"powershell -noprofile -executionpolicy bypass -file "%~dp0OfficeFixes\CheckLatestRelease.ps1""') do set "verifiedversion=%%$"
	echo "!verifiedversion!" | >nul findstr /r "[0-9].[0-9]" 			&& set "latestversion=!verifiedversion!"
	echo "!verifiedversion!" | >nul findstr /r "[0-9].[0-9][0-9]"		&& set "latestversion=!verifiedversion!"
	echo "!verifiedversion!" | >nul findstr /r "[0-9][0-9].[0-9]"		&& set "latestversion=!verifiedversion!"
	echo "!verifiedversion!" | >nul findstr /r "[0-9][0-9].[0-9][0-9]"	&& set "latestversion=!verifiedversion!"
	
	
	if defined latestVersion if !latestVersion! GTR !CurrentVersion! (
		echo:
		echo Found new Release.
		echo:
		echo Current Release :: !CurrentVersion!
		echo Latest Release  :: !latestVersion!
		echo:
		echo Please Update version to latest version
		echo https://github.com/maorosh123/OfficeRTool/releases/
		echo:
		pause
	)
	
	if /i "%*" 	EQU "-debug" (
		echo on
		set "SingleNul="
		set "SingleNulV1="
		set "SingleNulV2="
		set "SingleNulV3="
		set "MultiNul="
		set "TripleNul="
		set "debugMode=on"
	)

	set debugMode=
	set inidownpath=
	set inidownarch=
	set inidownlang=
	set DontSaveToIni=true
	set AutoSaveToIni=true
	
	set "OfficeRToolpath=%~dp0"
	set "OfficeRToolpath=%OfficeRToolpath:~0,-1%"
	set "OfficeRToolname=%~n0.cmd"
	
	color 0F
	mode con cols=140 lines=45
	
	title OfficeRTool - 2022/APR/23 -
	set "pswindowtitle=$Host.UI.RawUI.WindowTitle = 'Administrator: OfficeRTool - 2022/APR/23 -'"
	
	echo "%~dp0"|%SingleNul% findstr /L "%% # & ^ ^^ @ $ ~ ! ( )" && (
	echo.
	Echo Invalid path: "%~dp0"
	Echo Remove special symbols: "%% # & ^ @ $ ~ ! ( )"
	if not defined debugMode pause
	exit /b
	) || cd /d "%OfficeRToolpath%"
	
	echo. >"OfficeFixes\dummyfile" && %SingleNul% del /q "OfficeFixes\dummyfile" || (
		cls
		echo.
		echo ERROR ### Read Only Folder
		echo.
		if not defined debugMode pause
		exit /b
	)
	
	echo.
	set "missingFiles="
	set "binDLL=A64.dll SvcTrigger.xml x64.dll x86.dll"
	for %%# in (!binDLL!) do if not exist "OfficeFixes\bin\%%#" echo OfficeFixes\bin\%%# IS Missing & set "missingFiles=true"
	if defined missingFiles ( echo. & pause & exit /b ) else ( cls )
	
	set OSPP_HKLM=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform
	set OSPP_USER=HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform
	set XSPP_HKLM_X32=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
	set XSPP_HKLM_X64=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform
	set XSPP_USER=HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform

	rem Run as administrator, AveYo: ps\VBS version
	>nul fltmc || ( set "_=call "%~dpfx0" %*"
		powershell -nop -c start cmd -args '/d/x/r',$env:_ -verb runas || (
		mshta vbscript:execute^("createobject(""shell.application"").shellexecute(""cmd"",""/d/x/r "" &createobject(""WScript.Shell"").Environment(""PROCESS"")(""_""),,""runas"",1)(window.close)"^))|| (
		cls & echo:& echo Script elavation failed& pause)
		exit )
		
	set "WSH_Disabled=" & for %%$ in (HKCU, HKLM) do 2>nul reg query "%%$\Software\Microsoft\Windows Script Host\Settings" /v "Enabled" | >nul find /i "0x0" && set "WSH_Disabled=***"
	if defined WSH_Disabled for %%$ in (HKCU, HKLM) do %MultiNul% REG DELETE "%%$\SOFTWARE\Microsoft\Windows Script Host\Settings" /f /v Enabled
	set "WSH_Disabled=" & for %%$ in (HKCU, HKLM) do 2>nul reg query "%%$\Software\Microsoft\Windows Script Host\Settings" /v "Enabled" | >nul find /i "0x0" && set "WSH_Disabled=***"
	if defined WSH_Disabled (
		cls
		echo.
		echo ERROR ### Windows script host is disabled
		echo.
		if not defined debugMode pause
		exit /b
	)
	
	%MultiNul% del /q latest*.txt
	
	if /i "%*" EQU "-debug" (
		echo on
		set "SingleNul="
		set "SingleNulV1="
		set "SingleNulV2="
		set "SingleNulV3="
		set "MultiNul="
		set "TripleNul="
		set "debugMode=on"
		call :debugMode
		exit /b
	)
	
	rem Check if ANSI Colors is supported
	rem https://ss64.com/nt/syntax-ansi.html
	
	for /F "tokens=4,6 delims=[]. " %%a in ('ver') do (
		if %%a GEQ 10 (
			if %%b GEQ 10586 (
			
				rem [ANSI colors are available by default in Windows version 1909 or newer]
				if %%b GEQ 18363 goto :UseANSIColors
				
				rem [from 10.0.10586 to 10.0.18362.239 we need VirtualTerminalLevel Enabled -> 0x1]
				2>nul reg query HKCU\Console /v VirtualTerminalLevel | >nul find /i "0x1" && goto :UseANSIColors
				
				rem [we add it now, later in next run we will using ANSI code]
				>nul 2>&1 reg add HKCU\Console /f /v VirtualTerminalLevel /t REG_DWORD /d 0x1
			)
		)
	)
	
	rem lean xp+ color macros by AveYo:  %<%:af " hello "%>>%  &  %<%:cf " w\"or\"ld "%>%   for single \ / " use .%ESC%\  .%ESC%/  \"%ESC%\"
	for /f "delims=:" %%$ in ('"echo;prompt $h$s$h:|cmd /d"') do set "ESC=%%$"
	set "<=pushd "%appdata%"&2>nul findstr /c:\ /a"
	set ">=\..\c nul&set /p s=%ESC%%ESC%%ESC%%ESC%%ESC%%ESC%%ESC%<nul&popd&echo;"
	set ">>=%>:~0,-6%"
	set /p s=\<nul>"%appdata%\c"
	set "ANSI_COLORS="
	
	:: BACKGROUND COLORS
	set "B_Black=0" & set "B_Red=4" & set "B_Green=2"
	set "B_Yellow=6" & set "B_Blue=1" & set "B_Magenta=5"
	set "B_White=7" & set "B_Gray=8" & set "B_Aqua=3"
	set "BB_Black=0" & set "BB_Red=C" & set "BB_Green=A"
	set "BB_Yellow=E" & set "BB_Blue=9" & set "BB_Magenta=D"
	set "BB_White=F" & set "BB_Gray=8" & set "BB_Aqua=B"

	:: FORGROUND COLORS
	set "F_Black=0" & set "F_Red=4" & set "F_Green=2"
	set "F_Yellow=6" & set "F_Blue=1" & set "F_Magenta=5"
	set "F_White=7" & set "F_Gray=8" & set "F_Aqua=3"
	set "FF_Black=0" & set "FF_Red=C" & set "FF_Green=A"
	set "FF_Yellow=E" & set "FF_Blue=9" & set "FF_Magenta=D"
	set "FF_White=F" & set "FF_Gray=8" & set "FF_Aqua=B"
	goto :debugMode
	
	:UseANSIColors
	
	rem ANSI Colors in standard Windows 10 shell
	rem https://gist.github.com/mlocati/fdabcaeb8071d5c75a2d51712db24011
	
	set "ANSI_COLORS=*"
	
	:: BASIC CHARS
	:: ALT 0,2,7 --> 
	:: WORK WITH NOTPAD++
	set "<=["
	set ">=[0m"

	:: STYLES
	set "Reset=0m" & set "Bold=1m"
	set "Underline=4m" & set "Inverse=7m"

	:: BACKGROUND COLORS
	set "B_Black=30m" & set "B_Red=31m" & set "B_Green=32m"
	set "B_Yellow=33m" & set "B_Blue=34m" & set "B_Magenta=35m"
	set "B_Cyan=36m" & set "B_White=37m"
	set "BB_Black=90m" & set "BB_Red=91m" & set "BB_Green=92m"
	set "BB_Yellow=93m" & set "BB_Blue=94m" & set "BB_Magenta=95m"
	set "BB_Cyan=96m" & set "BB_White=97m"

	:: FOREGROUND COLORS
	set "F_Black=40m" & set "F_Red=41m" & set "F_Green=42m"
	set "F_Yellow=43m" & set "F_Blue=44m" & set "F_Magenta=45m"
	set "F_Cyan=46m" & set "F_White=47m"
	set "FF_Black=100m" & set "FF_Red=101m" & set "FF_Green=102m"
	set "FF_Yellow=103m" & set "FF_Blue=104m" & set "FF_Magenta=105m"
	set "FF_Cyan=106m" & set "FF_White=107m"

:debugMode

::===============================================================================================================
::===============================================================================================================
	cls
	mode con cols=140 lines=45
	color 0F
	echo:
::===============================================================================================================
:: DEFINE SYSTEM ENVIRONMENT
::===============================================================================================================
	for /F "tokens=6 delims=[]. " %%A in ('ver') do set /a win=%%A
	if %win% LSS 7601 (echo:)&&(echo:)&&(echo Unsupported Windows detected)&&(echo:)&&(echo Minimum OS must be Windows 7 SP1 or better)&&(echo:)&&(goto:TheEndIsNear)
	
	call :query "AddressWidth" "Win32_Processor"
	for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set "tmpX=%%g"
	((set winx=win_x%tmpX: =%)&&(set "repairplatform=x%tmpX: =%"))
	
	call :CheckSystemLanguage
	set "repairlang=!o16lang!"
	
	set "sls=SoftwareLicensingService"
	set "slp=SoftwareLicensingProduct"
	set "osps=OfficeSoftwareProtectionService"
	set "ospp=OfficeSoftwareProtectionProduct"
	
	call :query "version" "%sls%"
	for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set slsVer=%%g
	set "slsversion=%slsVer: =%"
	
	if %win% LSS 9200 (
		call :query "version" "%osps%"
		for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set ospsVer=%%g
		set "ospsversion=%ospsVer: =%"
	)
	
	call :Get-WinUserLanguageList_Warper
	
	cd /D "%OfficeRToolpath%"
	if not exist OfficeRTool.ini (
		set "CreateIniFile="
		if not defined DontSaveToIni	set CreateIniFile=***
		if defined AutoSaveToIni 		set CreateIniFile=***
		if defined CreateIniFile (
			>OfficeRTool.ini 2>&1 echo. && (
				%SingleNul% del /q OfficeRTool.ini
				>>OfficeRTool.ini echo --------------------------------
				>>OfficeRTool.ini echo :: default download-path
				>>OfficeRTool.ini echo not set
				>>OfficeRTool.ini echo --------------------------------
				>>OfficeRTool.ini echo :: default download-language
				>>OfficeRTool.ini echo not set
				>>OfficeRTool.ini echo --------------------------------
				>>OfficeRTool.ini echo :: default download-architecture
				>>OfficeRTool.ini echo not set
				>>OfficeRTool.ini echo --------------------------------
			)
		)
	)
	
::===============================================================================================================
::===============================================================================================================

:Office16VnextInstall

	set "DloadLP="
	set "DloadImg="
	set "createIso="
	set "OnlineInstaller="
	set "downpath=not set"
	set "checknewVersion="
	set "o16updlocid=not set"
	set "o16arch=not set"
    set "o16lang=en-US"
	set "langtext=Default Language"
    set "o16lcid=1033"
	
	cd /D "%OfficeRToolpath%"
	SET /a countx=0
	if exist OfficeRTool.ini (
		for /F "tokens=*" %%a in (OfficeRTool.ini) do (
			SET /a countx=!countx! + 1
			set var!countx!=%%a
		)
		if !countx! GEQ 10 call :UpdateLangFromIni
	)
	
	call :CleanRegistryKeys
	%MultiNul% del /q latest*.txt
	%MultiNul% reg add "%XSPP_USER%" /f /v KeyManagementServiceName /t REG_SZ /d "0.0.0.0"
	%MultiNul% reg add "%XSPP_HKLM_X32%" /f /v KeyManagementServiceName /t REG_SZ /d "0.0.0.0"
	%MultiNul% reg add "%XSPP_HKLM_X64%" /f /v KeyManagementServiceName /t REG_SZ /d "0.0.0.0"
	
	cls
	
    echo:
	call :PrintTitle "================== OFFICE DOWNLOAD AND INSTALL ============================="
		
	echo:
	call :Print "[H] SCRUB OFFICE" "%BB_Blue%"
	echo:
	call :Print "[R] RESET - REPAIR OFFICE" "%BB_Blue%"
	echo:
	call :Print "[K] START KMS ACTIVATION" "%BB_Yellow%"
	echo:
	call :Print "[A] SHOW CURRENT ACTIVATION STATUS" "%BB_Yellow%"
	echo:
	call :Print "[C] CONVERT RETAIL LICENSE TO VOLUME LICENSE" "%BB_Yellow%"
	echo:
	call :Print "[N] INSTALL OFFICE FROM ONLINE INSTALL PACKAGE" "%BB_Green%"
	echo:
	call :Print "[O] CREATE OFFICE ONLINE WEB-INSTALLER PACKAGE SETUP FILE" "%BB_Green%"
	echo:
	call :Print "[L] CREATE OFFICE ONLINE WEB-INSTALLER LANGUAGE PACK SETUP FILE" "%BB_Green%"
	echo:
	call :Print "[M] DOWNLOAD OFFICE OFFLINE INSTALL IMAGE" "%BB_Red%"
	echo:
	call :Print "[D] DOWNLOAD OFFICE OFFLINE INSTALL PACKAGE" "%BB_Red%"
	echo:
	call :Print "[I] INSTALL OFFICE FROM OFFLINE INSTALL PACKAGE-IMAGE" "%BB_Red%"
	echo:
	call :Print "[S] CREATE ISO IMAGE FROM OFFLINE INSTALL PACKAGE-IMAGE" "%BB_Red%"
	echo:
	call :Print "[F] CHECK FOR NEW VERSION" "%BB_Blue%"
	echo:
	call :Print "[V] ENABLE VISUAL UI [WITH LTSC LOGO]" "%BB_Blue%"
	echo:
	call :Print "[X] ENABLE VISUAL UI [WITH 365  LOGO]" "%BB_Blue%"
	echo:
	call :Print "[T] DISABLE ACQUISITION AND SENDING OF TELEMETRY DATA" "%BB_Blue%"
	echo:
	call :Print "[U] CHANGE OFFICE UPDATE-PATH (SWITCH DISTRIBUTION CHANNEL)" "%BB_Blue%"
	echo:
	call :Print "[E] END - STOP AND LEAVE PROGRAM" "%BB_Magenta%"
	echo:
	
	if defined debugMode (echo 00Y | choice)
    CHOICE /C DSICKAUTOREHVXNFML /N /M "YOUR CHOICE ?"
    if %errorlevel%==1 goto:DownloadO16Offline
	if %errorlevel%==2 set "createIso=defined"&goto:InstallO16
	if %errorlevel%==3 goto:InstallO16
    if %errorlevel%==4 goto:Convert16Activate
	if %errorlevel%==5 goto:KMSActivation_ACT_WARPER
	if %errorlevel%==6 goto:CheckActivationStatus
    if %errorlevel%==7 goto:ChangeUpdPath
	if %errorlevel%==8 goto:DisableTelemetry
	if %errorlevel%==9 goto:DownloadO16Online
	if %errorlevel%==10 goto:ResetRepair
	if %errorlevel%==11 goto:TheEndIsNear
	if %errorlevel%==12 goto:Scrub
	if %errorlevel%==13 (set logo=LTSC&goto:EnableVisualUI)
	if %errorlevel%==14 (set logo=365&goto:EnableVisualUI)
	if %errorlevel%==15 (set "OnlineInstaller=defined"&goto:InstallO16)
	if %errorlevel%==16 goto:CheckPlease
	if %errorlevel%==17 (set "DloadImg=defined"&goto:DownloadO16Online)
	if %errorlevel%==18 (set "DloadLP=defined"&goto:DownloadO16Online)
	goto:Office16VnextInstall
	
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ 
 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
 _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

:GenerateIMGLink
	if "%of16install%" EQU "1" (
		echo $OfficeDownloadURL='https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!o16lang!/ProPlusRetail.img'
		echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_2016_PROPLUS_Retail.ISO'
	)
	
	if "%of19install%" EQU "1" (
		echo $OfficeDownloadURL='https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!o16lang!/ProPlus2019Retail.img'
		echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_2019_PROPLUS_Retail.ISO'
	)
	
	if "%of21install%" EQU "1" (
		echo $OfficeDownloadURL='https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!o16lang!/ProPlus2021Retail.img'
		echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_2021_PROPLUS_Retail.ISO'
	)
	
	
	if "%pr16install%" EQU "1" (
		echo $OfficeDownloadURL='https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!o16lang!/ProjectProRetail.img'
		echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_2016_PROJECT_PRO_Retail.ISO'
	)
	
	if "%pr19install%" EQU "1" (
		echo $OfficeDownloadURL='https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!o16lang!/ProjectPro2019Retail.img'
		echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_2019_PROJECT_PRO_Retail.ISO'
	)
	
	if "%pr21install%" EQU "1" (
		echo $OfficeDownloadURL='https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!o16lang!/ProjectPro2021Retail.img'
		echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_2021_PROJECT_PRO_Retail.ISO'
	)
	
	if "%vi16install%" EQU "1" (
		echo $OfficeDownloadURL='https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!o16lang!/VISIOProRetail.img'
		echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_2016_VISIO_PRO_Retail.ISO'
	)
	
	if "%vi19install%" EQU "1" (
		echo $OfficeDownloadURL='https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!o16lang!/VISIOPro2019Retail.img'
		echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_2019_VISIO_PRO_Retail.ISO'
	)
	
	if "%vi21install%" EQU "1" (
		echo $OfficeDownloadURL='https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/!o16lang!/VISIOPro2021Retail.img'
		echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_2021_VISIO_PRO_Retail.ISO'
	)
	
	echo if (Test-Path($OfficeDownloadFile)){
	echo 	Remove-Item $OfficeDownloadFile
	echo }
	echo Write-Host
	echo Write-Host 'Download !o16lang! !WebProduct! Offline image file'
	echo try {
	echo 	Start-BitsTransfer -Source $OfficeDownloadURL -Destination $OfficeDownloadFile -ea 1
	echo }
	echo catch {}
	echo if (-not(Test-Path($OfficeDownloadFile))){
	echo 	try {
	echo 		(new-object Net.WebClient).DownloadFile($OfficeDownloadURL, $OfficeDownloadFile)
	echo 	}
	echo 	catch {
	echo 		Start $OfficeDownloadURL
	echo 	}
	echo }
	echo if (Test-Path($OfficeDownloadFile)){
	echo 	If ((Get-Item $OfficeDownloadFile).length -gt 1024kb) {
	echo 		Write-Host
	echo 		Write-Host 'image File generated on your desktop.'
	echo 	}
	echo 	Else {
	echo 		Write-Host
	echo 		Write-Host 'ERROR ### Check your office configuration'
	echo 		Remove-Item $OfficeDownloadFile
	echo 	}
	echo }
goto :eof

:GenerateSetupLink
	echo $OfficeDownloadURL='https://c2rsetup.officeapps.live.com/c2r/download.aspx^?ProductreleaseID=!WebProduct!^&language=!o16lang!^&platform=!o16arch!'
	echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_!WebProduct!_!o16arch!_online_installer.exe'
	echo if (Test-Path($OfficeDownloadFile)){
	echo 	Remove-Item $OfficeDownloadFile
	echo }
	echo Write-Host
	echo Write-Host 'Download !o16lang! !WebProduct! !o16arch! online installer Setup file'
	echo try {
	echo 	Start-BitsTransfer -Source $OfficeDownloadURL -Destination $OfficeDownloadFile -ea 1
	echo }
	echo catch {}
	echo if (-not(Test-Path($OfficeDownloadFile))){
	echo 	try {
	echo 		(new-object Net.WebClient).DownloadFile($OfficeDownloadURL, $OfficeDownloadFile)
	echo 	}
	echo 	catch {
	echo 		Start $OfficeDownloadURL
	echo 	}
	echo }
	echo if (Test-Path($OfficeDownloadFile)){
	echo 	If ((Get-Item $OfficeDownloadFile).length -gt 1024kb) {
	echo 		Write-Host
	echo 		Write-Host 'Setup File generated on your desktop.'
	echo 	}
	echo 	Else {
	echo 		Write-Host
	echo 		Write-Host 'ERROR ### Check your office configuration'
	echo 		Remove-Item $OfficeDownloadFile
	echo 	}
	echo }
goto :eof

:GenerateLPLink
	echo $OfficeDownloadURL='https://c2rsetup.officeapps.live.com/c2r/download.aspx^?ProductreleaseID=languagepack^&language=!o16lang!^&platform=!o16arch!^&source=O16LAP^&version=O16GA'
	echo $OfficeDownloadFile=$env:USERPROFILE+'\Desktop\!o16lang!_languagepack_!o16arch!_online_installer.exe'
	echo if (Test-Path($OfficeDownloadFile)){
	echo 	Remove-Item $OfficeDownloadFile
	echo }
	echo Write-Host
	echo Write-Host 'Download !o16lang! !o16arch! online LP Setup file'
	echo try {
	echo 	Start-BitsTransfer -Source $OfficeDownloadURL -Destination $OfficeDownloadFile -ea 1
	echo }
	echo catch {}
	echo if (-not(Test-Path($OfficeDownloadFile))){
	echo 	try {
	echo 		(new-object Net.WebClient).DownloadFile($OfficeDownloadURL, $OfficeDownloadFile)
	echo 	}
	echo 	catch {
	echo 		Start $OfficeDownloadURL
	echo 	}
	echo }
	echo if (Test-Path($OfficeDownloadFile)){
	echo 	If ((Get-Item $OfficeDownloadFile).length -gt 1024kb) {
	echo 		Write-Host
	echo 		Write-Host 'Setup File generated on your desktop.'
	echo 	}
	echo 	Else {
	echo 		Write-Host
	echo 		Write-Host 'ERROR ### Check your office configuration'
	echo 		Remove-Item $OfficeDownloadFile
	echo 	}
	echo }
goto :eof

:ChoiceLangSelect
	echo:
	echo ### Language selection ###
	echo:
    echo -^> Afrikaans, Albanian, Amharic, Arabic, Armenian, Assamese, Azerbaijani Latin
	echo -^> Bangla Bangladesh, Bangla Bengali India, Basque Basque, Belarusian, Bosnian, Bulgarian
	echo -^> Catalan, Catalan Valencia, Chinese Simplified, Chinese Traditional, Croatian, Czech
	echo -^> Danish, Dari, Dutch - English, English UK, Estonian
	echo -^> Filipino, Finnish, French, French Canada
	echo -^> Galician, Georgian, German, Greek, Gujarati
	echo -^> Hausa Nigeria, Hebrew, Hindi, Hungarian
	echo -^> Icelandic, Igbo, Indonesian, Irish, Italian, IsiXhosa, IsiZulu - Japanese
	echo -^> Kannada, Kazakh, Khmer, KiSwahili, Konkani, Korean, Kyrgyz
	echo -^> Latvian, Lithuanian, Luxembourgish
	echo -^> Macedonian, Malay Latin, Malayalam, Maltese, Maori, Marathi, Mongolian
	echo -^> Nepali, Norwedian Nynorsk, Norwegian Bokmal - Odia
	echo -^> Pashto, Persian, Polish, Portuguese Portugal, Portuguese Brazilian, Punjabi Gurmukhi
	echo -^> Quechua - Romanian, Romansh, Russian
	echo -^> Scottish Gaelic, Serbian, Serbian Bosnia, Serbian Serbia, Sindhi Arabic, Sinhala, Slovak, Slovenian,
	echo    Spanish, Spanish Mexico, Swedish, Sesotho sa Leboa, Setswana
	echo -^> Tamil, Tatar Cyrillic, Telugu, Thai, Turkish, Turkmen
	echo -^> Ukrainian, Urdu, Uyghur, Uzbek - Vietnamese - Welsh, Wolof - Yoruba
	goto :eof

:Print
	if defined ANSI_COLORS (
		call :PrintANSI %1 %2
	) else (
		call :PrintAncient %1 "0%~2"
	)
	goto :eof
	
:PrintANSI
	echo %<%%~2%~1%>%
	goto :eof
	
:PrintAncient
	%<%:%~2 %1%>%
	goto :eof
	
:PrintVersionInfo
	set "VALUE=%<%%FF_Black:m=;%%BB_White%"
	set "PROPERTY=%<%%FF_Blue:m=;%%BB_WHITE:m=;%%Bold%"
	
	if defined ANSI_COLORS (
		echo %PROPERTY% %~1 %>%%VALUE% %~2 %>%%PROPERTY% %~3 %>%%VALUE% %~4 %>%%PROPERTY% %~5 %>%%VALUE% %~6 %>%
	) else (
		%<%:9f " %~1 "%>>% & %<%:8f " %~2 "%>>% & %<%:9f " %~3 "%>>% & %<%:8f " %~4 "%>>% & %<%:9f " %~5 "%>>% & %<%:8f " %~6 "%>%
	)
	goto :eof

:PrintTitle
	if defined ANSI_COLORS 		call :PrintANSI 	%* "%FF_Magenta:m=;%%B_Yellow:m=;%%Bold%"
	if not defined ANSI_COLORS 	call :PrintAncient	%* "5E"
	goto :eof
	
:CheckPlease
	cls
	echo:
	echo *** Checking public Office distribution channels for new updates
	echo:
	echo:
	set "checknewVersion=defined"
	call :CheckNewVersion Current 492350f6-3a01-4f97-b9c0-c7c6ddf67d60
	call :CheckNewVersion CurrentPreview 64256afe-f5d9-4f86-8936-8840a6a4f5be
	call :CheckNewVersion BetaChannel 5440fd1f-7ecb-4221-8110-145efaa6372f
	call :CheckNewVersion MonthlyEnterprise 55336b82-a18d-4dd6-b5f6-9e5095c314a6	
	call :CheckNewVersion SemiAnnual 7ffbc6bf-bc32-4f92-8982-f9dd17fd3114
	call :CheckNewVersion SemiAnnualPreview b8f9b850-328d-4355-9145-c59439a0c4cf
	call :CheckNewVersion PerpetualVL2019 f2e724c1-748f-4b47-8fb8-8e0d210e9208
	call :CheckNewVersion PerpetualVL2021 5030841d-c919-4594-8d2d-84ae4f96e58e
	call :CheckNewVersion DogfoodDevMain ea4a4090-de26-49d7-93c1-91bff9e53fc3
	((echo:)&&(echo:)&&(echo:)&&(pause))
	goto:Office16VnextInstall
	
:KMSActivation_ACT_WARPER
	cls
	echo:
	call :PrintTitle "================== ACTIVATE OFFICE PRODUCTS ===================="
	echo.
	call :startKMSACTIVATION
	call :CleanRegistryKeys
	call :UpdateRegistryKeys %KMSHostIP% %KMSPort%
	call :CheckOfficeApplications
	
	if "%_ProPlusRetail%" EQU "YES" ((echo Activating Office Professional Plus 2016)&&(call :Office16Activate d450596f-894d-49e0-966a-fd39ed4c4c64))
	if "%_ProPlusVolume%" EQU "YES" ((echo Activating Office Professional Plus 2016)&&(call :Office16Activate d450596f-894d-49e0-966a-fd39ed4c4c64))
	if "%_ProPlus2019Retail%" EQU "YES" ((echo Activating Office Professional Plus 2019)&&(call :Office16Activate 85dd8b5f-eaa4-4af3-a628-cce9e77c9a03))
	if "%_ProPlus2019Volume%" EQU "YES" ((echo Activating Office Professional Plus 2019)&&(call :Office16Activate 85dd8b5f-eaa4-4af3-a628-cce9e77c9a03))
	if "%_ProPlus2021Retail%" EQU "YES" ((echo Activating Office Professional Plus 2021)&&(call :Office16Activate fbdb3e18-a8ef-4fb3-9183-dffd60bd0984))
	if "%_ProPlus2021Volume%" EQU "YES" ((echo Activating Office Professional Plus 2021)&&(call :Office16Activate fbdb3e18-a8ef-4fb3-9183-dffd60bd0984))
	if "%_ProPlusSPLA2021Volume%" EQU "YES" ((echo Activating Office Professional Plus 2021)&&(call :Office16Activate fbdb3e18-a8ef-4fb3-9183-dffd60bd0984))
	if "%_StandardRetail%" EQU "YES" ((echo Activating Office Standard 2016)&&(call :Office16Activate dedfa23d-6ed1-45a6-85dc-63cae0546de6))
	if "%_StandardVolume%" EQU "YES" ((echo Activating Office Standard 2016)&&(call :Office16Activate dedfa23d-6ed1-45a6-85dc-63cae0546de6))
	if "%_Standard2019Retail%" EQU "YES" ((echo Activating Office Standard 2019)&&(call :Office16Activate 6912a74b-a5fb-401a-bfdb-2e3ab46f4b02))
	if "%_Standard2019Volume%" EQU "YES" ((echo Activating Office Standard 2019)&&(call :Office16Activate 6912a74b-a5fb-401a-bfdb-2e3ab46f4b02))
	if "%_Standard2021Retail%" EQU "YES" ((echo Activating Office Standard 2021)&&(call :Office16Activate 080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3))
	if "%_Standard2021Volume%" EQU "YES" ((echo Activating Office Standard 2021)&&(call :Office16Activate 080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3))
	if "%_StandardSPLA2021Volume%" EQU "YES" ((echo Activating Office Standard 2021)&&(call :Office16Activate 080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3))
	if "%_O365ProPlusRetail%" EQU "YES" ((echo Activating Microsoft 365 Apps for Enterprise)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_O365BusinessRetail%" EQU "YES" ((echo Activating Microsoft 365 Apps for Business)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_O365HomePremRetail%" EQU "YES" ((echo Activating Microsoft 365 Home Premium retail)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_O365SmallBusPremRetail%" EQU "YES" ((echo Activating Microsoft 365 Small Business retail)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_ProfessionalRetail%" EQU "YES" ((echo Activating Professional Retail)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_Professional2019Retail%" EQU "YES" ((echo Activating Professional 2019 Retail)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_Professional2021Retail%" EQU "YES" ((echo Activating Professional 2021 Retail)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_HomeBusinessRetail%" EQU "YES" ((echo Activating Microsoft Home And Business )&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_HomeBusiness2019Retail%" EQU "YES" ((echo Activating Microsoft Home And Business 2019 )&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_HomeBusiness2021Retail%" EQU "YES" ((echo Activating Microsoft Home And Business 2021 )&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_HomeStudentRetail%" EQU "YES" ((echo Activating Microsoft Home And Student )&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_HomeStudent2019Retail%" EQU "YES" ((echo Activating Microsoft Home And Student 2019 )&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_HomeStudent2021Retail%" EQU "YES" ((echo Activating Microsoft Home And Student 2021 )&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_MondoRetail%" EQU "YES" ((echo Activating Office Mondo Grande Suite)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_MondoVolume%" EQU "YES" ((echo Activating Office Mondo Grande Suite)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_PersonalRetail%" EQU "YES" ((echo Activating Office Personal 2016 Retail)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_Personal2019Retail%" EQU "YES" ((echo Activating Office Personal 2019 Retail)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_Personal2021Retail%" EQU "YES" ((echo Activating Office Personal 2021 Retail)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_WordRetail%" EQU "YES" ((echo Activating Word 2016 SingleApp)&&(call :Office16Activate bb11badf-d8aa-470e-9311-20eaf80fe5cc))
	if "%_ExcelRetail%" EQU "YES" ((echo Activating Excel 2016 SingleApp)&&(call :Office16Activate c3e65d36-141f-4d2f-a303-a842ee756a29))
	if "%_PowerPointRetail%" EQU "YES" ((echo Activating PowerPoint 2016 SingleApp)&&(call :Office16Activate d70b1bba-b893-4544-96e2-b7a318091c33))
	if "%_AccessRetail%" EQU "YES" ((echo Activating Access 2016 SingleApp)&&(call :Office16Activate 67c0fc0c-deba-401b-bf8b-9c8ad8395804))
	if "%_OutlookRetail%" EQU "YES" ((echo Activating Outlook 2016 SingleApp)&&(call :Office16Activate ec9d9265-9d1e-4ed0-838a-cdc20f2551a1))
	if "%_PublisherRetail%" EQU "YES" ((echo Activating Publisher 2016 Single App)&&(call :Office16Activate 041a06cb-c5b8-4772-809f-416d03d16654))
	if "%_OneNoteRetail%" EQU "YES" ((echo Activating OneNote 2016 SingleApp)&&(call :Office16Activate d8cace59-33d2-4ac7-9b1b-9b72339c51c8))
	if "%_OneNoteVolume%" EQU "YES" ((echo Activating OneNote 2016 SingleApp)&&(call :Office16Activate d8cace59-33d2-4ac7-9b1b-9b72339c51c8))
	if "%_OneNote2021Retail%" EQU "YES" ((echo Activating OneNote 2021 SingleApp)&&(call :Office16Activate d8cace59-33d2-4ac7-9b1b-9b72339c51c8))
	if "%_SkypeForBusinessRetail%" EQU "YES" ((echo Activating Skype For Business 2016 SingleApp)&&(call :Office16Activate 83e04ee1-fa8d-436d-8994-d31a862cab77))
	if "%_Word2019Retail%" EQU "YES" ((echo Activating Word 2019 SingleApp)&&(call :Office16Activate 059834fe-a8ea-4bff-b67b-4d006b5447d3))
	if "%_Excel2019Retail%" EQU "YES" ((echo Activating Excel 2019 SingleApp)&&(call :Office16Activate 237854e9-79fc-4497-a0c1-a70969691c6b))
	if "%_PowerPoint2019Retail%" EQU "YES" ((echo Activating PowerPoint 2019 SingleApp)&&(call :Office16Activate 3131fd61-5e4f-4308-8d6d-62be1987c92c))
	if "%_Access2019Retail%" EQU "YES" ((echo Activating Access 2019 SingleApp)&&(call :Office16Activate 9e9bceeb-e736-4f26-88de-763f87dcc485))
	if "%_Outlook2019Retail%" EQU "YES" ((echo Activating Outlook 2019 SingleApp)&&(call :Office16Activate c8f8a301-19f5-4132-96ce-2de9d4adbd33))
	if "%_Publisher2019Retail%" EQU "YES" ((echo Activating Publisher 2019 Single App)&&(call :Office16Activate 9d3e4cca-e172-46f1-a2f4-1d2107051444))
	if "%_SkypeForBusiness2019Retail%" EQU "YES" ((echo Activating Skype For Business 2019 SingleApp)&&(call :Office16Activate 734c6c6e-b0ba-4298-a891-671772b2bd1b))
	if "%_Word2019Volume%" EQU "YES" ((echo Activating Word 2019 SingleApp)&&(call :Office16Activate 059834fe-a8ea-4bff-b67b-4d006b5447d3))
	if "%_Excel2019Volume%" EQU "YES" ((echo Activating Excel 2019 SingleApp)&&(call :Office16Activate 237854e9-79fc-4497-a0c1-a70969691c6b))
	if "%_PowerPoint2019Volume%" EQU "YES" ((echo Activating PowerPoint 2019 SingleApp)&&(call :Office16Activate 3131fd61-5e4f-4308-8d6d-62be1987c92c))
	if "%_Access2019Volume%" EQU "YES" ((echo Activating Access 2019 SingleApp)&&(call :Office16Activate 9e9bceeb-e736-4f26-88de-763f87dcc485))
	if "%_Outlook2019Volume%" EQU "YES" ((echo Activating Outlook 2019 SingleApp)&&(call :Office16Activate c8f8a301-19f5-4132-96ce-2de9d4adbd33))
	if "%_Publisher2019Volume%" EQU "YES" ((echo Activating Publisher 2019 Single App)&&(call :Office16Activate 9d3e4cca-e172-46f1-a2f4-1d2107051444))
	if "%_SkypeForBusiness2019Volume%" EQU "YES" ((echo Activating Skype For Business 2019 SingleApp)&&(call :Office16Activate 734c6c6e-b0ba-4298-a891-671772b2bd1b))
	if "%_Word2021Retail%" EQU "YES" ((echo Activating Word 2021 SingleApp)&&(call :Office16Activate abe28aea-625a-43b1-8e30-225eb8fbd9e5))
	if "%_Excel2021Retail%" EQU "YES" ((echo Activating Excel 2021 SingleApp)&&(call :Office16Activate ea71effc-69f1-4925-9991-2f5e319bbc24))
	if "%_PowerPoint2021Retail%" EQU "YES" ((echo Activating PowerPoint 2021 SingleApp)&&(call :Office16Activate 6e166cc3-495d-438a-89e7-d7c9e6fd4dea))
	if "%_Access2021Retail%" EQU "YES" ((echo Activating Access 2021 SingleApp)&&(call :Office16Activate 1fe429d8-3fa7-4a39-b6f0-03dded42fe14))
	if "%_Outlook2021Retail%" EQU "YES" ((echo Activating Outlook 2021 SingleApp)&&(call :Office16Activate a5799e4c-f83c-4c6e-9516-dfe9b696150b))
	if "%_Publisher2021Retail%" EQU "YES" ((echo Activating Publisher 2021 Single App)&&(call :Office16Activate aa66521f-2370-4ad8-a2bb-c095e3e4338f))
	if "%_SkypeForBusiness2021Retail%" EQU "YES" ((echo Activating Skype For Business 2021 SingleApp)&&(call :Office16Activate SkypeForBusiness2021))
	if "%_Word2021Volume%" EQU "YES" ((echo Activating Word 2021 SingleApp)&&(call :Office16Activate abe28aea-625a-43b1-8e30-225eb8fbd9e5))
	if "%_Excel2021Volume%" EQU "YES" ((echo Activating Excel 2021 SingleApp)&&(call :Office16Activate ea71effc-69f1-4925-9991-2f5e319bbc24))
	if "%_PowerPoint2021Volume%" EQU "YES" ((echo Activating PowerPoint 2021 SingleApp)&&(call :Office16Activate 6e166cc3-495d-438a-89e7-d7c9e6fd4dea))
	if "%_Access2021Volume%" EQU "YES" ((echo Activating Access 2021 SingleApp)&&(call :Office16Activate 1fe429d8-3fa7-4a39-b6f0-03dded42fe14))
	if "%_Outlook2021Volume%" EQU "YES" ((echo Activating Outlook 2021 SingleApp)&&(call :Office16Activate a5799e4c-f83c-4c6e-9516-dfe9b696150b))
	if "%_Publisher2021Volume%" EQU "YES" ((echo Activating Publisher 2021 Single App)&&(call :Office16Activate aa66521f-2370-4ad8-a2bb-c095e3e4338f))
	if "%_SkypeForBusiness2021Volume%" EQU "YES" ((echo Activating Skype For Business 2021 SingleApp)&&(call :Office16Activate 1f32a9af-1274-48bd-ba1e-1ab7508a23e8))
	if "%_VisioProRetail%" EQU "YES" ((echo Activating Visio Professional 2016)&&(call :Office16Activate 6bf301c1-b94a-43e9-ba31-d494598c47fb))
	if "%_AppxVisio%" EQU "YES" ((echo Activating Visio Professional UWP Appx)&&(call :Office16Activate 6bf301c1-b94a-43e9-ba31-d494598c47fb))
	if "%_ProjectProRetail%" EQU "YES" ((echo Activating Project Professional 2016)&&(call :Office16Activate 4f414197-0fc2-4c01-b68a-86cbb9ac254c))
	if "%_AppxProject%" EQU "YES" ((echo Activating Project Professional UWP Appx)&&(call :Office16Activate 4f414197-0fc2-4c01-b68a-86cbb9ac254c))
	if "%_VisioPro2019Retail%" EQU "YES" ((echo Activating Visio Professional 2019)&&(call :Office16Activate 5b5cf08f-b81a-431d-b080-3450d8620565))
	if "%_VisioPro2019Volume%" EQU "YES" ((echo Activating Visio Professional 2019)&&(call :Office16Activate 5b5cf08f-b81a-431d-b080-3450d8620565))
	if "%_ProjectPro2019Retail%" EQU "YES" ((echo Activating Project Professional 2019)&&(call :Office16Activate 2ca2bf3f-949e-446a-82c7-e25a15ec78c4))
	if "%_ProjectPro2019Volume%" EQU "YES" ((echo Activating Project Professional 2019)&&(call :Office16Activate 2ca2bf3f-949e-446a-82c7-e25a15ec78c4))
	if "%_VisioPro2021Retail%" EQU "YES" ((echo Activating Visio Professional 2021)&&(call :Office16Activate fb61ac9a-1688-45d2-8f6b-0674dbffa33c))
	if "%_VisioPro2021Volume%" EQU "YES" ((echo Activating Visio Professional 2021)&&(call :Office16Activate fb61ac9a-1688-45d2-8f6b-0674dbffa33c))
	if "%_ProjectPro2021Retail%" EQU "YES" ((echo Activating Project Professional 2021)&&(call :Office16Activate 76881159-155c-43e0-9db7-2d70a9a3a4ca))
	if "%_ProjectPro2021Volume%" EQU "YES" ((echo Activating Project Professional 2021)&&(call :Office16Activate 76881159-155c-43e0-9db7-2d70a9a3a4ca))
	if "%_UWPappINSTALLED%" EQU "YES" ((echo Activating Office UUP Apps)&&(call :Office16Activate 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce))
	if "%_VisioStdRetail%" EQU "YES" ((echo Activating Visio Standard 2016)&&(call :Office16Activate aa2a7821-1827-4c2c-8f1d-4513a34dda97))
	if "%_VisioStdVolume%" EQU "YES" ((echo Activating Visio Standard 2016)&&(call :Office16Activate aa2a7821-1827-4c2c-8f1d-4513a34dda97))
	if "%_VisioStdXVolume%" EQU "YES" ((echo Activating Visio Standard 2016 C2R)&&(call :Office16Activate 361fe620-64f4-41b5-ba77-84f8e079b1f7))
	if "%_VisioStd2019Retail%" EQU "YES" ((echo Activating Visio Standard 2019)&&(call :Office16Activate e06d7df3-aad0-419d-8dfb-0ac37e2bdf39))
	if "%_VisioStd2019Volume%" EQU "YES" ((echo Activating Visio Standard 2019)&&(call :Office16Activate e06d7df3-aad0-419d-8dfb-0ac37e2bdf39))
	if "%_VisioStd2021Retail%" EQU "YES" ((echo Activating Visio Standard 2021)&&(call :Office16Activate 72fce797-1884-48dd-a860-b2f6a5efd3ca))
	if "%_VisioStd2021Volume%" EQU "YES" ((echo Activating Visio Standard 2021)&&(call :Office16Activate 72fce797-1884-48dd-a860-b2f6a5efd3ca))
	if "%_ProjectStdRetail%" EQU "YES" ((echo Activating Project Standard 2016)&&(call :Office16Activate da7ddabc-3fbe-4447-9e01-6ab7440b4cd4))
	if "%_ProjectStdVolume%" EQU "YES" ((echo Activating Project Standard 2016)&&(call :Office16Activate da7ddabc-3fbe-4447-9e01-6ab7440b4cd4))
	if "%_ProjectStdXVolume%" EQU "YES" ((echo Activating Project Standard 2016 C2R)&&(call :Office16Activate cbbaca45-556a-4416-ad03-bda598eaa7c8))
	if "%_VisioProXVolume%" EQU "YES" ((echo Activating Visio Professional 2016 C2R)&&(call :Office16Activate 829b8110-0e6f-4349-bca4-42803577788d))
	if "%_ProjectProXVolume%" EQU "YES" ((echo Activating Project Professional 2016 C2R)&&(call :Office16Activate b234abe3-0857-4f9c-b05a-4dc314f85557))
	if "%_ProjectStd2019Retail%" EQU "YES" ((echo Activating Project Standard 2019)&&(call :Office16Activate 1777f0e3-7392-4198-97ea-8ae4de6f6381))
	if "%_ProjectStd2019Volume%" EQU "YES" ((echo Activating Project Standard 2019)&&(call :Office16Activate 1777f0e3-7392-4198-97ea-8ae4de6f6381))
	if "%_ProjectStd2021Retail%" EQU "YES" ((echo Activating Project Standard 2021)&&(call :Office16Activate 6dd72704-f752-4b71-94c7-11cec6bfc355))
	if "%_ProjectStd2021Volume%" EQU "YES" ((echo Activating Project Standard 2021)&&(call :Office16Activate 6dd72704-f752-4b71-94c7-11cec6bfc355))
	
	call :STOPKMSActivation
	timeout /t 4
	goto:Office16VnextInstall
	
::===============================================================================================================
::===============================================================================================================
:EnableVisualUI
	cls
	echo:
	call :PrintTitle "================== ENABLE VISUAL UI ===================="
	set "root="
	if exist "%ProgramFiles%\Microsoft Office\root" set "root=%ProgramFiles%\Microsoft Office\root"
	if exist "%ProgramFiles(x86)%\Microsoft Office\root" set "root=%ProgramFiles(x86)%\Microsoft Office\root"
	
	if not defined root (
		echo.
		echo Error ### Fail to find integrator.exe Tool
		echo.
		if not defined debugMode pause
		goto:Office16VnextInstall
	) else (
		if not exist "!root!\Integration\integrator.exe" (
			echo.
			echo Error ### Fail to find integrator.exe Tool
			echo.
			if not defined debugMode pause
			goto:Office16VnextInstall
		)
	)

	echo !logo! |%SingleNul% find /i "LTSC" && (
		echo.
		echo -- Integrate Professional 2021 Retail License
		"!root!\Integration\integrator" /I /License PRIDName=Professional2021Retail.16 PidKey=G7R2D-6NQ7C-CX62B-9YR9J-DGRYH
	)
	
	echo !logo! |%SingleNul% find /i "365" && (
		echo.
		echo -- Integrate Mondo 2016 Retail License
		"!root!\Integration\integrator" /I /License PRIDName=MondoRetail.16 PidKey=2N6B3-BXW6B-W2XBT-VVQ64-7H7DH
	)
	
	echo !logo! |%SingleNul% find /i "MONDO" && (
		echo.
		echo -- Integrate Mondo 2016 Volume License
		"!root!\Integration\integrator" /I /License PRIDName=MondoVolume.16 PidKey=HFTND-W9MK4-8B7MJ-B6C4G-XQBR2
	)

	echo -- Clean Registry Keys
	for /f "tokens=3,4,5,6,7,8,9,10 delims=-" %%A in ('whoami /user ^| find /i "S-1-5"') do (
		%MultiNul% reg delete "HKEY_USERS\S-%%A-%%B-%%C-%%D-%%E-%%F-%%G\SOFTWARE\Microsoft\Office" /f
		%MultiNul% reg delete "HKEY_USERS\S-%%A-%%B-%%C-%%D-%%E-%%F-%%G\SOFTWARE\Wow6432Node\Microsoft\Office" /f
		%MultiNul% reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides" /f
	)

	echo -- Install Visual UI Registry Keys
	call :reg_own "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\CurrentSkuIdAggregationForApp" "" S-1-5-32-544 "" Allow SetValue
	for %%# in (Word, Excel, Powerpoint, Access, Outlook, Publisher, OneNote, project, visio) do (
	  %MultiNul% reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\%%#" /f /v "Microsoft.Office.UXPlatform.FluentSVRefresh" /t REG_SZ /d "true"
	  %MultiNul% reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\%%#" /f /v "Microsoft.Office.UXPlatform.RibbonTouchOptimization" /t REG_SZ /d "true"
	  %MultiNul% reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\%%#" /f /v "Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu" /t REG_SZ /d "true"
	  %MultiNul% reg add "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\CurrentSkuIdAggregationForApp" /f /v "%%#" /t REG_SZ /d "{FBDB3E18-A8EF-4FB3-9183-DFFD60BD0984},{CE5FFCAF-75DA-4362-A9CB-00D2689918AA},"
	)
	call :reg_own "HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\CurrentSkuIdAggregationForApp" "" S-1-5-32-544 "" Deny SetValue

	echo -- Done.
	echo.

	echo Note:
	echo To initiate the Visual Refresh,
	echo it may be required to start some Office apps
	echo a couple of times.
	echo.
	echo Many Thanks to Xtreme21, Krakatoa, rioachim
	echo for helping make and debug this script
	echo.
	if not defined debugMode pause
	
	goto:Office16VnextInstall

:reg_own #key [optional] all user owner access permission  :        call :reg_own "HKCU\My" "" S-1-5-32-544 "" Allow FullControl
	powershell -nop -c $A='%~1','%~2','%~3','%~4','%~5','%~6';iex(([io.file]::ReadAllText('%~f0')-split':Own1\:.*')[1])&exit/b:Own1:
	$D1=[uri].module.gettype('System.Diagnostics.Process')."GetM`ethods"(42) |where {$_.Name -eq 'SetPrivilege'} #`:no-ev-warn
	'SeSecurityPrivilege','SeTakeOwnershipPrivilege','SeBackupPrivilege','SeRestorePrivilege'|foreach {$D1.Invoke($null, @("$_",2))}
	$path=$A[0]; $rk=$path-split'\\',2; $HK=gi -lit Registry::$($rk[0]) -fo; $s=$A[1]; $sps=[Security.Principal.SecurityIdentifier]
	$u=($A[2],'S-1-5-32-544')[!$A[2]];$o=($A[3],$u)[!$A[3]];$w=$u,$o |% {new-object $sps($_)}; $old=!$A[3];$own=!$old; $y=$s-eq'all'
	$rar=new-object Security.AccessControl.RegistryAccessRule( $w[0], ($A[5],'FullControl')[!$A[5]], 1, 0, ($A[4],'Allow')[!$A[4]] )
	$x=$s-eq'none';function Own1($k){$t=$HK.OpenSubKey($k,2,'TakeOwnership');if($t){0,4|%{try{$o=$t.GetAccessControl($_)}catch{$old=0}
	};if($old){$own=1;$w[1]=$o.GetOwner($sps)};$o.SetOwner($w[0]);$t.SetAccessControl($o); $c=$HK.OpenSubKey($k,2,'ChangePermissions')
	$p=$c.GetAccessControl(2);if($y){$p.SetAccessRuleProtection(1,1)};$p.ResetAccessRule($rar);if($x){$p.RemoveAccessRuleAll($rar)}
	$c.SetAccessControl($p);if($own){$o.SetOwner($w[1]);$t.SetAccessControl($o)};if($s){$subkeys=$HK.OpenSubKey($k).GetSubKeyNames()
	foreach($n in $subkeys){Own1 "$k\$n"}}}};Own1 $rk[1];if($env:VO){get-acl Registry::$path|fl} #:Own1: lean & mean snippet by AveYo

:SCRUB
	cls
	echo.
	call :PrintTitle "================ Remove Office 11 - 12 - 14 - 15 - 16 - C2R ================"
	echo ____________________________________________________________________________
	echo.
	echo.
	set "xzzyz5="
	set /p xzzyz5=Press Enter to continue, Any key to back to MAIN MENU ^>
	if defined xzzyz5 goto:Office16VnextInstall
	
	cls
	echo.
	call :PrintTitle "================ Remove Office 11 - 12 - 14 - 15 - 16 - C2R ================"
	echo ____________________________________________________________________________
	echo.
	echo.
	
	set result="%temp%\result"
	if exist "%windir%\SysWOW64\cscript.exe" set cscript="%windir%\SysWOW64\cscript.exe"
	if exist "%windir%\system32\cscript.exe" set cscript="%windir%\system32\cscript.exe"
	
	echo Process :: Clean Keys ^& Licences
	%MultiNul% "%OfficeRToolpath%\OfficeFixes\win_x32\cleanospp.exe"
	!cscript! "%OfficeRToolpath%\OfficeFixes\vbs\OLicenseCleanup.vbs" //nologo //b /QUIET
	echo ....................................................................................................
	
	echo Process :: Clean Registry ^& Folders
	>!result! 2>&1 dir "%ProgramFiles%\Microsoft Office*" /ad /b && (
		for /f "tokens=*" %%# in ('type !result!') do %MultiNul% call :DestryFolder "%ProgramFiles%\%%#"
	)
	
	>!result! 2>&1 dir "%ProgramFiles(x86)%\Microsoft Office*" /ad /b && (
		for /f "tokens=*" %%# in ('type !result!') do %MultiNul% call :DestryFolder "%ProgramFiles(x86)%\%%#"
	)
	
	%MultiNul% del /q !result!
	for /f "tokens=3,4,5,6,7,8,9,10 delims=-" %%A in ('whoami /user ^| find /i "S-1-5"') do (set "GUID=S-%%A-%%B-%%C-%%D-%%E-%%F-%%G")
	for %%$ in (HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER,HKEY_USERS) do (
		echo "%%$" |%SingleNul% find /i "HKEY_USERS" && (
			%SingleNulV2% reg query "%%$\!GUID!\SOFTWARE\Microsoft" /f office 			  | >>!result! find /i "%%$"
			%SingleNulV2% reg query "%%$\!GUID!\SOFTWARE\Wow6432Node\Microsoft" /f office | >>!result! find /i "%%$"
		) || (
			%SingleNulV2% reg query "%%$\SOFTWARE\microsoft" /f office 					  | >>!result! find /i "%%$"
			%SingleNulV2% reg query "%%$\SOFTWARE\WOW6432Node\microsoft" /f office 		  | >>!result! find /i "%%$"
		)
	)
	if exist !result! (
		for /f "tokens=*" %%$ in ('type !result!') do (
			%MultiNul% reg delete "%%$" /f
		)
		%MultiNul% del /q !result!
	)
	echo ....................................................................................................

	echo Process :: OffScrubC2R.vbs
	!cscript! "%OfficeRToolpath%\OfficeFixes\vbs\OffScrubC2R.vbs" //nologo //b ALL /NoCancel /Force /OSE /Quiet /NoReboot /Passive
	echo ....................................................................................................

	for %%G in (OffScrub_O16msi.vbs,OffScrub_O15msi.vbs,OffScrub10.vbs,OffScrub07.vbs,OffScrub03.vbs) do (
		echo Process :: %%G
		!cscript! "%OfficeRToolpath%\OfficeFixes\vbs\%%G" //nologo //b ALL /NoCancel /Force /OSE /Quiet /NoReboot /Passive
		echo.
	)
	
	echo Process :: Clean Leftover
	>!result! 2>&1 dir "%ProgramFiles%\Microsoft Office*" /ad /b && (
		for /f "tokens=*" %%# in ('type !result!') do %MultiNul% call :DestryFolder "%ProgramFiles%\%%#"
	)
	
	>!result! 2>&1 dir "%ProgramFiles(x86)%\Microsoft Office*" /ad /b && (
		for /f "tokens=*" %%# in ('type !result!') do %MultiNul% call :DestryFolder "%ProgramFiles(x86)%\%%#"
	)
	
	%MultiNul% del /q !result!
	for /f "tokens=3,4,5,6,7,8,9,10 delims=-" %%A in ('whoami /user ^| find /i "S-1-5"') do (set "GUID=S-%%A-%%B-%%C-%%D-%%E-%%F-%%G")
	for %%$ in (HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER,HKEY_USERS) do (
		echo "%%$" |%SingleNul% find /i "HKEY_USERS" && (
			%SingleNulV2% reg query "%%$\!GUID!\SOFTWARE\Microsoft" /f office 			  | >>!result! find /i "%%$"
			%SingleNulV2% reg query "%%$\!GUID!\SOFTWARE\Wow6432Node\Microsoft" /f office | >>!result! find /i "%%$"
		) || (
			%SingleNulV2% reg query "%%$\SOFTWARE\microsoft" /f office 					  | >>!result! find /i "%%$"
			%SingleNulV2% reg query "%%$\SOFTWARE\WOW6432Node\microsoft" /f office 		  | >>!result! find /i "%%$"
		)
	)
	if exist !result! (
		for /f "tokens=*" %%$ in ('type !result!') do (
			%MultiNul% reg delete "%%$" /f
		)
		%MultiNul% del /q !result!
	)

	echo.
	if not defined debugMode pause
	goto :Office16VnextInstall

:DestryFolder
	%MultiNul% rd/s/q "%temp%"
	%MultiNul% md "%temp%"
	set targetFolder=%*
	if exist %targetFolder% (
		rd /s /q %targetFolder%
		if exist %targetFolder% (
			for /f "tokens=*" %%g in ('dir /b/s /a-d %targetFolder%') do move /y "%%g" "%temp%"
			rd /s /q %targetFolder%
		)
	)
	goto :eof

:CheckNewVersion
	set "o16build=not set "
	set "o16latestbuild=not set"
	if /i '%1' EQU 'CURRENT' 			set "ADDIN=          "
	if /i '%1' EQU 'CurrentPreview' 	set "ADDIN=   "
	if /i '%1' EQU 'BetaChannel' 		set "ADDIN=      "
	if /i '%1' EQU 'MonthlyEnterprise' 	set "ADDIN="
	if /i '%1' EQU 'SemiAnnual' 		set "ADDIN=       "
	if /i '%1' EQU 'SemiAnnualPreview' 	set "ADDIN="
	if /i '%1' EQU 'PerpetualVL2019' 	set "ADDIN=  "
	if /i '%1' EQU 'PerpetualVL2021' 	set "ADDIN=  "
	if /i '%1' EQU 'DogfoodDevMain' 	set "ADDIN=   "
	if "%1" EQU "Manual_Override" goto:CheckNewVersionSkip1
	if exist "%OfficeRToolpath%\latest_%1_build.txt" (
		set /p "o16build=" <"%OfficeRToolpath%\latest_%1_build.txt"
		)
	set "o16build=%o16build:~0,-1%"
:CheckNewVersionSkip1
	if %win% GEQ 9200 "%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --no-verbose --output-document="%TEMP%\VersionDescriptor.txt" --tries=20 "https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData/?audienceFFN=%2" %MultiNul%
	if %win% LSS 9200 "%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --no-verbose --output-document="%TEMP%\VersionDescriptor.txt" --tries=20 "https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData/?audienceFFN=%2&osver=Client|6.1.0" %MultiNul%
	if %errorlevel% GEQ 1 goto:ErrCheckNewVersion1
	type "%TEMP%\VersionDescriptor.txt" | find "AvailableBuild" >"%TEMP%\found_office_build.txt"
	if %errorlevel% GEQ 1 goto:ErrCheckNewVersion2
	set /p "o16latestbuild=" <"%TEMP%\found_office_build.txt"
	set "o16latestbuild=%o16latestbuild:~21,16%"
	set "spaces=     "
	if "%o16latestbuild:~15,1%" EQU " " ((set "o16latestbuild=%o16latestbuild:~0,14%")&&(set "spaces=       "))
	if "%o16latestbuild:~0,3%" NEQ "16." goto:ErrCheckNewVersion3
	if "%1" EQU "Manual_Override" goto:CheckNewVersionSkip2a
	if "!o16build!" NEQ "!o16latestbuild!" (
		if defined checknewVersion Call :PrintVersionInfo "CHANNEL ::" "%1%ADDIN%" "ID ::" "%2" "VERSION ::" "%o16latestbuild%"
		echo !o16latestbuild! >"%OfficeRToolpath%\latest_%1_build.txt"
		echo !o16build! >>"%OfficeRToolpath%\latest_%1_build.txt"
		goto:CheckNewVersionSkip2b
	)
:CheckNewVersionSkip2a
if defined checknewVersion powershell -noprofile -command "%pswindowtitle%"; Write-Host "'   'Last known good Build:' '" -foreground "White" -nonewline; Write-Host "%o16latestbuild%'%spaces%'" -foreground "Green" -nonewline; Write-Host "No newer Build available" -foreground "White"
:CheckNewVersionSkip2b
	del /f /q "%TEMP%\VersionDescriptor.txt"
	del /f /q "%TEMP%\found_office_build.txt"
	set "buildcheck=ok"
	goto:eof
:ErrCheckNewVersion1
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** ERROR checking: * %1 * channel" -foreground "Red"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** No response from Office content delivery server" -foreground "Red"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** Check Internet connection and/or Channel-ID" -foreground "White"
	set "buildcheck=not ok"
	goto:eof
:ErrCheckNewVersion2
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** ERROR checking: * %1 * " -foreground "Red"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** No Build / Version number found in file: * VersionDescriptor.txt" -foreground "Red"
	copy "%TEMP%\VersionDescriptor.txt" "%TEMP%\%1_VersionDescriptor.txt" %MultiNul%
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** Check file "%TEMP%\%1_VersionDescriptor.txt" -foreground "White"
	set "buildcheck=not ok"
	goto:eof
:ErrCheckNewVersion3
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** ERROR checking: * %1 * " -foreground "Red"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** Unsupported Build / Version number detected: * !o16latestbuild! *" -foreground "Red"
	copy "%TEMP%\VersionDescriptor.txt" "%TEMP%\%1_VersionDescriptor.txt" %MultiNul%
	copy "%TEMP%\found_office_build.txt" "%TEMP%\%1_found_office_build.txt" %MultiNul%
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** Check file "%TEMP%\%1_VersionDescriptor.txt" -foreground "White"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** Check file "%TEMP%\%1_found_office_build.txt" -foreground "White"
	set "buildcheck=not ok"
	goto:eof
::===============================================================================================================
::===============================================================================================================

:DownloadO16Offline
    cd /D "%OfficeRToolpath%"
    set "installtrigger=0"
	set "channeltrigger=0"
	set "o16updlocid=not set"
	set "o16build=not set"
	cls
	echo:
	call :PrintTitle "=========================== Selected Configuration ========================="
	echo:
	echo DownloadPath: "!inidownpath!"
    echo:
	if "!o16updlocid!" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" echo Channel-ID:    !o16updlocid! (Current) && goto:DownOfflineContinue
	if "!o16updlocid!" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" echo Channel-ID:    !o16updlocid! (CurrentPreview) && goto:DownOfflineContinue
	if "!o16updlocid!" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" echo Channel-ID:    !o16updlocid! (BetaChannel) && goto:DownOfflineContinue
	if "!o16updlocid!" EQU "55336b82-a18d-4dd6-b5f6-9e5095c314a6" echo Channel-ID:    !o16updlocid! (MonthlyEnterprise) && goto:DownOfflineContinue
	if "!o16updlocid!" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" echo Channel-ID:    !o16updlocid! (SemiAnnual) && goto:DownOfflineContinue
	if "!o16updlocid!" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" echo Channel-ID:    !o16updlocid! (SemiAnnualPreview) && goto:DownOfflineContinue
	if "!o16updlocid!" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" echo Channel-ID:    !o16updlocid! (PerpetualVL2019) && goto:DownOfflineContinue
	if "!o16updlocid!" EQU "f2e724c1-748f-4b47-8fb8-8e0d210e9208" echo Channel-ID:    !o16updlocid! (PerpetualVL2021) && goto:DownOfflineContinue
	if "!o16updlocid!" EQU "5030841d-c919-4594-8d2d-84ae4f96e58e" echo Channel-ID:    !o16updlocid! (P) && goto:DownOfflineContinue
	if "!o16updlocid!" EQU "not set" echo Channel-ID:    not set && goto:DownOfflineContinue
	echo Channel-ID:    !o16updlocid! (Manual_Override)
::===============================================================================================================
:DownOfflineContinue
	echo:
	echo Office build:  !o16build!
	echo:
	echo Language:      !o16lang! (!langtext!)
    echo:
	echo Architecture:  !o16arch!
    echo ____________________________________________________________________________
	echo:
	echo Set new Office Package download path or press return for
	
	echo "!downpath!" | %SingleNul% find /i "not set" && (
		set "downpath=%USERPROFILE%\desktop"
	)
	set /p downpath=Set Office Package Download Path ^= "!downpath!" ^>
	set "downpath=!downpath:"=!"
	if /i "!downpath!" EQU "X" (set "downpath=not set")&&(goto:Office16VnextInstall)
	set "downdrive=!downpath:~0,2!"
	if "!downdrive:~-1!" NEQ ":" (echo:)&&(echo Unknown Drive "!downdrive!" - Drive not found)&&(echo Enter correct driveletter:\directory or enter "X" to exit)&&(echo:)&&(pause)&&(set "downpath=not set")&&(goto:DownloadO16Offline)
	cd /d !downdrive!\ %MultiNul%
	if errorlevel 1 (echo:)&&(echo Unknown Drive "!downdrive!" - Drive not found)&&(echo Enter correct driveletter:\directory or enter "X" to exit)&&(echo:)&&(pause)&&(set "downpath=not set")&&(goto:DownloadO16Offline)
	set "downdrive=!downpath:~0,3!"
	if "!downdrive:~-1!" EQU "\" (set "downpath=!downdrive!!downpath:~3!") else (set "downpath=!downdrive:~0,2!\!downpath:~2!")
	if "!downpath:~-1!" EQU "\" set "downpath=!downpath:~0,-1!"
::===============================================================================================================
	cd /D "%OfficeRToolpath%"
	set "installtrigger=0"
	echo:
	if "%inidownpath%" NEQ "!downpath!" ((echo Office install package download path changed)&&(echo old path "%inidownpath%" -- new path "!downpath!")&&(echo:))
	if defined AutoSaveToIni goto :_xv5
	if defined DontSaveToIni goto :SkipDownPathSave
	if "%inidownpath%" NEQ "!downpath!" set /p installtrigger=Save new path to OfficeRTool.ini? (1/0) ^>
	if "%installtrigger%" EQU "0" goto:SkipDownPathSave
	if /I "%installtrigger%" EQU "X" goto:SkipDownPathSave
	:_xv5
	set "inidownpath=!downpath!"
	echo -------------------------------->OfficeRTool.ini
	echo ^:^: default download-path>>OfficeRTool.ini
	echo %inidownpath%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-language>>OfficeRTool.ini
	echo %inidownlang%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-architecture>>OfficeRTool.ini
	echo %inidownarch%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo Download path saved.
::===============================================================================================================
:SkipDownPathSave
	echo:
	echo "Public known" standard distribution channels
	echo Channel Name                                    - Internal Naming   Index-#
	echo ___________________________________________________________________________
	echo:
	echo Current (Retail/RTM)                            - (Production::CC)      (1)
	echo CurrentPreview (Office Insider SLOW)            - (Insiders::CC)        (2)
	echo BetaChannel (Office Insider FAST)               - (Insiders::DEVMAIN)   (3)
	echo MonthlyEnterprise                               - (Production::MEC)     (4)
	echo SemiAnnual (Business)                           - (Production::DC)      (5)
	echo SemiAnnualPreview (Business Insider)            - (Insiders::FRDC)      (6)
	echo PerpetualVL2019                                 - (Production:LTSC)     (7)
	echo PerpetualVL2021                                 - (Production:LTSC2021) (8)
	echo Manual_Override (set identifier for Channel-ID's not public known)      (M)
	echo Exit to Main Menu                                                       (X)
	echo:
	set /p channeltrigger=Set Channel-Index-# (1,2,3,4,5,6,M) or X or press return for Current ^>
	if "%channeltrigger%" EQU "1" goto:ChanSel1
	if "%channeltrigger%" EQU "2" goto:ChanSel2
	if "%channeltrigger%" EQU "3" goto:ChanSel3
	if "%channeltrigger%" EQU "4" goto:ChanSel4
	if "%channeltrigger%" EQU "5" goto:ChanSel5
	if "%channeltrigger%" EQU "6" goto:ChanSel6
	if "%channeltrigger%" EQU "7" goto:ChanSel7
	if "%channeltrigger%" EQU "8" goto:ChanSel8
	if /I "%channeltrigger%" EQU "M" ((set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:ChanSelMan))
	if /I "%channeltrigger%" EQU "X" ((set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall))
	goto:ChanSel1
::===============================================================================================================
:ChanSel1
	set "o16updlocid=492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
	call :CheckNewVersion Current !o16updlocid!
	set "o16build=!o16latestbuild!"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel2
	set "o16updlocid=64256afe-f5d9-4f86-8936-8840a6a4f5be"
	call :CheckNewVersion CurrentPreview !o16updlocid!
	set "o16build=!o16latestbuild!"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel3
	set "o16updlocid=5440fd1f-7ecb-4221-8110-145efaa6372f"
	call :CheckNewVersion BetaChannel !o16updlocid!
	set "o16build=!o16latestbuild!"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel4
	set "o16updlocid=55336b82-a18d-4dd6-b5f6-9e5095c314a6"
	call :CheckNewVersion MonthlyEnterprise !o16updlocid!
	set "o16build=!o16latestbuild!"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel5
	set "o16updlocid=7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
	call :CheckNewVersion SemiAnnual !o16updlocid!
	set "o16build=!o16latestbuild!"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel6
	set "o16updlocid=b8f9b850-328d-4355-9145-c59439a0c4cf"
	call :CheckNewVersion SemiAnnualPreview !o16updlocid!
	set "o16build=!o16latestbuild!"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel7
	set "o16updlocid=f2e724c1-748f-4b47-8fb8-8e0d210e9208"
	call :CheckNewVersion PerpetualVL2019 !o16updlocid!
	set "o16build=!o16latestbuild!"
	goto:ChannelSelected
::===============================================================================================================
:ChanSel8
	set "o16updlocid=5030841d-c919-4594-8d2d-84ae4f96e58e"
	call :CheckNewVersion PerpetualVL2021 !o16updlocid!
	set "o16build=!o16latestbuild!"
	goto:ChannelSelected
::===============================================================================================================
:ChanSelMan
    echo:
	echo "Microsoft Internal Use Only" Beta/Testing distribution channels
	echo Internal Naming           Channel-ID:                               Index-#
	echo ___________________________________________________________________________
	echo:
    echo Dogfood::DevMain     ---^> ea4a4090-de26-49d7-93c1-91bff9e53fc3         (1)
    echo Dogfood::CC          ---^> f3260cf1-a92c-4c75-b02e-d64c0a86a968         (2)
    echo Dogfood::DCEXT       ---^> c4a7726f-06ea-48e2-a13a-9d78849eb706         (3)
    echo Dogfood::FRDC        ---^> 834504cc-dc55-4c6d-9e71-e024d0253f6d         (4)
    echo Microsoft::CC        ---^> 5462eee5-1e97-495b-9370-853cd873bb07         (5)
    echo Microsoft::DC        ---^> f4f024c8-d611-4748-a7e0-02b6e754c0fe         (6)
    echo Microsoft::DevMain   ---^> b61285dd-d9f7-41f2-9757-8f61cba4e9c8         (7)
    echo Microsoft::FRDC      ---^> 9a3b7ff2-58ed-40fd-add5-1e5158059d1c         (8)
	echo Microsoft::LTSC2021  ---^> 86752282-5841-4120-ac80-db03ae6b5fdb         (9) 
 	echo Insiders::LTSC       ---^> 2e148de9-61c8-4051-b103-4af54baffbb4         (A)
	echo Insiders::LTSC2021   ---^> 12f4f6ad-fdea-4d2a-a90f-17496cc19a48         (B)
	echo Insiders::MEC        ---^> 0002c1ba-b76b-4af9-b1ee-ae2ad587371f         (C)
	echo Exit to Main Menu                                                      (X)
    echo:
	set /p o16updlocid=Set Channel (enter Channel-ID or Index-#) ^>
	if "!o16updlocid!" EQU "not set" goto:DownloadO16Offline
	if /I "!o16updlocid!" EQU "X" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)
	if "!o16updlocid!" EQU "0" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)
	if "!o16updlocid!" EQU "1" set "o16updlocid=ea4a4090-de26-49d7-93c1-91bff9e53fc3"
	if "!o16updlocid!" EQU "2" set "o16updlocid=f3260cf1-a92c-4c75-b02e-d64c0a86a968"
    if "!o16updlocid!" EQU "3" set "o16updlocid=c4a7726f-06ea-48e2-a13a-9d78849eb706"
    if "!o16updlocid!" EQU "4" set "o16updlocid=834504cc-dc55-4c6d-9e71-e024d0253f6d
    if "!o16updlocid!" EQU "5" set "o16updlocid=5462eee5-1e97-495b-9370-853cd873bb07"
    if "!o16updlocid!" EQU "6" set "o16updlocid=f4f024c8-d611-4748-a7e0-02b6e754c0fe"
    if "!o16updlocid!" EQU "7" set "o16updlocid=b61285dd-d9f7-41f2-9757-8f61cba4e9c8"
    if "!o16updlocid!" EQU "8" set "o16updlocid=2e148de9-61c8-4051-b103-4af54baffbb4"
    if "!o16updlocid!" EQU "9" set "o16updlocid=86752282-5841-4120-ac80-db03ae6b5fdb"
	if /I "!o16updlocid!" EQU "A" set "o16updlocid=2e148de9-61c8-4051-b103-4af54baffbb4"
	if /I "!o16updlocid!" EQU "B" set "o16updlocid=12f4f6ad-fdea-4d2a-a90f-17496cc19a48"
	if /I "!o16updlocid!" EQU "C" set "o16updlocid=0002c1ba-b76b-4af9-b1ee-ae2ad587371f"
	echo Channel-ID:   !o16updlocid! (Manual_Override) && PAUSE
	call :CheckNewVersion Manual_Override !o16updlocid!
	set "o16build=!o16latestbuild!"
::===============================================================================================================
:ChannelSelected
	set "o16downloadloc=officecdn.microsoft.com.edgesuite.net/pr/!o16updlocid!/Office/Data"
	echo:
	if "%buildcheck%" EQU "not ok" ((pause)&&(set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall))
	set "o16buildCKS=!o16build!"
    set /p o16build=Set Office Build - or press return for !o16build! ^>
	echo "!o16build!" | >nul findstr /r "16.[0-9].[0-9][0-9][0-9][0-9][0-9].[0-9][0-9][0-9][0-9][0-9]" || (set "o16build=!o16buildCKS!" & goto :ChannelSelected)
	if "!o16build!" EQU "not set" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)
	if /I "!o16build!" EQU "X" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)
::===============================================================================================================
:LangSelect
	call :ChoiceLangSelect
	:LangSelect_
	echo:
	if /i "!o16lang!" EQU "not set" call :CheckSystemLanguage
    set /p o16lang=Set Language Value - or press return for !o16lang! ^>
	set "o16lang=!o16lang:, =!"
	set "o16lang=!o16lang:,=!"
	if defined o16lang if /i "x!o16lang:~0,1!" EQU "x " set "o16lang=!o16lang:~1!"
	if defined o16lang if /i "!o16lang:-1!x" EQU " x" set "o16lang=!o16lang:~0,-1!"
	call :SetO16Language
	if defined langnotfound (
		set "o16lang=not set"
		goto:LangSelect_
	)
::===============================================================================================================
	cd /D "%OfficeRToolpath%"
	set "installtrigger=0"
	if "%inidownlang%" NEQ "!o16lang!" ((echo:)&&(echo Office install package download language changed)&&(echo old language "%inidownlang%" -- new language "!o16lang!")&&(echo:))
	if defined AutoSaveToIni goto :_x35
	if defined DontSaveToIni goto :ArchSelect
	if "%inidownlang%" NEQ "!o16lang!" set /p installtrigger=Save new language to OfficeRTool.ini? (1/0) ^>
	if "%installtrigger%" EQU "0" goto:ArchSelect
	if /I "%installtrigger%" EQU "X" goto:ArchSelect
	:_x35
	set "inidownlang=!o16lang!"
	echo -------------------------------->OfficeRTool.ini
	echo ^:^: default download-path>>OfficeRTool.ini
	echo %inidownpath%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-language>>OfficeRTool.ini
	echo %inidownlang%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-architecture>>OfficeRTool.ini
	echo %inidownarch%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo Download language saved.
	
::===============================================================================================================
:ArchSelect
	if /i '!o16arch!' EQU 'not set' (
		if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set sBit=86)
		if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set sBit=64)
		if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set sBit=64
		if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set sBit=64
		set "o16arch=x!sBit!"
		if defined inidownarch (
			echo !inidownarch! | %SingleNul% find /i "not set" && set "o16arch=x!sBit!" || set "o16arch=!inidownarch!"
		)
	)
	
	echo:
	set /p o16arch=Set architecture to download (x86 or x64 or Multi) - or press return for !o16arch! ^>
	if /i "!o16arch!" EQU "x86" goto:SkipArchSelect
	if /i "!o16arch!" EQU "x64" goto:SkipArchSelect
	if /i "!o16arch!" EQU "Multi" goto:SkipArchSelect
	set "o16arch=not set"
	goto:ArchSelect
::===============================================================================================================
:SkipArchSelect
	cd /D "%OfficeRToolpath%"
	set "installtrigger=0"
	echo:
	if "%inidownarch%" NEQ "!o16arch!" ((echo Office install package download architecture changed)&&(echo old architecture "%inidownarch%" -- new architecture "!o16arch!")&&(echo:))
	if defined AutoSaveToIni goto :_x35xf
	if defined DontSaveToIni goto :SkipDownArchSave
	if "%inidownarch%" NEQ "!o16arch!" set /p installtrigger=Save new architecture to OfficeRTool.ini? (1/0) ^>
	if "%installtrigger%" EQU "0" goto:SkipDownArchSave
	if /I "%installtrigger%" EQU "X" goto:SkipDownArchSave
	:_x35xf
	set "inidownarch=!o16arch!"
	echo -------------------------------->OfficeRTool.ini
	echo ^:^: default download-path>>OfficeRTool.ini
	echo %inidownpath%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-language>>OfficeRTool.ini
	echo %inidownlang%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo ^:^: default download-architecture>>OfficeRTool.ini
	echo %inidownarch%>>OfficeRTool.ini
	echo -------------------------------->>OfficeRTool.ini
	echo Download architecture saved.
::===============================================================================================================
:SkipDownArchSave
	set "multiMode="
	if /i "!o16arch!" EQU "Multi" (
		set multiMode=TRUE
		set "o16arch=x86"
		call :SkipDownArchSave_
		set "o16arch=x64"
		call :SkipDownArchSave_
		goto :Office16VnextInstall
	)

:SkipDownArchSave_
	cls
    echo:
	call :PrintTitle "========================= Pending Download (SUMMARY) ========================="
    echo:
    echo DownloadPath: !downpath!
    echo:
	if "!o16updlocid!" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" echo Channel-ID:   !o16updlocid! (Current) && goto:PendDownContinue
	if "!o16updlocid!" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" echo Channel-ID:   !o16updlocid! (CurrentPreview) && goto:PendDownContinue
	if "!o16updlocid!" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" echo Channel-ID:   !o16updlocid! (BetaChannel) && goto:PendDownContinue
	if "!o16updlocid!" EQU "55336b82-a18d-4dd6-b5f6-9e5095c314a6" echo Channel-ID:   !o16updlocid! (MonthlyEnterprise) && goto:PendDownContinue
	if "!o16updlocid!" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" echo Channel-ID:   !o16updlocid! (SemiAnnual) && goto:PendDownContinue
	if "!o16updlocid!" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" echo Channel-ID:   !o16updlocid! (SemiAnnualPreview) && goto:PendDownContinue
	if "!o16updlocid!" EQU "f2e724c1-748f-4b47-8fb8-8e0d210e9208" echo Channel-ID:   !o16updlocid! (PerpetualVL2019) && goto:PendDownContinue
	if "!o16updlocid!" EQU "5030841d-c919-4594-8d2d-84ae4f96e58e" echo Channel-ID:   !o16updlocid! (PerpetualVL2021) && goto:PendDownContinue
	if "!o16updlocid!" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" echo Channel-ID:   !o16updlocid! (DogfoodDevMain) && goto:PendDownContinue
	echo Channel-ID:   !o16updlocid! (Manual_Override)
::===============================================================================================================
:PendDownContinue
	set "installtrigger=0"
	echo Office Build: !o16build!
	echo Language:     !o16lang! (%langtext%)
    echo Architecture: !o16arch!
    echo ____________________________________________________________________________
	echo:
	if defined multiMode goto:PendDownContinue_
	set /p installtrigger=(Enter) to download, (R)estart download, (E)xit to main menu ^>
	if /i "%installtrigger%" EQU "R" (goto:DownloadO16Offline)
	if /i "%installtrigger%" EQU "E" (set "o16updlocid=not set")&&(set "o16build=not set")&&(goto:Office16VnextInstall)

:PendDownContinue_
::===============================================================================================================
::===============================================================================================================
:Office16VNextDownload
	cls
	echo:
	call :PrintTitle "================== DOWNLOADING OFFICE OFFLINE SETUP PACKAGE ================"
	echo:
	if "!o16updlocid!" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" set "downbranch=Current" && goto:ContVNextDownload
	if "!o16updlocid!" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" set "downbranch=CurrentPreview" && goto:ContVNextDownload
	if "!o16updlocid!" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" set "downbranch=BetaChannel" && goto:ContVNextDownload
	if "!o16updlocid!" EQU "55336b82-a18d-4dd6-b5f6-9e5095c314a6" set "downbranch=MonthlyEnterprise" && goto:ContVNextDownload
	if "!o16updlocid!" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" set "downbranch=SemiAnnual" && goto:ContVNextDownload
	if "!o16updlocid!" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" set "downbranch=SemiAnnualPreview" && goto:ContVNextDownload
	if "!o16updlocid!" EQU "f2e724c1-748f-4b47-8fb8-8e0d210e9208" set "downbranch=PerpetualVL2019" && goto:ContVNextDownload
	if "!o16updlocid!" EQU "5030841d-c919-4594-8d2d-84ae4f96e58e" set "downbranch=PerpetualVL2021" && goto:ContVNextDownload
	if "!o16updlocid!" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" set "downbranch=DogfoodDevMain" && goto:ContVNextDownload
	set "downbranch=Manual_Override"
::===============================================================================================================
:ContVNextDownload
	cd /d "%downdrive%\" %MultiNul%
	md "!downpath!" %MultiNul%
	cd /d "!downpath!" %MultiNul%
	set "directory-prefix=!o16lang!_Office_!downbranch!_!o16arch!_v!o16build!"
	if defined multiMode set "directory-prefix=!o16lang!_Office_!downbranch!_x86_x64_v!o16build!"
	if "!o16arch!" EQU "x64" goto:X64DOWNLOAD
::===============================================================================================================
::	Download x86/32bit Office setup files
	
	echo.
	echo Download File :: v32.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/v32.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/v32.cab"
	
	echo.
	echo Download File :: v32_!o16build!.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/v32_!o16build!.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/v32_!o16build!.cab"
	
	echo.
	echo Download File :: !o16build!/i320.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/i320.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/i320.cab"
	
	echo.
	echo Download File :: !o16build!/i32%o16lcid%.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/i32%o16lcid%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/i32%o16lcid%.cab"
	
	echo.
	echo Download File :: !o16build!/s320.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/s320.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/s320.cab"
	
	echo.
	echo Download File :: !o16build!/s32%o16lcid%.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/s32%o16lcid%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/s32%o16lcid%.cab"
	
	echo.
	echo Download File :: !o16build!/sp32%o16lcid%.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/sp32%o16lcid%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/sp32%o16lcid%.cab"
	
	echo.
	echo Download File :: !o16build!/stream.x86.!o16lang!.dat
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/stream.x86.!o16lang!.dat
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/stream.x86.!o16lang!.dat"
	
	echo.
	echo Download File :: !o16build!/stream.x86.x-none.dat
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/stream.x86.x-none.dat
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/stream.x86.x-none.dat"
	
	echo.
	echo Download File :: !o16build!/stream.x86.!o16lang!.dat.cat
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/stream.x86.!o16lang!.dat.cat
	
	echo.
	echo Download File :: !o16build!/stream.x86.x-none.dat.cat
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/stream.x86.x-none.dat.cat
	goto:GENERALDOWNLOAD

::===============================================================================================================
::	Download x64/64bit Office setup files
:X64DOWNLOAD


	echo.
	echo Download File :: v64.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/v64.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/v64.cab"
	
	echo.
	echo Download File :: v64_!o16build!.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/v64_!o16build!.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/v64_!o16build!.cab"
	
	echo.
	echo Download File :: !o16build!/s640.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/s640.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/s640.cab"
	
	echo.
	echo Download File :: !o16build!/s64%o16lcid%.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/s64%o16lcid%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/s64%o16lcid%.cab"
	
	echo.
	echo Download File :: !o16build!/sp64%o16lcid%.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/sp64%o16lcid%.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/sp64%o16lcid%.cab"
	
	echo.
	echo Download File :: !o16build!/stream.x64.!o16lang!.dat
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/stream.x64.!o16lang!.dat
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/stream.x64.!o16lang!.dat"
	
	echo.
	echo Download File :: !o16build!/stream.x64.x-none.dat
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/stream.x64.x-none.dat
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/stream.x64.x-none.dat"
	
	echo.
	echo Download File :: !o16build!/stream.x64.!o16lang!.dat.cat
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/stream.x64.!o16lang!.dat.cat
	
	echo.
	echo Download File :: !o16build!/stream.x64.x-none.dat.cat
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/stream.x64.x-none.dat.cat
	
::===============================================================================================================	
:: Download setup file(s) used in both x86 and x64 architectures
:GENERALDOWNLOAD
	
	echo.
	echo Download File :: !o16build!/i640.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/i640.cab
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/i640.cab"
	
	echo.
	echo Download File :: !o16build!/i64%o16lcid%.cab
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --show-progress --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/i64%o16lcid%.cab	
	if %errorlevel% GEQ 1 call :WgetError "http://%o16downloadloc%/!o16build!/i64%o16lcid%.cab"
	
	echo.
	echo Download File :: !o16build!/i640.cab.cat
	"%OfficeRToolpath%\OfficeFixes\%winx%\wget.exe" --quiet --retry-connrefused --continue --tries=20 --force-directories --no-host-directories --cut-dirs=2 --directory-prefix=%directory-prefix% http://%o16downloadloc%/!o16build!/i640.cab.cat
::===============================================================================================================	
	echo ____________________________________________________________________________
	if "%downbranch%" EQU "Current" echo Current>%directory-prefix%\package.info
	if "%downbranch%" EQU "CurrentPreview" echo CurrentPreview>%directory-prefix%\package.info
	if "%downbranch%" EQU "BetaChannel" echo BetaChannel>%directory-prefix%\package.info
	if "%downbranch%" EQU "MonthlyEnterprise" echo MonthlyEnterprise>%directory-prefix%\package.info
	if "%downbranch%" EQU "SemiAnnual" echo SemiAnnual>%directory-prefix%\package.info
	if "%downbranch%" EQU "SemiAnnualPreview" echo SemiAnnualPreview>%directory-prefix%\package.info
	if "%downbranch%" EQU "PerpetualVL2019" echo PerpetualVL2019>%directory-prefix%\package.info
	if "%downbranch%" EQU "PerpetualVL2021" echo PerpetualVL2021>%directory-prefix%\package.info
	if "%downbranch%" EQU "DogfoodDevMain" echo DogfoodDevMain>%directory-prefix%\package.info
	if "%downbranch%" EQU "Manual_Override" echo Manual_Override>%directory-prefix%\package.info
	echo !o16build!>>%directory-prefix%\package.info
	echo !o16lang!>>%directory-prefix%\package.info
	if defined multiMode (
		echo Multi>>%directory-prefix%\package.info
	) else (
		echo !o16arch!>>%directory-prefix%\package.info
	)
	echo !o16updlocid!>>%directory-prefix%\package.info
	echo:
	echo:
	
	if defined multiMode (
		goto :eof
	) else (
		timeout /t 4
		goto :Office16VnextInstall
	)
::===============================================================================================================
::===============================================================================================================
:WgetError
	set "errortrigger=0"
	powershell -noprofile -command "%pswindowtitle%"; Write-Host "*** ERROR downloading: %1" -foreground "Red"
	echo:
	set /p errortrigger=Cancel Download now? (1/0) ^>
	if "%errortrigger%" EQU "1" (
		if exist "!downpath!\%directory-prefix%" rd "!downpath!\%directory-prefix%" /S /Q
		goto:Office16VnextInstall
	)
	echo:
	goto :eof
::===============================================================================================================
::===============================================================================================================
:DownloadO16Online
    cd /D "%OfficeRToolpath%"
    cls
	echo:
						set "tt=DOWNLOAD OFFICE ONLINE SETUP FILE"
	if defined DloadLP  set "tt=DOWNLOAD OFFICE ONLINE LP FILE"
	if defined DloadImg set "tt=DOWNLOAD OFFICE OFFLINE INSTALL IMAGE"
	call :PrintTitle "============= !tt! ==============="
	echo ____________________________________________________________________________
    
	set "txt="
	set "of16install=0"
    set "pr16install=0"
    set "vi16install=0"
    set "of19install=0"
    set "pr19install=0"
    set "vi19install=0"
	set "of21install=0"
    set "pr21install=0"
    set "vi21install=0"
	set "WebProduct=not set"
	set "installtrigger=O"
	if defined DloadLP goto :ArchSelectXYYY
	if not defined DloadImg if not defined DloadLP set "txt=setup.exe "
	echo:
	set /p installtrigger=Generate Office 2016 products !txt!download-link (1=YES/0=NO) ^>
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if "%installtrigger%" EQU "1" goto:WEBOFF2016
	echo:
	set /p installtrigger=Generate Office 2019 products !txt!download-link (1=YES/0=NO) ^>
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if "%installtrigger%" EQU "1" goto:WEBOFF2019
	echo:
	set /p installtrigger=Generate Office 2021 products !txt!download-link (1=YES/0=NO) ^>
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if "%installtrigger%" EQU "1" goto:WEBOFF2021
	goto:DownloadO16Online
:WEBOFF2016
	echo:
	echo ____________________________________________________________________________
	echo:
    set /p of16install=Set Office Professional Plus 2016 Install (1/0) ^>
	if "%of16install%" EQU "1" (set "WebProduct=ProPlusRetail")&&(goto:ArchSelectXYYY)
    echo:
    set /p pr16install=Set Project Professional 2016 Install (1/0) ^>
	if "%pr16install%" EQU "1" (set "WebProduct=ProjectProRetail")&&(goto:ArchSelectXYYY)
    echo:
    set /p vi16install=Set Visio Professional 2016 Install (1/0) ^>
	if "%vi16install%" EQU "1" (set "WebProduct=VisioProRetail")&&(goto:ArchSelectXYYY)
	goto:WEBOFFNOTHING
:WEBOFF2019
	echo:
	echo ____________________________________________________________________________
	echo:
    set /p of19install=Set Office Professional Plus 2019 Install (1/0) ^>
	if "%of19install%" EQU "1" (set "WebProduct=ProPlus2019Retail")&&(goto:ArchSelectXYYY)
    echo:
    set /p pr19install=Set Project Professional 2019 Install (1/0) ^>
	if "%pr19install%" EQU "1" (set "WebProduct=ProjectPro2019Retail")&&(goto:ArchSelectXYYY)
    echo:
    set /p vi19install=Set Visio Professional 2019 Install (1/0) ^>
	if "%vi19install%" EQU "1" (set "WebProduct=VisioPro2019Retail")&&(goto:ArchSelectXYYY)
	goto:WEBOFFNOTHING
:WEBOFF2021
	echo:
	echo ____________________________________________________________________________
	echo:
    set /p of21install=Set Office Professional Plus 2021 Install (1/0) ^>
	if "%of21install%" EQU "1" (set "WebProduct=ProPlus2021Retail")&&(goto:ArchSelectXYYY)
    echo:
    set /p pr21install=Set Project Professional 2021 Install (1/0) ^>
	if "%pr21install%" EQU "1" (set "WebProduct=ProjectPro2021Retail")&&(goto:ArchSelectXYYY)
    echo:
    set /p vi21install=Set Visio Professional 2021 Install (1/0) ^>
	if "%vi21install%" EQU "1" (set "WebProduct=VisioPro2021Retail")&&(goto:ArchSelectXYYY)
	goto:WEBOFFNOTHING
:WEBOFFNOTHING
	echo:
	echo ____________________________________________________________________________
	echo:
	echo Nothing selected - Returning to Main Menu now
	echo:
	if not defined debugMode pause
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================

:ArchSelectXYYY
	if defined DloadImg set "o16arch=Multi" & goto:WebLangSelect
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set sBit=86)
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set sBit=64)
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set sBit=64
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set sBit=64
	
	set "o16arch=x!sBit!"
	if defined inidownarch ((echo "!inidownarch!" | %SingleNul% find /i "not set") || set "o16arch=!inidownarch!")
	if /i '!o16arch!' EQU 'Multi' set "o16arch=x!sBit!"	
	if /i 'x!sBit!' NEQ '!o16arch!' (if /i '!sBit!' EQU '86' (set "o16arch=x!sBit!"))
	
	echo:
	set /p o16arch=Set architecture to download (x86 or x64) - or press return for !o16arch! ^>
	if /i "!o16arch!" EQU "x86" goto:WebLangSelect
	if /i "!o16arch!" EQU "x64" goto:WebLangSelect
	set "o16arch=not set"
	goto:ArchSelectXYYY

:WebLangSelect
	call :ChoiceLangSelect
:WebLangSelect_WOOB
	echo:
	if /i "!o16lang!" EQU "not set" call :CheckSystemLanguage
    set /p o16lang=Set Language Value - or press return for !o16lang! ^>
	set "o16lang=!o16lang:, =!"
	set "o16lang=!o16lang:,=!"
	if defined o16lang if /i "x!o16lang:~0,1!" EQU "x " set "o16lang=!o16lang:~1!"
	if defined o16lang if /i "!o16lang:-1!x" EQU " x" set "o16lang=!o16lang:~0,-1!"
	call :SetO16Language
	if defined langnotfound (
		set "o16lang=not set"
		goto:WebLangSelect_WOOB
	)
::===============================================================================================================
    echo:
    echo ____________________________________________________________________________
	echo:
						set "tt=Pending Online SETUP Downlad (SUMMARY)"
	if defined DloadLP  set "tt=Pending Online LP Downlad (SUMMARY)"
	if defined DloadImg set "tt=Pending Online IMAGE Downlad (SUMMARY)"
	call :PrintTitle "========================= !tt! ========================="
    echo:
	
    if "%of16install%" EQU "1" echo Download Office 2016 ?      : YES
    if "%pr16install%" EQU "1" echo Download Project 2016 ?     : YES
    if "%vi16install%" EQU "1" echo Download Visio 2016 ?       : YES
    if "%of19install%" EQU "1" echo Download Office 2019 ?      : YES
    if "%pr19install%" EQU "1" echo Download Project 2019 ?     : YES
    if "%vi19install%" EQU "1" echo Download Visio 2019 ?       : YES
	if "%of21install%" EQU "1" echo Download Office 2021 ?      : YES
    if "%pr21install%" EQU "1" echo Download Project 2021 ?     : YES
    if "%vi21install%" EQU "1" echo Download Visio 2021 ?       : YES
    echo:
    echo Install Architecture ?     : !o16arch!
    echo Install Language ?         : !o16lang!
    echo ____________________________________________________________________________
	echo:
    set /p installtrigger=(ENTER) to download, (R)estart download, (E)xit to main menu ? ^>
    if /i "%installtrigger%" EQU "R" goto:DownloadO16Online
	if /I "%installtrigger%" EQU "E" goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:OfficeWebInstall
    cls
	echo:
						set "tt=DOWNLOAD OFFICE ONLINE SETUP INSTALLER"
	if defined DloadImg set "tt=DOWNLOAD OFFICE OFFLINE IMAGE INSTALLER"
	if defined DloadLP  set "tt=DOWNLOAD OFFICE ONLINE LANGUAGE PACK INSTALLER"
	call :PrintTitle "========================= !tt! ========================="
	echo:
	if defined DloadImg (
	
		call :GenerateIMGLink > "%temp%\tmp.ps1"
		powershell -noprofile -executionpolicy bypass -file "%temp%\tmp.ps1"
		
		set "ISO="
		set "OfficeVer="
		set "Zz=%~dp0OfficeFixes\win_x32\7z.exe"
		
		if "%of16install%" EQU "1" set "ISO=%USERPROFILE%\Desktop\!o16lang!_2016_PROPLUS_Retail.ISO"
		if "%of19install%" EQU "1" set "ISO=%USERPROFILE%\Desktop\!o16lang!_2019_PROPLUS_Retail.ISO"
		if "%of21install%" EQU "1" set "ISO=%USERPROFILE%\Desktop\!o16lang!_2021_PROPLUS_Retail.ISO"
		if "%pr16install%" EQU "1" set "ISO=%USERPROFILE%\Desktop\!o16lang!_2016_PROJECT_PRO_Retail.ISO"
		if "%pr19install%" EQU "1" set "ISO=%USERPROFILE%\Desktop\!o16lang!_2019_PROJECT_PRO_Retail.ISO"
		if "%pr21install%" EQU "1" set "ISO=%USERPROFILE%\Desktop\!o16lang!_2021_PROJECT_PRO_Retail.ISO"
		if "%vi16install%" EQU "1" set "ISO=%USERPROFILE%\Desktop\!o16lang!_2016_VISIO_PRO_Retail.ISO"
		if "%vi19install%" EQU "1" set "ISO=%USERPROFILE%\Desktop\!o16lang!_2019_VISIO_PRO_Retail.ISO"
		if "%vi21install%" EQU "1" set "ISO=%USERPROFILE%\Desktop\!o16lang!_2021_VISIO_PRO_Retail.ISO"
		set "forCmd="!Zz!" l "!ISO!" "Office\Data""
		goto:NEXTBLAT
		
	)
	
	if defined DloadLP (
		call :GenerateLPLink > "%temp%\tmp.ps1"
		powershell -noprofile -executionpolicy bypass -file "%temp%\tmp.ps1"
		goto:TheFinalCountDown
	)

	if not defined DloadLP if not defined DloadImg (
	
		call :GenerateSetupLink > "%temp%\tmp.ps1"
		powershell -noprofile -executionpolicy bypass -file "%temp%\tmp.ps1"
		goto:TheFinalCountDown
	)
	
:NEXTBLAT
	if exist "%ISO%" for /f "tokens=4 skip=19 delims= " %%$ in ('"%forCmd%"') do if not defined OfficeVer set "OfficeVer=%%$"
	for %%I in ("%ISO%") do set "SourcePath=%%~dpI"
	for %%I in ("%ISO%") do set "newFile=%%~nI_v!OfficeVer:~12!.iso"
	if defined OfficeVer if not exist "!SourcePath!!newFile!" ren "%ISO%" "!newFile!"

:TheFinalCountDown
    echo ____________________________________________________________________________
	echo:
    echo:
	timeout /t 8
    goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:CheckActivationStatus
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	set "CDNBaseUrl=not set"
	set "UpdateUrl=not set"
	set "UpdateBranch=not set"
	cls
	powershell.exe -command "& {$pshost = Get-Host;$pswindow = $pshost.UI.RawUI;$newsize = $pswindow.BufferSize;$newsize.height = 100;$pswindow.buffersize = $newsize;}"
	echo:
	call :PrintTitle "================== SHOW CURRENT ACTIVATION STATUS =========================="
	echo:
	echo Office installation path:
	echo %installpath16%
	echo:
	if "%ProPlusVLFound%" EQU "YES" ((set "ChannelName=Native Volume (VLSC)")&&(set "UpdateUrl=Windows Update")&&(goto:CheckActCont))
	if "%StandardVLFound%" EQU "YES" ((set "ChannelName=Native Volume (VLSC)")&&(set "UpdateUrl=Windows Update")&&(goto:CheckActCont))
	if "%ProjectProVLFound%" EQU "YES" ((set "ChannelName=Native Volume (VLSC)")&&(set "UpdateUrl=Windows Update")&&(goto:CheckActCont))
	if "%VisioProVLFound%" EQU "YES" ((set "ChannelName=Native Volume (VLSC)")&&(set "UpdateUrl=Windows Update")&&(goto:CheckActCont))
	if "%_UWPappINSTALLED%" EQU "YES" ((set "ChannelName=Microsoft (Apps) Store")&&(set "UpdateUrl=Microsoft (Apps) Store")&&(goto:CheckActCont))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "CDNBaseUrl" 2^>nul') DO (Set "CDNBaseUrl=%%B")
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "UpdateUrl" 2^>nul') DO (Set "UpdateUrl=%%B")
	call:DecodeChannelName %UpdateUrl%
::===============================================================================================================
:CheckActCont
	echo Distribution-Channel:
	echo %ChannelName%
	echo:
	echo Updates-Url:
	echo %UpdateUrl%
	echo ____________________________________________________________________________
	echo:
	if "%_ProPlusRetail%" EQU "YES" ((echo Office Professional Plus 2016 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_ProPlusVolume%" EQU "YES" ((echo Office Professional Plus 2016 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus))
	if "%_ProPlus2019Retail%" EQU "YES" ((echo Office Professional Plus 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus2019))
	if "%_ProPlus2019Volume%" EQU "YES" ((echo Office Professional Plus 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus2019))
	if "%_ProPlus2021Retail%" EQU "YES" ((echo Office Professional Plus 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus2021))
	if "%_ProPlus2021Volume%" EQU "YES" ((echo Office Professional Plus 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus2021))
	if "%_ProPlusSPLA2021Volume%" EQU "YES" ((echo Office Professional Plus 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProPlus2021))
	if "%_StandardRetail%" EQU "YES" ((echo Office Standard 2016 --- ProductVersion: %o16version%)&&(echo:)&&(call :CheckKMSActivation Standard))
	if "%_StandardVolume%" EQU "YES" ((echo Office Standard 2016 --- ProductVersion: %o16version%)&&(echo:)&&(call :CheckKMSActivation Standard))
	if "%_Standard2019Retail%" EQU "YES" ((echo Office Standard 2019 --- ProductVersion: %o16version%)&&(echo:)&&(call :CheckKMSActivation Standard2019))
	if "%_Standard2019Volume%" EQU "YES" ((echo Office Standard 2019 --- ProductVersion: %o16version%)&&(echo:)&&(call :CheckKMSActivation Standard2019))
	if "%_Standard2021Retail%" EQU "YES" ((echo Office Standard 2021 --- ProductVersion: %o16version%)&&(echo:)&&(call :CheckKMSActivation Standard2021))
	if "%_Standard2021Volume%" EQU "YES" ((echo Office Standard 2021 --- ProductVersion: %o16version%)&&(echo:)&&(call :CheckKMSActivation Standard2021))
	if "%_StandardSPLA2021Volume%" EQU "YES" ((echo Office Standard 2021 --- ProductVersion: %o16version%)&&(echo:)&&(call :CheckKMSActivation Standard2021))
	if "%_O365ProPlusRetail%" EQU "YES" ((echo Microsoft 365 Apps for Enterprise --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_O365BusinessRetail%" EQU "YES" ((echo Microsoft 365 Apps for Business --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_O365HomePremRetail%" EQU "YES" ((echo Microsoft 365 Home Premium retail --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_O365SmallBusPremRetail%" EQU "YES" ((echo Microsoft 365 Small Business retail --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_ProfessionalRetail%" EQU "YES" ((echo Professional Retail --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_Professional2019Retail%" EQU "YES" ((echo Professional 2019 Retail --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_Professional2021Retail%" EQU "YES" ((echo Professional 2021 Retail --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_HomeBusinessRetail%" EQU "YES" ((echo Microsoft Home And Business --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_HomeBusiness2019Retail%" EQU "YES" ((echo Microsoft Home And Business 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_HomeBusiness2021Retail%" EQU "YES" ((echo Microsoft Home And Business 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_HomeStudentRetail%" EQU "YES" ((echo Microsoft Home And Student --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_HomeStudent2019Retail%" EQU "YES" ((echo Microsoft Home And Student 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_HomeStudent2021Retail%" EQU "YES" ((echo Microsoft Home And Student 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_MondoRetail%" EQU "YES" ((echo Office Mondo Grande Suite --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_MondoVolume%" EQU "YES" ((echo Office Mondo Grande Suite --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_PersonalRetail%" EQU "YES" ((echo Office Personal 2016 Retail --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_Personal2019Retail%" EQU "YES" ((echo Office Personal 2019 Retail --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_Personal2021Retail%" EQU "YES" ((echo Office Personal 2021 Retail --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_WordRetail%" EQU "YES" ((echo Word 2016 SingleApp ------------------ ProductVersion : %o16version%)&&(call :CheckKMSActivation Word))
	if "%_ExcelRetail%" EQU "YES" ((echo Excel 2016 SingleApp ----------------- ProductVersion : %o16version%)&&(call :CheckKMSActivation Excel))
	if "%_PowerPointRetail%" EQU "YES" ((echo PowerPoint 2016 SingleApp ------------ ProductVersion : %o16version%)&&(call :CheckKMSActivation PowerPoint))
	if "%_AccessRetail%" EQU "YES" ((echo Access 2016 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Access))
	if "%_OutlookRetail%" EQU "YES" ((echo Outlook 2016 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Outlook))
	if "%_PublisherRetail%" EQU "YES" ((echo Publisher 2016 Single App --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Publisher))
	if "%_OneNoteRetail%" EQU "YES" ((echo OneNote 2016 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation OneNote))
	if "%_OneNoteVolume%" EQU "YES" ((echo OneNote 2016 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation OneNote))
	if "%_OneNote2021Retail%" EQU "YES" ((echo OneNote 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation OneNote))
	if "%_SkypeForBusinessRetail%" EQU "YES" ((echo Skype For Business 2016 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation SkypeForBusiness))
	if "%_AppxWinword%" EQU "YES" ((echo Word UWP Appx --- ProductVersion : %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_AppxExcel%" EQU "YES" ((echo Excel UWP Appx --- ProductVersion : %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_AppxPowerPoint%" EQU "YES" ((echo PowerPoint UWP Appx - ProductVersion : %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_AppxAccess%" EQU "YES" ((echo Access UWP Appx ----- ProductVersion : %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_AppxOutlook%" EQU "YES" ((echo Outlook UWP Appx ---- ProductVersion : %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_AppxPublisher%" EQU "YES" ((echo Publisher UWP Appx -- ProductVersion : %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_AppxOneNote%" EQU "YES" ((echo OneNote UWP Appx ---- ProductVersion : %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_AppxSkypeForBusiness%" EQU "YES" ((echo Skype UWP Appx ------ ProductVersion : %o16version%)&&(call :CheckKMSActivation Mondo))
	if "%_Word2019Retail%" EQU "YES" ((echo Word 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Word2019))
	if "%_Excel2019Retail%" EQU "YES" ((echo Excel 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Excel2019))
	if "%_PowerPoint2019Retail%" EQU "YES" ((echo PowerPoint 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation PowerPoint2019))
	if "%_Access2019Retail%" EQU "YES" ((echo Access 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Access2019))
	if "%_Outlook2019Retail%" EQU "YES" ((echo Outlook 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Outlook2019))
	if "%_Publisher2019Retail%" EQU "YES" ((echo Publisher 2019 Single App --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Publisher2019))
	if "%_SkypeForBusiness2019Retail%" EQU "YES" ((echo Skype For Business 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation SkypeForBusiness2019))
	if "%_Word2019Volume%" EQU "YES" ((echo Word 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Word2019))
	if "%_Excel2019Volume%" EQU "YES" ((echo Excel 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Excel2019))
	if "%_PowerPoint2019Volume%" EQU "YES" ((echo PowerPoint 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation PowerPoint2019))
	if "%_Access2019Volume%" EQU "YES" ((echo Access 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Access2019))
	if "%_Outlook2019Volume%" EQU "YES" ((echo Outlook 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Outlook2019))
	if "%_Publisher2019Volume%" EQU "YES" ((echo Publisher 2019 Single App --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Publisher2019))
	if "%_SkypeForBusiness2019Volume%" EQU "YES" ((echo Skype For Business 2019 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation SkypeForBusiness2019))
	if "%_Word2021Retail%" EQU "YES" ((echo Word 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Word2021))
	if "%_Excel2021Retail%" EQU "YES" ((echo Excel 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Excel2021))
	if "%_PowerPoint2021Retail%" EQU "YES" ((echo PowerPoint 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation PowerPoint2021))
	if "%_Access2021Retail%" EQU "YES" ((echo Access 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Access2021))
	if "%_Outlook2021Retail%" EQU "YES" ((echo Outlook 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Outlook2021))
	if "%_Publisher2021Retail%" EQU "YES" ((echo Publisher 2021 Single App --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Publisher2021))
	if "%_SkypeForBusiness2021Retail%" EQU "YES" ((echo Skype For Business 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation SkypeForBusiness2021))
	if "%_Word2021Volume%" EQU "YES" ((echo Word 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Word2021))
	if "%_Excel2021Volume%" EQU "YES" ((echo Excel 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Excel2021))
	if "%_PowerPoint2021Volume%" EQU "YES" ((echo PowerPoint 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation PowerPoint2021))
	if "%_Access2021Volume%" EQU "YES" ((echo Access 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Access2021))
	if "%_Outlook2021Volume%" EQU "YES" ((echo Outlook 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Outlook2021))
	if "%_Publisher2021Volume%" EQU "YES" ((echo Publisher 2021 Single App --- ProductVersion: %o16version%)&&(call :CheckKMSActivation Publisher2021))
	if "%_SkypeForBusiness2021Volume%" EQU "YES" ((echo Skype For Business 2021 SingleApp --- ProductVersion: %o16version%)&&(call :CheckKMSActivation SkypeForBusiness2021))
	if "%_VisioProRetail%" EQU "YES" ((echo Visio Professional 2016 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioPro))
	if "%_AppxVisio%" EQU "YES" ((echo Visio Professional UWP Appx --- ProductVersion : %o16version%)&&(call :CheckKMSActivation VisioPro))
	if "%_ProjectProRetail%" EQU "YES" ((echo Project Professional 2016 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectPro))
	if "%_AppxProject%" EQU "YES" ((echo Project Professional UWP Appx - ProductVersion : %o16version%)&&(call :CheckKMSActivation ProjectPro))
	if "%_VisioPro2019Retail%" EQU "YES" ((echo Visio Professional 2019 ---- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioPro2019))
	if "%_VisioPro2019Volume%" EQU "YES" ((echo Visio Professional 2019 ---- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioPro2019))
	if "%_ProjectPro2019Retail%" EQU "YES" ((echo Project Professional 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectPro2019))
	if "%_ProjectPro2019Volume%" EQU "YES" ((echo Project Professional 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectPro2019))
	if "%_VisioPro2021Retail%" EQU "YES" ((echo Visio Professional 2021 ---- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioPro2021))
	if "%_VisioPro2021Volume%" EQU "YES" ((echo Visio Professional 2021 ---- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioPro2021))
	if "%_ProjectPro2021Retail%" EQU "YES" ((echo Project Professional 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectPro2021))
	if "%_ProjectPro2021Volume%" EQU "YES" ((echo Project Professional 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectPro2021))
	if "%_VisioStdRetail%" EQU "YES" ((echo Visio Standard 2016 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioSTD))
	if "%_VisioStdVolume%" EQU "YES" ((echo Visio Standard 2016 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioSTD))
	if "%_VisioStdXVolume%" EQU "YES" ((echo Visio Standard 2016 C2R --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioStdXC2R))
	if "%_VisioStd2019Retail%" EQU "YES" ((echo Visio Standard 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioSTD2019))
	if "%_VisioStd2019Volume%" EQU "YES" ((echo Visio Standard 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioStd2019))
	if "%_VisioStd2021Retail%" EQU "YES" ((echo Visio Standard 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioStd2021))
	if "%_VisioStd2021Volume%" EQU "YES" ((echo Visio Standard 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioStd2021))
	if "%_ProjectStdRetail%" EQU "YES" ((echo Project Standard 2016 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectStd))
	if "%_ProjectStdVolume%" EQU "YES" ((echo Project Standard 2016 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectStd))
	if "%_ProjectStdXVolume%" EQU "YES" ((echo Project Standard 2016 C2R --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectStdXC2R))
	if "%_VisioProXVolume%" EQU "YES" ((echo Visio Professional 2016 C2R --- ProductVersion: %o16version%)&&(call :CheckKMSActivation VisioProXC2R))
	if "%_ProjectProXVolume%" EQU "YES" ((echo Project Professional 2016 C2R --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectProXC2R))
	if "%_ProjectStd2019Retail%" EQU "YES" ((echo Project Standard 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectStd2019))
	if "%_ProjectStd2019Volume%" EQU "YES" ((echo Project Standard 2019 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectStd2019))
	if "%_ProjectStd2021Retail%" EQU "YES" ((echo Project Standard 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectStd2021))
	if "%_ProjectStd2021Volume%" EQU "YES" ((echo Project Standard 2021 --- ProductVersion: %o16version%)&&(call :CheckKMSActivation ProjectStd2021))
	
	echo:
	echo:
	if not defined debugMode pause
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:CheckKMSActivation
::===============================================================================================================
	set "Product=%1"
	set "LicStatus=9"
	set "PartialProductKey=XXXXX"
	set "LicStatusText=(---UNKNOWN---)           "
	set /a "GraceMin=0"
	set "EvalEndDate=00000000"
	set "activationtext=unknown"
	set "PartProdKey=not set"
	if %win% GEQ 9200	(
		
		set info=EvaluationEndDate,GracePeriodRemaining,ID,LicenseFamily,LicenseStatus,PartialProductKey
		set wmiSearch=Name like '%%%%!Product!%%%%' and PartialProductKey is not NULL
		
		call :Query "!info!" "!slp!" "!wmiSearch!"
		for /f "tokens=1,2,3,4,5,6,7,8 skip=3 delims=," %%g in ('type "%temp%\result"') do (
			set "EvalEndDate=%%g"
			set "GraceMin=%%h"
			set "ID=%%i"
			set "LicFamily=%%j"
			set "LicStatus=%%k"
			set "PartialProductKey=%%l"
		)
			
	)
	if %win% LSS 9200	(
		
		set info=EvaluationEndDate,GracePeriodRemaining,ID,LicenseFamily,LicenseStatus,PartialProductKey
		set wmiSearch=Name like '%%%%!Product!%%%%' and PartialProductKey is not NULL
		
		call :Query "!info!" "!ospp!" "!wmiSearch!"
		for /f "tokens=1,2,3,4,5,6,7,8 skip=3 delims=," %%g in ('type "%temp%\result"') do (
			set "EvalEndDate=%%g"
			set "GraceMin=%%h"
			set "ID=%%i"
			set "LicenseFamily=%%j"
			set "LicStatus=%%k"
			set "PartialProductKey=%%l"
		)
	)
	set /a GraceDays=!GraceMin!/1440
	set "EvalEndDate=!EvalEndDate:~0,8!"
	set "EvalEndDate=!EvalEndDate:~4,2!/!EvalEndDate:~6,2!/!EvalEndDate:~0,4!"
	if "!LicStatus!" EQU "0" (set "LicStatusText=(---UNLICENSED---)        ")
	if "!LicStatus!" EQU "1" (set "LicStatusText=(---LICENSED---)          ")
	if "!LicStatus!" EQU "2" (set "LicStatusText=(---OOB_GRACE---)         ")
	if "!LicStatus!" EQU "3" (set "LicStatusText=(---OOT_GRACE---)         ")
	if "!LicStatus!" EQU "4" (set "LicStatusText=(---NONGENUINE_GRACE---)  ")
	if "!LicStatus!" EQU "5" (set "LicStatusText=(---NOTIFICATIONS---)     ")
	if "!LicStatus!" EQU "6" (set "LicStatusText=(---EXTENDED_GRACE---)    ")
	echo:
	echo License Family: !LicFamily!
	echo:
	echo Activation status: !LicStatus!  !LicStatusText! PartialProductKey: !PartialProductKey!
	if "!EvalEndDate!" NEQ "01/01/1601" (set "activationtext=Product's activation is time-restricted")
	if "!EvalEndDate!" EQU "01/01/1601" (set "activationtext=Product is permanently activated")
	if !LicStatus! EQU 1 if !GraceMin! EQU 0 ((echo:)&&(echo Remaining Retail activation period: !activationtext!))
	if !LicStatus! GEQ 1 if !GraceDays! GEQ 1 (echo:)
	if !LicStatus! GEQ 1 if !GraceDays! GEQ 1 powershell -noprofile -command "!pswindowtitle!"; Write-Host "Remaining KMS activation period: '!GraceDays!' days left '-' License expires:' '" -nonewline; Get-Date -date $(Get-Date).AddMinutes(!GraceMin!) -Format (Get-Culture).DateTimeFormat.ShortDatePattern
	if "!EvalEndDate!" NEQ "00/00/0000" if "!EvalEndDate!" NEQ "01/01/1601" ((echo:)&&(echo Evaluation/Preview timebomb active - Office product end-of-life: !EvalEndDate!))
	echo ____________________________________________________________________________
	echo:
	goto :eof
::===============================================================================================================
::===============================================================================================================
:ChangeUpdPath
	call :CheckOfficeApplications
::===============================================================================================================
	set "CDNBaseUrl=not set"
	set "UpdateUrl=not set"
	set "UpdateBranch=not set"
	set "installtrigger=O"
	set "channeltrigger=O"
	set "restrictbuild=newest available"
	set "updatetoversion="
	cls
	echo:
	call :PrintTitle "================== CHANGE INSTALLED OFFICE UPDATE-PATH ====================="
	echo:
	if "%ProPlusVLFound%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for native VLSC Volume version)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	if "%StandardVLFound%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for native VLSC Volume version)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	if "%ProjectProVLFound%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for native VLSC Volume version)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	if "%VisioProVLFound%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for native VLSC Volume version)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	if "%_UWPappINSTALLED%" EQU "YES" ((echo:)&&(echo CHANGE OFFICE UPDATE-PATH is not possible for Office UWP Appx Store Apps)&&(echo:)&&(pause)&&(goto:Office16VnextInstall))
	
	if "%_ProPlusRetail%" EQU "YES"              (echo Office Professional Plus 2016 -------- ProductVersion : %o16version%)
	if "%_ProPlusVolume%" EQU "YES"              (echo Office Professional Plus 2016 -------- ProductVersion : %o16version%)
	if "%_ProPlus2019Retail%" EQU "YES"          (echo Office Professional Plus 2019 -------- ProductVersion : %o16version%)
	if "%_ProPlus2019Volume%" EQU "YES"          (echo Office Professional Plus 2019 -------- ProductVersion : %o16version%)
	if "%_ProPlus2021Retail%" EQU "YES"          (echo Office Professional Plus 2021 -------- ProductVersion : %o16version%)
	if "%_ProPlus2021Volume%" EQU "YES"          (echo Office Professional Plus 2021 -------- ProductVersion : %o16version%)
	if "%_ProPlusSPLA2021Volume%" EQU "YES"      (echo Office Professional Plus 2021 -------- ProductVersion : %o16version%)
	if "%_O365ProPlusRetail%" EQU "YES"          (echo Microsoft 365 Apps for Enterprise ---- ProductVersion : %o16version%)
	if "%_O365BusinessRetail%" EQU "YES"         (echo Microsoft 365 Apps for Business ------ ProductVersion : %o16version%)
	if "%_O365HomePremRetail%" EQU "YES" 		 (echo Microsoft 365 Home Premium retail ---- ProductVersion : %o16version%)
	if "%_O365SmallBusPremRetail%" EQU "YES" 	 (echo Microsoft 365 Small Business retail -- ProductVersion : %o16version%)
	if "%_ProfessionalRetail%" EQU "YES" 		 (echo Professional Retail ------------------ ProductVersion : %o16version%)
	if "%_Professional2019Retail%" EQU "YES" 	 (echo Professional 2019 Retail ------------- ProductVersion : %o16version%)
	if "%_Professional2021Retail%" EQU "YES" 	 (echo Professional 2021 Retail ------------- ProductVersion : %o16version%)
	if "%_HomeBusinessRetail%" EQU "YES" 		 (echo Microsoft Home And Business ---------- ProductVersion : %o16version%)
	if "%_HomeBusiness2019Retail%" EQU "YES" 	 (echo Microsoft Home And Business 2019 ----- ProductVersion : %o16version%)
	if "%_HomeBusiness2021Retail%" EQU "YES" 	 (echo Microsoft Home And Business 2021 ----- ProductVersion : %o16version%)
	if "%_HomeStudentRetail%" EQU "YES" 		 (echo Microsoft Home And Student ----------- ProductVersion : %o16version%)
	if "%_HomeStudent2019Retail%" EQU "YES" 	 (echo Microsoft Home And Student 2019 ------ ProductVersion : %o16version%)
	if "%_HomeStudent2021Retail%" EQU "YES" 	 (echo Microsoft Home And Student 2021 ------ ProductVersion : %o16version%)
	if "%_MondoRetail%" EQU "YES"                (echo Office Mondo Grande Suite ------------ ProductVersion : %o16version%)
	if "%_MondoVolume%" EQU "YES"                (echo Office Mondo Grande Suite ------------ ProductVersion : %o16version%)
	if "%_PersonalRetail%" EQU "YES"             (echo Office Personal 2016 Retail ---------- ProductVersion : %o16version%)
	if "%_Personal2019Retail%" EQU "YES"         (echo Office Personal 2019 Retail ---------- ProductVersion : %o16version%)
	if "%_Personal2021Retail%" EQU "YES"         (echo Office Personal 2021 Retail ---------- ProductVersion : %o16version%)
	if "%_WordRetail%" EQU "YES"                 (echo Word 2016 SingleApp ------------------ ProductVersion : %o16version%)
	if "%_ExcelRetail%" EQU "YES"                (echo Excel 2016 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_PowerPointRetail%" EQU "YES"           (echo PowerPoint 2016 SingleApp ------------ ProductVersion : %o16version%)
	if "%_AccessRetail%" EQU "YES"               (echo Access 2016 SingleApp ---------------- ProductVersion : %o16version%)
	if "%_OutlookRetail%" EQU "YES"              (echo Outlook 2016 SingleApp --------------- ProductVersion : %o16version%)
	if "%_PublisherRetail%" EQU "YES"            (echo Publisher 2016 SingleApp ------------- ProductVersion : %o16version%)
	if "%_OneNoteRetail%" EQU "YES"              (echo OneNote 2016 SingleApp --------------- ProductVersion : %o16version%)
	if "%_OneNoteVolume%" EQU "YES"              (echo OneNote 2016 SingleApp --------------- ProductVersion : %o16version%)
	if "%_OneNote2021Retail%" EQU "YES"          (echo OneNote 2021 SingleApp --------------- ProductVersion : %o16version%)
	if "%_SkypeForBusinessRetail%" EQU "YES"     (echo Skype 2016 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_AppxWinword%" EQU "YES"                (echo Word UWP Appx ------------------------ ProductVersion : %o16version%)
	if "%_AppxExcel%" EQU "YES"                  (echo Excel UWP Appx ----------------------- ProductVersion : %o16version%)
	if "%_AppxPowerPoint%" EQU "YES"             (echo PowerPoint UWP Appx ------------------ ProductVersion : %o16version%)
	if "%_AppxAccess%" EQU "YES"                 (echo Access UWP Appx ---------------------- ProductVersion : %o16version%)
	if "%_AppxOutlook%" EQU "YES"                (echo Outlook UWP Appx --------------------- ProductVersion : %o16version%)
	if "%_AppxPublisher%" EQU "YES"              (echo Publisher UWP Appx ------------------- ProductVersion : %o16version%)
	if "%_AppxOneNote%" EQU "YES"                (echo OneNote UWP Appx --------------------- ProductVersion : %o16version%)
	if "%_AppxSkypeForBusiness%" EQU "YES"       (echo Skype UWP Appx ----------------------- ProductVersion : %o16version%)
	if "%_Word2019Retail%" EQU "YES"             (echo Word 2019 SingleApp ------------------ ProductVersion : %o16version%)
	if "%_Excel2019Retail%" EQU "YES"            (echo Excel 2019 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_PowerPoint2019Retail%" EQU "YES"       (echo PowerPoint 2019 SingleApp ------------ ProductVersion : %o16version%)
	if "%_Access2019Retail%" EQU "YES"           (echo Access 2019 SingleApp ---------------- ProductVersion : %o16version%)
	if "%_Outlook2019Retail%" EQU "YES"          (echo Outlook 2019 SingleApp --------------- ProductVersion : %o16version%)
	if "%_Publisher2019Retail%" EQU "YES"        (echo Publisher 2019 SingleApp ------------- ProductVersion : %o16version%)
	if "%_SkypeForBusiness2019Retail%" EQU "YES" (echo Skype 2019 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_Word2019Volume%" EQU "YES"             (echo Word 2019 SingleApp ------------------ ProductVersion : %o16version%)
	if "%_Excel2019Volume%" EQU "YES"            (echo Excel 2019 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_PowerPoint2019Volume%" EQU "YES"       (echo PowerPoint 2019 SingleApp ------------ ProductVersion : %o16version%)
	if "%_Access2019Volume%" EQU "YES"           (echo Access 2019 SingleApp ---------------- ProductVersion : %o16version%)
	if "%_Outlook2019Volume%" EQU "YES"          (echo Outlook 2019 SingleApp --------------- ProductVersion : %o16version%)
	if "%_Publisher2019Volume%" EQU "YES"        (echo Publisher 2019 SingleApp ------------- ProductVersion : %o16version%)
	if "%_SkypeForBusiness2019Volume%" EQU "YES" (echo Skype 2019 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_Word2021Retail%" EQU "YES"             (echo Word 2021 SingleApp ------------------ ProductVersion : %o16version%)
	if "%_Excel2021Retail%" EQU "YES"            (echo Excel 2021 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_PowerPoint2021Retail%" EQU "YES"       (echo PowerPoint 2021 SingleApp ------------ ProductVersion : %o16version%)
	if "%_Access2021Retail%" EQU "YES"           (echo Access 2021 SingleApp ---------------- ProductVersion : %o16version%)
	if "%_Outlook2021Retail%" EQU "YES"          (echo Outlook 2021 SingleApp --------------- ProductVersion : %o16version%)
	if "%_Publisher2021Retail%" EQU "YES"        (echo Publisher 2021 SingleApp ------------- ProductVersion : %o16version%)
	if "%_SkypeForBusiness2021Retail%" EQU "YES" (echo Skype 2021 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_Word2021Volume%" EQU "YES"             (echo Word 2021 SingleApp ------------------ ProductVersion : %o16version%)
	if "%_Excel2021Volume%" EQU "YES"            (echo Excel 2021 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_PowerPoint2021Volume%" EQU "YES"       (echo PowerPoint 2021 SingleApp ------------ ProductVersion : %o16version%)
	if "%_Access2021Volume%" EQU "YES"           (echo Access 2021 SingleApp ---------------- ProductVersion : %o16version%)
	if "%_Outlook2021Volume%" EQU "YES"          (echo Outlook 2021 SingleApp --------------- ProductVersion : %o16version%)
	if "%_Publisher2021Volume%" EQU "YES"        (echo Publisher 2021 SingleApp ------------- ProductVersion : %o16version%)
	if "%_SkypeForBusiness2021Volume%" EQU "YES" (echo Skype 2021 SingleApp ----------------- ProductVersion : %o16version%)
	if "%_VisioProRetail%" EQU "YES"             (echo Visio professional 2016 -------------- ProductVersion : %o16version%)
	if "%_AppxVisio%" EQU "YES"                  (echo Visio professional UWP Appx ---------- ProductVersion : %o16version%)
	if "%_VisioPro2019Retail%" EQU "YES"         (echo Visio professional 2019 -------------- ProductVersion : %o16version%)
	if "%_VisioPro2019Volume%" EQU "YES"         (echo Visio professional 2019 -------------- ProductVersion : %o16version%)
	if "%_VisioPro2021Retail%" EQU "YES"         (echo Visio professional 2021 -------------- ProductVersion : %o16version%)
	if "%_VisioPro2021Volume%" EQU "YES"         (echo Visio professional 2021 -------------- ProductVersion : %o16version%)
	if "%_ProjectProRetail%" EQU "YES"           (echo Project professional 2016 ------------ ProductVersion : %o16version%)
	if "%_AppxProject%" EQU "YES"                (echo Project professional UWP Appx -------- ProductVersion : %o16version%)
	if "%_ProjectPro2019Retail%" EQU "YES"       (echo Project professional 2019 ------------ ProductVersion : %o16version%)
	if "%_ProjectPro2019Volume%" EQU "YES"       (echo Project professional 2019 ------------ ProductVersion : %o16version%)
	if "%_ProjectPro2021Retail%" EQU "YES"       (echo Project professional 2021 ------------ ProductVersion : %o16version%)
	if "%_ProjectPro2021Volume%" EQU "YES"       (echo Project professional 2021 ------------ ProductVersion : %o16version%)
	if "%_VisioStdRetail%" EQU "YES"       		 (echo Visio Standard 2016 ------------------ ProductVersion : %o16version%)
	if "%_VisioStdVolume%" EQU "YES"       		 (echo Visio Standard 2016 ------------------ ProductVersion : %o16version%)
	if "%_VisioStdXVolume%" EQU "YES"       	 (echo Visio Standard 2016 C2R -------------- ProductVersion : %o16version%)
	if "%_VisioStd2019Retail%" EQU "YES"     	 (echo Visio Standard 2019 ------------------ ProductVersion : %o16version%)
	if "%_VisioStd2019Volume%" EQU "YES"         (echo Visio Standard 2019 ------------------ ProductVersion : %o16version%)
	if "%_VisioStd2021Retail%" EQU "YES"         (echo Visio Standard 2021 ------------------ ProductVersion : %o16version%)
	if "%_VisioStd2021Volume%" EQU "YES"         (echo Visio Standard 2021 ------------------ ProductVersion : %o16version%)
	if "%_ProjectStdRetail%" EQU "YES"       	 (echo Project Standard 2016 ---------------- ProductVersion : %o16version%)
	if "%_ProjectStdVolume%" EQU "YES"       	 (echo Project Standard 2016 ---------------- ProductVersion : %o16version%)
	if "%_ProjectStdXVolume%" EQU "YES"		 	 (echo Project Standard 2016 C2R ------------ ProductVersion : %o16version%)
	if "%_ProjectProXVolume%" EQU "YES"		 	 (echo Project professional 2016 C2R -------- ProductVersion : %o16version%)
	if "%_VisioProXVolume%" EQU "YES"		 	 (echo Visio professional 2016 C2R ---------- ProductVersion : %o16version%)
	if "%_ProjectStd2019Retail%" EQU "YES"     	 (echo Project Standard 2019 ---------------- ProductVersion : %o16version%)
	if "%_ProjectStd2019Volume%" EQU "YES"       (echo Project Standard 2019 ---------------- ProductVersion : %o16version%)
	if "%_ProjectStd2021Retail%" EQU "YES"       (echo Project Standard 2021 ---------------- ProductVersion : %o16version%)
	if "%_ProjectStd2021Volume%" EQU "YES"       (echo Project Standard 2021 ---------------- ProductVersion : %o16version%)
	
	echo ____________________________________________________________________________
	echo:
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "CDNBaseUrl" 2^>nul') DO (Set "CDNBaseUrl=%%B")
	call:DecodeChannelName %CDNBaseUrl%
	echo Distribution-Channel:
	echo %ChannelName%
	echo:
	echo CDNBase-Url:
	echo %CDNBaseUrl%
	echo:
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "UpdateChannel" 2^>nul') DO (Set "UpdateUrl=%%B")
	call:DecodeChannelName %UpdateUrl%
	echo Updates-Channel:
	echo %ChannelName%
	echo:
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "UpdateUrl" 2^>nul') DO (Set "UpdateUrl=%%B")
	echo Updates-Url:
	echo %UpdateUrl%
	echo:
	echo Group-Policy defined UpdateBranch:
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate" /v "UpdateBranch" 2^>nul') DO (Set "UpdateBranch=%%B")
	echo %UpdateBranch%
	echo ____________________________________________________________________________
	echo:
	echo Possible Office Update-Channel ID VALUES:
	echo 1 = Current (Retail/RTM)
	echo 2 = CurrentPreview (Office Insider SLOW)
	echo 3 = BetaChannel (Office Insider FAST)
	echo 4 = MonthlyEnterprise
	echo 5 = SemiAnnual (Business)
	echo 6 = SemiAnnualPreview (Business Insider)
	echo 7 = PerpetualVL2019 (Office 2019 Volume)
	echo 8 = PerpetualVL2021 (Office 2021 Volume)
	echo 9 = DogfoodDevMain (MS Internal Use Only)
	echo X = exit to Main Menu
	echo:
	set /p channeltrigger=Set New Update-Channel-ID (1,2,3,4,5,6,7,8,9) or X ^>
	if "%channeltrigger%" EQU "1" (
		set "latestfile=latest_Current_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
		set "UpdateBranch=Current"
		if not exist !latestfile! call :CheckNewVersion Current 492350f6-3a01-4f97-b9c0-c7c6ddf67d60
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "2" (
		set "latestfile=latest_CurrentPreview_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be"
		set "UpdateBranch=CurrentPreview"
		if not exist !latestfile! call :CheckNewVersion CurrentPreview 64256afe-f5d9-4f86-8936-8840a6a4f5be
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "3" (
		set "latestfile=latest_BetaChannel_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f"
		set "UpdateBranch=BetaChannel"
		if not exist !latestfile! call :CheckNewVersion BetaChannel 5440fd1f-7ecb-4221-8110-145efaa6372f
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "4" (
		set "latestfile=latest_MonthlyEnterprise_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6"
		set "UpdateBranch=MonthlyEnterprise"
		if not exist !latestfile! call :CheckNewVersion MonthlyEnterprise 55336b82-a18d-4dd6-b5f6-9e5095c314a6
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "5" (
		set "latestfile=latest_SemiAnnual_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
		set "UpdateBranch=SemiAnnual"
		if not exist !latestfile! call :CheckNewVersion SemiAnnual 7ffbc6bf-bc32-4f92-8982-f9dd17fd3114
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "6" (
		set "latestfile=latest_SemiAnnualPreview_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf"
		set "UpdateBranch=SemiAnnualPreview"
		if not exist !latestfile! call :CheckNewVersion CSemiAnnualPreview b8f9b850-328d-4355-9145-c59439a0c4cf
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "7" (
		set "latestfile=latest_PerpetualVL2019_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/f2e724c1-748f-4b47-8fb8-8e0d210e9208"
		set "UpdateBranch=PerpetualVL2019"
		if not exist !latestfile! call :CheckNewVersion PerpetualVL2019 f2e724c1-748f-4b47-8fb8-8e0d210e9208
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "8" (
		set "latestfile=latest_PerpetualVL2021_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/5030841d-c919-4594-8d2d-84ae4f96e58e"
		set "UpdateBranch=PerpetualVL2021"
		if not exist !latestfile! call :CheckNewVersion PerpetualVL2021 5030841d-c919-4594-8d2d-84ae4f96e58e
		goto:UpdateChannelSel
	)
	if "%channeltrigger%" EQU "9" (
		set "latestfile=latest_DogfoodDevMain_build.txt"
		set "UpdateUrl=http://officecdn.microsoft.com/pr/ea4a4090-de26-49d7-93c1-91bff9e53fc3"
		set "UpdateBranch=not set"
		if not exist !latestfile! call :CheckNewVersion DogfoodDevMain ea4a4090-de26-49d7-93c1-91bff9e53fc3
		goto:UpdateChannelSel
	)
	if /I "%channeltrigger%" EQU "X" (goto:Office16VnextInstall)
	goto:ChangeUpdPath
::===============================================================================================================
:UpdateChannelSel
	echo:
	set /a countx=0
	cd /D "%OfficeRToolpath%"
	for /F "tokens=*" %%a in (!latestfile!) do (
		SET /a countx=!countx! + 1
		set var!countx!=%%a
	)
	set "o16upg1build=%var1%"
	set "o16upg2build=%var2%"
	echo Manually enter any build-nummer such as %o16upg2build%(prior build)
	echo or simply press return for updating to: %o16upg1build%(newest build)
	set /p restrictbuild=Set Office update build ^>
	if "%restrictbuild%" NEQ "newest available" set "updatetoversion=updatetoversion=%restrictbuild%"
	call :DecodeChannelName %UpdateUrl%
	echo ____________________________________________________________________________
	echo:
	echo New Update-Configuration will be set to:
	echo:
	echo Distribution-Channel : %ChannelName%
	echo Update To Version    : %restrictbuild%
	echo:
	set /p installtrigger=(ENTER) to proceed, (R)estart update, (E)xit to main menu ? ^>
    if /i "%installtrigger%" EQU "R" goto:ChangeUpdPath
	if /I "%installtrigger%" EQU "E" goto:Office16VnextInstall
::===============================================================================================================
:ChangeUpdateConf
	reg add HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v CDNBaseUrl /d %UpdateUrl% /f %MultiNul%
	reg add HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v UpdateUrl /d %UpdateUrl% /f %MultiNul%
	reg add HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v UpdateChannel /d %UpdateUrl% /f %MultiNul%
	reg add HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v UpdateChannelChanged /d True /f %MultiNul%
	if "%UpdateBranch%" EQU "not set" reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f %MultiNul%
	if "%UpdateBranch%" NEQ "not set" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %UpdateBranch% /f %MultiNul%
	reg delete HKLM\Software\Microsoft\Office\ClickToRun\Configuration /v UpdateToVersion /f %MultiNul%
	reg delete HKLM\Software\Microsoft\Office\ClickToRun\Updates /v UpdateToVersion /f %MultiNul%
	if "%restrictbuild%" NEQ "newest available" (("%CommonProgramFiles%\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user %updatetoversion% updatepromptuser=True displaylevel=True)&&(goto:Office16VnextInstall))
	"%CommonProgramFiles%\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user updatepromptuser=True displaylevel=True %MultiNul%
	%MultiNul% del /q latest*.txt
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:DecodeChannelName
	set "ChannelName=%1"
	set "ChannelName=%ChannelName:~-36%"
	if "%ChannelName%" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" (set "ChannelName=Current (Retail/RTM)")&&(goto:eof)
	if "%ChannelName%" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" (set "ChannelName=CurrentPreview (Office Insider SLOW)")&&(goto:eof)
	if "%ChannelName%" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" (set "ChannelName=BetaChannel (Office Insider FAST)")&&(goto:eof)
	if "%ChannelName%" EQU "55336b82-a18d-4dd6-b5f6-9e5095c314a6" (set "ChannelName=MonthlyEnterprise")&&(goto:eof)
	if "%ChannelName%" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" (set "ChannelName=SemiAnnual (Business)")&&(goto:eof)
	if "%ChannelName%" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" (set "ChannelName=SemiAnnualPreview (Business Insider)")&&(goto:eof)
	if "%ChannelName%" EQU "f2e724c1-748f-4b47-8fb8-8e0d210e9208" (set "ChannelName=PerpetualVL2019 (Office 2019 Volume)")&&(goto:eof)
	if "%ChannelName%" EQU "5030841d-c919-4594-8d2d-84ae4f96e58e" (set "ChannelName=PerpetualVL2021 (Office 2021 Volume)")&&(goto:eof)
	if "%ChannelName%" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" (set "ChannelName=DogfoodDevMain (MS Internal Use Only)")&&(goto:eof)
	set "ChannelName=Non_Standard_Channel (Manual_Override)"
	goto:eof
::===============================================================================================================
::===============================================================================================================
:DisableTelemetry
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	cls
	echo:
	call :PrintTitle "================== DISABLE ACQUISITION OF TELEMETRY DATA ==================="
	echo:
	echo Scheduler:  4 Office Telemetry related Tasks were set / changed
	schtasks /Change /TN "Microsoft\Office\Office Automatic Updates" /Disable %MultiNul%
	schtasks /Change /TN "Microsoft\Office\OfficeTelemetryAgentFallBack2016" /Disable %MultiNul%
	schtasks /Change /TN "Microsoft\Office\OfficeTelemetryAgentLogOn2016" /Disable %MultiNul%
	schtasks /Change /TN "Microsoft\Office\Office ClickToRun Service Monitor" /Disable %MultiNul%
	echo:
	echo Registry:  29 Office Telemetry related User Keys were set / changed
	REG ADD HKCU\Software\Microsoft\Office\Common\ClientTelemetry /v DisableTelemetry /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common /v sendcustomerdata /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\Feedback /v enabled /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\Feedback /v includescreenshot /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Outlook\Options\Mail /v EnableLogging /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Word\Options /v EnableLogging /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common /v qmenable /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common /v updatereliabilitydata /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\General /v shownfirstrunoptin /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\General /v skydrivesigninoption /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Common\ptwatson /v ptwoptin /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\Firstrun /v disablemovie /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM /v Enablelogging /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM /v EnableUpload /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM /v EnableFileObfuscation /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v accesssolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v olksolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v onenotesolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v pptsolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v projectsolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v publishersolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v visiosolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v wdsolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedapplications /v xlsolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v agave /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v appaddins /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v comaddins /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v documentfiles /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v templatefiles /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v templatefiles /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKCU\Software\Policies\Microsoft\office\16.0\common\privacy /v disconnectedstate /t REG_DWORD /d 2 /f %MultiNul%
	REG ADD HKCU\Software\Policies\Microsoft\office\16.0\common\privacy /v usercontentdisabled /t REG_DWORD /d 2 /f %MultiNul%
	REG ADD HKCU\Software\Policies\Microsoft\office\16.0\common\privacy /v downloadcontentdisabled /t REG_DWORD /d 2 /f %MultiNul%
	REG ADD HKCU\Software\Policies\Microsoft\office\16.0\common\privacy /v ControllerConnectedServicesEnabled /t REG_DWORD /d 2 /f %MultiNul%
	REG ADD HKCU\Software\Policies\Microsoft\office\16.0\common\clienttelemetry /v sendtelemetry /t REG_DWORD /d 3 /f %MultiNul%
	echo:
	echo Registry:  23 Office Telemetry related Machine Group Policies were set / changed
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common /v qmenable /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common /v updatereliabilitydata /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common\General /v shownfirstrunoptin /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common\General /v skydrivesigninoption /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Common\ptwatson /v ptwoptin /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\Firstrun /v disablemovie /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM /v Enablelogging /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM /v EnableUpload /t REG_DWORD /d 0 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM /v EnableFileObfuscation /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v accesssolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v olksolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v onenotesolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v pptsolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v projectsolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v publishersolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v visiosolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v wdsolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedapplications /v xlsolution /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v agave /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v appaddins /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v comaddins /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v documentfiles /t REG_DWORD /d 1 /f %MultiNul%
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\OSM\preventedsolutiontypes /v templatefiles /t REG_DWORD /d 1 /f %MultiNul%
	echo:
	echo Registry:  1 Office BING search service registry key was set to disabled
	REG ADD HKLM\Software\Policies\Microsoft\Office\16.0\common\officeupdate /v preventbinginstall /t REG_DWORD /d 1 /f > nul 2>&1
	echo ____________________________________________________________________________
	echo:
    echo:
	timeout /t 4
    goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:ResetRepair
	call :CheckOfficeApplications
::===============================================================================================================
::===============================================================================================================
	cls
	echo:
	call :PrintTitle "======================= RESET / REPAIR OFFICE =============================="
	echo:
::===============================================================================================================
    echo ____________________________________________________________________________
	echo:
	echo Removing Office xrm-license files...
	echo (Retail-/Grace-licenses will be refreshed by Office Quick-/Online-Repair)
	echo:
	"%OfficeRToolpath%\OfficeFixes\%winx%\cleanospp.exe" -Licenses
	echo ____________________________________________________________________________
	echo:
	echo Removing Office product keys...
	echo (Retail grace key will be installed after next Office apps start)
	echo:
	"%OfficeRToolpath%\OfficeFixes\%winx%\cleanospp.exe" -PKey
	echo ____________________________________________________________________________
	echo:
	echo Starting official Office repair program...
	echo (select option "QUICK REPAIR")
	"%CommonProgramFiles%\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=repair platform=%repairplatform% culture=%repairlang%
	echo:
    echo ____________________________________________________________________________
	echo:
::===============================================================================================================
	call :CheckOfficeApplications
::===============================================================================================================
	timeout /t 4
    goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:InstallO16
	set "of16install=0"
	set "of19install=0"
	set "of21install=0"
	set "of36install=0"
	set "ofbsinstall=0"
	set "mo16install=0"
	set "wd16disable=0"
	set "ex16disable=0"
	set "pp16disable=0"
	set "ac16disable=0"
	set "ol16disable=0"
	set "pb16disable=0"
	set "on16disable=0"
	set "st16disable=0"
	set "od16disable=0"
	set "bsbsdisable=0"
	set "vi16install=0"
	set "pr16install=0"
	set "vi19install=0"
	set "pr19install=0"
	set "vi21install=0"
	set "pr21install=0"
	set "wd16install=0"
	set "ex16install=0"
	set "pp16install=0"
	set "ac16install=0"
	set "ol16install=0"
	set "pb16install=0"
	set "on16install=0"
	set "on21install=0"
	set "sk16install=0"
	set "wd19install=0"
	set "ex19install=0"
	set "pp19install=0"
	set "ac19install=0"
	set "ol19install=0"
	set "pb19install=0"
	set "sk19install=0"
	set "wd21install=0"
	set "ex21install=0"
	set "pp21install=0"
	set "ac21install=0"
	set "ol21install=0"
	set "pb21install=0"
	set "sk21install=0"
	set "installtrigger=not set"
	set "createpackage=0"
	set "productstoadd=0"
	set "excludedapps=0"
	set "productkeys=0"
	set "type=Retail"
	set "downpath=not set"
:InstallO16Loop
	cls
	if defined OnlineInstaller goto :InstSuites
	echo:
	call :PrintTitle "================== SELECT OFFICE FULL SUITE / SINGLE APPS ================="
	echo:
	set "searchdirpattern=16."
	if defined inidownpath set "downpath=!inidownpath!"
	echo !downpath! | %SingleNul% find /i "not set" && (
		set "downpath=%USERPROFILE%\desktop"
	)
	set /p downpath=Set Office Package Download Path ^= "!downpath!" ^>
	set "downpath=!downpath:"=!"
	if /I "!downpath!" EQU "X" ((set "downpath=not set")&&(goto:Office16VnextInstall))
	set "downdrive=!downpath:~0,2!"
	if "!downdrive:~-1!" NEQ ":" (set "downpath=not set" & goto:InstallO16Loop)
	cd /d !downdrive! %MultiNul% || (set "downpath=not set" & goto:InstallO16Loop)
	set "downdrive=!downpath:~0,3!"
	if "!downdrive:~-1!" EQU "\" (set "downpath=!downdrive!!downpath:~3!") else (set "downpath=!downdrive:~0,2!\!downpath:~2!")
	if "!downpath:~-1!" EQU "\" set "downpath=!downpath:~0,-1!"
::===============================================================================================================
	cd /d "!downdrive!\" %MultiNul%
	cd /d "!downpath!" %MultiNul%
	set /a countx=0
	
	echo:
	echo List of available installation packages
	
	if exist "!downpath!\*!searchdirpattern!*" (
		for /F "tokens=*" %%a in ('dir "!downpath!" /ad /b ^| find /i "%searchdirpattern%"') do (
			if exist "%%a\Office\Data\16.*" (
				echo:
				SET /a countx=!countx! + 1
				set packagelist!countx!=%%a
				echo !countx!   %%a
			)
		)
	)
	
	set "Zz=%~dp0OfficeFixes\win_x32\7z.exe"
	set "forCmd=dir "!downpath!" /b | find /i ".iso""
	for /f "tokens=*" %%$ in ('"%forCmd%"') do (
		echo "%%$" | >nul find /i "16." && (
			set "PACKAGE=%%$"
			set "PACKAGE_FIXED=!PACKAGE:~0,-4!"
			2>nul "!Zz!" l "!PACKAGE!" "Office\Data" | >nul find /i "Office\Data\16." && (
				if not exist "!downpath!\!PACKAGE_FIXED!\Office\Data\16.*" (
					echo:
					SET /a countx=!countx! + 1
					set packagelist!countx!=!PACKAGE_FIXED!
					echo !countx!   !PACKAGE_FIXED!
				)
			)
		)
	)
	
	if %countx% GTR 0 goto:PackageFound
	echo.
	echo ERROR ### No install packages found
	%SingleNul% timeout 2
	goto :InstallO16Loop
::===============================================================================================================
:PackageFound
	echo:
	echo:
	set /a packnum=0
	set /p packnum=Enter package number ^>
	if /I "%packnum%" EQU "X" goto:Office16VnextInstall
	if %packnum% EQU 0 ((set "searchdirpattern=not set")&&(goto:InstallO16Loop))
	if %packnum% GTR %countx% ((set "searchdirpattern=not set")&&(goto:InstallO16Loop))
	echo:
	
	set "downpath=!downpath!\!packagelist%packnum%!"
	set "installpath=!downpath!"
	
	if not exist "%installpath%\Office\Data\16.*" (
		if not exist "%installpath%.iso" (
			(echo:)&&(echo ERROR: Missing files.)&&(echo:)&&(pause)&&(goto:InstallO16)
		)
		md "%installpath%" %MultiNul%
		%Zz% x -y -o"%installpath%" "%installpath%.iso" "Office" %MultiNul% || (
			rd/s/q "%installpath%" %MultiNul%
			(echo:)&&(echo ERROR: Fail to extract files from ISO.)&&(echo:)&&(pause)&&(goto:InstallO16)
		)
		if not exist "%installpath%\Office\Data\16.*" (
			rd/s/q "%installpath%" %MultiNul%
			(echo:)&&(echo ERROR: Missing files.)&&(echo:)&&(pause)&&(goto:InstallO16)
		)
	)
	
	if "%installpath:~-1%" EQU "\" set "installpath=%installpath:~0,-1%"
	set countx=0
	cd /d "!downpath!"
	if exist package.info (
		for /F "tokens=*" %%a in (package.info) do (
		set /a countx=!countx! + 1
		set var!countx!=%%a
	))
	
	if !countx! LSS 5 goto :GetInfoFromFolder
	
	set "distribchannel=%var1%"
	if /i "%distribchannel:~-1%" EQU " " set "distribchannel=%distribchannel:~0,-1%"
	set "o16build=%var2%"
	set "o16lang=%var3%"
	
	set "AskUser="
	if /i '!o16lang!' EQU 'Multi' (
		set AskUser=true
		call :GetInfoFromFolder
		set "o16lang=!lang!"
	)
	call :SetO16Language
		
	set "o16arch=%var4%"
	set "o16updlocid=%var5%"
	
:Pdhfsdj45X
	cd /D "%OfficeRToolpath%"
	if "%winx%" EQU "win_x32" if "!o16arch!" EQU "x64" ((echo:)&&(echo ERROR: You can't install x64/64bit Office on x86/32bit Windows)&&(echo:)&&(pause)&&(goto:InstallO16))
::===============================================================================================================
:InstSuites
	set "instmethod=XML"
	cd /D "%OfficeRToolpath%"
	cls
	echo:
	call :PrintTitle "================== SELECT OFFICE FULL SUITE - SINGLE APPS ================="
	echo:
::===============================================================================================================
:SelFullSuite
	echo:
	echo Select full Office Suite for install:
	echo:
	call :Print "0.) Single Apps Install (no full suite)" "%BB_Blue%"
	echo:
	call :Print "1.) Office Professional Plus 2016 Retail" "%BB_Blue%"
	echo:
	call :Print "2.) Office Professional Plus 2019 Volume" "%BB_Blue%"
	echo:
	call :Print "3.) Office Professional Plus 2021 Volume" "%BB_Blue%"
	
	echo:
	call :Print "4.) Office 2016 Mondo" "%BB_Red%"
	echo:
	call :Print "5.) Microsoft 365 Apps for Business" "%BB_Red%"
	echo:
	call :Print "6.) Microsoft 365 Apps for Enterprise" "%BB_Red%"
	
	echo:
	call :Print "7.) Visio-Project 2016 Retail" "%BB_Green%"
	echo:
	call :Print "8.) Visio-Project 2019 Volume" "%BB_Green%"
	echo:
	call :Print "9.) Visio-Project 2021 Volume" "%BB_Green%"
	
	echo:
	echo:
	set /p installtrigger=Enter 1..9,0 or x to exit ^>
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	
	if defined OnlineInstaller (
	
		set "o16latestbuild="
	
		set "o16updlocid=492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
		set "distribchannel=Current"
		call :CheckNewVersion Current !o16updlocid!
		set "o16build=!o16latestbuild!"
		
		echo !o16latestbuild!|>nul find /i "not set" && (
			pause
			goto :Office16VnextInstall
		)
		
		if /I "%installtrigger%" EQU "2" set "o16updlocid=f2e724c1-748f-4b47-8fb8-8e0d210e9208"
		if /I "%installtrigger%" EQU "8" set "o16updlocid=f2e724c1-748f-4b47-8fb8-8e0d210e9208"
		if /i "!o16updlocid!" EQU "f2e724c1-748f-4b47-8fb8-8e0d210e9208" (
			set "type=Volume"
			set "distribchannel=PerpetualVL2019"
			call :CheckNewVersion PerpetualVL2019 !o16updlocid!
			set "o16build=!o16latestbuild!"
		)
		
		if /I "%installtrigger%" EQU "3" set "o16updlocid=5030841d-c919-4594-8d2d-84ae4f96e58e"
		if /I "%installtrigger%" EQU "9" set "o16updlocid=5030841d-c919-4594-8d2d-84ae4f96e58e"
		if /i "!o16updlocid!" EQU "5030841d-c919-4594-8d2d-84ae4f96e58e" (
			set "type=Volume"
			set "distribchannel=PerpetualVL2021"
			call :CheckNewVersion PerpetualVL2021 !o16updlocid!
			set "o16build=!o16latestbuild!"
		)
	)
	
	if "%installtrigger%" EQU "0" (goto:SingleAppsInstall)
	if "%installtrigger%" EQU "1" ((set "type=Retail")&&(set "of16install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "2" ((set "type=Volume")&&(set "of19install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "3" (if %win% GEQ 9600 if %o16build:~5,5% GEQ 14000 ((set "type=Volume")&&(set "of21install=1")&&(goto:InstallExclusions)))
	if "%installtrigger%" EQU "4" ((set "type=Retail")&&(set "mo16install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "5" ((set "type=Retail")&&(set "ofbsinstall=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "6" ((set "type=Retail")&&(set "of36install=1")&&(goto:InstallExclusions))
	if "%installtrigger%" EQU "7" ((set "type=Retail")&&(goto:InstVi16Pr16))
	if "%installtrigger%" EQU "8" ((set "type=Volume")&&(goto:InstVi19Pr19))
	if "%installtrigger%" EQU "9" (if %win% GEQ 9600 if %o16build:~5,5% GEQ 14000 ((set "type=Volume")&&(goto:InstVi21Pr21)))
	goto:InstSuites
::===============================================================================================================
:SingleAppsInstall
	echo:
	set /p installtrigger=Which "version" for Single App to install (1=2016, 2=2019, 3=2021) ^>
	if /I "%installtrigger%" EQU "X" goto:Office16VnextInstall
	if "%installtrigger%" EQU "1" goto:SingleApps2016Install
	if "%installtrigger%" EQU "2" goto:SingleApps2019Install
	if %win% GEQ 9600 if "%installtrigger%" EQU "3" if %o16build:~5,5% GEQ 14000 goto:SingleApps2021Install
	goto:InstSuites
:SingleApps2016Install
	if defined OnlineInstaller (
		set "distribchannel=Current"
		set "o16updlocid=492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
		call :CheckNewVersion Current !o16updlocid!
		set "o16build=!o16latestbuild!"
	)
	echo:
	set /p wd16install=Set Word 2016 Single App Install (1/0) ^>
	if /I "%wd16install%" EQU "X" goto:Office16VnextInstall
	set /p ex16install=Set Excel 2016 Single App Install (1/0) ^>
	if /I "%ex16install%" EQU "X" goto:Office16VnextInstall
	set /p pp16install=Set Powerpoint 2016 Single App Install (1/0) ^>
	if /I "%pp16install%" EQU "X" goto:Office16VnextInstall
	set /p ac16install=Set Access 2016 Single App Install (1/0) ^>
	if /I "%ac16install%" EQU "X" goto:Office16VnextInstall
	set /p ol16install=Set Outlook 2016 Single App Install (1/0) ^>
	if /I "%ol16install%" EQU "X" goto:Office16VnextInstall
	set /p pb16install=Set Publisher 2016 Single App Install (1/0) ^>
	if /I "%pb16install%" EQU "X" goto:Office16VnextInstall
	set /p on16install=Set OneNote 2016 Single App Install (1/0) ^>
	if /I "%on16install%" EQU "X" goto:Office16VnextInstall
	set /p sk16install=Set Skype For Business 2016 Single App Install (1/0) ^>
	if /I "%sk16install%" EQU "X" goto:Office16VnextInstall
	goto:InstVi16Pr16
:SingleApps2019Install
	if defined OnlineInstaller (
		set "type=Volume"
		set "o16updlocid=f2e724c1-748f-4b47-8fb8-8e0d210e9208"
		set "distribchannel=PerpetualVL2019"
		call :CheckNewVersion PerpetualVL2019 !o16updlocid!
		set "o16build=!o16latestbuild!"
	)
	echo:
	set /p wd19install=Set Word 2019 Single App Install (1/0) ^>
	if /I "%wd19install%" EQU "X" goto:Office16VnextInstall
	set /p ex19install=Set Excel 2019 Single App Install (1/0) ^>
	if /I "%ex19install%" EQU "X" goto:Office16VnextInstall
	set /p pp19install=Set Powerpoint 2019 Single App Install (1/0) ^>
	if /I "%pp19install%" EQU "X" goto:Office16VnextInstall
	set /p ac19install=Set Access 2019 Single App Install (1/0) ^>
	if /I "%ac19install%" EQU "X" goto:Office16VnextInstall
	set /p ol19install=Set Outlook 2019 Single App Install (1/0) ^>
	if /I "%ol19install%" EQU "X" goto:Office16VnextInstall
	set /p pb19install=Set Publisher 2019 Single App Install (1/0) ^>
	if /I "%pb19install%" EQU "X" goto:Office16VnextInstall
	set /p sk19install=Set Skype For Business 2019 Single App Install (1/0) ^>
	if /I "%sk19install%" EQU "X" goto:Office16VnextInstall
	goto:InstVi19Pr19
:SingleApps2021Install
	if defined OnlineInstaller (
		set "type=Volume"
		set "o16updlocid=5030841d-c919-4594-8d2d-84ae4f96e58e"
		set "distribchannel=PerpetualVL2021"
		call :CheckNewVersion PerpetualVL2021 !o16updlocid!
		set "o16build=!o16latestbuild!"
	)
	echo:
	set /p wd21install=Set Word 2021 Single App Install (1/0) ^>
	if /I "%wd21install%" EQU "X" goto:Office16VnextInstall
	set /p ex21install=Set Excel 2021 Single App Install (1/0) ^>
	if /I "%ex21install%" EQU "X" goto:Office16VnextInstall
	set /p pp21install=Set Powerpoint 2021 Single App Install (1/0) ^>
	if /I "%pp21install%" EQU "X" goto:Office16VnextInstall
	set /p ac21install=Set Access 2021 Single App Install (1/0) ^>
	if /I "%ac21install%" EQU "X" goto:Office16VnextInstall
	set /p ol21install=Set Outlook 2021 Single App Install (1/0) ^>
	if /I "%ol21install%" EQU "X" goto:Office16VnextInstall
	set /p pb21install=Set Publisher 2021 Single App Install (1/0) ^>
	if /I "%pb21install%" EQU "X" goto:Office16VnextInstall
	set /p on21install=Set OneNote 2021 Single App Install (1/0) ^>
	if /I "%on21install%" EQU "X" goto:Office16VnextInstall
	set /p sk21install=Set Skype For Business 2021 Single App Install (1/0) ^>
	if /I "%sk21install%" EQU "X" goto:Office16VnextInstall
	goto:InstVi21Pr21
::===============================================================================================================
:InstallExclusions
	if "%mo16install%" EQU "1" ((set "of16install=0")&&(set "of19install=0")&&(set "of21install=0")&&(set "of36install=0")&&(set "ofbsinstall=0"))
	if "%of16install%" EQU "1" ((set "mo16install=0")&&(set "of19install=0")&&(set "of21install=0")&&(set "of36install=0")&&(set "ofbsinstall=0"))
	if "%of19install%" EQU "1" ((set "mo16install=0")&&(set "of16install=0")&&(set "of21install=0")&&(set "of36install=0")&&(set "ofbsinstall=0"))
	if "%of21install%" EQU "1" ((set "mo16install=0")&&(set "of16install=0")&&(set "of19install=0")&&(set "of36install=0")&&(set "ofbsinstall=0"))
	if "%of36install%" EQU "1" ((set "mo16install=0")&&(set "of16install=0")&&(set "of21install=0")&&(set "of19install=0")&&(set "ofbsinstall=0"))
	if "%ofbsinstall%" EQU "1" ((set "mo16install=0")&&(set "of16install=0")&&(set "of21install=0")&&(set "of19install=0")&&(set "of36install=0"))
	echo:
	echo Full Suite Install Exclusion List - Disable not needed Office Programs
	set /p wd16disable=Disable Word Install  (1/0) ^>
	if /I "%wd16disable%" EQU "X" goto:Office16VnextInstall
	set /p ex16disable=Disable Excel Install (1/0) ^>
	if /I "%ex16disable%" EQU "X" goto:Office16VnextInstall
	set /p pp16disable=Disable Powerpoint Install (1/0) ^>
	if /I "%pp16disable%" EQU "X" goto:Office16VnextInstall
	set /p ac16disable=Disable Access Install (1/0) ^>
	if /I "%ac16disable%" EQU "X" goto:Office16VnextInstall
	set /p ol16disable=Disable Outlook Install (1/0) ^>
	if /I "%ol16disable%" EQU "X" goto:Office16VnextInstall
	set /p pb16disable=Disable Publisher Install (1/0) ^>
	if /I "%pb16disable%" EQU "X" goto:Office16VnextInstall
	set /p on16disable=Disable OneNote Install (1/0) ^>
	if /I "%on16disable%" EQU "X" goto:Office16VnextInstall
	set /p st16disable=Disable Teams and Skype for Business Install (1/0) ^>
	if /I "%st16disable%" EQU "X" goto:Office16VnextInstall
	set /p od16disable=Disable OneDrive For Business Install (1/0) ^>
	if /I "%od16disable%" EQU "X" goto:Office16VnextInstall
	set /p bsbsdisable=Disable Bing Search Background Service Install (1/0) ^>
	if /I "%bsbsdisable%" EQU "X" goto:Office16VnextInstall
::===============================================================================================================
	if "%of36install%" EQU "1" goto:InstVi19Pr19
	if "%ofbsinstall%" EQU "1" goto:InstVi19Pr19
	if "%of19install%" EQU "1" goto:InstVi19Pr19
	if "%of21install%" EQU "1" goto:InstVi21Pr21
:InstVi16Pr16
	echo:
	set /p vi16install=Set Visio 2016 Install (1/0) ^>
	set /p pr16install=Set Project 2016 Install (1/0) ^>
	goto:InstViPrEnd
:InstVi19Pr19
	echo:
	set /p vi19install=Set Visio 2019 Install (1/0) ^>
	set /p pr19install=Set Project 2019 Install (1/0) ^>
	goto:InstViPrEnd
:InstVi21Pr21
	echo:
	set /p vi21install=Set Visio 2021 Install (1/0) ^>
	set /p pr21install=Set Project 2021 Install (1/0) ^>
::===============================================================================================================
:InstViPrEnd
	echo ____________________________________________________________________________
	echo:
::===============================================================================================================
	if "!o16updlocid!" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" (echo Source Channel: "Current" - !o16build! -Setup-)&&(goto:PendSetupContinue)
	if "!o16updlocid!" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" (echo Source Channel: "CurrentPreview" - !o16build! -Setup-)&&(goto:PendSetupContinue)
	if "!o16updlocid!" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" (echo Source Channel: "BetaChannel" - !o16build! -Setup-)&&(goto:PendSetupContinue)
	if "!o16updlocid!" EQU "55336b82-a18d-4dd6-b5f6-9e5095c314a6" (echo Source Channel: "MonthlyEnterprise" - !o16build! -Setup-)&&(goto:PendSetupContinue)
	if "!o16updlocid!" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" (echo Source Channel: "SemiAnnual" - !o16build! -Setup-)&&(goto:PendSetupContinue)
	if "!o16updlocid!" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" (echo Source Channel: "SemiAnnualPreview" - !o16build! -Setup-)&&(goto:PendSetupContinue)
	if "!o16updlocid!" EQU "f2e724c1-748f-4b47-8fb8-8e0d210e9208" (echo Source Channel: "PerpetualVL2019" - !o16build! -Setup-)&&(goto:PendSetupContinue)
	if "!o16updlocid!" EQU "5030841d-c919-4594-8d2d-84ae4f96e58e" (echo Source Channel: "PerpetualVL2021" - !o16build! -Setup-)&&(goto:PendSetupContinue)
	if "!o16updlocid!" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" (echo Source Channel: "DogfoodDevMain" - !o16build! -Setup-)&&(goto:PendSetupContinue)
	echo "Manual_Override:" !o16updlocid! - !o16build! -Setup-
::===============================================================================================================
:PendSetupContinue
	cls
	echo:
	call :PrintTitle "The following programs are selected for install"
	echo:
	if "%wd16install%" EQU "1" goto:PendSetupSingleApp
	if "%ex16install%" EQU "1" goto:PendSetupSingleApp
	if "%pp16install%" EQU "1" goto:PendSetupSingleApp
	if "%ac16install%" EQU "1" goto:PendSetupSingleApp
	if "%ol16install%" EQU "1" goto:PendSetupSingleApp
	if "%pb16install%" EQU "1" goto:PendSetupSingleApp
	if "%on16install%" EQU "1" goto:PendSetupSingleApp
	if "%on21install%" EQU "1" goto:PendSetupSingleApp
	if "%sk16install%" EQU "1" goto:PendSetupSingleApp
	if "%wd19install%" EQU "1" goto:PendSetupSingleApp
	if "%ex19install%" EQU "1" goto:PendSetupSingleApp
	if "%pp19install%" EQU "1" goto:PendSetupSingleApp
	if "%ac19install%" EQU "1" goto:PendSetupSingleApp
	if "%ol19install%" EQU "1" goto:PendSetupSingleApp
	if "%pb19install%" EQU "1" goto:PendSetupSingleApp
	if "%sk19install%" EQU "1" goto:PendSetupSingleApp
	if "%wd21install%" EQU "1" goto:PendSetupSingleApp
	if "%ex21install%" EQU "1" goto:PendSetupSingleApp
	if "%pp21install%" EQU "1" goto:PendSetupSingleApp
	if "%ac21install%" EQU "1" goto:PendSetupSingleApp
	if "%ol21install%" EQU "1" goto:PendSetupSingleApp
	if "%pb21install%" EQU "1" goto:PendSetupSingleApp
	if "%sk21install%" EQU "1" goto:PendSetupSingleApp
::===============================================================================================================
	if "%of16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Office Professional Plus 2016" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%of19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Office Professional Plus 2019 Volume" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%of21install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Office Professional Plus 2021 Volume" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%of36install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Microsoft 365 Apps for Enterprise" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%ofbsinstall%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Microsoft 365 Apps for Business" -foreground "Green")&&(goto:PendSetupFullSuite)
	if "%mo16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Mondo 2016 Grande Suite" -foreground "Green")&&(goto:PendSetupFullSuite)
	goto:PendSetupProjectVisio
:PendSetupFullSuite
	if "%wd16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Word" -foreground "Red")
	if "%wd16disable%" EQU "0" (echo --^> Enabled:  Word)
	if "%ex16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Excel" -foreground "Red")
	if "%ex16disable%" EQU "0" (echo --^> Enabled:  Excel)
	if "%pp16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Powerpoint" -foreground "Red")
	if "%pp16disable%" EQU "0" (echo --^> Enabled:  PowerPoint)
	if "%ac16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Access" -foreground "Red")
	if "%ac16disable%" EQU "0" (echo --^> Enabled:  Access)
	if "%ol16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Outlook" -foreground "Red")
	if "%ol16disable%" EQU "0" (echo --^> Enabled:  Outlook)
	if "%pb16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Publisher" -foreground "Red")
	if "%pb16disable%" EQU "0" (echo --^> Enabled:  Publisher)
	if "%on16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: OneNote" -foreground "Red")
	if "%on16disable%" EQU "0" (echo --^> Enabled:  OneNote)
	if "%st16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Teams / Skype For Business" -foreground "Red")
	if "%st16disable%" EQU "0" (echo --^> Enabled:  Teams / Skype For Business)
	if "%od16disable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: OneDrive For Business" -foreground "Red")
	if "%od16disable%" EQU "0" (echo --^> Enabled:  OneDrive For Business)
	if "%bsbsdisable%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "--> Disabled: Bing Search Background Service" -foreground "Red")
	if "%bsbsdisable%" EQU "0" (echo --^> Enabled:  Bing Search Background Service)
	goto:PendSetupProjectVisio
::===============================================================================================================
:PendSetupSingleApp	
	if "%wd16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Word 2016 Single App" -foreground "Green")
	if "%ex16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Excel 2016 Single App" -foreground "Green")
	if "%pp16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "PowerPoint 2016 Single App" -foreground "Green")
	if "%ac16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Access 2016 Single App" -foreground "Green")
	if "%ol16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Outlook 2016 Single App" -foreground "Green")
	if "%pb16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Publisher 2016 Single App" -foreground "Green")
	if "%on16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "OneNote 2016 Single App" -foreground "Green")
	if "%on21install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "OneNote 2021 Single App" -foreground "Green")
	if "%sk16install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Skype For Business 2016 Single App" -foreground "Green")
	if "%wd19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Word 2019 Single App" -foreground "Green")
	if "%ex19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Excel 2019 Single App" -foreground "Green")
	if "%pp19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "PowerPoint 2019 Single App" -foreground "Green")
	if "%ac19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Access 2019 Single App" -foreground "Green")
	if "%ol19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Outlook 2019 Single App" -foreground "Green")
	if "%pb19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Publisher 2019 Single App" -foreground "Green")
	if "%sk19install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Skype For Business 2019 Single App" -foreground "Green")
	if "%wd21install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Word 2021 Single App" -foreground "Green")
	if "%ex21install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Excel 2021 Single App" -foreground "Green")
	if "%pp21install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "PowerPoint 2021 Single App" -foreground "Green")
	if "%ac21install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Access 2021 Single App" -foreground "Green")
	if "%ol21install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Outlook 2021 Single App" -foreground "Green")
	if "%pb21install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Publisher 2021 Single App" -foreground "Green")
	if "%sk21install%" EQU "1" (powershell -noprofile -command "%pswindowtitle%"; Write-Host "Skype For Business 2021 Single App" -foreground "Green")
::===============================================================================================================
:PendSetupProjectVisio
	if "%vi16install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Visio Professional 2016" -foreground "Green")
	if "%pr16install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Project Professional 2016" -foreground "Green")
	if "%vi19install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Visio Professional 2019 Volume" -foreground "Green")
	if "%pr19install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Project Professional 2019 Volume" -foreground "Green")
	if "%vi21install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Visio Professional 2021 Volume" -foreground "Green")
	if "%pr21install%" EQU "1" (echo:)&&(powershell -noprofile -command "%pswindowtitle%"; Write-Host "Project Professional 2021 Volume" -foreground "Green")
::===============================================================================================================
if not defined OnlineInstaller goto :OnlineInstaller_NEXT
	:ChannelSelected_xd
	echo:
	set "o16buildCKS=!o16build!"
    set /p o16build=Set Office Build - or press return for !o16build! ^>
	echo "!o16build!" | >nul findstr /r "16.[0-9].[0-9][0-9][0-9][0-9][0-9].[0-9][0-9][0-9][0-9][0-9]" || (set "o16build=!o16buildCKS!" & goto :ChannelSelected_xd)
:OnlineInstaller_Language_MENU
	call :ChoiceLangSelect
	:OnlineInstaller_Language_MENU_Loop
	echo:
	if /i "!o16lang!" EQU "not set" call :CheckSystemLanguage
	set /p o16lang=Set Language Value - or press return for !o16lang! ^>
	set "o16lang=!o16lang:, =!"
	set "o16lang=!o16lang:,=!"
	if defined o16lang if /i "x!o16lang:~0,1!" EQU "x " set "o16lang=!o16lang:~1!"
	if defined o16lang if /i "!o16lang:-1!x" EQU " x" set "o16lang=!o16lang:~0,-1!"
	call :SetO16Language
	if defined langnotfound (
		set "o16lang=not set"
		goto:OnlineInstaller_Language_MENU_Loop
	)

:OnlineInstaller_ARCH_MENU
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF NOT DEFINED PROCESSOR_ARCHITEW6432 set sBit=86)
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'x86' 		(IF DEFINED PROCESSOR_ARCHITEW6432 set sBit=64)
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'AMD64' 	set sBit=64
	if /i '%PROCESSOR_ARCHITECTURE%' EQU 'IA64' 	set sBit=64
	
	set "o16arch=x!sBit!"
	if defined inidownarch ((echo "!inidownarch!" | %SingleNul% find /i "not set") || set "o16arch=!inidownarch!")
	if /i '!o16arch!' EQU 'Multi' set "o16arch=x!sBit!"	
	if /i 'x!sBit!' NEQ '!o16arch!' (if /i '!sBit!' EQU '86' (set "o16arch=x!sBit!"))
	
	echo:
	set /p o16arch=Set architecture to download (x86 or x64) - or press return for !o16arch! ^>
	if /i "!o16arch!" EQU "x86" goto :OnlineInstaller_NEXT
	if /i "!o16arch!" EQU "x64" goto :OnlineInstaller_NEXT
	goto :OnlineInstaller_ARCH_MENU
	
:OnlineInstaller_NEXT
	
	REM cls
	echo:
	echo ____________________________________________________________________________
    echo:
	call :PrintTitle "========================= Pending Install (SUMMARY) ========================="
    echo:
	if "!o16updlocid!" EQU "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" echo Channel-ID:   !o16updlocid! (Current) && goto:ZetaXX112
	if "!o16updlocid!" EQU "64256afe-f5d9-4f86-8936-8840a6a4f5be" echo Channel-ID:   !o16updlocid! (CurrentPreview) && goto:ZetaXX112
	if "!o16updlocid!" EQU "5440fd1f-7ecb-4221-8110-145efaa6372f" echo Channel-ID:   !o16updlocid! (BetaChannel) && goto:ZetaXX112
	if "!o16updlocid!" EQU "55336b82-a18d-4dd6-b5f6-9e5095c314a6" echo Channel-ID:   !o16updlocid! (MonthlyEnterprise) && goto:ZetaXX112
	if "!o16updlocid!" EQU "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" echo Channel-ID:   !o16updlocid! (SemiAnnual) && goto:ZetaXX112
	if "!o16updlocid!" EQU "b8f9b850-328d-4355-9145-c59439a0c4cf" echo Channel-ID:   !o16updlocid! (SemiAnnualPreview) && goto:ZetaXX112
	if "!o16updlocid!" EQU "f2e724c1-748f-4b47-8fb8-8e0d210e9208" echo Channel-ID:   !o16updlocid! (PerpetualVL2019) && goto:ZetaXX112
	if "!o16updlocid!" EQU "5030841d-c919-4594-8d2d-84ae4f96e58e" echo Channel-ID:   !o16updlocid! (PerpetualVL2021) && goto:ZetaXX112
	if "!o16updlocid!" EQU "ea4a4090-de26-49d7-93c1-91bff9e53fc3" echo Channel-ID:   !o16updlocid! (DogfoodDevMain) && goto:ZetaXX112
	echo Channel-ID:   !o16updlocid! (Manual_Override)
::===============================================================================================================
:ZetaXX112
	echo Office Build: !o16build!
	echo Language:     !o16lang! (%langtext%)
    echo Architecture: !o16arch!
    echo ____________________________________________________________________________
	echo:
	
								set "menuX=(ENTER) to Install, (C)reate install Package, (R)estart installation, (E)xit to main menu ? >"
	if defined createIso 		set "menuX=(ENTER) to Create ISO, (R)estart installation, (E)xit to main menu ? >"
	if defined OnlineInstaller 	set "menuX=(ENTER) to Install, (R)estart installation, (E)xit to main menu ? >"
	set /p installtrigger=!menuX!
	
	cls
	echo:
	if /i "%installtrigger%" EQU "C" set "createpackage=1"
	if /i "%installtrigger%" EQU "R" goto:InstallO16
	if /i "%installtrigger%" EQU "E" goto:Office16VnextInstall

::===============================================================================================================
:OfficeC2RXMLInstall

	REM cls
    echo:
	call :PrintTitle "================= INSTALL OFFICE FULL SUITE / SINGLE APPS =================="
    echo:
	if /i "!o16arch!" EQU "Multi" goto:ZETAXX13
	
    if /i "!o16arch!" EQU "x64" (set "o16a=64") else (set "o16a=32")
	if /i "%instmethod%" EQU "XML" echo Creating setup files "setup.exe", "configure%o16a%.xml" and "start_setup.cmd"
	if /i "%instmethod%" EQU "C2R" echo Creating setup file "start_setup.cmd"
	
	if defined OnlineInstaller (
		set "downpath=%temp%"
		set "installpath=%temp%"
	)
	
	echo:
    echo in Installpath: "%installpath%"
    echo:
	
	if /i "%instmethod%" EQU "XML" set "oxml=!downpath!\configure%o16a%.xml"
	if /i "%instmethod%" EQU "XML" copy "%OfficeRToolpath%\OfficeFixes\setup.exe" "!downpath!" /Y %MultiNul%
	if /i "%instmethod%" EQU "XML" (set "channel= channel="%distribchannel%"")
	if /i "%instmethod%" EQU "C2R" if exist "!downpath!\setup*.exe" del /s /q "!downpath!\setup*.exe" %MultiNul%
	if /i "%distribchannel%" EQU "Manual_Override" (set "channel=")
	if /i "%distribchannel%" EQU "DogfoodDevMain" (set "channel=")
	if exist "!downpath!\configure*.xml" del /s /q "!downpath!\configure*.xml" %MultiNul%
	set "obat=!downpath!\start_setup.cmd"
	copy "%OfficeRToolpath%\OfficeFixes\start_setup.cmd" "!downpath!" /Y %MultiNul%
	
	if "%instmethod%" EQU "C2R" goto:CreateC2RConfig
	if "%instmethod%" EQU "XML" goto:CreateXMLConfig
	goto:InstallO16
	
:ZETAXX13
	if /i "%instmethod%" EQU "XML" copy "%OfficeRToolpath%\OfficeFixes\setup.exe" "!downpath!" /Y %MultiNul%
	if /i "%instmethod%" EQU "XML" (set "channel= channel="%distribchannel%"")
	if /i "%instmethod%" EQU "C2R" if exist "!downpath!\setup*.exe" del /s /q "!downpath!\setup*.exe" %MultiNul%
	if /i "%distribchannel%" EQU "Manual_Override" (set "channel=")
	if /i "%distribchannel%" EQU "DogfoodDevMain" (set "channel=")
	if exist "!downpath!\configure*.xml" del /s /q "!downpath!\configure*.xml" %MultiNul%
	set "obat=!downpath!\start_setup.cmd"
	copy "%OfficeRToolpath%\OfficeFixes\start_setup.cmd" "!downpath!" /Y %MultiNul%
	
	set "o16a=32"
	set "oxml=!downpath!\configure32.xml"
	call :generateXML
	
	set "o16a=64"
	set "oxml=!downpath!\configure64.xml"
	call :generateXML
		
	goto:CreateStartSetupBatch
	
::===============================================================================================================
:CreateXMLConfig
	call :generateXML
	goto:CreateStartSetupBatch
::===============================================================================================================
::===============================================================================================================
:CreateC2RConfig
	if "%mo16install%" EQU "1" (
		set "productstoadd=!productstoadd!^^|Mondo%type%.16_%%instlang%%_x-none"
		set "productID=Mondo%type%"
		)
    if "%of16install%" EQU "1" (
		set "productstoadd=!productstoadd!^^|ProPlus%type%.16
		_%%instlang%%_x-none"
		set "productID=ProPlus%type%"
		)
    if "%of19install%" EQU "1" (
		set "productstoadd=!productstoadd!^^|ProPlus2019%type%.16_%%instlang%%_x-none"
		set "productID=ProPlus2019%type%"
		)
    if "%of21install%" EQU "1" (
		set "productstoadd=!productstoadd!^^|ProPlus2021%type%.16_%%instlang%%_x-none"
		set "productID=ProPlus2021%type%"
		)
	if "%of36install%" EQU "1" (
		set "productstoadd=!productstoadd!^^|O365ProPlus%type%.16_%%instlang%%_x-none"
		set "productID=O365ProPlus%type%"
		)
	if "%ofbsinstall%" EQU "1" (
        set "productstoadd=!productstoadd!^^|O365Business%type%.16_%%instlang%%_x-none"
        set "productID=O365Business%type%"
		)
		if "%wd16disable%" EQU "1" set "excludedapps=!excludedapps!,word"
		if "%ex16disable%" EQU "1" set "excludedapps=!excludedapps!,excel"
		if "%pp16disable%" EQU "1" set "excludedapps=!excludedapps!,powerpoint"
		if "%ac16disable%" EQU "1" set "excludedapps=!excludedapps!,access"
		if "%ol16disable%" EQU "1" set "excludedapps=!excludedapps!,outlook"
		if "%pb16disable%" EQU "1" set "excludedapps=!excludedapps!,publisher"
		if "%on16disable%" EQU "1" set "excludedapps=!excludedapps!,onenote"
		if "%st16disable%" EQU "1" set "excludedapps=!excludedapps!,lync"
		if "%st16disable%" EQU "1" set "excludedapps=!excludedapps!,teams"
		if "%od16disable%" EQU "1" set "excludedapps=!excludedapps!,groove"
		if "%od16disable%" EQU "1" set "excludedapps=!excludedapps!,onedrive"
		if "%bsbsdisable%" EQU "1" set "excludedapps=!excludedapps!,bing"
    )
	if "!excludedapps:~0,2!" EQU "0," (set "excludedapps=%productID%.excludedapps.16^=!excludedapps:~2!") else (set "excludedapps=")
::===============================================================================================================		
    if "%vi16install%" EQU "1" set "productstoadd=!productstoadd!^^|VisioPro%type%.16_%%instlang%%_x-none"
	if "%pr16install%" EQU "1" set "productstoadd=!productstoadd!^^|ProjectPro%type%.16_%%instlang%%_x-none" 
    if "%vi19install%" EQU "1" set "productstoadd=!productstoadd!^^|VisioPro2019%type%.16_%%instlang%%_x-none"
	if "%pr19install%" EQU "1" set "productstoadd=!productstoadd!^^|ProjectPro2019%type%.16_%%instlang%%_x-none"
    if "%vi21install%" EQU "1" set "productstoadd=!productstoadd!^^|VisioPro2021%type%.16_%%instlang%%_x-none"
	if "%pr21install%" EQU "1" set "productstoadd=!productstoadd!^^|ProjectPro2021%type%.16_%%instlang%%_x-none"
::===============================================================================================================
    if "%wd16install%" EQU "1" set "productstoadd=!productstoadd!^^|Word%type%.16_%%instlang%%_x-none"
	if "%wd19install%" EQU "1" set "productstoadd=!productstoadd!^^|Word2019%type%.16_%%instlang%%_x-none"
	if "%wd21install%" EQU "1" set "productstoadd=!productstoadd!^^|Word2021%type%.16_%%instlang%%_x-none"
	if "%ex16install%" EQU "1" set "productstoadd=!productstoadd!^^|Excel%type%.16_%%instlang%%_x-none"
    if "%ex19install%" EQU "1" set "productstoadd=!productstoadd!^^|Excel2019%type%.16_%%instlang%%_x-none"
	if "%ex21install%" EQU "1" set "productstoadd=!productstoadd!^^|Excel2021%type%.16_%%instlang%%_x-none"
	if "%pp16install%" EQU "1" set "productstoadd=!productstoadd!^^|PowerPoint%type%.16_%%instlang%%_x-none"
    if "%pp19install%" EQU "1" set "productstoadd=!productstoadd!^^|PowerPoint2019%type%.16_%%instlang%%_x-none"
    if "%pp21install%" EQU "1" set "productstoadd=!productstoadd!^^|PowerPoint2021%type%.16_%%instlang%%_x-none"
	if "%ac16install%" EQU "1" set "productstoadd=!productstoadd!^^|Access%type%.16_%%instlang%%_x-none"
    if "%ac19install%" EQU "1" set "productstoadd=!productstoadd!^^|Access2019%type%.16_%%instlang%%_x-none"
    if "%ac21install%" EQU "1" set "productstoadd=!productstoadd!^^|Access2021%type%.16_%%instlang%%_x-none"
    if "%ol16install%" EQU "1" set "productstoadd=!productstoadd!^^|Outlook%type%.16_%%instlang%%_x-none"
    if "%ol19install%" EQU "1" set "productstoadd=!productstoadd!^^|Outlook2019%type%.16_%%instlang%%_x-none"
    if "%ol21install%" EQU "1" set "productstoadd=!productstoadd!^^|Outlook2021%type%.16_%%instlang%%_x-none"
	if "%pb16install%" EQU "1" set "productstoadd=!productstoadd!^^|Publisher%type%.16_%%instlang%%_x-none"
    if "%pb19install%" EQU "1" set "productstoadd=!productstoadd!^^|Publisher2019%type%.16_%%instlang%%_x-none"
    if "%pb21install%" EQU "1" set "productstoadd=!productstoadd!^^|Publisher2021%type%.16_%%instlang%%_x-none"
	if "%on16install%" EQU "1" set "productstoadd=!productstoadd!^^|OneNote%type%.16_%%instlang%%_x-none"
    if "%sk16install%" EQU "1" set "productstoadd=!productstoadd!^^|SkypeForBusiness%type%.16_%%instlang%%_x-none"
	if "%sk19install%" EQU "1" set "productstoadd=!productstoadd!^^|SkypeForBusiness2019%type%.16_%%instlang%%_x-none"
	if "%sk21install%" EQU "1" set "productstoadd=!productstoadd!^^|SkypeForBusiness2021%type%.16_%%instlang%%_x-none"
	if "%on21install%" EQU "1" set "productstoadd=!productstoadd!^^|OneNote2021%type%.16_%%instlang%%_x-none"
::===============================================================================================================
:CreateStartSetupBatch
	if "%distribchannel%" EQU "Current" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "CurrentPreview" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "BetaChannel" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "MonthlyEnterprise" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
		if "%distribchannel%" EQU "SemiAnnual" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "SemiAnnualPreview" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "PerpetualVL2019" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "PerpetualVL2021" (
		echo :: Set Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "DogfoodDevMain" (
		echo :: Remove Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f ^>nul 2^>^&1 >>"%obat%"
	)
	if "%distribchannel%" EQU "Manual_Override" (
		echo :: Remove Group Policy value "UpdateBranch" in registry for "%distribchannel%" >>"%obat%"
		echo reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f ^>nul 2^>^&1 >>"%obat%"
	)
	echo ^:^:=============================================================================================================== >>"%obat%"
	if "%instmethod%" EQU "C2R" echo start "" /MIN "%%CommonProgramFiles%%\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" deliverymechanism=%%instupdlocid%% platform=%%instarch1%% productreleaseid=none forcecentcheck= culture=%%instlang%% defaultplatform=False storeid= lcid=%%instlcid%% b= forceappshutdown=True piniconstotaskbar=False scenariosubtype=ODT scenario=unknown updatesenabled.16=True acceptalleulas.16=True updatebaseurl.16=http://officecdn.microsoft.com/pr/%%instupdlocid%% cdnbaseurl.16=http://officecdn.microsoft.com/pr/%%instupdlocid%% version.16=%%instversion%% mediatype.16=Local baseurl.16=%%installfolder%% sourcetype.16=Local flt.downloadappvcab=unknown flt.useclientcabmanager=unknown flt.useexptransportinplacepl=unknown flt.useaddons=unknown flt.useofficehelperaddon=unknown flt.useonedriveclientaddon=unknown productstoadd=!productstoadd:~3! !excludedapps! >>"%obat%"
	if "%instmethod%" EQU "XML" echo "powershell" start setup.exe -WorkingDirectory '%%installfolder%%' -Args '/configure configure%%instarch2%%.xml' -Verb RunAs -WindowStyle hidden >>"%obat%"
	echo exit >>"%obat%"
	echo ^:^:=============================================================================================================== >>"%obat%"
	if defined createIso (
		set "createIso="
		
		%MultiNul% "%OfficeRToolpath%\OfficeFixes\win_x32\oscdimg.exe" -m -u1 "!downpath!" "!downpath!.iso" || (
			echo.
			echo ERROR ### Iso Creation failed
			echo.
		)
		goto:InstDone
	)
	if /i "%createpackage%" EQU "1" goto:InstDone
::===============================================================================================================
	if "%distribchannel%" EQU "Current" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f %MultiNul%
	if "%distribchannel%" EQU "CurrentPreview" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f %MultiNul%
	if "%distribchannel%" EQU "BetaChannel" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f %MultiNul%
	if "%distribchannel%" EQU "MonthlyEnterprise" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f %MultiNul%
	if "%distribchannel%" EQU "SemiAnnual" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f %MultiNul%
	if "%distribchannel%" EQU "SemiAnnualPreview" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f %MultiNul%
	if "%distribchannel%" EQU "PerpetualVL2019" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f %MultiNul%
	if "%distribchannel%" EQU "PerpetualVL2021" reg add HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /d %distribchannel% /f %MultiNul%
	if "%distribchannel%" EQU "DogfoodDevMain" reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f %MultiNul%
	if "%distribchannel%" EQU "Manual_Override" reg delete HKLM\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate /v UpdateBranch /f %MultiNul%
	cd /D "%installpath%"
	
	if defined OnlineInstaller (
		"powershell" start '%temp%\setup.exe' -Args '/configure !oxml!' -Verb RunAs -WindowStyle hidden
		goto :InstDone
	)
	
	start "" /MIN "%obat%"
::===============================================================================================================
:InstDone
	echo ____________________________________________________________________________
    echo:
	echo:
	timeout /t 4
    goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:CheckOfficeApplications
	set "_ProPlusRetail=NO"
	set "_ProPlusVolume=NO"
	set "_ProPlus2019Retail=NO"
	set "_ProPlus2021Retail=NO"
	set "_O365ProPlusRetail=NO"
	set "_O365BusinessRetail=NO"
	set "_O365HomePremRetail=NO"
	set "_O365SmallBusPremRetail=NO"
	set "_ProfessionalRetail=NO"
	set "_Professional2019Retail=NO"
	set "_Professional2021Retail=NO"	
	set "_HomeBusinessRetail=NO"
	set "_HomeBusiness2019Retail=NO"
	set "_HomeBusiness2021Retail=NO"
	set "_HomeStudentRetail=NO"
	set "_HomeStudent2019Retail=NO"
	set "_HomeStudent2021Retail=NO"
	set "_MondoRetail=NO"
	set "_MondoVolume=NO"
	set "_PersonalRetail=NO"
	set "_Personal2019Retail=NO"
	set "_Personal2021Retail=NO"
	set "_StandardRetail=NO"
	set "_StandardVolume=NO"
	set "_Standard2019Retail=NO"
	set "_Standard2019Volume=NO"
	set "_Standard2021Retail=NO"
	set "_Standard2021Volume=NO"
	set "_StandardSPLA2021Volume=NO"
	set "_VisioProRetail=NO"
	set "_ProjectProRetail=NO"
	set "_VisioPro2019Retail=NO"
	set "_ProjectPro2019Retail=NO"	
	set "_VisioStdRetail=NO"
	set "_VisioStdVolume=NO"
	set "_VisioStdXVolume=NO"
	set "_VisioStd2019Retail=NO"
	set "_VisioStd2019Volume=NO"
	set "_VisioStd2021Retail=NO"
	set "_VisioStd2021Volume=NO"
	set "_ProjectStdRetail=NO"
	set "_ProjectStdVolume=NO"
	set "_ProjectStdXVolume=NO"
	set "_ProjectProXVolume=NO"
	set "_VisioProXVolume=NO"
	set "_ProjectStd2019Retail=NO"
	set "_ProjectStd2019Volume=NO"
	set "_ProjectStd2021Retail=NO"
	set "_ProjectStd2021Volume=NO"
	set "_VisioPro2021Retail=NO"
	set "_ProjectPro2021Retail=NO"
	set "_WordRetail=NO"
	set "_ExcelRetail=NO"
	set "_PowerPointRetail=NO"
	set "_AccessRetail=NO"
	set "_OutlookRetail=NO"
	set "_PublisherRetail=NO"
	set "_OneNoteRetail=NO"
	set "_OneNoteVolume=NO"
	set "_OneNote2021Retail=NO"
	set "_SkypeForBusinessRetail=NO"
	set "_Word2019Retail=NO"
	set "_Excel2019Retail=NO"
	set "_PowerPoint2019Retail=NO"
	set "_Access2019Retail=NO"
	set "_Outlook2019Retail=NO"
	set "_Publisher2019Retail=NO"
	set "_SkypeForBusiness2019Retail=NO"
	set "_Word2021Retail=NO"
	set "_Excel2021Retail=NO"
	set "_PowerPoint2021Retail=NO"
	set "_Access2021Retail=NO"
	set "_Outlook2021Retail=NO"
	set "_Publisher2021Retail=NO"
	set "_SkypeForBusiness2021Retail=NO"
	set "_ProPlusVolume=NO"
	set "_ProPlus2019Volume=NO"
	set "_ProPlus2021Volume=NO"
	set "_ProPlusSPLA2021Volume=NO"
	set "_O365ProPlusVolume=NO"
	set "_O365BusinessVolume=NO"
	set "_MondoVolume=NO"
	set "_VisioProVolume=NO"
	set "_ProjectProVolume=NO"
	set "_VisioPro2019Volume=NO"
	set "_ProjectPro2019Volume=NO"
	set "_VisioPro2021Volume=NO"
	set "_ProjectPro2021Volume=NO"
	set "_WordVolume=NO"
	set "_ExcelVolume=NO"
	set "_PowerPointVolume=NO"
	set "_AccessVolume=NO"
	set "_OutlookVolume=NO"
	set "_PublisherVolume=NO"
	set "_SkypeForBusinessVolume=NO"
	set "_Word2019Volume=NO"
	set "_Excel2019Volume=NO"
	set "_PowerPoint2019Volume=NO"
	set "_Access2019Volume=NO"
	set "_Outlook2019Volume=NO"
	set "_Publisher2019Volume=NO"
	set "_SkypeForBusiness2019Volume=NO"
	set "_Word2021Volume=NO"
	set "_Excel2021Volume=NO"
	set "_PowerPoint2021Volume=NO"
	set "_Access2021Volume=NO"
	set "_Outlook2021Volume=NO"
	set "_Publisher2021Volume=NO"
	set "_SkypeForBusiness2021Volume=NO"
	set "_UWPappINSTALLED=NO"
	set "_AppxWinword=NO"
	set "_AppxExcel=NO"
	set "_AppxPowerPoint=NO"
	set "_AppxAccess=NO"
	set "_AppxPublisher=NO"
	set "_AppxOutlook=NO"
	set "_AppxSkypeForBusiness=NO"
	set "_AppxOneNote=NO"
	set "_AppxVisio=NO"
	set "_AppxProject=NO"
	set "ProPlusVLFound=NO"
	set "StandardVLFound=NO"
	set "ProjectProVLFound=NO"
	set "VisioProVLFound=NO"
	set "installpath16=not set"
	set "officepath3=not set"
	set "o16version=not set"
	set "o16arch=not set"
	reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "InstallationPath" %MultiNul%
	if %errorlevel% EQU 0 goto:CheckOffice16C2R
	reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" %MultiNul%
	if %errorlevel% EQU 0 goto:CheckOfficeVL32onW64
	reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" %MultiNul%
	if %errorlevel% EQU 0 goto:CheckOfficeVL32W32orVL64W64
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\msoxmled.exe" %MultiNul%
	if %errorlevel% EQU 0 ((set "_UWPappINSTALLED=YES")&&(goto:CheckAppxOffice16UWP))
	(echo:) && (echo Supported Office 2016/2019/2021 product not found) && (echo:) && (pause) && (goto:Office16VnextInstall)
	goto:eof
::===============================================================================================================
::===============================================================================================================
:CheckAppxOffice16UWP
	for /F "tokens=9 delims=\_() " %%A IN ('reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\msoxmled.exe" /ve 2^>nul') DO (set "o16version=%%A")
	for /F "tokens=4" %%A IN ('reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\msoxmled.exe" /ve 2^>nul') DO (set "installpath16=%%A")
	set "installpath16=C:\Program !installpath16!"
	set "installpath16=!installpath16:~0,-35!"
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\winword.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxWinword=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxExcel=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxPowerPoint=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\msaccess.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxAccess=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\mspub.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxPublisher=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxOutlook=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\lync.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxSkypeForBusiness=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\onenote.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxOneNote=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\visio.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxVisio=YES")
	reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\winproj.exe" %MultiNul%
	if %errorlevel% EQU 0 (set "_AppxProject=YES")
	goto:eof
::===============================================================================================================
::===============================================================================================================
:CheckOffice16C2R
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "Platform" 2^>nul') DO (set "o16arch=%%B")
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "InstallationPath" 2^>nul') DO (Set "installpath16=%%B")
	set "officepath3=%installpath16%\Office16"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v "ProductReleaseIds" 2^>nul') DO (Set "Office16AppsInstalled=%%B")
	for /F "tokens=1,2,3,4,5,6,7,8,9,10,11,12,13 delims=," %%A IN ("%Office16AppsInstalled%") DO (
	set "_%%A=YES"
	set "_%%B=YES"
	set "_%%C=YES"
	set "_%%D=YES"
	set "_%%E=YES"
	set "_%%F=YES"
	set "_%%G=YES"
	set "_%%H=YES"
	set "_%%I=YES"
	set "_%%J=YES"
	set "_%%K=YES"
	set "_%%L=YES"
	set "_%%M=YES"
	)
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs" /v "ActiveConfiguration" 2^>nul') DO (set "o16activeconf=%%B")
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\%o16activeconf%" /v "Modifier" 2^>nul') DO (set "o16version=%%B")
	set "o16version=%o16version:~0,16%"
	if "%o16version:~15,1%" EQU "|" (set "o16version=%o16version:~0,14%")
	
	if "%o16version:~4,1%" EQU "|" (
		for /F "tokens=2,*" %%A IN ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v "VersionToReport" 2^>nul') DO (set "o16version=%%B")
	)
	goto:eof
::===============================================================================================================
::===============================================================================================================
:CheckOfficeVL32onW64
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0011-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProPlusVLFound=%%B") %MultiNul%
	if "%ProPlusVLFound:~-39%" EQU "Microsoft Office Professional Plus 2016" ((set "ProPlusVLFound=YES")&&(set "_ProPlusRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0012-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "StandardVLFound=%%B") %MultiNul%
	if "%StandardVLFound:~-30%" EQU "Microsoft Office Standard 2016" ((set "StandardVLFound=YES")&&(set "_StandardRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-003B-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProjectProVLFound=%%B") %MultiNul%
	if "%ProjectProVLFound:~-35%" EQU "Microsoft Project Professional 2016" ((set "ProjectProVLFound=YES")&&(set "_ProjectProRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0051-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "VisioProVLFound=%%B") %MultiNul%
	if "%VisioProVLFound:~-33%" EQU "Microsoft Visio Professional 2016" ((set "VisioProVLFound=YES")&&(set "_VisioProRetail=YES"))
	if "%_ProPlusRetail%" EQU "YES" goto:OfficeVL32onW64Path
	if "%_StandardRetail%" EQU "YES" goto:OfficeVL32onW64Path
	if "%_ProjectProRetail%" EQU "YES" goto:OfficeVL32onW64Path
	if "%_VisioProRetail%" EQU "YES" goto:OfficeVL32onW64Path
	goto:Office16VnextInstall
::===============================================================================================================
:OfficeVL32onW64Path
	set "o16arch=x86"
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" 2^>nul') DO (Set "installpath16=%%B") %MultiNul%
	set "officepath3=%installpath16%"
	set "checkversionpath=%CommonProgramFiles(x86)%"
	set "checkversionpath=%checkversionpath:\=\\%"
	>%temp%\result cscript "OfficeFixes\KMS Helper.vbs" "/DATA_FILE" "%checkversionpath%\\Microsoft Shared\\OFFICE16\\MSO.dll"
	for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set o16version=%%g
	goto:eof
::===============================================================================================================
:CheckOfficeVL32W32orVL64W64
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0011-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProPlusVLFound=%%B") %MultiNul%
	if "%ProPlusVLFound:~-39%" EQU "Microsoft Office Professional Plus 2016" ((set "ProPlusVLFound=YES")&&(set "_ProPlusRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0012-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "StandardVLFound=%%B") %MultiNul%
	if "%StandardVLFound:~-30%" EQU "Microsoft Office Standard 2016" ((set "StandardVLFound=YES")&&(set "_StandardRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-003B-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "ProjectProVLFound=%%B") %MultiNul%
	if "%ProjectProVLFound:~-35%" EQU "Microsoft Project Professional 2016" ((set "ProjectProVLFound=YES")&&(set "_ProjectProRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0051-0000-0000-0000000FF1CE" /ve 2^>nul') DO (Set "VisioProVLFound=%%B") %MultiNul%
	if "%VisioProVLFound:~-33%" EQU "Microsoft Visio Professional 2016" ((set "VisioProVLFound=YES")&&(set "_VisioProRetail=YES"))
	if "%_ProPlusRetail%" EQU "YES" (set "o16arch=x86")&&(goto:OfficeVL32VL64Path)
	if "%_StandardRetail%" EQU "YES" (set "o16arch=x86")&&(goto:OfficeVL32VL64Path)
	if "%_ProjectProRetail%" EQU "YES" (set "o16arch=x86")&&(goto:OfficeVL32VL64Path)
	if "%_VisioProRetail%" EQU "YES" (set "o16arch=x86")&&(goto:OfficeVL32VL64Path)
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0011-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "ProPlusVLFound=%%B") %MultiNul%
	if "%ProPlusVLFound:~-39%" EQU "Microsoft Office Professional Plus 2016" ((set "ProPlusVLFound=YES")&&(set "_ProPlusRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0012-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "StandardVLFound=%%B") %MultiNul%
	if "%StandardVLFound:~-30%" EQU "Microsoft Office Standard 2016" ((set "StandardVLFound=YES")&&(set "_StandardRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-003B-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "ProjectProVLFound=%%B") %MultiNul%
	if "%ProjectProVLFound:~-35%" EQU "Microsoft Project Professional 2016" ((set "ProjectProVLFound=YES")&&(set "_ProjectProRetail=YES"))
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\InstalledPackages\90160000-0051-0000-1000-0000000FF1CE" /ve 2^>nul') DO (Set "VisioProVLFound=%%B") %MultiNul%
	if "%VisioProVLFound:~-33%" EQU "Microsoft Visio Professional 2016" ((set "VisioProVLFound=YES")&&(set "_VisioProRetail=YES"))
	if "%_ProPlusRetail%" EQU "YES" (set "o16arch=x64")&&(goto:OfficeVL32VL64Path)
	if "%_StandardRetail%" EQU "YES" (set "o16arch=x64")&&(goto:OfficeVL32VL64Path)
	if "%_ProjectProRetail%" EQU "YES" (set "o16arch=x64")&&(goto:OfficeVL32VL64Path)
	if "%_VisioProRetail%" EQU "YES" (set "o16arch=x64")&&(goto:OfficeVL32VL64Path)
	goto:Office16VnextInstall
::===============================================================================================================
:OfficeVL32VL64Path
	for /F "tokens=2,*" %%A IN ('reg query "HKLM\Software\Microsoft\Office\16.0\Common\InstallRoot" /v "Path" 2^>nul') DO (Set "installpath16=%%B") %MultiNul%
	set "officepath3=%installpath16%"
	set "checkversionpath=%CommonProgramFiles%"
	set "checkversionpath=%checkversionpath:\=\\%"
	>%temp%\result cscript "OfficeFixes\KMS Helper.vbs" "/DATA_FILE" "%checkversionpath%\\Microsoft Shared\\Office16\\MSO.dll"
	for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set o16version=%%g
	goto:eof
::===============================================================================================================
::===============================================================================================================
:Convert16Activate
::===============================================================================================================
	
	call :CheckOfficeApplications
::===============================================================================================================
	cls
	echo:
	call :PrintTitle "================== CONVERT / CHANGE OFFICE TO VOLUME ========================="
	echo:
	echo Installation path:
	echo "%installpath16%"
	echo ____________________________________________________________________________
	echo:
	echo Office Suites:
	set /a countx=0
	echo:
	if "%_ProPlusRetail%" EQU "YES" 				   ((echo Office Professional Plus 2016              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ProPlusVolume%" EQU "YES" 				   ((echo Office Professional Plus 2016              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ProPlus2019Retail%" EQU "YES" 			   ((echo Office Professional Plus 2019              = "FOUND")&&(set /a countx=!countx! + 1)) else if "%_ProPlus2019Volume%" EQU "YES" (
													    (echo Office Professional Plus 2019              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ProPlus2021Retail%" EQU "YES" 			   ((echo Office Professional Plus 2021              = "FOUND")&&(set /a countx=!countx! + 1)) else if "%_ProPlus2021Volume%" EQU "YES" (
														(echo Office Professional Plus 2021              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ProPlusSPLA2021Volume%" EQU "YES" 		   ((echo Office Professional Plus 2021              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_StandardRetail%" EQU "YES" 				   ((echo Office Standard 2016                       = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_StandardVolume%" EQU "YES" 				   ((echo Office Standard 2016                       = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Standard2019Retail%" EQU "YES" 			   ((echo Office Standard 2019                       = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Standard2019Volume%" EQU "YES" 			   ((echo Office Standard 2019                       = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Standard2021Retail%" EQU "YES" 			   ((echo Office Standard 2021                       = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Standard2021Volume%" EQU "YES" 			   ((echo Office Standard 2021                       = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_StandardSPLA2021Volume%" EQU "YES" 		   ((echo Office Standard 2021                       = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_O365ProPlusRetail%" EQU "YES" 			   ((echo Microsoft 365 Apps for Enterprise          = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_O365BusinessRetail%" EQU "YES" 			   ((echo Microsoft 365 Apps for Business            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_O365HomePremRetail%" EQU "YES" 			   ((echo Microsoft 365 Home Premium retail          = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_O365SmallBusPremRetail%" EQU "YES" 		   ((echo Microsoft 365 Small Business retail        = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ProfessionalRetail%" EQU "YES" 			   ((echo Professional 2016 Retail                   = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Professional2019Retail%" EQU "YES" 		   ((echo Professional 2019 Retail                   = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Professional2021Retail%" EQU "YES" 		   ((echo Professional 2021 Retail                   = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_HomeBusinessRetail%" EQU "YES" 			   ((echo Microsoft Home And Business                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_HomeBusiness2019Retail%" EQU "YES" 		   ((echo Microsoft Home And Business 2019           = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_HomeBusiness2021Retail%" EQU "YES" 		   ((echo Microsoft Home And Business 2021           = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_HomeStudentRetail%" EQU "YES" 			   ((echo Microsoft Home And Student                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_HomeStudent2019Retail%" EQU "YES" 		   ((echo Microsoft Home And Student 2019            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_HomeStudent2021Retail%" EQU "YES" 		   ((echo Microsoft Home And Student 2021            = "FOUND")&&(set /a countx=!countx! + 1))	
	if "%_MondoRetail%" EQU "YES" 					   ((echo Office Mondo Grande Suite                  = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_MondoVolume%" EQU "YES" 					   ((echo Office Mondo Grande Suite                  = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PersonalRetail%" EQU "YES" 				   ((echo Office Personal 2016 Retail                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Personal2019Retail%" EQU "YES" 			   ((echo Office Personal 2019 Retail                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Personal2021Retail%" EQU "YES" 			   ((echo Office Personal 2021 Retail                = "FOUND")&&(set /a countx=!countx! + 1))
	if !countx! EQU 0 									(echo Office Full Suite installation             = "NOT FOUND")
	echo ____________________________________________________________________________
	echo:
	echo Additional Apps:
	set /a countx=0
	if "%_VisioProRetail%" EQU "YES" ((echo:)&&			(echo Visio Pro 2016                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxVisio%" EQU "YES" ((echo:)&&				(echo Visio Pro UWP Appx                         = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_VisioPro2019Retail%" EQU "YES" ((echo:)&&		(echo Visio Pro 2019                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_VisioPro2019Volume%" EQU "YES" ((echo:)&&		(echo Visio Pro 2019                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_VisioPro2021Retail%" EQU "YES" ((echo:)&&		(echo Visio Pro 2021                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_VisioPro2021Volume%" EQU "YES" ((echo:)&&		(echo Visio Pro 2021                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ProjectProRetail%" EQU "YES" ((echo:)&&		(echo Project Pro 2016                           = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_AppxProject%" EQU "YES" ((echo:)&&			(echo Project Pro UWP Appx                       = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectPro2019Retail%" EQU "YES" ((echo:)&&	(echo Project Pro 2019                           = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectPro2019Volume%" EQU "YES" ((echo:)&&	(echo Project Pro 2019                           = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectPro2021Retail%" EQU "YES" ((echo:)&&	(echo Project Pro 2021                           = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectPro2021Volume%" EQU "YES" ((echo:)&&	(echo Project Pro 2021                           = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_VisioStdRetail%" EQU "YES" ((echo:)&&			(echo Visio Standard 2016                        = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_VisioStdVolume%" EQU "YES" ((echo:)&&			(echo Visio Standard 2016                        = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_VisioStdXVolume%" EQU "YES" ((echo:)&&		(echo Visio Standard 2016 C2R                    = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_VisioStd2019Retail%" EQU "YES" ((echo:)&&		(echo Visio Standard 2019                        = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_VisioStd2019Volume%" EQU "YES" ((echo:)&&		(echo Visio Standard 2019                        = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_VisioStd2021Retail%" EQU "YES" ((echo:)&&		(echo Visio Standard 2021                        = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_VisioStd2021Volume%" EQU "YES" ((echo:)&&		(echo Visio Standard 2021                        = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectStdRetail%" EQU "YES" ((echo:)&&		(echo Project Standard 2016                      = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectStdVolume%" EQU "YES" ((echo:)&&		(echo Project Standard 2016                      = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectStdXVolume%" EQU "YES" ((echo:)&&		(echo Project Standard 2016 C2R                  = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectProXVolume%" EQU "YES" ((echo:)&&		(echo Project Professional 2016 C2R              = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_VisioProXVolume%" EQU "YES" ((echo:)&&		(echo Visio Professional 2016 C2R                = "FOUND")&&(set /a countx=!countx! + 2))	
	if "%_ProjectStd2019Retail%" EQU "YES" ((echo:)&&	(echo Project Standard 2019                      = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectStd2019Volume%" EQU "YES" ((echo:)&&	(echo Project Standard 2019                      = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectStd2021Retail%" EQU "YES" ((echo:)&&	(echo Project Standard 2021                      = "FOUND")&&(set /a countx=!countx! + 2))
	if "%_ProjectStd2021Volume%" EQU "YES" ((echo:)&&	(echo Project Standard 2021                      = "FOUND")&&(set /a countx=!countx! + 2))
	if !countx! EQU 0 ((echo:)&&						(echo Visio and Project Installation             = "NOT FOUND"))
	echo ____________________________________________________________________________
	echo:
	echo Single Apps:
	set /a countx=0
	if "%_WordRetail%" EQU "YES" ((echo:)&&				(echo Word 2016                                  = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_ExcelRetail%" EQU "YES" ((echo:)&&			(echo Excel 2016                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PowerPointRetail%" EQU "YES" ((echo:)&&		(echo PowerPoint 2016                            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AccessRetail%" EQU "YES" ((echo:)&&			(echo Access 2016                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_OutlookRetail%" EQU "YES" ((echo:)&&			(echo Outlook 2016                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PublisherRetail%" EQU "YES" ((echo:)&&		(echo Publisher 2016                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_OneNoteRetail%" EQU "YES" ((echo:)&&			(echo OneNote 2016                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_OneNoteVolume%" EQU "YES" ((echo:)&&			(echo OneNote 2016                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_OneNote2021Retail%" EQU "YES" ((echo:)&&		(echo OneNote 2021                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_SkypeForBusinessRetail%" EQU "YES" ((echo:)&&	(echo Skype 2016                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxWinword%" EQU "YES" ((echo:)&&			(echo Word UWP Appx                              = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxExcel%" EQU "YES" ((echo:)&&				(echo Excel UWP Appx                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxPowerPoint%" EQU "YES" ((echo:)&&			(echo PowerPoint UWP Appx                        = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxAccess%" EQU "YES" ((echo:)&&				(echo Access UWP Appx                            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxOutlook%" EQU "YES" ((echo:)&&			(echo Outlook UWP Appx                           = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxPublisher%" EQU "YES" ((echo:)&&			(echo Publisher UWP Appx                         = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxOneNote%" EQU "YES" ((echo:)&&			(echo OneNote UWP Appx                           = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_AppxSkypeForBusiness%" EQU "YES" ((echo:)&&	(echo Skype UWP Appx                             = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Word2019Retail%" EQU "YES" ((echo:)&&			(echo Word 2019                                  = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Excel2019Retail%" EQU "YES" ((echo:)&&		(echo Excel 2019                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PowerPoint2019Retail%" EQU "YES" ((echo:)&&	(echo PowerPoint 2019                            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Access2019Retail%" EQU "YES" ((echo:)&&		(echo Access 2019                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Outlook2019Retail%" EQU "YES" ((echo:)&&		(echo Outlook 2019                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Publisher2019Retail%" EQU "YES" ((echo:)&&	(echo Publisher 2019                             = "FOUND")&&(set /a countx=!countx! + 1))
if "%_SkypeForBusiness2019Retail%" EQU "YES" ((echo:)&& (echo Skype 2019                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Word2019Volume%" EQU "YES" ((echo:)&&			(echo Word 2019                                  = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Excel2019Volume%" EQU "YES" ((echo:)&&		(echo Excel 2019                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PowerPoint2019Volume%" EQU "YES" ((echo:)&&	(echo PowerPoint 2019                            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Access2019Volume%" EQU "YES" ((echo:)&&		(echo Access 2019                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Outlook2019Volume%" EQU "YES" ((echo:)&&		(echo Outlook 2019                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Publisher2019Volume%" EQU "YES" ((echo:)&&	(echo Publisher 2019                             = "FOUND")&&(set /a countx=!countx! + 1))
if "%_SkypeForBusiness2019Volume%" EQU "YES" ((echo:)&&	(echo Skype 2019                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Word2021Retail%" EQU "YES" ((echo:)&&			(echo Word 2021                                  = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Excel2021Retail%" EQU "YES" ((echo:)&&		(echo Excel 2021                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PowerPoint2021Retail%" EQU "YES" ((echo:)&&	(echo PowerPoint 2021                            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Access2021Retail%" EQU "YES" ((echo:)&&		(echo Access 2021                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Outlook2021Retail%" EQU "YES" ((echo:)&&		(echo Outlook 2021                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Publisher2021Retail%" EQU "YES" ((echo:)&&	(echo Publisher 2021                             = "FOUND")&&(set /a countx=!countx! + 1))
if "%_SkypeForBusiness2021Retail%" EQU "YES" ((echo:)&&	(echo Skype 2021                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Word2021Volume%" EQU "YES" ((echo:)&&			(echo Word 2021                                  = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Excel2021Volume%" EQU "YES" ((echo:)&&		(echo Excel 2021                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_PowerPoint2021Volume%" EQU "YES" ((echo:)&&	(echo PowerPoint 2021                            = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Access2021Volume%" EQU "YES" ((echo:)&&		(echo Access 2021                                = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Outlook2021Volume%" EQU "YES" ((echo:)&&		(echo Outlook 2021                               = "FOUND")&&(set /a countx=!countx! + 1))
	if "%_Publisher2021Volume%" EQU "YES" ((echo:)&&	(echo Publisher 2021                             = "FOUND")&&(set /a countx=!countx! + 1))
if "%_SkypeForBusiness2021Volume%" EQU "YES" ((echo:)&&	(echo Skype 2021                                 = "FOUND")&&(set /a countx=!countx! + 1))
	if !countx! EQU 0 ((echo:)&&						(echo Single Apps installation                   = "NOT FOUND"))
	echo ____________________________________________________________________________
	echo:
	echo:
	if not defined debugMode pause
::===============================================================================================================
	cls
	echo:
	call :PrintTitle "==== CLEANUP (Removing Office Retail-Trial-Grace Keys and Licenses) ====="
	echo:
	"%OfficeRToolpath%\OfficeFixes\%winx%\cleanospp.exe" -PKey
	echo ____________________________________________________________________________
	echo:
	"%OfficeRToolpath%\OfficeFixes\%winx%\cleanospp.exe" -Licenses
	echo ____________________________________________________________________________
	echo:
	echo:
::===============================================================================================================
::	WIN10 Insider/Preview activation failure workaround
	if exist "%windir%\System32\spp\store_test\2.0\tokens.dat"	(
		echo:
		echo ____________________________________________________________________________
		echo:
		echo Windows 10 Insider detected. Workaround for KMS activation needed^!
		echo:
		echo Stopping license service "sppsvc"...
		call :StopService "sppsvc"
		echo:
		echo Deleting license file "tokens.dat"...
		del "%windir%\System32\spp\store_test\2.0\tokens.dat" %MultiNul%
		echo:
		echo Starting license service "sppsvc"...
		call :StartService "sppsvc"
		echo:
		echo Recreating license file "tokens.dat". Please wait...
		cscript //Nologo //B %SystemRoot%\System32\slmgr.vbs /rilc
		cscript //Nologo //B %SystemRoot%\System32\slmgr /ato
		echo:
		echo License file successfully recreated
		echo ____________________________________________________________________________
		echo:
		echo:
	)
	timeout /t 4
::===============================================================================================================
	cls
	echo:
	call :PrintTitle "================== CONVERT / CHANGE OFFICE TO VOLUME ========================="
	echo:
::===============================================================================================================
	call :Office16ConversionLoop
::===============================================================================================================	
	del "%TEMP%\ONAME_CHANGE*.REG" %MultiNul%
	
::	Change name of installed Office 2019 Retail products from Retail to Volume
	if "%_Word2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_Excel2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_PowerPoint2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_Access2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_Outlook2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_Publisher2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_SkypeForBusiness2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_ProPlus2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_VisioPro2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_ProjectPro2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_ProjectStd2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_VisioStd2019Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	
	if exist "%TEMP%\ONAME_CHANGE1.REG"	(
		powershell -noprofile -command "& {Get-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE1.REG" | ForEach-Object { $_ -replace '2019Retail', '2019Volume' } | Set-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE2.REG"}" %MultiNul%
		reg delete HKLM\Software\Microsoft\Office\ClickToRun\Configuration /f %MultiNul%
		reg import "%TEMP%\ONAME_CHANGE2.REG" %MultiNul%
		del "%TEMP%\ONAME_CHANGE*.REG" %MultiNul%
	)
	
::	Change name of installed Office 2021 Retail products from Retail to Volume
	if "%_VisioStd2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_ProjectStd2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_Word2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_Excel2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_PowerPoint2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_Access2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_Outlook2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_Publisher2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_SkypeForBusiness2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_ProPlus2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_VisioPro2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	if "%_ProjectPro2021Retail%" EQU "YES" reg export HKLM\Software\Microsoft\Office\ClickToRun\Configuration "%TEMP%\ONAME_CHANGE1.REG" /Y %MultiNul%
	
	if exist "%TEMP%\ONAME_CHANGE1.REG"	(
		powershell -noprofile -command "& {Get-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE1.REG" | ForEach-Object { $_ -replace '2021Retail', '2021Volume' } | Set-Content -Encoding Unicode "%TEMP%\ONAME_CHANGE2.REG"}" %MultiNul%
		reg delete HKLM\Software\Microsoft\Office\ClickToRun\Configuration /f %MultiNul%
		reg import "%TEMP%\ONAME_CHANGE2.REG" %MultiNul%
		del "%TEMP%\ONAME_CHANGE*.REG" %MultiNul%
	)
	
::===============================================================================================================
	goto:Office16VnextInstall
::===============================================================================================================
::===============================================================================================================
:Office16Activate
	set /a "GraceMin=0"
	if %win% GEQ 9200 (
		set "ID=%1"
		set "subKey=0ff1ce15-a989-479d-af46-f275c6370663"
		call :UpdateRegistryKeys %KMSHostIP% %KMSPort%
		
		set "lastErr="
		set "activationCMD=cscript //nologo "OfficeFixes\KMS Helper.vbs" "/ACTIVATE" "%slp%" "%1""
		
		REM call :Query "GracePeriodRemaining" "%slp%" "ID Like '%%%%%1%%%%'"
		REM for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set GraceMin=%%g
		REM if !GraceMin! EQU 259200 (echo Activation !ID! successful) else (echo Activation !ID! failed)
	)
	if %win% LSS 9200 (
		set "ID=%1"
		set "subKey=0ff1ce15-a989-479d-af46-f275c6370663"
		call :UpdateRegistryKeys %KMSHostIP% %KMSPort%
		
		set "lastErr="
		set "activationCMD=cscript "OfficeFixes\KMS Helper.vbs" "/ACTIVATE" "%ospp%" "%1""
		
		REM call :Query "GracePeriodRemaining" "%ospp%" "ID Like '%%%%%1%%%%'"
		REM for /f "tokens=1 skip=3 delims=," %%g in ('type "%temp%\result"') do set GraceMin=%%g
		REM if !GraceMin! EQU 259200 (echo Activation !ID! successful) else (echo Activation !ID! failed)
	)
	
	for /f "tokens=1,2 delims=: " %%x in ('"!activationCMD!"') do set "lastErr=%%y"
	if /i '!lastErr!' EQU '0' (echo Activation !ID! successful) else (echo Activation !ID! failed & echo Error Number !lastErr!)
	call :CleanRegistryKeys
	echo ________________________________________________________________
	echo:
	goto:eof
::===============================================================================================================
::===============================================================================================================
:SetO16Language

	%MultiNul% del /q "%temp%\tmp"
	set langnotfound=***
	>"%temp%\tmp" call :Language_List
	
	rem %%g=English %%h=1033 %%i=en-us %%j:0409	
	for /f "tokens=1,2,3,4 delims=*" %%g in ('type "%temp%\tmp"') do (
		
		if /i '!o16lang!' EQU '%%i' (
			set langtext=%%g
			set o16lcid=%%h
			set langnotfound=
		)
		
		if /i '!o16lang!' EQU '%%g' (
			set langtext=%%g
			set o16lcid=%%h
			set o16lang=%%i
			set langnotfound=
		)
	)
	
	%MultiNul% del /q "%temp%\tmp"
    goto:eof
::===============================================================================================================
::===============================================================================================================
:ConvertOffice16
	cls
	echo:
	call :PrintTitle "================= %1 found ========================================"
	echo:
	
	echo %1 | %SingleNul% find /i "mondo" && (
		set "root="
		if exist "%ProgramFiles%\Microsoft Office\root"			set "root=%ProgramFiles%\Microsoft Office\root"
		if exist "%ProgramFiles(x86)%\Microsoft Office\root"	set "root=%ProgramFiles(x86)%\Microsoft Office\root"
		if defined root (
			if exist "!root!\Integration\integrator.exe" (
				"!root!\Integration\integrator" /I /License PRIDName=MondoVolume.16 PidKey=HFTND-W9MK4-8B7MJ-B6C4G-XQBR2
			)
		)
	)
	
	if "%3" EQU "_AE2" goto :ConvertOffice2021_AE2
::================================================================================================================
	if %win% GEQ 9200 (
		echo %1 |%SingleNul% find /i "C2R" && (
			cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_ClientC2R-ul.xrm-ms"
			cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_ClientC2R-ul-oob.xrm-ms"
			cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_ClientC2R-ppd.xrm-ms"
			
			REM cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAKC2R-pl.xrm-ms"
			REM cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAKC2R-ul-phn.xrm-ms"
			REM cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAKC2R-ul-oob.xrm-ms"
			REM cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAKC2R-ppd.xrm-ms"
		) || (
			cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul.xrm-ms"
			cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul-oob.xrm-ms"
			cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ppd.xrm-ms"
			
			REM cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-pl.xrm-ms"
			REM cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ul-phn.xrm-ms"
			REM cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ul-oob.xrm-ms"
			REM cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ppd.xrm-ms"
		)
		
	)
	if %win% LSS 9200 (
		echo %1 |%SingleNul% find /i "C2R" && (
			cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_ClientC2R-ul.xrm-ms"
			cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_ClientC2R-ul-oob.xrm-ms"
			cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_ClientC2R-ppd.xrm-ms"
			
			REM cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAKC2R-pl.xrm-ms"
			REM cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAKC2R-ul-phn.xrm-ms"
			REM cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAKC2R-ul-oob.xrm-ms"
			REM cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAKC2R-ppd.xrm-ms"
		) || (
			cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul.xrm-ms"
			cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul-oob.xrm-ms"
			cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ppd.xrm-ms"
			
			REM cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-pl.xrm-ms"
			REM cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ul-phn.xrm-ms"
			REM cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ul-oob.xrm-ms"
			REM cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK%2-ppd.xrm-ms"
		)
    )
    echo ____________________________________________________________________________
	echo:
	echo:
	timeout /t 4
	goto:eof
::===============================================================================================================
:ConvertOffice2021_AE2
	if %win% GEQ 9200 (
	
		cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul.xrm-ms"
		cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul-oob.xrm-ms"
		cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ppd.xrm-ms"
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE-pl.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE-ul-phn.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE-ul-oob.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE-ppd.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE1-pl.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE1-ul-phn.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE1-ul-oob.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE1-ppd.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE2-pl.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE2-ul-phn.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE2-ul-oob.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE2-ppd.xrm-ms"
		REM if exist !licenseFile! cscript "%windir%\system32\slmgr.vbs" /ilc !licenseFile!
		
	)
	if %win% LSS 9200 (
	
		cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul.xrm-ms"
		cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ul-oob.xrm-ms"
		cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_KMS_Client%2-ppd.xrm-ms"
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE-pl.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE-ul-phn.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE-ul-oob.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE-ppd.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE1-pl.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE1-ul-phn.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE1-ul-oob.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE1-ppd.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE2-pl.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE2-ul-phn.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE2-ul-oob.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
		
		REM set licenseFile="%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\%1VL_MAK_AE2-ppd.xrm-ms"
		REM if exist !licenseFile! cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:!licenseFile!
	)
    echo ____________________________________________________________________________
	echo:
	echo:
	timeout /t 4
	goto:eof
::===============================================================================================================
::===============================================================================================================
:ConvertGeneral16
	cls
	echo:
	call :PrintTitle "================= Office General Client found =============================="
	echo:
	if %win% GEQ 9200    (    
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\pkeyconfig-office.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-stil.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul-oob.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root-bridge-test.xrm-ms"
	cscript "%windir%\system32\slmgr.vbs" /ilc "%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-bridge-office.xrm-ms"
	)
	if %win% LSS 9200    (    
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\pkeyconfig-office.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-stil.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-ul-oob.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-root-bridge-test.xrm-ms"
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inslic:"%OfficeRToolpath%\OfficeFixes\ospp\Licenses16\client-issuance-bridge-office.xrm-ms"
	)
	echo ____________________________________________________________________________
	echo:
	echo:
	timeout /t 4
	goto:eof
::===============================================================================================================
::===============================================================================================================
:Office16ConversionLoop
	call :ConvertGeneral16
	if "%_ProPlusRetail%" EQU "YES" call :ConvertOffice16 ProPlus
	if "%_ProPlusVolume%" EQU "YES" call :ConvertOffice16 ProPlus
	if "%_ProPlus2019Retail%" EQU "YES" call :ConvertOffice16 ProPlus2019 _AE
	if "%_ProPlus2019Volume%" EQU "YES" call :ConvertOffice16 ProPlus2019 _AE
	if "%_ProPlus2021Retail%" EQU "YES" call :ConvertOffice16 ProPlus2021 _AE _AE2
	if "%_ProPlus2021Volume%" EQU "YES" call :ConvertOffice16 ProPlus2021 _AE _AE2
	if "%_ProPlusSPLA2021Volume%" EQU "YES" call :ConvertOffice16 ProPlus2021 _AE _AE2
	if "%_O365ProPlusRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_O365BusinessRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_O365HomePremRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_O365SmallBusPremRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_ProfessionalRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_Professional2019Retail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_Professional2021Retail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_HomeBusinessRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_HomeBusiness2019Retail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_HomeBusiness2021Retail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_HomeStudentRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_HomeStudent2019Retail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_HomeStudent2021Retail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_MondoRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_MondoVolume%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_PersonalRetail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_Personal2019Retail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_Personal2021Retail%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_UWPappINSTALLED%" EQU "YES" call :ConvertOffice16 Mondo
	if "%_StandardRetail%" EQU "YES" call :ConvertOffice16 Standard
	if "%_StandardVolume%" EQU "YES" call :ConvertOffice16 Standard
	if "%_Standard2019Retail%" EQU "YES" call :ConvertOffice16 Standard2019 _AE
	if "%_Standard2019Volume%" EQU "YES" call :ConvertOffice16 Standard2019 _AE
	if "%_Standard2021Retail%" EQU "YES" call :ConvertOffice16 Standard2021 _AE _AE2
	if "%_Standard2021Volume%" EQU "YES" call :ConvertOffice16 Standard2021 _AE _AE2
	if "%_StandardSPLA2021Volume%" EQU "YES" call :ConvertOffice16 Standard2021 _AE _AE2
	if "%_WordRetail%" EQU "YES" call :ConvertOffice16 Word
	if "%_ExcelRetail%" EQU "YES" call :ConvertOffice16 Excel
	if "%_PowerPointRetail%" EQU "YES" call :ConvertOffice16 PowerPoint
	if "%_AccessRetail%" EQU "YES" call :ConvertOffice16 Access
	if "%_OutlookRetail%" EQU "YES" call :ConvertOffice16 Outlook
	if "%_PublisherRetail%" EQU "YES" call :ConvertOffice16 Publisher
	if "%_OneNoteRetail%" EQU "YES" call :ConvertOffice16 OneNote
	if "%_OneNoteVolume%" EQU "YES" call :ConvertOffice16 OneNote
	if "%_OneNote2021Retail%" EQU "YES" call :ConvertOffice16 OneNote
	if "%_SkypeForBusinessRetail%" EQU "YES" call :ConvertOffice16 SkypeForBusiness
	if "%_Word2019Retail%" EQU "YES" call :ConvertOffice16 Word2019 _AE
	if "%_Word2021Retail%" EQU "YES" call :ConvertOffice16 Word2021 _AE _AE2
	if "%_Excel2019Retail%" EQU "YES" call :ConvertOffice16 Excel2019 _AE
	if "%_Excel2021Retail%" EQU "YES" call :ConvertOffice16 Excel2021 _AE _AE2
	if "%_PowerPoint2019Retail%" EQU "YES" call :ConvertOffice16 PowerPoint2019 _AE
	if "%_PowerPoint2021Retail%" EQU "YES" call :ConvertOffice16 PowerPoint2021 _AE _AE2
	if "%_Access2019Retail%" EQU "YES" call :ConvertOffice16 Access2019 _AE
	if "%_Access2021Retail%" EQU "YES" call :ConvertOffice16 Access2021 _AE _AE2
	if "%_Outlook2019Retail%" EQU "YES" call :ConvertOffice16 Outlook2019 _AE
	if "%_Outlook2021Retail%" EQU "YES" call :ConvertOffice16 Outlook2021 _AE _AE2
	if "%_Publisher2019Retail%" EQU "YES" call :ConvertOffice16 Publisher2019 _AE
	if "%_Publisher2021Retail%" EQU "YES" call :ConvertOffice16 Publisher2021 _AE _AE2
	if "%_SkypeForBusiness2019Retail%" EQU "YES" call :ConvertOffice16 SkypeForBusiness2019 _AE
	if "%_SkypeForBusiness2021Retail%" EQU "YES" call :ConvertOffice16 SkypeForBusiness2021 _AE _AE2
	if "%_Word2019Volume%" EQU "YES" call :ConvertOffice16 Word2019 _AE
	if "%_Word2021Volume%" EQU "YES" call :ConvertOffice16 Word2021 _AE _AE2
	if "%_Excel2019Volume%" EQU "YES" call :ConvertOffice16 Excel2019 _AE
	if "%_Excel2021Volume%" EQU "YES" call :ConvertOffice16 Excel2021 _AE _AE2
	if "%_PowerPoint2019Volume%" EQU "YES" call :ConvertOffice16 PowerPoint2019 _AE
	if "%_PowerPoint2021Volume%" EQU "YES" call :ConvertOffice16 PowerPoint2021 _AE _AE2
	if "%_Access2019Volume%" EQU "YES" call :ConvertOffice16 Access2019 _AE
	if "%_Access2021Volume%" EQU "YES" call :ConvertOffice16 Access2021 _AE _AE2
	if "%_Outlook2019Volume%" EQU "YES" call :ConvertOffice16 Outlook2019 _AE
	if "%_Outlook2021Volume%" EQU "YES" call :ConvertOffice16 Outlook2021 _AE _AE2
	if "%_Publisher2019Volume%" EQU "YES" call :ConvertOffice16 Publisher2019 _AE
	if "%_Publisher2021Volume%" EQU "YES" call :ConvertOffice16 Publisher2021 _AE _AE2
	if "%_SkypeForBusiness2019Volume%" EQU "YES" call :ConvertOffice16 SkypeForBusiness2019 _AE
	if "%_SkypeForBusiness2021Volume%" EQU "YES" call :ConvertOffice16 SkypeForBusiness2021 _AE _AE2
	if "%_VisioProRetail%" EQU "YES" call :ConvertOffice16 VisioPro
	if "%_AppxVisio%" EQU "YES" call :ConvertOffice16 VisioPro
	if "%_VisioPro2019Retail%" EQU "YES" call :ConvertOffice16 VisioPro2019 _AE
	if "%_VisioPro2019Volume%" EQU "YES" call :ConvertOffice16 VisioPro2019 _AE
	if "%_VisioPro2021Retail%" EQU "YES" call :ConvertOffice16 VisioPro2021 _AE _AE2
	if "%_VisioPro2021Volume%" EQU "YES" call :ConvertOffice16 VisioPro2021 _AE _AE2
	if "%_AppxProject%" EQU "YES" call :ConvertOffice16 ProjectPro
	if "%_ProjectProRetail%" EQU "YES" call :ConvertOffice16 ProjectPro
	if "%_ProjectPro2019Retail%" EQU "YES" call :ConvertOffice16 ProjectPro2019 _AE
	if "%_ProjectPro2019Volume%" EQU "YES" call :ConvertOffice16 ProjectPro2019 _AE
	if "%_ProjectPro2021Retail%" EQU "YES" call :ConvertOffice16 ProjectPro2021 _AE _AE2
	if "%_ProjectPro2021Volume%" EQU "YES" call :ConvertOffice16 ProjectPro2021 _AE _AE2
	if "%_VisioStdRetail%" EQU "YES" call :ConvertOffice16 VisioStd
	if "%_VisioStdVolume%" EQU "YES" call :ConvertOffice16 VisioStd
	if "%_VisioStdXVolume%" EQU "YES" call :ConvertOffice16 VisioStdXC2R
	if "%_VisioStd2019Retail%" EQU "YES" call :ConvertOffice16 VisioStd2019 _AE
	if "%_VisioStd2019Volume%" EQU "YES" call :ConvertOffice16 VisioStd2019 _AE
	if "%_VisioStd2021Retail%" EQU "YES" call :ConvertOffice16 VisioStd2021 _AE _AE2
	if "%_VisioStd2021Volume%" EQU "YES" call :ConvertOffice16 VisioStd2021 _AE _AE2
	if "%_ProjectStdRetail%" EQU "YES" call :ConvertOffice16 ProjectStd
	if "%_ProjectStdVolume%" EQU "YES" call :ConvertOffice16 ProjectStd
	if "%_ProjectStdXVolume%" EQU "YES" call :ConvertOffice16 ProjectStdXC2R
	if "%_ProjectProXVolume%" EQU "YES" call :ConvertOffice16 ProjectProXC2R
	if "%_VisioProXVolume%" EQU "YES" call :ConvertOffice16 VisioProXC2R
	if "%_ProjectStd2019Retail%" EQU "YES" call :ConvertOffice16 ProjectStd2019 _AE
	if "%_ProjectStd2019Volume%" EQU "YES" call :ConvertOffice16 ProjectStd2019 _AE
	if "%_ProjectStd2021Retail%" EQU "YES" call :ConvertOffice16 ProjectStd2021 _AE _AE2
	if "%_ProjectStd2021Volume%" EQU "YES" call :ConvertOffice16 ProjectStd2021 _AE _AE2
	
	cls
	echo:
	call :PrintTitle "================= INSTALLING GVLK =========================================="
	echo:
	if %win% GEQ 9200 if "%_ProPlusRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office Professional Plus 2016"
	if %win% LSS 9200 if "%_ProPlusRetail%" EQU "YES" call :OfficeGVLKInstall "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office Professional Plus 2016"
	if %win% GEQ 9200 if "%_ProPlusVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office Professional Plus 2016"
	if %win% LSS 9200 if "%_ProPlusVolume%" EQU "YES" call :OfficeGVLKInstall "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99","Office Professional Plus 2016"
	if %win% GEQ 9200 if "%_ProPlus2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP","Office Professional Plus 2019"
	if %win% LSS 9200 if "%_ProPlus2019Retail%" EQU "YES" call :OfficeGVLKInstall "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP","Office Professional Plus 2019"
	if %win% GEQ 9200 if "%_ProPlus2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP","Office Professional Plus 2019"
	if %win% LSS 9200 if "%_ProPlus2019Volume%" EQU "YES" call :OfficeGVLKInstall "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP","Office Professional Plus 2019"
	if %win% GEQ 9200 if "%_ProPlus2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH","Office Professional Plus 2021"
	if %win% LSS 9200 if "%_ProPlus2021Retail%" EQU "YES" call :OfficeGVLKInstall "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH","Office Professional Plus 2021"
	if %win% GEQ 9200 if "%_ProPlus2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH","Office Professional Plus 2021"
	if %win% LSS 9200 if "%_ProPlus2021Volume%" EQU "YES" call :OfficeGVLKInstall "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH","Office Professional Plus 2021"
	if %win% GEQ 9200 if "%_ProPlusSPLA2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH","Office Professional Plus 2021"
	if %win% LSS 9200 if "%_ProPlusSPLA2021Volume%" EQU "YES" call :OfficeGVLKInstall "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH","Office Professional Plus 2021"
	if %win% GEQ 9200 if "%_O365ProPlusRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft 365 Apps for Enterprise"
	if %win% LSS 9200 if "%_O365ProPlusRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft 365 Apps for Enterprise"
	if %win% GEQ 9200 if "%_O365BusinessRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft 365 Apps for Business"
	if %win% LSS 9200 if "%_O365BusinessRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft 365 Apps for Business"
	if %win% GEQ 9200 if "%_O365HomePremRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft 365 Home Premium retail"
	if %win% LSS 9200 if "%_O365HomePremRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft 365 Home Premium retail"
	if %win% GEQ 9200 if "%_O365SmallBusPremRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft 365 Small Business retail"
	if %win% LSS 9200 if "%_O365SmallBusPremRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft 365 Small Business retail"
	if %win% GEQ 9200 if "%_ProfessionalRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Professional 2016 Retail"
	if %win% LSS 9200 if "%_ProfessionalRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Professional 2016 Retail"
	if %win% GEQ 9200 if "%_Professional2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Professional 2019 Retail"
	if %win% LSS 9200 if "%_Professional2019Retail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Professional 2019 Retail"
	if %win% GEQ 9200 if "%_Professional2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Professional 2021 Retail"
	if %win% LSS 9200 if "%_Professional2021Retail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Professional 2021 Retail"
	if %win% GEQ 9200 if "%_HomeBusinessRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Business"
	if %win% LSS 9200 if "%_HomeBusinessRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Business"
	if %win% GEQ 9200 if "%_HomeBusiness2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Business 2019"
	if %win% LSS 9200 if "%_HomeBusiness2019Retail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Business 2019"	
	if %win% GEQ 9200 if "%_HomeBusiness2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Business 2021"
	if %win% LSS 9200 if "%_HomeBusiness2021Retail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Business 2021"
	if %win% GEQ 9200 if "%_HomeStudentRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Student"
	if %win% LSS 9200 if "%_HomeStudentRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Student"
	if %win% GEQ 9200 if "%_HomeStudent2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Student 2019"
	if %win% LSS 9200 if "%_HomeStudent2019Retail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Student 2019"	
	if %win% GEQ 9200 if "%_HomeStudent2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Student 2021"
	if %win% LSS 9200 if "%_HomeStudent2021Retail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Microsoft Home And Student 2021"
	if %win% GEQ 9200 if "%_MondoRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Mondo 2016 Grande Suite"
	if %win% LSS 9200 if "%_MondoRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Mondo 2016 Grande Suite"
	if %win% GEQ 9200 if "%_MondoVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Mondo 2016 Grande Suite"
	if %win% LSS 9200 if "%_MondoVolume%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Mondo 2016 Grande Suite"
	if %win% GEQ 9200 if "%_PersonalRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Personal 2016 Retail"
	if %win% LSS 9200 if "%_PersonalRetail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Personal 2016 Retail"
	if %win% GEQ 9200 if "%_Personal2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Personal 2019 Retail"
	if %win% LSS 9200 if "%_Personal2019Retail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Personal 2019 Retail"
	if %win% GEQ 9200 if "%_Personal2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Personal 2021 Retail"
	if %win% LSS 9200 if "%_Personal2021Retail%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office Personal 2021 Retail"
	if %win% GEQ 9200 if "%_UWPappINSTALLED%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office UWP Appxs"
	if %win% LSS 9200 if "%_UWPappINSTALLED%" EQU "YES" call :OfficeGVLKInstall "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2","Office UWP Appxs"
	if %win% GEQ 9200 if "%_StandardRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"JNRGM-WHDWX-FJJG3-K47QV-DRTFM","Office Standard 2016"
	if %win% LSS 9200 if "%_StandardRetail%" EQU "YES" call :OfficeGVLKInstall "JNRGM-WHDWX-FJJG3-K47QV-DRTFM","Office Standard 2016"
	if %win% GEQ 9200 if "%_StandardVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"JNRGM-WHDWX-FJJG3-K47QV-DRTFM","Office Standard 2016"
	if %win% LSS 9200 if "%_StandardVolume%" EQU "YES" call :OfficeGVLKInstall "JNRGM-WHDWX-FJJG3-K47QV-DRTFM","Office Standard 2016"
	if %win% GEQ 9200 if "%_Standard2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK","Office Standard 2019"
	if %win% LSS 9200 if "%_Standard2019Retail%" EQU "YES" call :OfficeGVLKInstall "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK","Office Standard 2019"
	if %win% GEQ 9200 if "%_Standard2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK","Office Standard 2019"
	if %win% LSS 9200 if "%_Standard2019Volume%" EQU "YES" call :OfficeGVLKInstall "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK","Office Standard 2019"
	if %win% GEQ 9200 if "%_Standard2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3","Office Standard 2021"
	if %win% LSS 9200 if "%_Standard2021Retail%" EQU "YES" call :OfficeGVLKInstall "KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3","Office Standard 2021"
	if %win% GEQ 9200 if "%_Standard2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3","Office Standard 2021"
	if %win% LSS 9200 if "%_Standard2021Volume%" EQU "YES" call :OfficeGVLKInstall "KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3","Office Standard 2021"
	if %win% GEQ 9200 if "%_StandardSPLA2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3","Office Standard 2021"
	if %win% LSS 9200 if "%_StandardSPLA2021Volume%" EQU "YES" call :OfficeGVLKInstall "KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3","Office Standard 2021"
	if %win% GEQ 9200 if "%_WordRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6","Word 2016"
	if %win% LSS 9200 if "%_WordRetail%" EQU "YES" call :OfficeGVLKInstall "WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6","Word 2016"
	if %win% GEQ 9200 if "%_ExcelRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"9C2PK-NWTVB-JMPW8-BFT28-7FTBF","Excel 2016"
	if %win% LSS 9200 if "%_ExcelRetail%" EQU "YES" call :OfficeGVLKInstall "9C2PK-NWTVB-JMPW8-BFT28-7FTBF","Excel 2016"
	if %win% GEQ 9200 if "%_PowerPointRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6","PowerPoint 2016"
	if %win% LSS 9200 if "%_PowerPointRetail%" EQU "YES" call :OfficeGVLKInstall "J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6","PowerPoint 2016"
	if %win% GEQ 9200 if "%_AccessRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW","Access 2016"
	if %win% LSS 9200 if "%_AccessRetail%" EQU "YES" call :OfficeGVLKInstall "GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW","Access 2016"
	if %win% GEQ 9200 if "%_OutlookRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"R69KK-NTPKF-7M3Q4-QYBHW-6MT9B","Outlook 2016"
	if %win% LSS 9200 if "%_OutlookRetail%" EQU "YES" call :OfficeGVLKInstall "R69KK-NTPKF-7M3Q4-QYBHW-6MT9B","Outlook 2016"
	if %win% GEQ 9200 if "%_PublisherRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"F47MM-N3XJP-TQXJ9-BP99D-8K837","Publisher 2016"
	if %win% LSS 9200 if "%_PublisherRetail%" EQU "YES" call :OfficeGVLKInstall "F47MM-N3XJP-TQXJ9-BP99D-8K837","Publisher 2016"
	if %win% GEQ 9200 if "%_OneNoteRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2016"
	if %win% LSS 9200 if "%_OneNoteRetail%" EQU "YES" call :OfficeGVLKInstall "DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2016"
	if %win% GEQ 9200 if "%_OneNoteVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2016"
	if %win% LSS 9200 if "%_OneNoteVolume%" EQU "YES" call :OfficeGVLKInstall "DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2016"
	if %win% GEQ 9200 if "%_OneNote2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2021"
	if %win% LSS 9200 if "%_OneNote2021Retail%" EQU "YES" call :OfficeGVLKInstall "DR92N-9HTF2-97XKM-XW2WJ-XW3J6","OneNote 2021"
	if %win% GEQ 9200 if "%_SkypeForBusinessRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"869NQ-FJ69K-466HW-QYCP2-DDBV6","Skype For Business 2016"
	if %win% LSS 9200 if "%_SkypeForBusinessRetail%" EQU "YES" call :OfficeGVLKInstall "869NQ-FJ69K-466HW-QYCP2-DDBV6","Skype For Business 2016"
	if %win% GEQ 9200 if "%_Word2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"PBX3G-NWMT6-Q7XBW-PYJGG-WXD33","Word 2019"
	if %win% LSS 9200 if "%_Word2019Retail%" EQU "YES" call :OfficeGVLKInstall "PBX3G-NWMT6-Q7XBW-PYJGG-WXD33","Word 2019"
	if %win% GEQ 9200 if "%_Excel2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"TMJWT-YYNMB-3BKTF-644FC-RVXBD","Excel 2019"
	if %win% LSS 9200 if "%_Excel2019Retail%" EQU "YES" call :OfficeGVLKInstall "TMJWT-YYNMB-3BKTF-644FC-RVXBD","Excel 2019"
	if %win% GEQ 9200 if "%_PowerPoint2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"RRNCX-C64HY-W2MM7-MCH9G-TJHMQ","PowerPoint 2019"
	if %win% LSS 9200 if "%_PowerPoint2019Retail%" EQU "YES" call :OfficeGVLKInstall "RRNCX-C64HY-W2MM7-MCH9G-TJHMQ","PowerPoint 2019"
	if %win% GEQ 9200 if "%_Access2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT","Access 2019"
	if %win% LSS 9200 if "%_Access2019Retail%" EQU "YES" call :OfficeGVLKInstall "9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT","Access 2019"
	if %win% GEQ 9200 if "%_Outlook2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK","Outlook 2019"
	if %win% LSS 9200 if "%_Outlook2019Retail%" EQU "YES" call :OfficeGVLKInstall "7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK","Outlook 2019"
	if %win% GEQ 9200 if "%_Publisher2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"G2KWX-3NW6P-PY93R-JXK2T-C9Y9V","Publisher 2019"
	if %win% LSS 9200 if "%_Publisher2019Retail%" EQU "YES" call :OfficeGVLKInstall "G2KWX-3NW6P-PY93R-JXK2T-C9Y9V","Publisher 2019"
	if %win% GEQ 9200 if "%_SkypeForBusiness2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NCJ33-JHBBY-HTK98-MYCV8-HMKHJ","Skype For Business 2019"
	if %win% GEQ 9200 if "%_Word2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"TN8H9-M34D3-Y64V9-TR72V-X79KV","Word 2021"
	if %win% LSS 9200 if "%_Word2021Retail%" EQU "YES" call :OfficeGVLKInstall "TN8H9-M34D3-Y64V9-TR72V-X79KV","Word 2021"
	if %win% GEQ 9200 if "%_Excel2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NWG3X-87C9K-TC7YY-BC2G7-G6RVC","Excel 2021"
	if %win% LSS 9200 if "%_Excel2021Retail%" EQU "YES" call :OfficeGVLKInstall "NWG3X-87C9K-TC7YY-BC2G7-G6RVC","Excel 2021"
	if %win% GEQ 9200 if "%_PowerPoint2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"TY7XF-NFRBR-KJ44C-G83KF-GX27K","PowerPoint 2021"
	if %win% LSS 9200 if "%_PowerPoint2021Retail%" EQU "YES" call :OfficeGVLKInstall "TY7XF-NFRBR-KJ44C-G83KF-GX27K","PowerPoint 2021"
	if %win% GEQ 9200 if "%_Access2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"WM8YG-YNGDD-4JHDC-PG3F4-FC4T4","Access 2021"
	if %win% LSS 9200 if "%_Access2021Retail%" EQU "YES" call :OfficeGVLKInstall "WM8YG-YNGDD-4JHDC-PG3F4-FC4T4","Access 2019"
	if %win% GEQ 9200 if "%_Outlook2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"C9FM6-3N72F-HFJXB-TM3V9-T86R9","Outlook 2021"
	if %win% LSS 9200 if "%_Outlook2021Retail%" EQU "YES" call :OfficeGVLKInstall "C9FM6-3N72F-HFJXB-TM3V9-T86R9","Outlook 2021"
	if %win% GEQ 9200 if "%_Publisher2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ","Publisher 2021"
	if %win% LSS 9200 if "%_Publisher2021Retail%" EQU "YES" call :OfficeGVLKInstall "2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ","Publisher 2021"
	if %win% GEQ 9200 if "%_SkypeForBusiness2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HWCXN-K3WBT-WJBKY-R8BD9-XK29P","Skype For Business 2021"
	if %win% LSS 9200 if "%_SkypeForBusiness2021Retail%" EQU "YES" call :OfficeGVLKInstall "HWCXN-K3WBT-WJBKY-R8BD9-XK29P","Skype For Business 2021"
	if %win% GEQ 9200 if "%_Word2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"TN8H9-M34D3-Y64V9-TR72V-X79KV","Word 2021"
	if %win% LSS 9200 if "%_Word2021Volume%" EQU "YES" call :OfficeGVLKInstall "TN8H9-M34D3-Y64V9-TR72V-X79KV","Word 2021"
	if %win% GEQ 9200 if "%_Excel2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NWG3X-87C9K-TC7YY-BC2G7-G6RVC","Excel 2021"
	if %win% LSS 9200 if "%_Excel2021Volume%" EQU "YES" call :OfficeGVLKInstall "NWG3X-87C9K-TC7YY-BC2G7-G6RVC","Excel 2021"
	if %win% GEQ 9200 if "%_PowerPoint2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"TY7XF-NFRBR-KJ44C-G83KF-GX27K","PowerPoint 2021"
	if %win% LSS 9200 if "%_PowerPoint2021Volume%" EQU "YES" call :OfficeGVLKInstall "TY7XF-NFRBR-KJ44C-G83KF-GX27K","PowerPoint 2021"
	if %win% GEQ 9200 if "%_Access2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"WM8YG-YNGDD-4JHDC-PG3F4-FC4T4","Access 2021"
	if %win% LSS 9200 if "%_Access2021Volume%" EQU "YES" call :OfficeGVLKInstall "WM8YG-YNGDD-4JHDC-PG3F4-FC4T4","Access 2021"
	if %win% GEQ 9200 if "%_Outlook2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"C9FM6-3N72F-HFJXB-TM3V9-T86R9","Outlook 2021"
	if %win% LSS 9200 if "%_Outlook2021Volume%" EQU "YES" call :OfficeGVLKInstall "C9FM6-3N72F-HFJXB-TM3V9-T86R9","Outlook 2021"
	if %win% GEQ 9200 if "%_Publisher2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ","Publisher 2021"
	if %win% LSS 9200 if "%_Publisher2021Volume%" EQU "YES" call :OfficeGVLKInstall "2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ","Publisher 2021"
	if %win% GEQ 9200 if "%_SkypeForBusiness2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"HWCXN-K3WBT-WJBKY-R8BD9-XK29P","Skype For Business 2021"
	if %win% LSS 9200 if "%_SkypeForBusiness2021Volume%" EQU "YES" call :OfficeGVLKInstall "HWCXN-K3WBT-WJBKY-R8BD9-XK29P","Skype For Business 2021"
	if %win% GEQ 9200 if "%_VisioProRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","Visio Professional 2016"
	if %win% LSS 9200 if "%_VisioProRetail%" EQU "YES" call :OfficeGVLKInstall "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","Visio Professional 2016"
	if %win% GEQ 9200 if "%_ProjectProRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"YG9NW-3K39V-2T3HJ-93F3Q-G83KT","Project Professional 2016"
	if %win% LSS 9200 if "%_ProjectProRetail%" EQU "YES" call :OfficeGVLKInstall "YG9NW-3K39V-2T3HJ-93F3Q-G83KT","Project Professional 2016"
	if %win% GEQ 9200 if "%_AppxProject%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"YG9NW-3K39V-2T3HJ-93F3Q-G83KT","ProjectPro UWP Appx"
	if %win% LSS 9200 if "%_AppxProject%" EQU "YES" call :OfficeGVLKInstall "YG9NW-3K39V-2T3HJ-93F3Q-G83KT","ProjectPro UWP Appx"
	if %win% GEQ 9200 if "%_AppxVisio%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","VisioPro UWP Appx"
	if %win% LSS 9200 if "%_AppxVisio%" EQU "YES" call :OfficeGVLKInstall "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK","VisioPro UWP Appx"
	if %win% GEQ 9200 if "%_VisioPro2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"9BGNQ-K37YR-RQHF2-38RQ3-7VCBB","Visio Professional 2019"
	if %win% LSS 9200 if "%_VisioPro2019Retail%" EQU "YES" call :OfficeGVLKInstall "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB","Visio Professional 2019"
	if %win% GEQ 9200 if "%_VisioPro2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"9BGNQ-K37YR-RQHF2-38RQ3-7VCBB","Visio Professional 2019"
	if %win% LSS 9200 if "%_VisioPro2019Volume%" EQU "YES" call :OfficeGVLKInstall "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB","Visio Professional 2019"
	if %win% GEQ 9200 if "%_ProjectPro2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B","Project Professional 2019"
	if %win% LSS 9200 if "%_ProjectPro2019Retail%" EQU "YES" call :OfficeGVLKInstall "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B","Project Professional 2019"
	if %win% GEQ 9200 if "%_ProjectPro2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B","Project Professional 2019"
	if %win% LSS 9200 if "%_ProjectPro2019Volume%" EQU "YES" call :OfficeGVLKInstall "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B","Project Professional 2019"
	if %win% GEQ 9200 if "%_VisioPro2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4","Visio Professional 2021"
	if %win% LSS 9200 if "%_VisioPro2021Retail%" EQU "YES" call :OfficeGVLKInstall "KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4","Visio Professional 2021"
	if %win% GEQ 9200 if "%_VisioPro2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4","Visio Professional 2021"
	if %win% LSS 9200 if "%_VisioPro2021Volume%" EQU "YES" call :OfficeGVLKInstall "KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4","Visio Professional 2021"
	if %win% GEQ 9200 if "%_ProjectPro2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"FTNWT-C6WBT-8HMGF-K9PRX-QV9H8","Project Professional 2021"
	if %win% LSS 9200 if "%_ProjectPro2021Retail%" EQU "YES" call :OfficeGVLKInstall "FTNWT-C6WBT-8HMGF-K9PRX-QV9H8","Project Professional 2021"
	if %win% GEQ 9200 if "%_ProjectPro2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"FTNWT-C6WBT-8HMGF-K9PRX-QV9H8","Project Professional 2021"
	if %win% LSS 9200 if "%_ProjectPro2021Volume%" EQU "YES" call :OfficeGVLKInstall "FTNWT-C6WBT-8HMGF-K9PRX-QV9H8","Project Professional 2021"
	if %win% GEQ 9200 if "%_VisioStdRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"7WHWN-4T7MP-G96JF-G33KR-W8GF4","Visio Standard"
	if %win% LSS 9200 if "%_VisioStdRetail%" EQU "YES" call :OfficeGVLKInstall "7WHWN-4T7MP-G96JF-G33KR-W8GF4","Visio Standard"
	if %win% GEQ 9200 if "%_VisioStdVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"7WHWN-4T7MP-G96JF-G33KR-W8GF4","Visio Standard"
	if %win% LSS 9200 if "%_VisioStdVolume%" EQU "YES" call :OfficeGVLKInstall "7WHWN-4T7MP-G96JF-G33KR-W8GF4","Visio Standard"
	if %win% GEQ 9200 if "%_VisioStdXVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"NY48V-PPYYH-3F4PX-XJRKJ-W4423","Visio Standard C2R"
	if %win% LSS 9200 if "%_VisioStdXVolume%" EQU "YES" call :OfficeGVLKInstall "NY48V-PPYYH-3F4PX-XJRKJ-W4423","Visio Standard C2R"
	if %win% GEQ 9200 if "%_VisioStd2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2","Visio Standard 2019"
	if %win% LSS 9200 if "%_VisioStd2019Retail%" EQU "YES" call :OfficeGVLKInstall "7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2","Visio Standard 2019"
	if %win% GEQ 9200 if "%_VisioStd2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2","Visio Standard 2019"
	if %win% LSS 9200 if "%_VisioStd2019Volume%" EQU "YES" call :OfficeGVLKInstall "7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2","Visio Standard 2019"
	if %win% GEQ 9200 if "%_VisioStd2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"MJVNY-BYWPY-CWV6J-2RKRT-4M8QG","Visio Standard 2021"
	if %win% LSS 9200 if "%_VisioStd2021Retail%" EQU "YES" call :OfficeGVLKInstall "MJVNY-BYWPY-CWV6J-2RKRT-4M8QG","Visio Standard 2021"
	if %win% GEQ 9200 if "%_VisioStd2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"MJVNY-BYWPY-CWV6J-2RKRT-4M8QG","Visio Standard 2021"
	if %win% LSS 9200 if "%_VisioStd2021Volume%" EQU "YES" call :OfficeGVLKInstall "MJVNY-BYWPY-CWV6J-2RKRT-4M8QG","Visio Standard 2021"
	if %win% GEQ 9200 if "%_ProjectStdRetail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"GNFHQ-F6YQM-KQDGJ-327XX-KQBVC","Project Standard 2016"
	if %win% LSS 9200 if "%_ProjectStdRetail%" EQU "YES" call :OfficeGVLKInstall "GNFHQ-F6YQM-KQDGJ-327XX-KQBVC","Project Standard"
	if %win% GEQ 9200 if "%_ProjectStdVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"GNFHQ-F6YQM-KQDGJ-327XX-KQBVC","Project Standard 2016"
	if %win% LSS 9200 if "%_ProjectStdVolume%" EQU "YES" call :OfficeGVLKInstall "GNFHQ-F6YQM-KQDGJ-327XX-KQBVC","Project Standard"
	if %win% GEQ 9200 if "%_ProjectStdXVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"D8NRQ-JTYM3-7J2DX-646CT-6836M","Project Standard 2016 C2R"
	if %win% LSS 9200 if "%_ProjectStdXVolume%" EQU "YES" call :OfficeGVLKInstall "D8NRQ-JTYM3-7J2DX-646CT-6836M","Project Standard C2R"
	if %win% GEQ 9200 if "%_ProjectStd2019Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"C4F7P-NCP8C-6CQPT-MQHV9-JXD2M","Project Standard 2019"
	if %win% LSS 9200 if "%_ProjectStd2019Retail%" EQU "YES" call :OfficeGVLKInstall "C4F7P-NCP8C-6CQPT-MQHV9-JXD2M","Project Standard 2019"
	if %win% GEQ 9200 if "%_ProjectStd2019Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"C4F7P-NCP8C-6CQPT-MQHV9-JXD2M","Project Standard 2019"
	if %win% LSS 9200 if "%_ProjectStd2019Volume%" EQU "YES" call :OfficeGVLKInstall "C4F7P-NCP8C-6CQPT-MQHV9-JXD2M","Project Standard 2019"
	if %win% GEQ 9200 if "%_ProjectStd2021Retail%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T","Project Standard 2021"
	if %win% LSS 9200 if "%_ProjectStd2021Retail%" EQU "YES" call :OfficeGVLKInstall "J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T","Project Standard 2021"
	if %win% GEQ 9200 if "%_ProjectStd2021Volume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T","Project Standard 2021"
	if %win% LSS 9200 if "%_ProjectStd2021Volume%" EQU "YES" call :OfficeGVLKInstall "J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T","Project Standard 2021"
	if %win% GEQ 9200 if "%_ProjectProXVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"WGT24-HCNMF-FQ7XH-6M8K7-DRTW9","Project Professional 2016 C2R"
	if %win% LSS 9200 if "%_ProjectProXVolume%" EQU "YES" call :OfficeGVLKInstall "WGT24-HCNMF-FQ7XH-6M8K7-DRTW9","Project Professional 2016 C2R"
	if %win% GEQ 9200 if "%_VisioProXVolume%" EQU "YES" call :OfficeGVLKInstall "%sls%",%slsversion%,"69WXN-MBYV6-22PQG-3WGHK-RM6XC","Visio Professional 2016 C2R"
	if %win% LSS 9200 if "%_VisioProXVolume%" EQU "YES" call :OfficeGVLKInstall "69WXN-MBYV6-22PQG-3WGHK-RM6XC","Visio Professional 2016 C2R"
	
	echo:
	echo:
	timeout /t 4
	goto:eof
::===============================================================================================================
::===============================================================================================================
:OfficeGVLKInstall
	echo:
	if %win% GEQ 9200    (
    echo %4
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inpkey:%~3 %MultiNul%
    if %errorlevel% EQU 0 ((echo:)&&(echo Successfully installed %~3)&&(echo:))
    if %errorlevel% NEQ 0 ((echo:)&&(echo Installing %~3 failed)&&(echo:))
    )
    if %win% LSS 9200    (
    echo %2
	cscript "%OfficeRToolpath%\OfficeFixes\ospp\ospp.vbs" /inpkey:%~1 %MultiNul%
    if %errorlevel% EQU 0 ((echo:)&&(echo Successfully installed %~1)&&(echo:))
    if %errorlevel% NEQ 0 ((echo:)&&(echo Installing %~1 Failed)&&(echo:))
    )
	echo ____________________________________________________________________________
	goto:eof
::===============================================================================================================
::===============================================================================================================
:TheEndIsNear
	echo:
	echo:
	echo Ending OfficeRTool ...
	timeout /t 4
	exit
::===============================================================================================================

:CleanRegistryKeys

rem OSPP.VBS Nethood
rem OSPP.VBS Nethood
rem OSPP.VBS Nethood

%MultiNul% reg delete "%OSPP_USER%" /f /v KeyManagementServiceName
%MultiNul% reg delete "%OSPP_USER%" /f /v KeyManagementServicePort
%MultiNul% reg delete "%OSPP_USER%" /f /v DisableDnsPublishing
%MultiNul% reg delete "%OSPP_USER%" /f /v DisableKeyManagementServiceHostCaching

%MultiNul% reg delete "%OSPP_HKLM%" /f /v KeyManagementServiceName
%MultiNul% reg delete "%OSPP_HKLM%" /f /v KeyManagementServicePort
%MultiNul% reg delete "%OSPP_HKLM%" /f /v DisableDnsPublishing
%MultiNul% reg delete "%OSPP_HKLM%" /f /v DisableKeyManagementServiceHostCaching

rem SLMGR.VBS Nethood
rem SLMGR.VBS Nethood
rem SLMGR.VBS Nethood

%MultiNul% reg delete "%XSPP_USER%" /f /v KeyManagementServiceName
%MultiNul% reg delete "%XSPP_USER%" /f /v KeyManagementServicePort
%MultiNul% reg delete "%XSPP_USER%" /f /v DisableDnsPublishing
%MultiNul% reg delete "%XSPP_USER%" /f /v DisableKeyManagementServiceHostCaching

%MultiNul% reg delete "%XSPP_HKLM_X32%" /f /v KeyManagementServiceName
%MultiNul% reg delete "%XSPP_HKLM_X32%" /f /v KeyManagementServicePort
%MultiNul% reg delete "%XSPP_HKLM_X32%" /f /v DisableDnsPublishing
%MultiNul% reg delete "%XSPP_HKLM_X32%" /f /v DisableKeyManagementServiceHostCaching

%MultiNul% reg delete "%XSPP_HKLM_X64%" /f /v KeyManagementServiceName
%MultiNul% reg delete "%XSPP_HKLM_X64%" /f /v KeyManagementServicePort
%MultiNul% reg delete "%XSPP_HKLM_X64%" /f /v DisableDnsPublishing
%MultiNul% reg delete "%XSPP_HKLM_X64%" /f /v DisableKeyManagementServiceHostCaching

rem WMI Nethood -- Create SubKey under SPP KEY
rem WMI Nethood -- Create SubKey under SPP KEY
rem WMI Nethood -- Create SubKey under SPP KEY

for %%# in (55c92734-d682-4d71-983e-d6ec3f16059f, 0ff1ce15-a989-479d-af46-f275c6370663, 59a52881-a989-479d-af46-f275c6370663) do (
	%MultiNul% reg delete "%XSPP_USER%\%%#" /f
	%MultiNul% reg delete "%XSPP_HKLM_X32%\%%#" /f
	%MultiNul% reg delete "%XSPP_HKLM_X64%\%%#" /f
)
goto :eof

:UpdateLangFromIni
	set "inidownpath=!var3!"
	if "%inidownpath:~-1%" EQU " " set "inidownpath=%inidownpath:~0,-1%"
	set "downpath=!inidownpath!"
	set "inidownlang=!var6!"
	if "%inidownlang:~-1%" EQU " " set "inidownlang=%inidownlang:~0,-1%"
	set "o16lang=!inidownlang!"
	set "inidownarch=!var9!"
	if "%inidownarch:~-1%" EQU " " set "inidownarch=%inidownarch:~0,-1%"
	set "o16arch=!inidownarch!"
	call :UpdateSystemLanguge
	goto :eof

:updateRegistryKeys

rem OSPP.VBS Nethood
rem OSPP.VBS Nethood
rem OSPP.VBS Nethood

%MultiNul% reg add "%OSPP_USER%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
%MultiNul% reg add "%OSPP_USER%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
%MultiNul% reg add "%OSPP_USER%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
%MultiNul% reg add "%OSPP_USER%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

%MultiNul% reg add "%OSPP_HKLM%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
%MultiNul% reg add "%OSPP_HKLM%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
%MultiNul% reg add "%OSPP_HKLM%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
%MultiNul% reg add "%OSPP_HKLM%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

rem SLMGR.VBS Nethood
rem SLMGR.VBS Nethood
rem SLMGR.VBS Nethood

%MultiNul% reg add "%XSPP_USER%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
%MultiNul% reg add "%XSPP_USER%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
%MultiNul% reg add "%XSPP_USER%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
%MultiNul% reg add "%XSPP_USER%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

%MultiNul% reg add "%XSPP_HKLM_X32%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
%MultiNul% reg add "%XSPP_HKLM_X32%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
%MultiNul% reg add "%XSPP_HKLM_X32%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
%MultiNul% reg add "%XSPP_HKLM_X32%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

%MultiNul% reg add "%XSPP_HKLM_X64%" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
%MultiNul% reg add "%XSPP_HKLM_X64%" /f /v KeyManagementServicePort /t REG_SZ /d "%2"
%MultiNul% reg add "%XSPP_HKLM_X64%" /f /v DisableDnsPublishing /t REG_DWORD /d 0
%MultiNul% reg add "%XSPP_HKLM_X64%" /f /v DisableKeyManagementServiceHostCaching /t REG_DWORD /d 0

rem WMI Nethood -- Create SubKey under SPP KEY
rem WMI Nethood -- Create SubKey under SPP KEY
rem WMI Nethood -- Create SubKey under SPP KEY

%MultiNul% reg add "%XSPP_USER%\!subKey!\!Id!" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
%MultiNul% reg add "%XSPP_USER%\!subKey!\!Id!" /f /v KeyManagementServicePort /t REG_SZ /d "%2"

%MultiNul% reg add "%XSPP_HKLM_X32%\!subKey!\!Id!" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
%MultiNul% reg add "%XSPP_HKLM_X32%\!subKey!\!Id!" /f /v KeyManagementServicePort /t REG_SZ /d "%2"

%MultiNul% reg add "%XSPP_HKLM_X64%\!subKey!\!Id!" /f /v KeyManagementServiceName /t REG_SZ /d "%1"
%MultiNul% reg add "%XSPP_HKLM_X64%\!subKey!\!Id!" /f /v KeyManagementServicePort /t REG_SZ /d "%2"

goto :eof

:Query
%MultiNul% del /q "%temp%\result"
if /i '%3' EQU '' (
	>"%temp%\result" cscript "OfficeFixes\KMS Helper.vbs" "/QUERY_BASIC" %1 %2
) else (
	>"%temp%\result" cscript "OfficeFixes\KMS Helper.vbs" "/QUERY_ADVENCED" %1 %2 %3
)
goto :eof

:Channel_List
echo Current*492350f6-3a01-4f97-b9c0-c7c6ddf67d60
echo CurrentPreview*64256afe-f5d9-4f86-8936-8840a6a4f5be
echo BetaChannel*5440fd1f-7ecb-4221-8110-145efaa6372f
echo MonthlyEnterprise*55336b82-a18d-4dd6-b5f6-9e5095c314a6
echo SemiAnnual*7ffbc6bf-bc32-4f92-8982-f9dd17fd3114
echo SemiAnnualPreview*b8f9b850-328d-4355-9145-c59439a0c4cf
echo PerpetualVL2019*f2e724c1-748f-4b47-8fb8-8e0d210e9208
echo PerpetualVL2021*5030841d-c919-4594-8d2d-84ae4f96e58e
echo DogfoodDevMain*ea4a4090-de26-49d7-93c1-91bff9e53fc3
echo Manual_Override*ea4a4090-de26-49d7-93c1-91bff9e53fc3
echo Manual_Override*f3260cf1-a92c-4c75-b02e-d64c0a86a968
echo Manual_Override*c4a7726f-06ea-48e2-a13a-9d78849eb706
echo Manual_Override*834504cc-dc55-4c6d-9e71-e024d0253f6d
echo Manual_Override*5462eee5-1e97-495b-9370-853cd873bb07
echo Manual_Override*f4f024c8-d611-4748-a7e0-02b6e754c0fe
echo Manual_Override*b61285dd-d9f7-41f2-9757-8f61cba4e9c8
echo Manual_Override*9a3b7ff2-58ed-40fd-add5-1e5158059d1c
echo Manual_Override*86752282-5841-4120-ac80-db03ae6b5fdb
echo Manual_Override*2e148de9-61c8-4051-b103-4af54baffbb4
echo Manual_Override*12f4f6ad-fdea-4d2a-a90f-17496cc19a48
echo Manual_Override*0002c1ba-b76b-4af9-b1ee-ae2ad587371f
goto :eof

:Language_List
echo Afrikaans*1078*af-za*0436
echo Albanian*1052*sq-al*041c
echo Amharic*1118*am-et*045e
echo Arabic*1025*ar-sa*0401
echo Armenian*1067*hy-am*042b
echo Assamese*1101*as-in*044d
echo Azerbaijani Latin*1068*az-latn-az*042c
echo Bangla Bangladesh*2117*bn-bd*0845
echo Bangla Bengali India*1093*bn-in*0445
echo Basque Basque*1069*eu-es*042d
echo Belarusian*1059*be-by*0423
echo Bosnian*5146*bs-latn-ba*0141a
echo Bulgarian*1026*bg-bg*0402
echo Catalan Valencia*2051*ca-es-valencia*0803
echo Catalan*1027*ca-es*0403
echo Chinese Simplified*2052*zh-cn*0804
echo Chinese Traditional*1028*zh-tw*0404
echo Croatian*1050*hr-hr*041a
echo Czech*1029*cs-cz*0405
echo Danish*1030*da-dk*0406
echo Dari*1164*prs-af*048c
echo Dutch*1043*nl-nl*0413
echo English UK*2057*en-GB*0809
echo English*1033*en-us*0409
echo Estonian*1061*et-ee*0425
echo Filipino*1124*fil-ph*0464
echo Finnish*1035*fi-fi*040b
echo French Canada*3084*fr-CA*0C0C
echo French*1036*fr-fr*040c
echo Galician*1110*gl-es*0456
echo Georgian*1079*ka-ge*0437
echo German*1031*de-de*0407
echo Greek*1032*el-gr*0408
echo Gujarati*1095*gu-in*0447
echo Hausa Nigeria*1128*ha-Latn-NG*0468
echo Hebrew*1037*he-il*040d
echo Hindi*1081*hi-in*0439
echo Hungarian*1038*hu-hu*040e
echo Icelandic*1039*is-is*040f
echo Igbo*1136*ig-NG*0470
echo Indonesian*1057*id-id*0421
echo Irish*2108*ga-ie*083c
echo Italian*1040*it-it*0410
echo Japanese*1041*ja-jp*0411
echo Kannada*1099*kn-in*044b
echo Kazakh*1087*kk-kz*043f
echo Khmer*1107*km-kh*0453
echo KiSwahili*1089*sw-ke*0441
echo Konkani*1111*kok-in*0457
echo Korean*1042*ko-kr*0412
echo Kyrgyz*1088*ky-kg*0440
echo Latvian*1062*lv-lv*0426
echo Lithuanian*1063*lt-lt*0427
echo Luxembourgish*1134*lb-lu*046e
echo Macedonian*1071*mk-mk*042f
echo Malay Latin*1086*ms-my*043e
echo Malayalam*1100*ml-in*044c
echo Maltese*1082*mt-mt*043a
echo Maori*1153*mi-nz*0481
echo Marathi*1102*mr-in*044e
echo Mongolian*1104*mn-mn*0450
echo Nepali*1121*ne-np*0461
echo Norwedian Nynorsk*2068*nn-no*0814
echo Norwegian Bokmal*1044*nb-no*0414
echo Odia*1096*or-in*0448
echo Pashto*1123*ps-AF*0463
echo Persian*1065*fa-ir*0429
echo Polish*1045*pl-pl*0415
echo Portuguese Brazilian*1046*pt-br*0416
echo Portuguese Portugal*2070*pt-pt*0816
echo Punjabi Gurmukhi*1094*pa-in*0446
echo Quechua*3179*quz-pe*0c6b
echo Romanian*1048*ro-ro*0418
echo Romansh*1047*rm-CH*0417
echo Russian*1049*ru-ru*0419
echo Setswana*1074*tn-ZA*0432
echo Scottish Gaelic*1169*gd-gb*0491
echo Serbian Bosnia*7194*sr-cyrl-ba*01c1a
echo Serbian Serbia*10266*sr-cyrl-rs*0281a
echo Serbian*9242*sr-latn-rs*0241a
echo Sesotho sa Leboa*1132*nso-ZA*046C
echo Sindhi Arabic*2137*sd-arab-pk*0859
echo Sinhala*1115*si-lk*045b
echo Slovak*1051*sk-sk*041b
echo Slovenian*1060*sl-si*0424
echo Spanish*3082*es-es*0c0a
echo Spanish Mexico*2058*es-MX*080A
echo Swedish*1053*sv-se*041d
echo Tamil*1097*ta-in*0449
echo Tatar Cyrillic*1092*tt-ru*0444
echo Telugu*1098*te-in*044a
echo Thai*1054*th-th*041e
echo Turkish*1055*tr-tr*041f
echo Turkmen*1090*tk-tm*0442
echo Ukrainian*1058*uk-ua*0422
echo Urdu*1056*ur-pk*0420
echo Uyghur*1152*ug-cn*0480
echo Uzbek*1091*uz-latn-uz*0443
echo Vietnamese*1066*vi-vn*042a
echo Welsh*1106*cy-gb*0452
echo Wolof*1160*wo-SN*0488
echo Yoruba*1130*yo-NG*046A
echo isiXhosa*1076*xh-ZA*0434
echo isiZulu*1077*zu-ZA*0435
goto :eof

:Get-WinUserLanguageList_Warper
call :Get-WinUserLanguageList
if defined SysLanIdHex call :convertLanHexToDec
goto :eof
:Get-WinUserLanguageList
set xVal=
set SysLanCD=
set SysLanIdHex=
%MultiNul% reg query "HKEY_CURRENT_USER\Control Panel\International\User Profile" /v Languages || goto :eof
for /f "tokens=3 delims= " %%g in ('reg query "HKEY_CURRENT_USER\Control Panel\International\User Profile" /v Languages ^| find /i "REG_MULTI_SZ"') do set xVal=%%g
if defined xVal 		(for /f "tokens=1 delims=\0" %%g in ('echo !xVal!') do set SysLanIdHex=%%g)
if defined SysLanIdHex 	(for /f "tokens=1 delims= " %%g in ('reg query "HKEY_CURRENT_USER\Control Panel\International\User Profile\!SysLanIdHex!" ^| find /i "000"') do set SysLanIdHex=%%g)
if defined SysLanIdHex 	(for /f "tokens=1 delims=:" %%g in ('echo !SysLanIdHex!') do set SysLanIdHex=%%g)
goto :eof
:convertLanHexToDec
:: %%g=English %%h=1033 %%i=en-us %%j:0409
>"%temp%\tmp" call :Language_List
for /f "tokens=1,2,3,4 delims=*" %%g in ('type "%temp%\tmp"') do (
	if 	/i '%%j' EQU '!SysLanIdHex!' (
		set SysLanCD=%%i
		goto :convertLanHexToDec_
	)	
)
:convertLanHexToDec_
goto :eof

:CheckSystemLanguage
set var=&set var=%*
if not defined var (

	:: Using HKCR :: PreferredUILanguages Value
	%MultiNul% reg query "HKEY_CURRENT_USER\Control Panel\Desktop" /v PreferredUILanguages && (
		REM echo Using HKCR :: PreferredUILanguages Value
		for /f "tokens=1,3" %%g in ('reg query "HKEY_CURRENT_USER\Control Panel\Desktop" /v PreferredUILanguages') do (
			if /i '%%g' EQU 'PreferredUILanguages' call :CheckSystemLanguage %%h
		)
		goto :eof
	)
	
	:: Using HKLM:: PreferredUILanguages Value
	%MultiNul% reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\MUI\Settings" /v PreferredUILanguages && (
		REM echo Using HKLM:: PreferredUILanguages Value
		for /f "tokens=1,3" %%g in ('reg query "HKEY_LOCAL_MACHINE\SYSTEM\SYSTEM\CurrentControlSet\Control\MUI\Settings" /v PreferredUILanguages') do (
			if /i '%%g' EQU 'PreferredUILanguages' call :CheckSystemLanguage %%h
		)
		goto :eof
	)
	
	:: using Get-WinUserLanguageList Function Cmd Warper
	if defined SysLanCD (
		call :CheckSystemLanguage %SysLanCD%
		goto :eof
	)
	
	:: using dism :: get-intl
	REM echo using dism :: get-intl
	for /f "tokens=4,6" %%g  in ('dism /online /get-intl') do (
		if /i '%%g' EQU 'language' call :CheckSystemLanguage %%h
	)
	goto :eof
)
if defined var (
		
	:: %%g=English %%h=1033 %%i=en-us %%j:0409
	>"%temp%\tmp" call :Language_List
	for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
		if /i '%var%' EQU '%%i' (
			set langtext=%%g
			set o16lcid=%%h
			set o16lang=%%i
			goto :CheckSystemLanguage_
		)
	)
)
:CheckSystemLanguage_
goto :eof

:UpdateSystemLanguge
rem %%g=English %%h=1033 %%i=en-us %%j:0409
set "langFound="
>"%temp%\tmp" call :Language_List
for /f "tokens=1,2,3,4 delims=*" %%g in ('type "%temp%\tmp"') do (
	if /i "!o16lang!" EQU "%%i" (
		set langtext=%%g
		set o16lcid=%%h
		set langFound=***
	)
)
if not defined langFound (
	set "o16lang=en-US"
	set "langtext=Default Language"
    set "o16lcid=1033"
)
goto :eof

:GetInfoFromFolder
	for %%g in (lLngID, multiLang, MultiL, MultiM, x32, x64, lang, version, ChannelName, ChannelID, channeltrigger) do set %%g=
	if exist office\data\v64.cab (set "x64=*"&set "lBit=64")
	if exist office\data\v32.cab (set "x32=*"&set "lBit=32")
	if not defined x64 (
		if not defined x32 (
			echo:
			echo Download incomplete - Package unusable - Redo download
			echo:
			if not defined debugMode pause
			goto:InstallO16Loop
		)
	)
	
	%MultiNul% dir /ad /b "office\data\16*" && (
		for /f "tokens=*" %%g in ('dir /ad /b "office\data\16*"') do set "version=%%g"
	)
	
	echo "!version!" | >nul findstr /r "16.[0-9].[0-9][0-9][0-9][0-9][0-9].[0-9][0-9][0-9][0-9][0-9]" || (
		(echo:)&&(echo Download incomplete - Package unusable - Redo download)&&(echo:)&&(pause)&&(goto:InstallO16Loop)
	)
	
	if not defined version (echo:)&&(echo Download incomplete - Package unusable - Redo download)&&(echo:)&&(pause)&&(goto:InstallO16Loop)
	if defined x64 (
		if defined x32 (
			set "x64="
			set "MultiM=XXX"
			set "lBit=32"
		)
	)
	
	%MultiNul% del /q "%temp%\tmp"
	>"%temp%\tmp" call :Language_List
	
	for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (	
		if exist "office\data\!version!\i%lBit%%%h.cab" (
			set "lLngID=%%h"
			if not defined multiLang (
				set "multiLang=!lLngID!"
			) else (
				(echo '!multiLang!' |>nul find /i "!lLngID!") || (set "multiLang=!multiLang!,!lLngID!")
			)
		)
	)

	if defined lLngID (
		if defined multiLang (
		
			set "count=0"
			set "countVal="
			if /i '!lLngID!' NEQ '!multiLang!' (
				
				echo:
				echo Multi Language Found.
				echo _____________________
				echo:
				for %%# in (!multiLang!) do (
					set /a count+=1
					call :FindLngId %%#
					echo Language [!count!] :: [%%#] !langIdName!
					set "lang_!count!=%%#"
					set "countVal=!countVal!!count!"
				)
				
				echo:
				CHOICE /C !countVal! /M "Select Language ID ::"
				
				FOR /L %%# IN (1,1,!count!) DO (			
					if /i '!errorlevel!' EQU '%%#' (
						set "lLngID=!lang_%%#!"
					)
				)
				
				set "MultiL=XXX"
			)
		)
	) else (
		(echo:)&&(echo Download incomplete - Package unusable - Redo download)&&(echo:)&&(pause)&&(goto:InstallO16Loop)
	)
	
	call :ConvertIDtoXXFormat !lLngID!
	set "lang=!langId_New!"
	
	if defined AskUser (
		set "AskUser="
		goto :eof
	)
	
	if defined x32 set "CabFile=Office\Data\v32.cab"
	if defined x64 set "CabFile=Office\Data\v64.cab"
	
	(>nul expand !CabFile! -F:VersionDescriptor.xml "%temp%") || (
		(echo:)&&(echo Download incomplete - Package unusable - Redo download)&&(echo:)&&(pause)&&(goto:InstallO16Loop)
	)
	
	if not exist "%temp%\VersionDescriptor.xml" (
		(echo:)&&(echo Download incomplete - Package unusable - Redo download)&&(echo:)&&(pause)&&(goto:InstallO16Loop)
	)
	
	set "DeliveryMechanism="
	for /f "tokens=*" %%# in ('type "%temp%\VersionDescriptor.xml"^|find /i "DeliveryMechanism"') do set "DeliveryMechanism=%%#"
	
	if not defined DeliveryMechanism (
		(echo:)&&(echo Download incomplete - Package unusable - Redo download)&&(echo:)&&(pause)&&(goto:InstallO16Loop)
	)
	set "DeliveryMechanism=!DeliveryMechanism:~28,-2!"
	
	rem %%g Name, %%h Channel
	>"%temp%\tmp" call :Channel_List
	for /f "tokens=1,2 delims=*" %%g in ('type "%temp%\tmp"') do (
		(echo !DeliveryMechanism! | >nul find /i "%%h") && (
			set "ChannelName=%%g"
			set "ChannelID=%%h"
		)
	)
	
	if not defined ChannelName (
		(echo:)&&(echo Download incomplete - Package unusable - Redo download)&&(echo:)&&(pause)&&(goto:InstallO16Loop)
	)
	
	%TripleNul% echo.>package.info && (
	
		>package.info echo !ChannelName!
		>>package.info echo !version!
		
		if defined MultiL (
			>>package.info echo Multi
		) else (
			>>package.info echo !lang!
		)
		if defined MultiM (
			>>package.info echo Multi
		) else (
			if defined x32 (
				>>package.info echo x86
			) else (
				>>package.info echo x64
			)
		)
		>>package.info echo !ChannelID!
		
	) || (
		echo.
		echo ERROR ### Fail to write package.info File
		echo.
		if not defined debugMode pause
		goto:InstallO16Loop
	)
	
	set "distribchannel=!ChannelName!"
	if "%distribchannel:~-1%" EQU " " set "distribchannel=%distribchannel:~0,-1%"
	echo set "o16build=!version!"
	set "o16build=!version!"
	set "o16lang=!lang!"
	call :SetO16Language
	
	if defined MultiM (
		set "o16arch=Multi"
	) else (
		if defined x32 (
			set "o16arch=x86"
		) else (
			set "o16arch=x64"
		)
	)
	set "o16updlocid=!ChannelID!"
	goto :Pdhfsdj45X
	
:FindLngId
	set "langIdName="
	set var=&set var=%*
	if defined var (

		:: %%g=English %%h=1033 %%i=en-us %%j:0409
		>"%temp%\tmp" call :Language_List
		for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
			if /i '%var%' EQU '%%h' (
				set "langIdName=%%g"
				goto :FindLngId_
			)
		)
	)
	:FindLngId_
	goto :eof
	
:ConvertIDtoXXFormat
	set "langId_New="
	set var=&set var=%*
	if defined var (

		:: %%g=English %%h=1033 %%i=en-us %%j:0409
		>"%temp%\tmp" call :Language_List
		for /f "tokens=1,2,3 delims=*" %%g in ('type "%temp%\tmp"') do (
			if /i '%var%' EQU '%%h' (
				set "langId_New=%%i"
				goto :FindLngId_
			)
		)
	)
	:FindLngId_
	goto :eof
	
:generateXML
	 >"%oxml%" echo ^<Configuration^>
	>>"%oxml%" echo     ^<Add OfficeClientEdition="%o16a%" Version="!o16build!"%channel% ^> 
	
	if "%mo16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Mondo%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
		if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%of16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="ProPlus%type%"^> 
		>>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
		if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
		>>"%oxml%" echo         ^</Product^> 
	)
    if "%of19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="ProPlus2019%type%"^> 
		>>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
		if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
		>>"%oxml%" echo         ^</Product^> 
	)
  if "%of21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="ProPlus2021%type%"^> 
		>>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
		if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
		>>"%oxml%" echo         ^</Product^> 
	)
   if "%of36install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="O365ProPlus%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
		if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
		>>"%oxml%" echo         ^</Product^> 
	)
    if "%ofbsinstall%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="O365Business%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
		if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
		>>"%oxml%" echo         ^</Product^> 
	)
    if "%vi16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="VisioPro%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
    	if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
	    >>"%oxml%" echo         ^</Product^> 
	)
    if "%vi19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="VisioPro2019%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
    	if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
	    >>"%oxml%" echo         ^</Product^> 
	)
    if "%vi21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="VisioPro2021%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
    	if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
	    >>"%oxml%" echo         ^</Product^> 
	)
    if "%pr16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="ProjectPro%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
    	if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
	    >>"%oxml%" echo         ^</Product^> 
	)
    if "%pr19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="ProjectPro2019%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
    	if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
	    >>"%oxml%" echo         ^</Product^> 
	)
    if "%pr21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="ProjectPro2021%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
    	if "%wd16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Word"/^> 
		if "%ex16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Excel"/^> 
		if "%pp16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="PowerPoint"/^> 
		if "%ac16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Access"/^> 
		if "%ol16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Outlook"/^> 
		if "%pb16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Publisher"/^> 
		if "%on16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneNote"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Lync"/^> 
		if "%st16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Teams"/^> 
		if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Groove"/^> 
        if "%od16disable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="OneDrive"/^> 
		if "%bsbsdisable%" EQU "1" >>"%oxml%" echo             ^<ExcludeApp ID="Bing"/^> 
	    >>"%oxml%" echo         ^</Product^> 
	)
   if "%wd16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Word%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%wd19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Word2019%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%wd21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Word2021%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%ex16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Excel%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%ex19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Excel2019%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%ex21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Excel2021%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%pp16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="PowerPoint%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%pp19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="PowerPoint2019%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%pp21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="PowerPoint2021%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%ac16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Access%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%ac19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Access2019%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%ac21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Access2021%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%ol16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Outlook%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
	if "%ol19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Outlook2019%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
	if "%ol21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Outlook2021%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
	if "%pb16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Publisher%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%pb19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Publisher2019%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%pb21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="Publisher2021%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%on16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="OneNote%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%sk16install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="SkypeForBusiness%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%sk19install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="SkypeForBusiness2019%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
	if "%on21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="OneNote2021Retail"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    if "%sk21install%" EQU "1" (
        >>"%oxml%" echo         ^<Product ID="SkypeForBusiness2021%type%"^> 
        >>"%oxml%" echo             ^<Language ID="!o16lang!"/^> 
        >>"%oxml%" echo         ^</Product^> 
	)
    >>"%oxml%" echo     ^</Add^> 
	>>"%oxml%" echo     ^<Property Name="ForceAppsShutdown" Value="True" /^> 
	>>"%oxml%" echo     ^<Property Name="PinIconsToTaskbar" Value="False" /^> 
    >>"%oxml%" echo     ^<Display Level="Full" AcceptEula="True" /^> 
	>>"%oxml%" echo     ^<Updates Enabled="True" UpdatePath="http://officecdn.microsoft.com/pr/!o16updlocid!"%channel% /^> 
	>>"%oxml%" echo ^</Configuration^>
	goto :eof
	
Rem abbodi1406 KMS VL ALL LOCAL ACTIVATION
Rem abbodi1406 KMS VL ALL LOCAL ACTIVATION
Rem abbodi1406 KMS VL ALL LOCAL ACTIVATION

:STARTKMSActivation
set SSppHook=0
set KMSPort=1688
set KMSHostIP=0.0.0.0
set KMS_RenewalInterval=10080
set KMS_ActivationInterval=120
set KMS_HWID=0x3A1C049600B60076

set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set "_TaskEx=\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger"

if /i "%PROCESSOR_ARCHITECTURE%"=="amd64" set "xOS=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "xOS=A64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "xOS=x86"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "xOS=x64"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "xOS=A64"

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"
set _Hook="%SysPath%\SppExtComObjHook.dll"

for /f %%A in ('"2>nul dir /b /ad %SysPath%\spp\tokens\skus"') do (
	if %win% GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*GVLK*.xrm-ms" set SSppHook=1
	if %win% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VLKMS*.xrm-ms" set SSppHook=1
	if %win% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VL-BYPASS*.xrm-ms" set SSppHook=1
)
set OsppHook=1
sc query osppsvc %MultiNul%
if %errorlevel% EQU 1060 set OsppHook=0

set ESU_KMS=0
if %win% LSS 9200 for /f %%A in ('"2>nul dir /b /ad %SysPath%\spp\tokens\channels"') do (
  if exist "%SysPath%\spp\tokens\channels\%%A\*VL-BYPASS*.xrm-ms" set ESU_KMS=1
)
if %ESU_KMS% EQU 1 (set "adoff=and LicenseDependsOn is NULL"&set "addon=and LicenseDependsOn is not NULL") else (set "adoff="&set "addon=")
set ESU_EDT=0
if %ESU_KMS% EQU 1 for %%A in (Enterprise,EnterpriseE,EnterpriseN,Professional,ProfessionalE,ProfessionalN,Ultimate,UltimateE,UltimateN) do (
  if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%%A\*.xrm-ms" set ESU_EDT=1
)
if %ESU_EDT% EQU 1 set SSppHook=1
set ESU_ADD=0

if %win% GEQ 9200 (
	set OSType=Win8
	set SppVer=SppExtComObj.exe
) else if %win% GEQ 7600 (
	set OSType=Win7
	set SppVer=sppsvc.exe
) else (
	if not defined debugMode pause
	exit /b
)
if %OSType% EQU Win8 reg query "%IFEO%\sppsvc.exe" %MultiNul% && (
	reg delete "%IFEO%\sppsvc.exe" /f %MultiNul%
	call :StopService sppsvc
)
set _uRI=%KMS_RenewalInterval%
set _uAI=%KMS_ActivationInterval%
set _AUR=0
if exist %_Hook% dir /b /al %_Hook% %MultiNul% || (
  reg query "%IFEO%\%SppVer%" /v VerifierFlags %MultiNul% && set _AUR=1
  if %SSppHook% EQU 0 reg query "%IFEO%\osppsvc.exe" /v VerifierFlags %MultiNul% && set _AUR=1
)

if %_AUR% EQU 0 (
	set KMS_RenewalInterval=43200
	set KMS_ActivationInterval=43200
) else (
	set KMS_RenewalInterval=%_uRI%
	set KMS_ActivationInterval=%_uAI%
)
if %win% GEQ 9600 (
	reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /t REG_DWORD /d 1 /f %MultiNul%
	if %win% EQU 14393 reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoAcquireGT /t REG_DWORD /d 1 /f %MultiNul%
)

call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
  if exist "%SysPath%\%%#" del /f /q "%SysPath%\%%#" %MultiNul%
  if exist "%SystemRoot%\SysWOW64\%%#" del /f /q "%SystemRoot%\SysWOW64\%%#" %MultiNul%
)
set AclReset=0
set _cphk=0
if %_AUR% EQU 1 set _cphk=1
if %_cphk% EQU 1 (
	copy /y "%OfficeRToolpath%\OfficeFixes\bin\%xOS%.dll" %_Hook% %MultiNul%
	goto :skipsym
)
mklink %_Hook% "%OfficeRToolpath%\OfficeFixes\bin\%xOS%.dll" %MultiNul%
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% NEQ 0 goto :E_SYM
icacls %_Hook% /findsid *S-1-5-32-545 %SingleNulV2% | find /i "SppExtComObjHook.dll" %SingleNul% || (
	set AclReset=1
	icacls %_Hook% /grant *S-1-5-32-545:RX %MultiNul%
)
:skipsym
if %SSppHook% NEQ 0 call :CreateIFEOEntry %SppVer%
if %_AUR% EQU 1 (call :CreateIFEOEntry osppsvc.exe) else (if %OsppHook% NEQ 0 call :CreateIFEOEntry osppsvc.exe)
if %_AUR% EQU 1 if %OSType% EQU Win7 call :CreateIFEOEntry SppExtComObj.exe
if %_AUR% EQU 1 (
call :UpdateIFEOEntry %SppVer%
call :UpdateIFEOEntry osppsvc.exe
)
goto :eof

:StopKMSActivation
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %_AUR% EQU 0 call :RemoveHook
sc start sppsvc trigger=timer;sessionid=0 %MultiNul%
goto :eof

:StopService
sc query %1 | find /i "STOPPED" %SingleNul% || net stop %1 /y %MultiNul%
sc query %1 | find /i "STOPPED" %SingleNul% || sc stop %1 %MultiNul%
goto :eof


:RemoveHook
if %AclReset% EQU 1 icacls %_Hook% /reset %MultiNul%
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
	if exist "%SysPath%\%%#" del /f /q "%SysPath%\%%#" %MultiNul%
	if exist "%SystemRoot%\SysWOW64\%%#" del /f /q "%SystemRoot%\SysWOW64\%%#" %MultiNul%
)
for %%# in (SppExtComObj.exe,sppsvc.exe,osppsvc.exe) do reg query "%IFEO%\%%#" %MultiNul% && (
	call :RemoveIFEOEntry %%#
)
if %OSType% EQU Win8 schtasks /query /tn "%_TaskEx%" %MultiNul% && (
	schtasks /delete /f /tn "%_TaskEx%" %MultiNul%
)
goto :eof

:CreateIFEOEntry
reg delete "%IFEO%\%1" /f /v Debugger %MultiNul%
reg add "%IFEO%\%1" /f /v VerifierDlls /t REG_SZ /d "SppExtComObjHook.dll" %MultiNul%
reg add "%IFEO%\%1" /f /v VerifierDebug /t REG_DWORD /d 0x00000000 %MultiNul%
reg add "%IFEO%\%1" /f /v VerifierFlags /t REG_DWORD /d 0x80000000 %MultiNul%
reg add "%IFEO%\%1" /f /v GlobalFlag /t REG_DWORD /d 0x00000100 %MultiNul%
reg add "%IFEO%\%1" /f /v KMS_Emulation /t REG_DWORD /d 1 %MultiNul%
reg add "%IFEO%\%1" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% %MultiNul%
reg add "%IFEO%\%1" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% %MultiNul%

if /i %1 EQU SppExtComObj.exe if %win% GEQ 9600 (
	reg add "%IFEO%\%1" /f /v KMS_HWID /t REG_QWORD /d "%KMS_HWID%" %MultiNul%
)
goto :eof

:RemoveIFEOEntry
if /i %1 NEQ osppsvc.exe (
reg delete "%IFEO%\%1" /f %MultiNul%
goto :eof
)
if %OsppHook% EQU 0 (
reg delete "%IFEO%\%1" /f %MultiNul%
)
if %OsppHook% NEQ 0 for %%A in (Debugger,VerifierDlls,VerifierDebug,VerifierFlags,GlobalFlag,KMS_Emulation,KMS_ActivationInterval,KMS_RenewalInterval,Office2010,Office2013,Office2016,Office2019) do reg delete "%IFEO%\%1" /v %%A /f %MultiNul%
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "0.0.0.0" %MultiNul%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" %MultiNul%
goto :eof

:UpdateIFEOEntry
reg add "%IFEO%\%1" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% %MultiNul%
reg add "%IFEO%\%1" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% %MultiNul%
if /i %1 EQU SppExtComObj.exe if %win% GEQ 9600 (
reg add "%IFEO%\%1" /f /v KMS_HWID /t REG_QWORD /d "%KMS_HWID%" %MultiNul%
)
if /i %1 EQU sppsvc.exe (
reg add "%IFEO%\SppExtComObj.exe" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% %MultiNul%
reg add "%IFEO%\SppExtComObj.exe" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% %MultiNul%
)

:UpdateOSPPEntry
if /i %1 EQU osppsvc.exe (
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMSHostIP%" %MultiNul%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMSPort%" %MultiNul%
)
goto :eof

:StopService
:: Stop service based on parameter
sc query "%1" | find /i "STOPPED" %MultiNul% || (
  net stop "%1" /y %MultiNul%
)
sc query "%1" | find /i "STOPPED" %MultiNul% || (
  sc stop "%1" %MultiNul%
)
exit /b

:StartService
sc query "%~1" | %SingleNul% find /i "STOPPED" && %MultiNul% sc start "%~1"
sc query "%~1" | %SingleNul% find /i "RUNNING" || goto :StartService
goto:eof