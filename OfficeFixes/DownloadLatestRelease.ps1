<# Using Powershell To Retrieve Latest Package Url From Github Releases #>
<# https://copdips.com/2019/12/Using-Powershell-to-retrieve-latest-package-url-from-github-releases.html #>

try {
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
	Write-host
	Write-host Download Latest Release
	Invoke-WebRequest -Uri $realDownloadUrl -OutFile $OutputFile
} Catch {}