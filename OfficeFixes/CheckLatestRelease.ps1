<# Using Powershell To Retrieve Latest Package Url From Github Releases #>
<# https://copdips.com/2019/12/Using-Powershell-to-retrieve-latest-package-url-from-github-releases.html #>

try {
	$url = 'https://github.com/maorosh123/OfficeRTool/releases/latest'
	$request = [System.Net.WebRequest]::Create($url)
	$response = $request.GetResponse()
	$realTagUrl = $response.ResponseUri.OriginalString
	Write-Host $realTagUrl.split('/')[-1].Trim('v')
} Catch {}