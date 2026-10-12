# VDI environment check
# Collects what automation tools are available in this VDI and (optionally)
# tests whether a site can be reached with your Windows login.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File env-check.ps1
#   powershell -ExecutionPolicy Bypass -File env-check.ps1 -Url "https://site/reports"
# If scripts are blocked, open PowerShell and paste this whole file in instead.
# Output: vdi-env-report.txt on your Desktop. No passwords are collected.

param([string]$Url)

$report = [ordered]@{}
function Try-Get($block) { try { & $block } catch { "error: $($_.Exception.Message)" } }

$report['Date']            = Get-Date -Format s
$report['OS']              = Try-Get { (Get-CimInstance Win32_OperatingSystem).Caption }
$report['PowerShell']      = $PSVersionTable.PSVersion.ToString()
$report['LanguageMode']    = $ExecutionContext.SessionState.LanguageMode
$report['ExecutionPolicy'] = Try-Get { (Get-ExecutionPolicy -List | Out-String).Trim() }
$report['curl.exe']        = [bool](Get-Command curl.exe -ErrorAction SilentlyContinue)

$edge = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
$report['Edge'] = if (Test-Path $edge) { (Get-Item $edge).VersionInfo.ProductVersion } else { 'not found' }
$chrome = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
$report['Chrome'] = if (Test-Path $chrome) { (Get-Item $chrome).VersionInfo.ProductVersion } else { 'not found' }

$pad = @("${env:ProgramFiles(x86)}\Power Automate Desktop", "$env:ProgramFiles\Power Automate Desktop") | Where-Object { Test-Path $_ }
$padStore = Try-Get { Get-AppxPackage -Name '*PowerAutomate*' | Select-Object -ExpandProperty Version }
$report['PowerAutomateDesktop'] = if ($pad -or $padStore) { "installed $padStore $pad".Trim() } else { 'not found' }

$ahk = @("$env:ProgramFiles\AutoHotkey", "${env:ProgramFiles(x86)}\AutoHotkey", "$env:LOCALAPPDATA\Programs\AutoHotkey") | Where-Object { Test-Path $_ }
$report['AutoHotkey'] = if ($ahk) { $ahk -join ', ' } else { 'not found' }

$report['EdgeDevToolsPolicy'] = Try-Get {
    $v = (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -ErrorAction Stop).DeveloperToolsAvailability
    switch ($v) { 0 {'allowed (0)'} 1 {'allowed (1)'} 2 {'BLOCKED (2)'} default {'not set (allowed)'} }
}
$report['EdgeDownloadDirPolicy'] = Try-Get { (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -ErrorAction Stop).DownloadDirectory }
$report['DownloadsFolder'] = "$env:USERPROFILE\Downloads"
$report['VDIServices'] = Try-Get {
    (Get-Service -DisplayName '*Citrix*', '*Horizon*', '*VMware*', '*Remote Desktop*' -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty DisplayName | Sort-Object -Unique) -join '; '
}
$report['WinHttpProxy'] = Try-Get { (netsh winhttp show proxy | Out-String).Trim() }

if ($Url) {
    $report['SiteTest.Url'] = $Url
    try {
        $r = Invoke-WebRequest -Uri $Url -UseDefaultCredentials -UseBasicParsing -MaximumRedirection 10 -TimeoutSec 30
        $report['SiteTest.Status']    = $r.StatusCode
        $report['SiteTest.FinalUrl']  = $r.BaseResponse.ResponseUri.AbsoluteUri
        $report['SiteTest.LoginForm'] = [bool]($r.Content -match 'type\s*=\s*["'']?password')
        $report['SiteTest.LinkCount'] = @($r.Links).Count
        $files = @($r.Links | Where-Object { $_.href -match '\.(xlsx?|csv|pdf|zip|txt|docx?)(\?|$)' } | Select-Object -First 20 -ExpandProperty href)
        $report['SiteTest.FileLinks'] = if ($files) { $files -join "`n    " } else { 'none found in raw HTML (page may load them with JavaScript)' }
    } catch {
        $report['SiteTest.Error'] = $_.Exception.Message
    }
}

$out = Join-Path ([Environment]::GetFolderPath('Desktop')) 'vdi-env-report.txt'
($report.GetEnumerator() | ForEach-Object { "{0}: {1}" -f $_.Key, $_.Value }) | Set-Content -Path $out -Encoding UTF8
Get-Content $out
Write-Host "`nSaved to $out" -ForegroundColor Green
