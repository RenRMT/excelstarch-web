# Registers the GitHub Pages build of ExcelStarch with Excel desktop for the current user.
# No admin rights or Node.js needed. Run with:
#   powershell -ExecutionPolicy Bypass -File scripts\register-addin.ps1
# Pass -Unregister to remove it again. Restart Excel afterwards.

param([switch]$Unregister)

$manifestUrl = "https://renrmt.github.io/excelstarch-web/manifest.xml"
$installDir = Join-Path $env:LOCALAPPDATA "ExcelStarch"
$manifestPath = Join-Path $installDir "manifest.xml"
$regKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

if ($Unregister) {
  Remove-ItemProperty -Path $regKey -Name "ExcelStarch" -ErrorAction SilentlyContinue
  Write-Output "ExcelStarch unregistered. Restart Excel."
  return
}

New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Invoke-WebRequest -Uri $manifestUrl -OutFile $manifestPath -UseBasicParsing

New-Item -Path $regKey -Force | Out-Null
New-ItemProperty -Path $regKey -Name "ExcelStarch" -Value $manifestPath -PropertyType String -Force | Out-Null

Write-Output "ExcelStarch registered from $manifestPath. Restart Excel."
