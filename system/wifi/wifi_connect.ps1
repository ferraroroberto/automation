# Generate the local, untracked wifi_connect.xml consumed by wifi_connect.py.
#
# Values are supplied at call time — never edited into this file, which is
# tracked in git. Pass them as parameters, or set WIFI_SSID / WIFI_PASSWORD in
# the environment and omit them.
#
#   .\wifi_connect.ps1 -Ssid "MyNetwork" -Password "my-passphrase"
#   .\wifi_connect.ps1            # reads $env:WIFI_SSID / $env:WIFI_PASSWORD

param(
    [string]$Ssid     = $env:WIFI_SSID,
    [string]$Password = $env:WIFI_PASSWORD
)

if ([string]::IsNullOrWhiteSpace($Ssid)) {
    throw "SSID not supplied. Pass -Ssid or set the WIFI_SSID environment variable."
}
if ([string]::IsNullOrWhiteSpace($Password)) {
    throw "Password not supplied. Pass -Password or set the WIFI_PASSWORD environment variable."
}

$bytes = [System.Text.Encoding]::ASCII.GetBytes($Ssid)
$hex = [BitConverter]::ToString($bytes) -replace '-', ''
Write-Output "SSID: $Ssid"
Write-Output "Hex:  $hex"

$xmlFilePath = Join-Path $PSScriptRoot 'wifi_connect.xml'
if (-not (Test-Path $xmlFilePath)) {
    Copy-Item (Join-Path $PSScriptRoot 'wifi_connect.xml.sample') $xmlFilePath
}

$xml = [xml](Get-Content $xmlFilePath)
$xml.WLANProfile.name = $Ssid
$xml.WLANProfile.SSIDConfig.SSID.name = $Ssid
$xml.WLANProfile.SSIDConfig.SSID.hex = $hex
$xml.WLANProfile.MSM.security.sharedKey.keyMaterial = $Password
$xml.Save($xmlFilePath)

Write-Output "wifi_connect.xml updated successfully."
