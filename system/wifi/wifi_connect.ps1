# Update wifi_connect.xml with a new SSID and password.
# Edit the two values below, then run this script from PowerShell.

$ssid = "REDWIFI_dGYG"
$password = "HfUC6XyRXxb7b4NF"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($ssid)
$hex = [BitConverter]::ToString($bytes) -replace '-', ''
Write-Output "SSID: $ssid"
Write-Output "Hex:  $hex"

$xmlFilePath = Join-Path $PSScriptRoot 'wifi_connect.xml'
$xml = [xml](Get-Content $xmlFilePath)
$xml.WLANProfile.name = $ssid
$xml.WLANProfile.SSIDConfig.SSID.name = $ssid
$xml.WLANProfile.SSIDConfig.SSID.hex = $hex
$xml.WLANProfile.MSM.security.sharedKey.keyMaterial = $password
$xml.Save($xmlFilePath)

Write-Output "wifi_connect.xml updated successfully."
