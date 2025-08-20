# Replace "YourWiFiName" with your actual WiFi network name
$ssid = "REDWIFI_dGYG"
$bytes = [System.Text.Encoding]::ASCII.GetBytes($ssid)
$hex = [BitConverter]::ToString($bytes) -replace '-',''
Write-Output "SSID: $ssid"
Write-Output "Hex: $hex"

# Now let's update the XML file with these values
$xmlFilePath = "e:\onedrive\Documentos\Roberto\projects\automation\notion-automation\home\system\connectwifi.xml"
$xml = [xml](Get-Content $xmlFilePath)
$xml.WLANProfile.name = $ssid
$xml.WLANProfile.SSIDConfig.SSID.name = $ssid
$xml.WLANProfile.SSIDConfig.SSID.hex = $hex
$xml.WLANProfile.MSM.security.sharedKey.keyMaterial = "HfUC6XyRXxb7b4NF"
$xml.Save($xmlFilePath)

Write-Output "WiFi profile XML has been updated successfully!"
