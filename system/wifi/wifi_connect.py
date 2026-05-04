# source chatGPT > https://chatgpt.com/c/66e5815c-38a8-8009-a7d7-51cd9d956cff

import logging
import os
import subprocess
import xml.etree.ElementTree as ET

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger(__name__)

def connect_to_wifi():
    try:
        # Get the path to the current folder where the script is located
        script_directory = os.path.dirname(os.path.abspath(__file__))

        # Path to the Wi-Fi profile XML file
        xml_profile_path = os.path.join(script_directory, 'wifi_connect.xml')

        # Parse the XML file to get the SSID
        tree = ET.parse(xml_profile_path)
        root = tree.getroot()
        target_ssid = root.find('.//{http://www.microsoft.com/networking/WLAN/profile/v1}name').text
        log.info("Target WiFi SSID from XML profile: %s", target_ssid)

        # Step 1: Check current Wi-Fi connection status
        log.info("Checking current Wi-Fi connection status...")
        output = subprocess.check_output(['netsh', 'wlan', 'show', 'interfaces'], encoding='utf-8')
        log.info("Output of 'netsh wlan show interfaces':\n%s", output)

        # Parse the output to find the current SSID and interface name
        current_ssid = None
        interface_name = None
        for line in output.splitlines():
            if "SSID" in line and "BSSID" not in line:
                current_ssid = line.split(":")[1].strip()
            if "Name" in line and "wifi integrado" in line:
                interface_name = line.split(":")[1].strip()

        # Make sure we have an interface name, if not, throw an error
        if not interface_name:
            log.error("Error: No interface name detected. Make sure you have a Wi-Fi adapter connected.")
            return

        if current_ssid:
            log.info("Currently connected to: %s", current_ssid)
            if current_ssid == target_ssid:
                log.info("Already connected to %s. No action required.", target_ssid)
                return
            else:
                log.info("Connected to a different Wi-Fi: %s. Disconnecting...", current_ssid)
                disconnect_result = subprocess.run(['netsh', 'wlan', 'disconnect'], capture_output=True, text=True)
                log.info("Output of 'netsh wlan disconnect': %s", disconnect_result.stdout)
                if disconnect_result.returncode != 0:
                    log.error("Error disconnecting: %s", disconnect_result.stderr)
                    return

        # Step 2: Add the Wi-Fi profile using the XML file
        log.info("Adding profile from %s", xml_profile_path)
        add_profile_result = subprocess.run(['netsh', 'wlan', 'add', 'profile', f'filename={xml_profile_path}'], capture_output=True, text=True)
        log.info("Output of 'netsh wlan add profile': %s", add_profile_result.stdout)
        if add_profile_result.returncode != 0:
            log.error("Error adding profile: %s", add_profile_result.stderr)
            return

        # Step 3: Attempt to connect to the Wi-Fi network
        log.info("Attempting to connect to %s using interface %s...", target_ssid, interface_name)
        result = subprocess.run(['netsh', 'wlan', 'connect', f'name={target_ssid}', f'interface={interface_name}'], capture_output=True, text=True)
        log.info("Output of 'netsh wlan connect': %s", result.stdout)
        if result.returncode == 0:
            log.info("Successfully connected to %s.", target_ssid)
        else:
            log.error("Failed to connect to %s: %s", target_ssid, result.stderr)

    except Exception as e:
        log.error("Error: %s", e)

# Example usage
connect_to_wifi()
