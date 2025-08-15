# source chatGPT > https://chatgpt.com/c/66e5815c-38a8-8009-a7d7-51cd9d956cff

import os
import subprocess
import xml.etree.ElementTree as ET

def connect_to_wifi():
    try:
        # Get the path to the current folder where the script is located
        script_directory = os.path.dirname(os.path.abspath(__file__))

        # Path to the Wi-Fi profile XML file
        xml_profile_path = os.path.join(script_directory, 'connectwifi.xml')
        
        # Parse the XML file to get the SSID
        tree = ET.parse(xml_profile_path)
        root = tree.getroot()
        target_ssid = root.find('.//{http://www.microsoft.com/networking/WLAN/profile/v1}name').text
        print(f"Target WiFi SSID from XML profile: {target_ssid}")

        # Step 1: Check current Wi-Fi connection status
        print("Checking current Wi-Fi connection status...")
        output = subprocess.check_output(['netsh', 'wlan', 'show', 'interfaces'], encoding='utf-8')
        print(f"Output of 'netsh wlan show interfaces':\n{output}")

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
            print("Error: No interface name detected. Make sure you have a Wi-Fi adapter connected.")
            return

        if current_ssid:
            print(f"Currently connected to: {current_ssid}")
            if current_ssid == target_ssid:
                print(f"Already connected to {target_ssid}. No action required.")
                return
            else:
                print(f"Connected to a different Wi-Fi: {current_ssid}. Disconnecting...")
                disconnect_result = subprocess.run(['netsh', 'wlan', 'disconnect'], capture_output=True, text=True)
                print(f"Output of 'netsh wlan disconnect': {disconnect_result.stdout}")
                if disconnect_result.returncode != 0:
                    print(f"Error disconnecting: {disconnect_result.stderr}")
                    return

        # Step 2: Add the Wi-Fi profile using the XML file
        print(f"Adding profile from {xml_profile_path}")
        add_profile_result = subprocess.run(['netsh', 'wlan', 'add', 'profile', f'filename={xml_profile_path}'], capture_output=True, text=True)
        print(f"Output of 'netsh wlan add profile': {add_profile_result.stdout}")
        if add_profile_result.returncode != 0:
            print(f"Error adding profile: {add_profile_result.stderr}")
            return

        # Step 3: Attempt to connect to the Wi-Fi network
        print(f"Attempting to connect to {target_ssid} using interface {interface_name}...")
        result = subprocess.run(['netsh', 'wlan', 'connect', f'name={target_ssid}', f'interface={interface_name}'], capture_output=True, text=True)
        print(f"Output of 'netsh wlan connect': {result.stdout}")
        if result.returncode == 0:
            print(f"Successfully connected to {target_ssid}.")
        else:
            print(f"Failed to connect to {target_ssid}: {result.stderr}")

    except Exception as e:
        print(f"Error: {str(e)}")

# Example usage
connect_to_wifi()
