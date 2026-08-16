import logging
import os
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

# _netsh.py is a sibling; make it importable when this file is run as a script.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from _netsh import parse_ssid, run_netsh  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger(__name__)

def connect_to_wifi() -> None:
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
        interfaces = run_netsh(['wlan', 'show', 'interfaces'])
        if not interfaces.ran or not interfaces.succeeded or not interfaces.has_output:
            # The query did not answer, so the current state is *unknown* - which
            # is not the same as "not connected". Reconnecting blindly from here
            # would tear down a perfectly good connection.
            if not interfaces.ran:
                reason = f"netsh could not be run ({interfaces.error})"
            elif not interfaces.succeeded:
                reason = f"netsh exited {interfaces.returncode}: {interfaces.stderr.strip()}"
            else:
                reason = "netsh returned no readable output"
            log.error(
                "Error: current Wi-Fi state is unknown - %s. Aborting rather than "
                "assuming disconnected.", reason
            )
            return
        output = interfaces.stdout
        log.info("Output of 'netsh wlan show interfaces':\n%s", output)

        # Parse the output to find the current SSID and interface name
        current_ssid = parse_ssid(output)
        interface_name = None
        for line in output.splitlines():
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
                disconnect_result = run_netsh(['wlan', 'disconnect'])
                log.info("Output of 'netsh wlan disconnect': %s", disconnect_result.stdout)
                if not disconnect_result.succeeded:
                    log.error("Error disconnecting: %s", disconnect_result.stderr)
                    return

        # Step 2: Add the Wi-Fi profile using the XML file
        log.info("Adding profile from %s", xml_profile_path)
        add_profile_result = run_netsh(
            ['wlan', 'add', 'profile', f'filename={xml_profile_path}']
        )
        log.info("Output of 'netsh wlan add profile': %s", add_profile_result.stdout)
        if not add_profile_result.succeeded:
            log.error("Error adding profile: %s", add_profile_result.stderr)
            return

        # Step 3: Attempt to connect to the Wi-Fi network
        log.info("Attempting to connect to %s using interface %s...", target_ssid, interface_name)
        result = run_netsh(
            ['wlan', 'connect', f'name={target_ssid}', f'interface={interface_name}']
        )
        log.info("Output of 'netsh wlan connect': %s", result.stdout)
        if result.succeeded:
            log.info("Successfully connected to %s.", target_ssid)
        else:
            log.error("Failed to connect to %s: %s", target_ssid, result.stderr)

    except Exception as e:
        log.error("Error: %s", e)

if __name__ == "__main__":
    connect_to_wifi()
