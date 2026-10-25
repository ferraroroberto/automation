#!/usr/bin/env python3
"""
Network Scanner - ARP-based Local Network Discovery

Performs ARP scans on local network to discover active devices, including Access Points
that may not respond to ICMP ping requests. Essential for finding dynamically assigned
IP addresses on devices like Netgear routers in AP mode.

Why ARP Scan vs ICMP Ping:
- ARP operates at Layer 2 (data link layer) and discovers devices regardless of firewall settings
- ICMP ping operates at Layer 3 and can be blocked by host firewalls or network policies
- Access Points and routers often have ICMP disabled but must respond to ARP for network operation
- ARP scan reveals all active network interfaces, including those with firewall protection

Usage:
    python network_scanner.py
    # Scan default subnet (192.168.0.0/24):
    python network_scanner.py --subnet 192.168.0.0/24
    # Custom output file:
    python network_scanner.py --output network_scan.json
    # Enable debug logging:
    python network_scanner.py --debug

Requirements:
    - Administrative/root privileges required for raw socket operations
    - scapy library: pip install scapy
    - Windows: WinPcap or Npcap must be installed (see documentation)

Output:
    - Comprehensive console table showing IP, MAC, hostname, manufacturer, and device type
    - Detailed network scan summary with device counts and statistics
    - Special focus on Netgear devices for easy AP identification
    - JSON file export with complete device information and metadata

Example Output:
    IP Address       MAC Address         Hostname        Manufacturer       Device Type
    -------------------------------------------------------------------------------------
    192.168.0.1     aa:bb:cc:dd:ee:ff  router           Cisco Systems       Router/AP
    192.168.0.100   11:22:33:44:55:66  netgear-ap       Netgear Inc         Router/AP
    192.168.0.150   77:88:99:aa:bb:cc  iphone           Apple Inc           Mobile Device
    192.168.0.200   88:99:aa:bb:cc:dd  desktop-pc       Dell Inc            Computer

    📊 Network Scan Summary:
       • Total devices found: 4
       • Local network devices: 4
       • Unique manufacturers: 3

    🏭 Manufacturers:
       • Cisco Systems: 1 device(s)
       • Netgear Inc: 1 device(s)
       • Apple Inc: 1 device(s)
       • Dell Inc: 1 device(s)

    📱 Device Types:
       • Router/AP: 2 device(s)
       • Mobile Device: 1 device(s)
       • Computer: 1 device(s)

    🎯 Netgear Devices Found (1):
       • 192.168.0.100 - 11:22:33:44:55:66 (netgear-ap)
       💡 These are likely your routers/access points!

    🏷️ Devices with Hostnames (3):
       • router → 192.168.0.1 (Cisco Systems)
       • netgear-ap → 192.168.0.100 (Netgear Inc)
       • iphone → 192.168.0.150 (Apple Inc)

How to Run:
    Windows (PowerShell as Administrator):
        ⚠️ Requires WinPcap or Npcap installed first!
        .\.venv\Scripts\python.exe network_scanner.py

    Linux/macOS (with sudo):
        sudo ./.venv/bin/python network_scanner.py
"""

import argparse
import json
import logging
import os
import platform
import socket
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any

try:
    from scapy.all import ARP, Ether, srp
    SCAPY_AVAILABLE = True
except ImportError:
    SCAPY_AVAILABLE = False

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# mac_oui_database.py is a sibling; make it importable when this file is run as a script.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from mac_oui_database import MAC_OUI_DATABASE  # noqa: E402


def get_manufacturer(mac_address: str) -> str:
    """
    Look up manufacturer name from MAC address OUI.

    Args:
        mac_address: MAC address in format XX:XX:XX:XX:XX:XX

    Returns:
        Manufacturer name or "Unknown" if not found
    """
    if not mac_address or len(mac_address) < 8:
        return "Unknown"

    # Extract OUI (first 3 bytes) from MAC address
    oui = mac_address.lower()[:8]  # XX:XX:XX format

    return MAC_OUI_DATABASE.get(oui, "Unknown")


def get_hostname(ip_address: str) -> str:
    """
    Perform reverse DNS lookup to get hostname with timeout and threading.

    Args:
        ip_address: IP address to lookup

    Returns:
        Hostname or empty string if lookup fails
    """
    result = {"hostname": ""}

    def dns_lookup():
        """Threaded DNS lookup to prevent blocking."""
        try:
            # Use a very short timeout
            socket.setdefaulttimeout(0.5)  # 0.5 second timeout
            hostname = socket.gethostbyaddr(ip_address)[0]
            result["hostname"] = hostname
        except (socket.herror, socket.gaierror, socket.timeout):
            result["hostname"] = ""
        except Exception:
            result["hostname"] = ""
        finally:
            socket.setdefaulttimeout(None)

    # Start DNS lookup in a separate thread
    lookup_thread = threading.Thread(target=dns_lookup, daemon=True)
    lookup_thread.start()

    # Wait for thread to complete with a timeout
    lookup_thread.join(timeout=0.8)  # Wait max 0.8 seconds

    # If thread is still alive, the lookup timed out
    if lookup_thread.is_alive():
        logger.debug(f"DNS lookup for {ip_address} timed out")
        return ""

    return result["hostname"]


# Vendor-name substrings (lowercase) -> device type, checked in order.
VENDOR_DEVICE_TYPES = (
    ("Router/AP", ("netgear", "tp-link", "cisco", "asus")),
    ("Apple Device", ("apple",)),
    ("Gaming Console", ("sony interactive", "xbox")),
    ("IoT Device", ("raspberry",)),
    ("Smart Device", ("google inc", "amazon")),
    ("Virtual Machine", ("vmware", "virtualbox", "qemu")),
)


def get_device_type(mac_address: str, ip_address: str) -> str:
    """
    Attempt to identify device type based on MAC address patterns and IP.

    Args:
        mac_address: MAC address
        ip_address: IP address

    Returns:
        Device type hint
    """
    if not mac_address:
        return "Unknown"

    manufacturer = get_manufacturer(mac_address).lower()

    # Every vendor-derived class is matched against the resolved vendor name
    # (from MAC_OUI_DATABASE via get_manufacturer), not the raw MAC hex, so
    # the classes track the database instead of a second prefix list.
    for device_type, vendors in VENDOR_DEVICE_TYPES:
        if any(vendor in manufacturer for vendor in vendors):
            return device_type

    # Mobile devices (local IP ranges often indicate phones/tablets)
    ip_parts = ip_address.split('.')
    if len(ip_parts) == 4:
        third_octet = int(ip_parts[2])
        # Common mobile device ranges in home networks
        if third_octet >= 100:  # Many routers assign mobile devices to higher ranges
            return "Mobile Device"

    return "Computer/Network Device"


def get_device_status(ip_address: str) -> Dict[str, Any]:
    """
    Get additional device status information.

    Args:
        ip_address: IP address to check

    Returns:
        Dictionary with device status information
    """
    status = {
        "is_local": False,
        "subnet_hint": "",
        "timestamp": datetime.now().isoformat()
    }

    try:
        # Check if IP is in common local ranges
        ip_parts = ip_address.split('.')
        if len(ip_parts) == 4:
            first_octet = int(ip_parts[0])
            second_octet = int(ip_parts[1])

            if first_octet == 192 and second_octet == 168:
                status["is_local"] = True
                status["subnet_hint"] = "192.168.x.x (Home Network)"
            elif first_octet == 172 and second_octet >= 16 and second_octet <= 31:
                status["is_local"] = True
                status["subnet_hint"] = "172.16-31.x.x (Corporate Network)"
            elif first_octet == 10:
                status["is_local"] = True
                status["subnet_hint"] = "10.x.x.x (Corporate/Private Network)"

    except (ValueError, IndexError):
        pass

    return status


def check_privileges() -> bool:
    """
    Check if script has sufficient privileges for raw socket operations.

    Returns:
        True if privileges are sufficient, False otherwise

    Raises:
        RuntimeError: If privileges are insufficient and cannot be elevated
    """
    try:
        # Test privileged operation - create raw socket
        import socket
        test_socket = socket.socket(socket.AF_INET, socket.SOCK_RAW, socket.IPPROTO_ICMP)
        test_socket.close()
        logger.debug("✅ Administrative privileges confirmed")
        return True
    except PermissionError:
        error_msg = """
❌ Insufficient privileges for network scanning!

This script requires administrative/root privileges to perform ARP scanning.
Raw socket operations need elevated permissions to send and receive network packets.

How to run with proper privileges:

Windows (PowerShell):
    1. Right-click PowerShell and select "Run as Administrator"
    2. Navigate to the script directory
    3. Run: .\.venv\Scripts\python.exe network_scanner.py

Linux/macOS:
    1. Open terminal
    2. Run: sudo ./.venv/bin/python network_scanner.py
    3. Enter your password when prompted

VirtualBox/VM Users:
    - Ensure your VM has "Bridged Adapter" network mode
    - Host-only mode may not show all network devices
        """
        logger.error(error_msg)
        raise RuntimeError("Insufficient privileges for network scanning")
    except Exception as e:
        logger.warning(f"⚠️ Privilege check failed: {e}")
        return False


def arp_scan(subnet: str, timeout: int = 2, verbose: bool = True) -> List[Dict[str, str]]:
    """
    Perform ARP scan on specified subnet to discover active devices.

    Args:
        subnet: Network subnet in CIDR notation (e.g., "192.168.0.0/24")
        timeout: Timeout in seconds for ARP responses
        verbose: Whether to show progress messages

    Returns:
        List of dictionaries containing IP, MAC, and manufacturer info

    Raises:
        RuntimeError: If ARP scan fails
    """
    if not SCAPY_AVAILABLE:
        raise RuntimeError("scapy library is required but not installed. Run: pip install scapy")

    if verbose:
        logger.info(f"🔍 Scanning subnet: {subnet}")
        logger.info("⏳ Sending ARP requests and waiting for responses...")

    try:
        # Create ARP request packet
        arp_request = ARP(pdst=subnet)

        # Create Ethernet broadcast frame
        broadcast = Ether(dst="ff:ff:ff:ff:ff:ff")

        # Combine packets
        arp_request_broadcast = broadcast / arp_request

        # Send packets and receive responses
        # Set timeout and retry parameters
        answered, unanswered = srp(
            arp_request_broadcast,
            timeout=timeout,
            retry=2,
            verbose=0  # Suppress scapy verbose output
        )

        devices = []
        for sent, received in answered:
            ip_addr = received.psrc
            mac_addr = received.hwsrc

            # Gather comprehensive device information
            device_info = {
                "ip": ip_addr,
                "mac": mac_addr,
                "hostname": get_hostname(ip_addr),
                "manufacturer": get_manufacturer(mac_addr),
                "device_type": get_device_type(mac_addr, ip_addr),
                "status": get_device_status(ip_addr),
                "arp_details": {
                    "operation": received.op,
                    "hardware_type": received.hwtype,
                    "protocol_type": received.ptype,
                    "hardware_len": received.hwlen,
                    "protocol_len": received.plen
                }
            }
            devices.append(device_info)

        if verbose:
            logger.info(f"✅ ARP scan completed - found {len(devices)} active devices")

        # Sort by IP address for consistent output
        devices.sort(key=lambda x: [int(i) for i in x['ip'].split('.')])

        return devices

    except Exception as e:
        error_msg = str(e)

        # Handle Windows-specific winpcap issue
        if "winpcap is not installed" in error_msg or "pcap won't be used" in error_msg:
            if platform.system() == "Windows":
                logger.error("❌ Windows ARP scanning requires WinPcap or Npcap to be installed")
                logger.error("📥 Please install one of the following:")
                logger.error("   • Npcap: https://npcap.com/#download")
                logger.error("   • WinPcap: https://www.winpcap.org/install/")
                logger.error("   💡 Npcap is recommended for modern Windows versions")
                logger.error("")
                logger.error("🔄 Alternative: Use a Linux VM or WSL for network scanning")
                logger.error("   WSL command: sudo apt install python3-scapy")
            else:
                logger.error(f"❌ ARP scan failed: {e}")

        raise RuntimeError(f"ARP scan failed: {error_msg}")


def print_table(devices: List[Dict[str, Any]]) -> None:
    """
    Print comprehensive device information in a formatted table.

    Args:
        devices: List of device dictionaries with extended information
    """
    if not devices:
        logger.info("ℹ️ No active devices found on the network.")
        print("\nNo devices discovered. Possible causes:")
        print("- Network may be different from expected subnet")
        print("- Firewall blocking ARP responses")
        print("- Devices may be powered off or disconnected")
        print("- Virtual machine network configuration issues")
        return

    # Calculate column widths
    ip_width = max(len(device["ip"]) for device in devices) if devices else 15
    mac_width = max(len(device["mac"]) for device in devices) if devices else 17
    hostname_width = max(len(device.get("hostname", "")) for device in devices) if devices else 12
    manuf_width = max(len(device["manufacturer"]) for device in devices) if devices else 15
    type_width = max(len(device.get("device_type", "")) for device in devices) if devices else 12

    # Ensure minimum widths for readability
    ip_width = max(ip_width, 15)
    mac_width = max(mac_width, 17)
    hostname_width = max(hostname_width, 12)
    manuf_width = max(manuf_width, 15)
    type_width = max(type_width, 12)

    # Header
    header = f"{'IP Address'.ljust(ip_width)}  {'MAC Address'.ljust(mac_width)}  {'Hostname'.ljust(hostname_width)}  {'Manufacturer'.ljust(manuf_width)}  {'Device Type'.ljust(type_width)}"
    print(f"\n{header}")
    print("-" * len(header))

    for device in devices:
        hostname = device.get("hostname", "") or "-"
        device_type = device.get("device_type", "") or "Unknown"
        print(f"{device['ip'].ljust(ip_width)}  {device['mac'].ljust(mac_width)}  {hostname.ljust(hostname_width)}  {device['manufacturer'].ljust(manuf_width)}  {device_type.ljust(type_width)}")

    # Show detailed summary
    print(f"\n📊 Network Scan Summary:")
    print(f"   • Total devices found: {len(devices)}")

    # Count by manufacturer
    manufacturers = {}
    device_types = {}
    netgear_devices = []
    local_devices = 0

    for device in devices:
        manuf = device["manufacturer"]
        dev_type = device.get("device_type", "Unknown")

        manufacturers[manuf] = manufacturers.get(manuf, 0) + 1
        device_types[dev_type] = device_types.get(dev_type, 0) + 1

        if "netgear" in manuf.lower():
            netgear_devices.append(device)

        if device.get("status", {}).get("is_local", False):
            local_devices += 1

    print(f"   • Local network devices: {local_devices}")
    print(f"   • Unique manufacturers: {len([m for m in manufacturers.keys() if m != 'Unknown'])}")

    if manufacturers:
        print(f"\n🏭 Manufacturers:")
        for manuf, count in sorted(manufacturers.items(), key=lambda x: x[1], reverse=True):
            if manuf != "Unknown":
                print(f"   • {manuf}: {count} device(s)")

    if device_types:
        print(f"\n📱 Device Types:")
        for dev_type, count in sorted(device_types.items(), key=lambda x: x[1], reverse=True):
            print(f"   • {dev_type}: {count} device(s)")

    # Special Netgear focus
    if netgear_devices:
        print(f"\n🎯 Netgear Devices Found ({len(netgear_devices)}):")
        for device in netgear_devices:
            hostname = device.get("hostname", "") or "No hostname"
            print(f"   • {device['ip']} - {device['mac']} ({hostname})")
        print("   💡 These are likely your routers/access points!")
        print("   💡 Look for your R9000 by checking the web interface at http://[IP]:8080")

    # Show devices with hostnames
    devices_with_names = [d for d in devices if d.get("hostname")]
    if devices_with_names:
        print(f"\n🏷️ Devices with Hostnames ({len(devices_with_names)}):")
        for device in devices_with_names:
            print(f"   • {device['hostname']} → {device['ip']} ({device['manufacturer']})")


def ask_save_json(devices: List[Dict[str, str]], default_output: str, config: Dict[str, Any]) -> Optional[str]:
    """
    Ask user if they want to save scan results to JSON file using tkinter dialog.

    Args:
        devices: List of device dictionaries
        default_output: Default output file path
        config: Configuration dictionary

    Returns:
        Path to save file if user chooses to save, None otherwise
    """
    # Create root window (will be hidden)
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Ask user if they want to save
    result = messagebox.askyesno(
        "Save Results",
        f"Network scan found {len(devices)} devices.\n\nWould you like to save the results to a JSON file?"
    )

    if not result:
        root.destroy()
        return None

    # Show file dialog to choose save location
    file_path = filedialog.asksaveasfilename(
        title="Save Network Scan Results",
        defaultextension=".json",
        filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        initialfile=os.path.basename(default_output)
    )

    root.destroy()

    if file_path:  # User didn't cancel
        return file_path
    else:
        return None


def save_json(devices: List[Dict[str, str]], output_file: str, config: Dict[str, Any]) -> None:
    """
    Save device information to JSON file.

    Args:
        devices: List of device dictionaries
        output_file: Path to output JSON file
        config: Configuration dictionary
    """
    logger.debug(f"💾 Saving {len(devices)} devices to JSON file: {output_file}")

    # Prepare data for JSON output
    manufacturers = {}
    device_types = {}
    netgear_count = 0
    hostname_count = 0
    local_count = 0

    for device in devices:
        manuf = device["manufacturer"]
        dev_type = device.get("device_type", "Unknown")

        manufacturers[manuf] = manufacturers.get(manuf, 0) + 1
        device_types[dev_type] = device_types.get(dev_type, 0) + 1

        if "netgear" in manuf.lower():
            netgear_count += 1

        if device.get("hostname"):
            hostname_count += 1

        if device.get("status", {}).get("is_local", False):
            local_count += 1

    output_data = {
        "devices": devices,
        "scan_summary": {
            "total_devices": len(devices),
            "netgear_devices": netgear_count,
            "devices_with_hostnames": hostname_count,
            "local_network_devices": local_count,
            "subnet_scanned": config.get("default_subnet", "192.168.0.0/24"),
            "scan_timestamp": datetime.now().isoformat(),
            "manufacturers": manufacturers,
            "device_types": device_types
        }
    }

    # Add metadata if configured
    if config.get("output_format", {}).get("include_metadata", True):
        output_data["metadata"] = {
            "scanner_version": "1.0.0",
            "source": "ARP Network Scanner",
            "config_source": "network_scanner.json",
            "scapy_version": "Available" if SCAPY_AVAILABLE else "Not Available"
        }

    try:
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(
                output_data,
                f,
                indent=2 if config.get("output_format", {}).get("pretty_print", True) else None,
                ensure_ascii=config.get("output_format", {}).get("ensure_ascii", False)
            )

        logger.info(f"✅ Successfully saved {len(devices)} devices to {output_file}")

    except Exception as e:
        logger.error(f"❌ Failed to save JSON file: {e}")
        raise


def load_config() -> Dict[str, Any]:
    """
    Load configuration from JSON file with fallback defaults.

    Returns:
        Configuration dictionary
    """
    config_path = Path(__file__).parent / "network_scanner.json"

    try:
        if config_path.exists():
            logger.debug(f"📂 Loading configuration from: {config_path}")
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            logger.info("✅ Configuration loaded successfully")
            return config
        else:
            logger.warning(f"⚠️ Configuration file not found at {config_path}, using defaults")
    except Exception as e:
        logger.warning(f"⚠️ Failed to load configuration: {e}, using defaults")

    # Fallback default configuration
    return {
        "default_subnet": "192.168.0.0/24",
        "default_output": "network_scan.json",
        "scan_timeout": 2,
        "logging": {
            "level": "INFO",
            "format": "%(asctime)s - %(levelname)s - %(message)s"
        },
        "output_format": {
            "include_metadata": True,
            "pretty_print": True,
            "ensure_ascii": False
        },
        "arp_scan": {
            "timeout": 2,
            "retry_count": 2,
            "verbose_output": True
        }
    }


def _build_arg_parser(config: Dict[str, Any]) -> argparse.ArgumentParser:
    """Build the command-line parser, defaulting from the config."""
    parser = argparse.ArgumentParser(
        description="ARP-based network scanner for local device discovery",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python network_scanner.py
  python network_scanner.py --subnet 192.168.1.0/24
  python network_scanner.py --output my_scan.json --debug

Privileges Required:
  Windows: Run as Administrator
  Linux/macOS: Run with sudo

Dependencies:
  pip install scapy
            """
    )

    parser.add_argument(
        "--subnet", "-s",
        default=config["default_subnet"],
        help=f"Network subnet to scan in CIDR notation (default: {config['default_subnet']})"
    )

    parser.add_argument(
        "--output", "-o",
        default=config["default_output"],
        help=f"Output JSON file path (default: {config['default_output']})"
    )

    parser.add_argument(
        "--timeout", "-t",
        type=int,
        default=config.get("arp_scan", {}).get("timeout", 2),
        help="ARP response timeout in seconds (default: 2)"
    )

    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug logging"
    )
    return parser


def _log_netgear_guidance(devices: List[Dict[str, Any]]) -> None:
    """Log the Netgear devices found, or hints on why none were."""
    netgear_devices = [d for d in devices if "netgear" in d["manufacturer"].lower()]
    if netgear_devices:
        logger.info("🎯 Netgear devices found:")
        for device in netgear_devices:
            logger.info("   • %s - %s (%s)", device['ip'], device['mac'], device['manufacturer'])
        logger.info("   💡 The Netgear R9000 AP should be among these devices!")
    else:
        logger.warning("⚠️ No Netgear devices found.")
        logger.info("   This could mean:")
        logger.info("   • Your Netgear AP is on a different subnet")
        logger.info("   • The AP has a different MAC OUI than expected")
        logger.info("   • Try scanning other common subnets (192.168.1.0/24, 192.168.10.0/24, etc.)")


def main(config: Optional[Dict[str, Any]] = None) -> None:
    """
    Main function to perform network scanning.

    Args:
        config: Configuration dictionary (optional)
    """
    if config is None:
        config = load_config()

    args = None  # Guard against UnboundLocalError in the except block below if
                 # argument parsing itself raises before `args` is assigned.
    try:
        parser = _build_arg_parser(config)
        args = parser.parse_args()

        # Set logging level based on debug flag or config
        if args.debug:
            logger.setLevel(logging.DEBUG)
            logger.debug("🔍 Debug mode enabled")
        else:
            log_level = config.get("logging", {}).get("level", "INFO")
            logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))

        logger.info("🚀 Starting network discovery scan")
        logger.info("📡 Scanning for active network devices...")

        # Check privileges before attempting scan
        check_privileges()

        # Perform ARP scan
        devices = arp_scan(args.subnet, args.timeout)

        # Display results
        print_table(devices)

        # Ask user if they want to save results to JSON
        save_path = ask_save_json(devices, args.output, config)

        logger.info("✅ Network scan completed successfully!")
        logger.info("📊 Total devices found: %d", len(devices))

        if save_path:
            save_json(devices, save_path, config)
            logger.info("💾 Results saved to: %s", save_path)
        else:
            logger.info("💾 Results not saved (user cancelled)")

        _log_netgear_guidance(devices)

    except KeyboardInterrupt:
        logger.info("⏹️  Network scan cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"❌ Network scan failed: {e}")
        if args is not None and args.debug:
            logger.exception("Full traceback:")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
        logger.info("✅ Script completed successfully")
    except Exception as e:
        logger.error(f"❌ Script failed: {e}")
        sys.exit(1)