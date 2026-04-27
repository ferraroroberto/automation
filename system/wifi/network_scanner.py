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
        sudo ./venv/bin/python network_scanner.py

    Or activate virtual environment and run:
        source venv/bin/activate  # Linux/macOS
        .\.venv\Scripts\activate  # Windows PowerShell
        python network_scanner.py
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
from typing import Dict, List, Optional, Tuple, Any

try:
    from scapy.all import ARP, Ether, srp, conf
    SCAPY_AVAILABLE = True
except ImportError:
    SCAPY_AVAILABLE = False

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# MAC Address OUI (Organizationally Unique Identifier) database
# Mapping of OUI prefixes to manufacturer names
MAC_OUI_DATABASE = {
    # Netgear devices (comprehensive list)
    "00:14:6c": "Netgear Inc",
    "20:76:93": "Netgear Inc",
    "08:36:c9": "Netgear Inc",
    "9c:3d:cf": "Netgear Inc",
    "b0:39:56": "Netgear Inc",
    "00:0f:b5": "Netgear Inc",
    "00:09:5b": "Netgear Inc",
    "00:18:4d": "Netgear Inc",
    "00:24:b2": "Netgear Inc",
    "04:a1:51": "Netgear Inc",
    "08:bd:43": "Netgear Inc",
    "0c:60:76": "Netgear Inc",
    "10:da:43": "Netgear Inc",
    "1c:b7:2c": "Netgear Inc",
    "20:e5:2a": "Netgear Inc",
    "28:c6:8e": "Netgear Inc",
    "2c:b0:5d": "Netgear Inc",
    "30:46:9a": "Netgear Inc",
    "34:98:b5": "Netgear Inc",
    "40:5d:82": "Netgear Inc",
    "44:94:fc": "Netgear Inc",
    "4c:60:de": "Netgear Inc",
    "50:6a:03": "Netgear Inc",
    "6c:b0:ce": "Netgear Inc",
    "74:44:01": "Netgear Inc",
    "84:1b:5e": "Netgear Inc",
    "88:03:55": "Netgear Inc",
    "8c:3b:ad": "Netgear Inc",
    "94:10:3e": "Netgear Inc",
    "9c:c9:eb": "Netgear Inc",
    "a0:04:60": "Netgear Inc",
    "a4:2b:b0": "Netgear Inc",
    "ac:60:b6": "Netgear Inc",
    "b0:70:2d": "Netgear Inc",
    "bc:44:86": "Netgear Inc",
    "c0:25:06": "Netgear Inc",
    "c4:04:15": "Netgear Inc",
    "c8:d7:19": "Netgear Inc",
    "cc:40:d0": "Netgear Inc",
    "d0:7e:35": "Netgear Inc",
    "dc:ef:ca": "Netgear Inc",
    "e0:46:9a": "Netgear Inc",
    "e8:fc:af": "Netgear Inc",
    "ec:1a:59": "Netgear Inc",
    "f0:da:7c": "Netgear Inc",

    # Virtualization and Cloud
    "00:50:56": "VMware Inc",
    "00:0c:29": "VMware Inc",
    "00:05:69": "VMware Inc",
    "00:1c:14": "VMware Inc",
    "00:1c:42": "Parallels Inc",
    "08:00:27": "Oracle VirtualBox",
    "52:54:00": "QEMU/KVM",
    "02:42:ac": "Docker Container",
    "06:0e:2b": "AWS EC2",
    "0e:0e:0e": "Google Cloud",

    # Microsoft and Windows
    "00:03:ff": "Microsoft Corporation",
    "00:15:5d": "Microsoft Corporation",
    "00:50:f2": "Microsoft Corporation",
    "7c:ed:8d": "Microsoft Corporation",
    "8c:60:4f": "Microsoft Corporation",
    "9c:b6:54": "Microsoft Corporation",

    # Apple devices
    "28:cd:c1": "Apple Inc",
    "8c:85:90": "Apple Inc",
    "ac:bc:32": "Apple Inc",
    "f0:18:98": "Apple Inc",
    "60:33:4b": "Apple Inc",
    "24:ab:81": "Apple Inc",
    "68:5b:35": "Apple Inc",
    "a4:c3:61": "Apple Inc",
    "b8:17:c2": "Apple Inc",

    # Cisco and networking
    "00:50:cc": "Cisco Systems",
    "00:0a:f7": "Broadcom",
    "00:10:18": "Broadcom",
    "00:90:4c": "Epigram Inc",
    "00:01:42": "Parallels Inc",
    "00:1b:21": "Intel Corporate",
    "00:21:5c": "Intel Corporate",
    "00:26:c6": "Intel Corporate",
    "a0:36:9f": "Intel Corporate",
    "a0:36:9f": "Intel Corporate",
    "b8:8a:60": "Intel Corporate",
    "dc:53:7c": "Compal Broadband Networks",
    "e0:b9:ba": "Cisco Systems",

    # IoT and embedded devices
    "b8:27:eb": "Raspberry Pi Foundation",
    "dc:a6:32": "Raspberry Pi Foundation",
    "e4:5f:01": "Raspberry Pi Foundation",
    "d8:3a:dd": "Samsung Electronics",
    "00:12:fb": "Samsung Electronics",
    "bc:76:70": "Samsung Electronics",
    "8c:79:67": "Samsung Electronics",
    "b4:07:f9": "Samsung Electronics",
    "50:c8:e5": "Samsung Electronics",
    "00:0e:c6": "ASUSTek Computer Inc",
    "00:13:d4": "ASUSTek Computer Inc",
    "00:17:31": "ASUSTek Computer Inc",
    "00:1a:92": "ASUSTek Computer Inc",
    "00:1b:fc": "ASUSTek Computer Inc",
    "00:1d:60": "ASUSTek Computer Inc",
    "00:1e:8c": "ASUSTek Computer Inc",
    "00:22:15": "ASUSTek Computer Inc",
    "00:24:8c": "ASUSTek Computer Inc",
    "00:26:18": "ASUSTek Computer Inc",
    "04:92:26": "ASUSTek Computer Inc",
    "08:62:66": "ASUSTek Computer Inc",
    "10:bf:48": "ASUSTek Computer Inc",
    "14:da:e9": "ASUSTek Computer Inc",
    "18:31:bf": "ASUSTek Computer Inc",
    "1c:b7:2c": "ASUSTek Computer Inc",
    "20:16:d8": "ASUSTek Computer Inc",
    "24:5e:be": "ASUSTek Computer Inc",
    "2c:4d:54": "ASUSTek Computer Inc",
    "2c:56:dc": "ASUSTek Computer Inc",
    "2c:fd:a1": "ASUSTek Computer Inc",
    "30:85:a9": "ASUSTek Computer Inc",
    "38:2c:4a": "ASUSTek Computer Inc",
    "3c:7c:3f": "ASUSTek Computer Inc",
    "40:b0:76": "ASUSTek Computer Inc",
    "44:8a:5b": "ASUSTek Computer Inc",
    "48:5b:39": "ASUSTek Computer Inc",
    "4c:ed:de": "ASUSTek Computer Inc",
    "50:46:5d": "ASUSTek Computer Inc",
    "54:a0:50": "ASUSTek Computer Inc",
    "60:a4:4c": "ASUSTek Computer Inc",
    "64:5d:86": "ASUSTek Computer Inc",
    "68:3e:34": "ASUSTek Computer Inc",
    "70:4d:7b": "ASUSTek Computer Inc",
    "70:8b:cd": "ASUSTek Computer Inc",
    "74:05:a5": "ASUSTek Computer Inc",
    "74:d0:2b": "ASUSTek Computer Inc",
    "78:24:af": "ASUSTek Computer Inc",
    "78:8a:20": "ASUSTek Computer Inc",
    "7c:87:ce": "ASUSTek Computer Inc",
    "80:2a:a8": "ASUSTek Computer Inc",
    "84:00:d2": "ASUSTek Computer Inc",
    "84:16:f9": "ASUSTek Computer Inc",
    "84:4b:f5": "ASUSTek Computer Inc",
    "88:51:fb": "ASUSTek Computer Inc",
    "8c:0d:76": "ASUSTek Computer Inc",
    "8c:7c:92": "ASUSTek Computer Inc",
    "8c:a9:82": "ASUSTek Computer Inc",
    "90:55:de": "ASUSTek Computer Inc",
    "90:6d:c8": "ASUSTek Computer Inc",
    "94:44:52": "ASUSTek Computer Inc",
    "98:3b:16": "ASUSTek Computer Inc",
    "9c:5c:8e": "ASUSTek Computer Inc",
    "a0:2b:b8": "ASUSTek Computer Inc",
    "a8:5e:45": "ASUSTek Computer Inc",
    "ac:22:0b": "ASUSTek Computer Inc",
    "ac:60:b6": "ASUSTek Computer Inc",
    "ac:9e:17": "ASUSTek Computer Inc",
    "b0:6e:bf": "ASUSTek Computer Inc",
    "b0:70:2d": "ASUSTek Computer Inc",
    "b8:88:e3": "ASUSTek Computer Inc",
    "bc:54:36": "ASUSTek Computer Inc",
    "bc:ae:c5": "ASUSTek Computer Inc",
    "c4:56:00": "ASUSTek Computer Inc",
    "c4:57:6e": "ASUSTek Computer Inc",
    "c8:60:00": "ASUSTek Computer Inc",
    "cc:b0:da": "ASUSTek Computer Inc",
    "d0:17:c2": "ASUSTek Computer Inc",
    "d0:53:49": "ASUSTek Computer Inc",
    "d4:5d:64": "ASUSTek Computer Inc",
    "d8:50:e6": "ASUSTek Computer Inc",
    "dc:39:6f": "ASUSTek Computer Inc",
    "e0:3f:49": "ASUSTek Computer Inc",
    "e0:69:95": "ASUSTek Computer Inc",
    "e4:95:6e": "ASUSTek Computer Inc",
    "e8:94:f6": "ASUSTek Computer Inc",
    "ec:1a:59": "ASUSTek Computer Inc",
    "f0:79:59": "ASUSTek Computer Inc",
    "f4:6d:e2": "ASUSTek Computer Inc",
    "f8:32:e4": "ASUSTek Computer Inc",
    "fc:34:97": "ASUSTek Computer Inc",

    # TP-Link and networking equipment
    "00:21:27": "TP-Link Technologies",
    "00:27:19": "TP-Link Technologies",
    "00:31:92": "TP-Link Technologies",
    "00:5f:67": "TP-Link Technologies",
    "10:fe:ed": "TP-Link Technologies",
    "14:cc:20": "TP-Link Technologies",
    "1c:61:b4": "TP-Link Technologies",
    "20:dc:e6": "TP-Link Technologies",
    "24:69:a5": "TP-Link Technologies",
    "28:10:7b": "TP-Link Technologies",
    "2c:4d:79": "TP-Link Technologies",
    "30:de:4b": "TP-Link Technologies",
    "34:96:72": "TP-Link Technologies",
    "3c:46:d8": "TP-Link Technologies",
    "40:31:3c": "TP-Link Technologies",
    "44:b3:2d": "TP-Link Technologies",
    "48:22:54": "TP-Link Technologies",
    "4c:32:75": "TP-Link Technologies",
    "50:bd:5f": "TP-Link Technologies",
    "54:af:97": "TP-Link Technologies",
    "58:d9:d5": "TP-Link Technologies",
    "5c:63:bf": "TP-Link Technologies",
    "60:e3:27": "TP-Link Technologies",
    "64:70:02": "TP-Link Technologies",
    "68:ff:7b": "TP-Link Technologies",
    "6c:e8:74": "TP-Link Technologies",
    "70:4f:57": "TP-Link Technologies",
    "74:83:c2": "TP-Link Technologies",
    "78:11:dc": "TP-Link Technologies",
    "78:32:1b": "TP-Link Technologies",
    "78:55:17": "TP-Link Technologies",
    "78:8c:b5": "TP-Link Technologies",
    "78:d3:4f": "TP-Link Technologies",
    "7c:8b:ca": "TP-Link Technologies",
    "80:89:17": "TP-Link Technologies",
    "84:16:f9": "TP-Link Technologies",
    "88:5d:dd": "TP-Link Technologies",
    "8c:21:0a": "TP-Link Technologies",
    "90:9a:4a": "TP-Link Technologies",
    "90:f6:52": "TP-Link Technologies",
    "94:0c:6d": "TP-Link Technologies",
    "98:10:e8": "TP-Link Technologies",
    "98:de:d0": "TP-Link Technologies",
    "9c:21:6a": "TP-Link Technologies",
    "9c:a2:f4": "TP-Link Technologies",
    "a0:36:9f": "TP-Link Technologies",
    "a0:f3:c1": "TP-Link Technologies",
    "a8:15:4d": "TP-Link Technologies",
    "ac:15:a2": "TP-Link Technologies",
    "ac:60:b6": "TP-Link Technologies",
    "ac:84:c6": "TP-Link Technologies",
    "ac:84:c9": "TP-Link Technologies",
    "b0:4e:26": "TP-Link Technologies",
    "b0:95:75": "TP-Link Technologies",
    "b4:b0:24": "TP-Link Technologies",
    "bc:46:99": "TP-Link Technologies",
    "c0:4a:00": "TP-Link Technologies",
    "c4:12:f5": "TP-Link Technologies",
    "c4:e9:84": "TP-Link Technologies",
    "cc:32:e5": "TP-Link Technologies",
    "d4:21:22": "TP-Link Technologies",
    "d4:6a:6a": "TP-Link Technologies",
    "d8:0d:17": "TP-Link Technologies",
    "d8:5d:e2": "TP-Link Technologies",
    "dc:02:8e": "TP-Link Technologies",
    "e0:05:c5": "TP-Link Technologies",
    "e4:95:6e": "TP-Link Technologies",
    "e8:65:49": "TP-Link Technologies",
    "e8:9f:80": "TP-Link Technologies",
    "ec:0e:c4": "TP-Link Technologies",
    "ec:17:2f": "TP-Link Technologies",
    "ec:23:68": "TP-Link Technologies",
    "ec:26:ca": "TP-Link Technologies",
    "ec:33:bb": "TP-Link Technologies",
    "ec:41:18": "TP-Link Technologies",
    "f0:5c:19": "TP-Link Technologies",
    "f4:83:cd": "TP-Link Technologies",
    "f4:ec:38": "TP-Link Technologies",
    "f8:1a:67": "TP-Link Technologies",
    "f8:8e:85": "TP-Link Technologies",
    "fc:34:97": "TP-Link Technologies",

    # Google/Nest devices
    "00:1a:11": "Google Inc",
    "3c:5a:b4": "Google Inc",
    "54:60:09": "Google Inc",
    "94:eb:2c": "Google Inc",
    "a4:77:33": "Google Inc",
    "b4:43:0d": "Google Inc",
    "d8:6c:63": "Google Inc",
    "f4:f5:24": "Google Inc",

    # Amazon devices
    "00:fc:8b": "Amazon Technologies Inc",
    "18:74:2e": "Amazon Technologies Inc",
    "22:47:da": "Amazon Technologies Inc",
    "34:d2:70": "Amazon Technologies Inc",
    "40:b4:cd": "Amazon Technologies Inc",
    "44:65:0d": "Amazon Technologies Inc",
    "50:f5:da": "Amazon Technologies Inc",
    "68:9e:19": "Amazon Technologies Inc",
    "74:75:48": "Amazon Technologies Inc",
    "78:e1:03": "Amazon Technologies Inc",
    "84:d6:d0": "Amazon Technologies Inc",
    "ac:63:be": "Amazon Technologies Inc",
    "b4:7c:9c": "Amazon Technologies Inc",
    "b8:8d:12": "Amazon Technologies Inc",
    "c0:56:27": "Amazon Technologies Inc",
    "cc:7b:35": "Amazon Technologies Inc",
    "d0:07:90": "Amazon Technologies Inc",
    "dc:0b:1a": "Amazon Technologies Inc",
    "f0:27:2d": "Amazon Technologies Inc",
    "f0:d2:f1": "Amazon Technologies Inc",

    # Gaming consoles
    "00:03:0f": "Sony Interactive Entertainment",
    "00:0d:93": "Sony Interactive Entertainment",
    "00:19:c5": "Sony Interactive Entertainment",
    "00:1b:0a": "Sony Interactive Entertainment",
    "00:1d:0d": "Sony Interactive Entertainment",
    "00:1e:a9": "Sony Interactive Entertainment",
    "00:21:9e": "Sony Interactive Entertainment",
    "00:22:aa": "Sony Interactive Entertainment",
    "00:24:8d": "Sony Interactive Entertainment",
    "00:26:43": "Sony Interactive Entertainment",
    "00:73:96": "Sony Interactive Entertainment",
    "04:25:9c": "Sony Interactive Entertainment",
    "08:00:46": "Sony Interactive Entertainment",
    "0c:fe:45": "Sony Interactive Entertainment",
    "1c:7e:51": "Sony Interactive Entertainment",
    "1c:96:5a": "Sony Interactive Entertainment",
    "20:3a:ef": "Sony Interactive Entertainment",
    "24:21:ab": "Sony Interactive Entertainment",
    "2c:22:8b": "Sony Interactive Entertainment",
    "2c:56:dc": "Sony Interactive Entertainment",
    "30:52:cb": "Sony Interactive Entertainment",
    "34:31:11": "Sony Interactive Entertainment",
    "40:b8:37": "Sony Interactive Entertainment",
    "44:1c:a8": "Sony Interactive Entertainment",
    "48:0f:cf": "Sony Interactive Entertainment",
    "4c:21:d0": "Sony Interactive Entertainment",
    "50:1a:c5": "Sony Interactive Entertainment",
    "50:67:f0": "Sony Interactive Entertainment",
    "54:42:49": "Sony Interactive Entertainment",
    "58:48:22": "Sony Interactive Entertainment",
    "5c:96:9d": "Sony Interactive Entertainment",
    "60:5f:8d": "Sony Interactive Entertainment",
    "64:4b:f0": "Sony Interactive Entertainment",
    "64:a6:51": "Sony Interactive Entertainment",
    "68:96:7b": "Sony Interactive Entertainment",
    "6c:c2:17": "Sony Interactive Entertainment",
    "70:9e:29": "Sony Interactive Entertainment",
    "74:63:df": "Sony Interactive Entertainment",
    "78:18:81": "Sony Interactive Entertainment",
    "7c:04:d0": "Sony Interactive Entertainment",
    "7c:8b:ca": "Sony Interactive Entertainment",
    "84:25:3f": "Sony Interactive Entertainment",
    "84:89:ad": "Sony Interactive Entertainment",
    "84:a4:23": "Sony Interactive Entertainment",
    "88:03:55": "Sony Interactive Entertainment",
    "8c:41:f2": "Sony Interactive Entertainment",
    "90:03:b7": "Sony Interactive Entertainment",
    "94:44:52": "Sony Interactive Entertainment",
    "98:0c:a5": "Sony Interactive Entertainment",
    "9c:84:bf": "Sony Interactive Entertainment",
    "a4:15:66": "Sony Interactive Entertainment",
    "a4:77:33": "Sony Interactive Entertainment",
    "a8:75:9b": "Sony Interactive Entertainment",
    "ac:5f:3e": "Sony Interactive Entertainment",
    "b0:05:94": "Sony Interactive Entertainment",
    "b4:29:3d": "Sony Interactive Entertainment",
    "b8:8a:60": "Sony Interactive Entertainment",
    "bc:60:a7": "Sony Interactive Entertainment",
    "c4:3c:fe": "Sony Interactive Entertainment",
    "c8:63:f1": "Sony Interactive Entertainment",
    "cc:9e:a2": "Sony Interactive Entertainment",
    "d0:0f:6d": "Sony Interactive Entertainment",
    "d4:22:3f": "Sony Interactive Entertainment",
    "d4:8a:fc": "Sony Interactive Entertainment",
    "d8:0d:17": "Sony Interactive Entertainment",
    "dc:0b:1a": "Sony Interactive Entertainment",
    "dc:68:eb": "Sony Interactive Entertainment",
    "e4:1f:13": "Sony Interactive Entertainment",
    "e8:4e:ce": "Sony Interactive Entertainment",
    "f0:46:1c": "Sony Interactive Entertainment",
    "f4:43:8f": "Sony Interactive Entertainment",
    "f8:46:1c": "Sony Interactive Entertainment",
    "fc:0f:e6": "Sony Interactive Entertainment",

    # Microsoft Xbox
    "00:03:ff": "Microsoft Xbox",
    "00:0d:4c": "Microsoft Xbox",
    "00:11:dc": "Microsoft Xbox",
    "00:15:5d": "Microsoft Xbox",
    "00:1d:d1": "Microsoft Xbox",
    "00:22:48": "Microsoft Xbox",
    "00:25:ae": "Microsoft Xbox",
    "02:00:01": "Microsoft Xbox",
    "06:0e:2b": "Microsoft Xbox",
    "7c:ed:8d": "Microsoft Xbox",
    "8c:3b:ad": "Microsoft Xbox",
    "8c:6d:50": "Microsoft Xbox",
    "9c:d2:42": "Microsoft Xbox",
    "a8:40:7d": "Microsoft Xbox",
    "b8:76:3f": "Microsoft Xbox",
    "c4:1e:fa": "Microsoft Xbox",
    "d0:0e:a4": "Microsoft Xbox",
    "d8:6c:e9": "Microsoft Xbox",
    "e6:0a:7d": "Microsoft Xbox",
    "f8:2f:5b": "Microsoft Xbox",
}


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

    mac_lower = mac_address.lower()

    # Router/AP indicators
    if any(pattern in mac_lower for pattern in ["netgear", "tp-link", "cisco", "asus"]):
        return "Router/AP"

    # Apple devices
    if mac_lower.startswith(("28:cd:c1", "8c:85:90", "ac:bc:32", "f0:18:98")):
        return "Apple Device"

    # Gaming consoles
    if mac_lower.startswith(("00:03:0f", "00:0d:93", "00:03:ff", "00:0d:4c")):
        return "Gaming Console"

    # IoT/Smart devices
    if mac_lower.startswith(("b8:27:eb", "dc:a6:32", "e4:5f:01")):  # Raspberry Pi
        return "IoT Device"
    if mac_lower.startswith(("00:1a:11", "3c:5a:b4")):  # Google/Nest
        return "Smart Device"
    if mac_lower.startswith(("00:fc:8b", "18:74:2e")):  # Amazon
        return "Smart Device"

    # Virtualization
    if mac_lower.startswith(("00:50:56", "00:0c:29", "08:00:27", "52:54:00")):
        return "Virtual Machine"

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
    2. Run: sudo ./venv/bin/python network_scanner.py
    3. Enter your password when prompted

Alternatively, activate virtual environment first:
    source venv/bin/activate  # Linux/macOS
    sudo python network_scanner.py

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


def main(config: Optional[Dict[str, Any]] = None) -> None:
    """
    Main function to perform network scanning.

    Args:
        config: Configuration dictionary (optional)
    """
    if config is None:
        config = load_config()

    try:
        # Parse command line arguments
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

        # Special guidance for Netgear AP discovery
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

    except KeyboardInterrupt:
        logger.info("⏹️  Network scan cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"❌ Network scan failed: {e}")
        if args.debug:
            logger.exception("Full traceback:")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
        logger.info("✅ Script completed successfully")
    except Exception as e:
        logger.error(f"❌ Script failed: {e}")
        sys.exit(1)