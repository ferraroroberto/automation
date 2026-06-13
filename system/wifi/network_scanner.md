# Network Scanner - ARP-based Local Network Discovery

A professional-grade Python script for discovering active devices on your local network using ARP scanning. Specifically designed to find Access Points and routers that may not respond to traditional ping scans.

## 🎯 Purpose

This scanner is particularly useful for:
- Finding your Netgear R9000 Access Point's dynamically assigned IP address
- Discovering all active network devices regardless of firewall settings
- Identifying device manufacturers from MAC addresses
- Network troubleshooting and inventory

## 🔍 Why ARP Scan vs ICMP Ping

**ARP (Address Resolution Protocol)** operates at Layer 2 and discovers devices by their MAC addresses. This is more effective than ICMP ping because:

- **Firewall Bypass**: ARP requests cannot be blocked by host firewalls since they're essential for network communication
- **AP Discovery**: Access Points and routers often disable ICMP responses but must respond to ARP
- **Complete Coverage**: Finds all active network interfaces, even those with strict security policies
- **Hardware Level**: Works at the data link layer where network hardware identification occurs

## 📋 Requirements

- **Python 3.7+**
- **Administrative privileges** (required for raw socket operations)
- **scapy library**: `pip install scapy`

### Windows-Specific Requirements

**Important**: On Windows, you must install either **Npcap** or **WinPcap** for ARP scanning to work:

#### Option 1: Npcap (Recommended)
1. Download from: https://npcap.com/#download
2. Install with "WinPcap API-compatible mode" enabled
3. Restart your computer after installation

#### Option 2: WinPcap (Legacy)
1. Download from: https://www.winpcap.org/install/
2. Install the package
3. May require compatibility mode on Windows 10+

#### Alternative: Use WSL or Linux VM
If you prefer not to install WinPcap/Npcap on Windows:
- Use Windows Subsystem for Linux (WSL)
- Or run the scanner in a Linux virtual machine
- Command in WSL: `sudo apt install python3-scapy`

## 🚀 How to Run

### Windows (PowerShell)
```powershell
# As Administrator
.\.venv\Scripts\python.exe system\wifi\network_scanner.py
```

### Linux/macOS
```bash
# With sudo
sudo ./.venv/bin/python system/wifi/network_scanner.py
```

## 📊 Usage Examples

### Basic Scan (Default: 192.168.0.0/24)
```bash
python network_scanner.py
```

### Custom Subnet
```bash
python network_scanner.py --subnet 192.168.1.0/24
```

### Custom Output File
```bash
python network_scanner.py --output my_network_scan.json
```

### Debug Mode
```bash
python network_scanner.py --debug
```

### Combined Options
```bash
python network_scanner.py --subnet 192.168.10.0/24 --output office_scan.json --timeout 3
```

## 📄 Output Format

### Console Table
```
IP Address       MAC Address         Hostname        Manufacturer       Device Type
-------------------------------------------------------------------------------------
192.168.0.1     aa:bb:cc:dd:ee:ff  router           Cisco Systems       Router/AP
192.168.0.100   11:22:33:44:55:66  netgear-ap       Netgear Inc         Router/AP     ← Your AP!
192.168.0.150   77:88:99:aa:bb:cc  iphone           Apple Inc           Mobile Device
192.168.0.200   88:99:aa:bb:cc:dd  desktop-pc       Dell Inc            Computer
```

### Network Scan Summary
```
📊 Network Scan Summary:
   • Total devices found: 39
   • Local network devices: 37
   • Unique manufacturers: 8

🏭 Manufacturers:
   • Netgear Inc: 1 device(s)
   • Cisco Systems: 1 device(s)
   • Apple Inc: 12 device(s)
   • Samsung Electronics: 3 device(s)
   • Dell Inc: 1 device(s)

📱 Device Types:
   • Mobile Device: 15 device(s)
   • Router/AP: 3 device(s)
   • Computer: 6 device(s)
   • IoT Device: 2 device(s)

🎯 Netgear Devices Found (1):
   • 192.168.0.100 - 11:22:33:44:55:66 (netgear-ap)
   💡 These are likely your routers/access points!

🏷️ Devices with Hostnames (23):
   • router → 192.168.0.1 (Cisco Systems)
   • netgear-ap → 192.168.0.100 (Netgear Inc)
   • iphone → 192.168.0.150 (Apple Inc)
```

### JSON Export
```json
{
  "devices": [
    {
      "ip": "192.168.0.1",
      "mac": "aa:bb:cc:dd:ee:ff",
      "hostname": "router",
      "manufacturer": "Cisco Systems",
      "device_type": "Router/AP",
      "status": {
        "is_local": true,
        "subnet_hint": "192.168.x.x (Home Network)",
        "timestamp": "2025-12-29T12:52:04.123456"
      },
      "arp_details": {
        "operation": 2,
        "hardware_type": 1,
        "protocol_type": 2048,
        "hardware_len": 6,
        "protocol_len": 4
      }
    }
  ],
  "scan_summary": {
    "total_devices": 39,
    "netgear_devices": 1,
    "devices_with_hostnames": 23,
    "local_network_devices": 37,
    "subnet_scanned": "192.168.0.0/24",
    "scan_timestamp": "2025-12-29T12:52:04.123456",
    "manufacturers": {
      "Netgear Inc": 1,
      "Cisco Systems": 1,
      "Apple Inc": 12
    },
    "device_types": {
      "Router/AP": 3,
      "Mobile Device": 15,
      "Computer": 6
    }
  }
}
```

## ⚙️ Configuration

The script uses `network_scanner.json` for configuration:

```json
{
  "default_subnet": "192.168.0.0/24",
  "default_output": "network_scan.json",
  "scan_timeout": 2,
  "logging": {
    "level": "INFO"
  },
  "output_format": {
    "include_metadata": true,
    "pretty_print": true
  }
}
```

## 🔧 Command Line Options

| Option | Short | Default | Description |
|--------|-------|---------|-------------|
| `--subnet` | `-s` | 192.168.0.0/24 | Network subnet in CIDR notation |
| `--output` | `-o` | network_scan.json | Output JSON file path |
| `--timeout` | `-t` | 2 | ARP response timeout in seconds |
| `--debug` | | | Enable debug logging |

## 🛠️ Troubleshooting

### "winpcap is not installed" Error (Windows)
**Cause**: Windows requires WinPcap or Npcap for network packet operations.

**Solutions**:
1. **Install Npcap** (Recommended):
   - Download: https://npcap.com/#download
   - Install with "WinPcap API-compatible mode" enabled
   - Restart computer

2. **Install WinPcap** (Legacy):
   - Download: https://www.winpcap.org/install/
   - May need compatibility mode on Windows 10+

3. **Use WSL or Linux VM**:
   ```bash
   # In WSL/Ubuntu
   sudo apt update
   sudo apt install python3-scapy
   sudo python3 network_scanner.py
   ```

### "Insufficient privileges" Error
**Solution**: Run with administrative/root privileges
- Windows: Right-click PowerShell → "Run as Administrator"
- Linux/macOS: Use `sudo`

### "scapy library not installed" Error
**Solution**: Install scapy
```bash
pip install scapy
```

### No devices found
**Possible causes**:
- Wrong subnet (try 192.168.1.0/24, 192.168.10.0/24, 10.0.0.0/24)
- Virtual machine network configuration (use Bridged Adapter mode)
- Network isolation or VLAN configuration

### Virtual Machine Issues
- Ensure VM uses "Bridged Adapter" network mode
- Host-only mode may not show all network devices
- Check VM network settings match host network

## 📈 Features

- **ARP-based scanning** - Finds devices that ping cannot detect
- **MAC manufacturer lookup** - Identifies device vendors from extensive database
- **DNS hostname resolution** - Reverse DNS lookup for device names
- **Device type detection** - Intelligent classification (Router/AP, Mobile, IoT, Gaming, etc.)
- **Network analysis** - Local vs external network detection
- **Privilege checking** - Clear error messages for permission issues
- **Comprehensive reporting** - Detailed statistics and summaries
- **JSON export** - Complete structured data output with metadata
- **Comprehensive logging** - Debug and info level logging
- **Error handling** - Graceful failure with helpful messages
- **Netgear detection** - Special highlighting and focus for Netgear devices
- **Manufacturer statistics** - Breakdown by vendor and device type
- **Hostname discovery** - Devices with resolved names highlighted
- **Configurable** - JSON-based configuration file
- **Cross-platform** - Windows, Linux, and macOS support

## 🔐 Security Considerations

- Requires elevated privileges for network scanning
- Only scans local network (no internet exposure)
- No data transmission outside local subnet
- Results contain only IP and MAC address information

## 📝 File Structure

```
system/wifi/
├── network_scanner.py      # Main scanner script
├── network_scanner.json    # Configuration file
└── network_scanner.md      # This documentation
```

## 🎯 Finding Your Netgear R9000 AP

The script is specifically designed to help you locate your Netgear R9000 Access Point:

1. Run the scanner on your main router's subnet (usually 192.168.0.0/24)
2. Look for devices with "Netgear Inc" in the Manufacturer column
3. The AP's IP address will be listed alongside its MAC address
4. Use this IP to access the AP's web interface (usually http://[IP]:8080)

If no Netgear devices appear:
- Try scanning other common subnets
- Check if the AP is powered on and connected
- Verify the AP is in Access Point mode (not Router mode)