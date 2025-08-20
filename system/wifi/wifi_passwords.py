#!/usr/bin/env python3
"""
List all saved Wi-Fi networks and passwords on Windows.
Outputs to console and saves wifi_passwords.json in the current folder.

Usage:
    python wifi_passwords.py
    # or choose a different output file:
    python wifi_passwords.py --output my_wifi_list.json
    # enable debug mode:
    python wifi_passwords.py --debug
"""

import subprocess
import tempfile
import os
import glob
import xml.etree.ElementTree as ET
import json
import sys
import argparse
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# XML namespace for WLAN profiles
NS = {"ns": "http://www.microsoft.com/networking/WLAN/profile/v1"}

def export_profiles(tempdir: str) -> None:
    """
    Export all WLAN profiles with keys in clear text into tempdir as XML files.
    
    Args:
        tempdir: Temporary directory path for XML exports
        
    Raises:
        RuntimeError: If netsh export command fails
    """
    logger.debug(f"📂 Exporting WLAN profiles to temporary directory: {tempdir}")
    
    cmd = ["netsh", "wlan", "export", "profile", "key=clear", f"folder={tempdir}"]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    out = (proc.stdout or "") + (proc.stderr or "")
    
    if proc.returncode != 0:
        logger.error(f"❌ netsh export failed: {out.strip()}")
        raise RuntimeError(f"netsh export failed: {out.strip()}")
    
    logger.info("✅ WLAN profiles exported successfully")

def parse_profile_xmls(tempdir: str) -> List[Dict[str, Any]]:
    """
    Parse XML profile files and extract network information.
    
    Args:
        tempdir: Temporary directory containing XML files
        
    Returns:
        List of dictionaries containing network information
    """
    logger.debug(f"🔍 Parsing XML files in directory: {tempdir}")
    
    results = []
    xml_files = glob.glob(os.path.join(tempdir, "*.xml"))
    
    if not xml_files:
        logger.warning("⚠️ No XML profile files found")
        return results
    
    logger.info(f"📄 Found {len(xml_files)} XML profile files")
    
    for path in xml_files:
        try:
            root = ET.parse(path).getroot()
            
            # SSID/name can be in two places depending on schema
            ssid = (root.findtext("ns:name", namespaces=NS) 
                   or root.findtext("ns:SSIDConfig/ns:SSID/ns:name", namespaces=NS) 
                   or os.path.basename(path))

            auth = root.findtext(".//ns:authentication", namespaces=NS) or ""
            enc = root.findtext(".//ns:encryption", namespaces=NS) or ""
            pwd = root.findtext(".//ns:sharedKey/ns:keyMaterial", namespaces=NS) or ""

            results.append({
                "ssid": ssid,
                "authentication": auth,
                "encryption": enc,
                "password": pwd
            })
            
            logger.debug(f"✅ Parsed profile: {ssid}")
            
        except Exception as e:
            logger.warning(f"⚠️ Failed to parse {os.path.basename(path)}: {e}")
            results.append({
                "ssid": os.path.basename(path),
                "authentication": "",
                "encryption": "",
                "password": "",
                "error": str(e)
            })
    
    logger.info(f"✅ Successfully parsed {len(results)} profiles")
    return results

def print_table(rows: List[Dict[str, Any]]) -> None:
    """
    Print network information in a formatted table.
    
    Args:
        rows: List of network dictionaries
    """
    if not rows:
        logger.info("ℹ️ No saved Wi-Fi profiles found.")
        return
    
    # Calculate column widths
    w = max(len(r["ssid"]) for r in rows) if rows else 10
    
    print(f"\n{'SSID'.ljust(w)}  AUTHENTICATION      ENCRYPTION   PASSWORD")
    print("-" * (w + 50))
    
    for r in rows:
        print(f"{r['ssid'].ljust(w)}  {r['authentication'] or '-':<18} {r['encryption'] or '-':<11} {r['password'] or '-'}")

def save_json(rows: List[Dict[str, Any]], output_file: str, config: Dict[str, Any]) -> None:
    """
    Save network information to JSON file.
    
    Args:
        rows: List of network dictionaries
        output_file: Path to output JSON file
        config: Configuration dictionary
    """
    logger.debug(f"💾 Saving {len(rows)} networks to JSON file: {output_file}")
    
    # Prepare data for JSON output
    output_data = {
        "networks": rows
    }
    
    # Add metadata if configured
    if config.get("output_format", {}).get("include_metadata", True):
        output_data["metadata"] = {
            "total_networks": len(rows),
            "export_timestamp": str(Path(output_file).stat().st_mtime) if Path(output_file).exists() else "N/A",
            "source": "Windows WLAN Profile Export",
            "config_source": "wifi_passwords.json"
        }
    
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(
                output_data, 
                f, 
                indent=2 if config.get("output_format", {}).get("pretty_print", True) else None,
                ensure_ascii=config.get("output_format", {}).get("ensure_ascii", False)
            )
        
        logger.info(f"✅ Successfully saved {len(rows)} networks to {config['default_output']}")
        
    except Exception as e:
        logger.error(f"❌ Failed to save JSON file: {e}")
        raise

def load_config() -> Dict[str, Any]:
    """
    Load configuration from JSON file with fallback defaults.
    
    Returns:
        Configuration dictionary
    """
    config_path = Path(__file__).parent / "wifi_passwords.json"
    
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
        "default_output": "wifi_passwords.json",
        "temp_dir_prefix": "wifi_export_",
        "logging": {
            "level": "INFO",
            "format": "%(asctime)s - %(levelname)s - %(message)s"
        },
        "output_format": {
            "include_metadata": True,
            "pretty_print": True,
            "ensure_ascii": False
        },
        "xml_parsing": {
            "namespace": "http://www.microsoft.com/networking/WLAN/profile/v1",
            "fallback_to_filename": True
        }
    }

def main(config: Optional[Dict[str, Any]] = None) -> None:
    """
    Main function to export WiFi passwords.
    
    Args:
        config: Configuration dictionary (optional)
    """
    if config is None:
        config = load_config()
    
    try:
        # Parse command line arguments
        parser = argparse.ArgumentParser(
            description="Export saved Wi-Fi networks and passwords to JSON",
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
Examples:
  python wifi_passwords.py
  python wifi_passwords.py --output my_wifi_list.json
  python wifi_passwords.py --debug
            """
        )
        
        parser.add_argument(
            "--output", "-o",
            default=config["default_output"],
            help=f"Output JSON file path (default: {config['default_output']})"
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
        
        logger.info("🚀 Starting WiFi password export")
        
        # Create temporary directory and export profiles
        with tempfile.TemporaryDirectory(prefix=config["temp_dir_prefix"]) as tmp:
            logger.debug(f"📁 Created temporary directory: {tmp}")
            
            export_profiles(tmp)
            rows = parse_profile_xmls(tmp)
        
        # Display results and save to file
        print_table(rows)
        save_json(rows, args.output, config)
        
        print(f"\n✅ Export completed successfully!")
        print(f"📊 Total networks found: {len(rows)}")
        print(f"💾 Output saved to: {args.output}")
        
    except KeyboardInterrupt:
        logger.info("⏹️ Export cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"❌ Export failed: {e}")
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
