#!/usr/bin/env python3
"""
Wi-Fi BAT Generator (GUI).

Pick a saved Wi-Fi network from a Tkinter list, then save a self-contained
.bat file anywhere. The generated BAT embeds the full WLAN profile (SSID +
password + security) as base64-encoded XML; running it re-creates the
profile via `netsh wlan add profile` and connects via `netsh wlan connect`.
The BAT is idempotent: it skips reconnect if the target SSID is already
active.

Usage:
    python wifi_bat_generator.py
"""

import argparse
import base64
import logging
import os
import re
import subprocess
import sys
import tempfile
import tkinter as tk
from html import escape
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional

# Reuse the saved-profile exporter from the sibling module.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from wifi_passwords import export_profiles, parse_profile_xmls  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# Accept the same Windows console locales the rest of the suite assumes.
NETSH_ENCODINGS = ("utf-8", "cp1252", "cp850", "mbcs")

_FILENAME_SAFE_RE = re.compile(r"[^A-Za-z0-9._-]+")


def _run_netsh(args: List[str]) -> str:
    """Run a netsh command and return decoded stdout (best-effort decoding)."""
    proc = subprocess.run(
        ["netsh", *args],
        capture_output=True,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )
    raw = proc.stdout or b""
    for enc in NETSH_ENCODINGS:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def get_current_ssid() -> Optional[str]:
    """Return the SSID of the currently connected Wi-Fi, or None."""
    try:
        output = _run_netsh(["wlan", "show", "interfaces"])
    except FileNotFoundError:
        logger.error("❌ netsh not found - this tool only runs on Windows.")
        return None

    for line in output.splitlines():
        # SSID line contains "SSID" but not "BSSID"; the value follows the colon.
        if "SSID" in line and "BSSID" not in line:
            _, _, value = line.partition(":")
            ssid = value.strip()
            if ssid:
                return ssid
    return None


def list_saved_profiles() -> List[Dict[str, str]]:
    """Export all saved WLAN profiles to a tempdir and return parsed rows."""
    with tempfile.TemporaryDirectory(prefix="wifi_bat_") as tmp:
        export_profiles(tmp)
        return parse_profile_xmls(tmp)


def sanitize_filename(ssid: str) -> str:
    """Strip characters that are unsafe in filenames; keep alnum/._-."""
    cleaned = _FILENAME_SAFE_RE.sub("_", ssid).strip("_")
    return cleaned or "wifi"


def build_profile_xml(ssid: str, authentication: str, encryption: str, password: str) -> str:
    """Build a WLANProfile XML matching the schema used by wifi_connect.xml."""
    auth = (authentication or "").strip() or "open"
    enc = (encryption or "").strip() or "none"
    hex_ssid = ssid.encode("utf-8").hex().upper()

    if password and auth.lower() != "open":
        security = (
            "    <security>\n"
            "      <authEncryption>\n"
            f"        <authentication>{escape(auth)}</authentication>\n"
            f"        <encryption>{escape(enc)}</encryption>\n"
            "        <useOneX>false</useOneX>\n"
            "      </authEncryption>\n"
            "      <sharedKey>\n"
            "        <keyType>passPhrase</keyType>\n"
            "        <protected>false</protected>\n"
            f"        <keyMaterial>{escape(password)}</keyMaterial>\n"
            "      </sharedKey>\n"
            "    </security>\n"
        )
    else:
        security = (
            "    <security>\n"
            "      <authEncryption>\n"
            "        <authentication>open</authentication>\n"
            "        <encryption>none</encryption>\n"
            "        <useOneX>false</useOneX>\n"
            "      </authEncryption>\n"
            "    </security>\n"
        )

    return (
        '<?xml version="1.0"?>\n'
        '<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">\n'
        f"  <name>{escape(ssid)}</name>\n"
        "  <SSIDConfig>\n"
        "    <SSID>\n"
        f"      <hex>{hex_ssid}</hex>\n"
        f"      <name>{escape(ssid)}</name>\n"
        "    </SSID>\n"
        "  </SSIDConfig>\n"
        "  <connectionType>ESS</connectionType>\n"
        "  <connectionMode>auto</connectionMode>\n"
        "  <MSM>\n"
        f"{security}"
        "  </MSM>\n"
        '  <MacRandomization xmlns="http://www.microsoft.com/networking/WLAN/profile/v3">\n'
        "    <enableRandomization>false</enableRandomization>\n"
        "  </MacRandomization>\n"
        "</WLANProfile>\n"
    )


def build_bat(ssid: str, profile_xml: str) -> str:
    """Wrap a base64-encoded WLAN profile XML in a self-contained BAT script."""
    safe_name = sanitize_filename(ssid)
    encoded = base64.b64encode(profile_xml.encode("utf-8")).decode("ascii")
    # Wrap to 76 chars per line so each `echo` line stays well under cmd's 8KB limit.
    chunks = [encoded[i : i + 76] for i in range(0, len(encoded), 76)]
    echo_block = "\n".join(f"    echo {chunk}" for chunk in chunks)

    # The header carries the literal SSID; we double any % so the BAT survives
    # cmd.exe variable expansion when the SSID happens to contain "%".
    target_ssid_literal = ssid.replace("%", "%%")

    return (
        "@echo off\n"
        "REM Generated by wifi_bat_generator.py - self-contained Wi-Fi connector.\n"
        f'REM Target SSID: {target_ssid_literal}\n'
        "setlocal\n"
        f'set "TARGET_SSID={target_ssid_literal}"\n'
        f'set "PROFILE_XML=%TEMP%\\wifi_{safe_name}.xml"\n'
        f'set "PROFILE_B64=%TEMP%\\wifi_{safe_name}.b64"\n'
        "\n"
        "REM --- Skip if already connected to the target SSID ---\n"
        'set "CUR="\n'
        "for /f \"tokens=2 delims=:\" %%a in ("
        "'netsh wlan show interfaces ^| findstr /C:\"SSID\" ^| findstr /V \"BSSID\"'"
        ") do set \"CUR=%%a\"\n"
        'if defined CUR set "CUR=%CUR:~1%"\n'
        'if /I "%CUR%"=="%TARGET_SSID%" (\n'
        "    echo Already connected to %TARGET_SSID%.\n"
        "    goto :end\n"
        ")\n"
        "\n"
        "REM --- If profile already exists in any scope, skip add and connect ---\n"
        'netsh wlan show profile name="%TARGET_SSID%" >nul 2>&1\n'
        "if not errorlevel 1 (\n"
        "    echo Profile \"%TARGET_SSID%\" already exists, connecting...\n"
        "    goto :connect\n"
        ")\n"
        "\n"
        "REM --- Write base64 payload, decode to XML, add profile, connect ---\n"
        '> "%PROFILE_B64%" (\n'
        f"{echo_block}\n"
        ")\n"
        'certutil -f -decode "%PROFILE_B64%" "%PROFILE_XML%" >nul\n'
        "if errorlevel 1 ( echo Failed to decode profile. & goto :cleanup )\n"
        "\n"
        'netsh wlan add profile filename="%PROFILE_XML%" user=current\n'
        "if errorlevel 1 ( echo Could not add profile, attempting connect with existing profile... )\n"
        "\n"
        ":connect\n"
        'netsh wlan connect name="%TARGET_SSID%"\n'
        "if errorlevel 1 ( echo Failed to connect to %TARGET_SSID%. & goto :cleanup )\n"
        "\n"
        "echo Connected to %TARGET_SSID%.\n"
        "\n"
        ":cleanup\n"
        'del /q "%PROFILE_B64%" "%PROFILE_XML%" 2>nul\n'
        ":end\n"
        "endlocal\n"
        "pause\n"
    )


class WifiBatApp(tk.Tk):
    """Tk window: pick a saved Wi-Fi network, save a self-contained BAT."""

    COLUMNS = ("ssid", "auth", "encryption", "has_password")

    def __init__(self) -> None:
        super().__init__()
        self.title("Wi-Fi BAT Generator")
        self.geometry("720x440")
        self.minsize(560, 320)

        self._profiles: List[Dict[str, str]] = []
        self._current_ssid: Optional[str] = None

        self._build_ui()
        self.refresh()

    def _build_ui(self) -> None:
        header = ttk.Frame(self, padding=(12, 10, 12, 0))
        header.pack(fill=tk.X)
        self.current_var = tk.StringVar(value="Currently connected: …")
        ttk.Label(header, textvariable=self.current_var, font=("Segoe UI", 10, "bold")).pack(
            side=tk.LEFT
        )
        ttk.Button(header, text="Refresh", command=self.refresh).pack(side=tk.RIGHT)

        body = ttk.Frame(self, padding=12)
        body.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(body, columns=self.COLUMNS, show="headings", selectmode="browse")
        self.tree.heading("ssid", text="SSID")
        self.tree.heading("auth", text="Authentication")
        self.tree.heading("encryption", text="Encryption")
        self.tree.heading("has_password", text="Password")
        self.tree.column("ssid", width=260, anchor=tk.W)
        self.tree.column("auth", width=130, anchor=tk.W)
        self.tree.column("encryption", width=110, anchor=tk.W)
        self.tree.column("has_password", width=90, anchor=tk.CENTER)
        self.tree.tag_configure("active", background="#d4edda")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", lambda _e: self.save_bat())

        scroll = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.tree.yview)
        scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scroll.set)

        actions = ttk.Frame(self, padding=(12, 0, 12, 12))
        actions.pack(fill=tk.X)
        ttk.Button(actions, text="Save BAT…", command=self.save_bat).pack(side=tk.LEFT)
        ttk.Button(actions, text="Test selected", command=self.test_selected).pack(
            side=tk.LEFT, padx=(8, 0)
        )
        ttk.Button(actions, text="Quit", command=self.destroy).pack(side=tk.RIGHT)

        self.status_var = tk.StringVar(value="")
        ttk.Label(self, textvariable=self.status_var, padding=(12, 0, 12, 8)).pack(
            fill=tk.X
        )

    def refresh(self) -> None:
        self.status_var.set("Loading saved profiles…")
        self.update_idletasks()
        try:
            self._current_ssid = get_current_ssid()
            self._profiles = list_saved_profiles()
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to load saved profiles")
            messagebox.showerror("Wi-Fi BAT Generator", f"Failed to load saved profiles:\n{exc}")
            self.status_var.set("Error loading profiles.")
            return

        self.current_var.set(
            f"Currently connected: {self._current_ssid}"
            if self._current_ssid
            else "Currently connected: (none)"
        )

        for row in self.tree.get_children():
            self.tree.delete(row)
        self._profiles.sort(key=lambda p: p.get("ssid", "").lower())
        for profile in self._profiles:
            ssid = profile.get("ssid", "")
            tag = "active" if self._current_ssid and ssid == self._current_ssid else ""
            self.tree.insert(
                "",
                tk.END,
                values=(
                    ssid,
                    profile.get("authentication") or "-",
                    profile.get("encryption") or "-",
                    "yes" if profile.get("password") else "no",
                ),
                tags=(tag,) if tag else (),
            )
        self.status_var.set(f"{len(self._profiles)} saved profile(s) loaded.")

    def _selected_profile(self) -> Optional[Dict[str, str]]:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Wi-Fi BAT Generator", "Pick a network from the list first.")
            return None
        ssid = self.tree.item(sel[0], "values")[0]
        for profile in self._profiles:
            if profile.get("ssid") == ssid:
                return profile
        return None

    def _generate_bat_for(self, profile: Dict[str, str]) -> str:
        return build_bat(
            profile["ssid"],
            build_profile_xml(
                profile["ssid"],
                profile.get("authentication", ""),
                profile.get("encryption", ""),
                profile.get("password", ""),
            ),
        )

    def save_bat(self) -> None:
        profile = self._selected_profile()
        if not profile:
            return
        ssid = profile["ssid"]
        default_name = f"connect_{sanitize_filename(ssid)}.bat"
        target = filedialog.asksaveasfilename(
            title="Save Wi-Fi BAT",
            defaultextension=".bat",
            filetypes=[("Batch files", "*.bat"), ("All files", "*.*")],
            initialfile=default_name,
        )
        if not target:
            return
        try:
            Path(target).write_text(self._generate_bat_for(profile), encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("Wi-Fi BAT Generator", f"Failed to write BAT:\n{exc}")
            return
        self.status_var.set(f"Saved: {target}")
        messagebox.showinfo("Wi-Fi BAT Generator", f"Saved BAT for '{ssid}' to:\n{target}")

    def test_selected(self) -> None:
        profile = self._selected_profile()
        if not profile:
            return
        ssid = profile["ssid"]
        if not messagebox.askyesno(
            "Wi-Fi BAT Generator",
            f"This will run the generated BAT for '{ssid}' in a new console window.\n"
            "If you are not currently on that network, Windows may switch.\n\nProceed?",
        ):
            return
        bat_path = Path(tempfile.gettempdir()) / f"wifi_bat_test_{sanitize_filename(ssid)}.bat"
        try:
            bat_path.write_text(self._generate_bat_for(profile), encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("Wi-Fi BAT Generator", f"Failed to write test BAT:\n{exc}")
            return
        try:
            os.startfile(str(bat_path))  # type: ignore[attr-defined]  # noqa: S606
        except OSError as exc:
            messagebox.showerror("Wi-Fi BAT Generator", f"Failed to launch BAT:\n{exc}")
            return
        self.status_var.set(f"Launched test: {bat_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="GUI to generate self-contained Wi-Fi BAT files.")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging.")
    args = parser.parse_args()
    if args.debug:
        logger.setLevel(logging.DEBUG)
    if sys.platform != "win32":
        logger.warning("⚠️ This tool relies on `netsh` and only works on Windows.")
    app = WifiBatApp()
    app.mainloop()


if __name__ == "__main__":
    main()
