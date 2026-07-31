"""
Remap the Windows Copilot key back to Ctrl.

Most modern Windows keyboards (and laptops where the right-Ctrl was reflashed)
send the Copilot key as a chord — typically LShift + LWin + F23 — produced by
the firmware/driver. This script installs a Win32 low-level keyboard hook and
rewrites that chord into a real Ctrl press, holding Ctrl for as long as the
key is held (so Ctrl+C, Ctrl+V, etc. behave correctly).

Usage:
    python copilot_remap.py detect    # press the key once, save the chord
    python copilot_remap.py run       # run the remapper
    python copilot_remap.py show      # print the saved chord

No third-party dependencies. Uses only the stdlib (ctypes + Win32).
"""

from __future__ import annotations

import argparse
import ctypes
import json
import logging
import sys
import time
from ctypes import wintypes
from pathlib import Path

log = logging.getLogger(__name__)

CONFIG = Path(__file__).with_name("chord.json")
DEFAULT_TARGET_VK = 0xA3  # VK_RCONTROL

# ---- Win32 constants ---------------------------------------------------------

WH_KEYBOARD_LL = 13
WM_KEYDOWN, WM_KEYUP = 0x0100, 0x0101
WM_SYSKEYDOWN, WM_SYSKEYUP = 0x0104, 0x0105
WM_QUIT = 0x0012

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002

# Tag we set on injected events so our own hook ignores them.
INJECTED_TAG = 0xC0F11070

VK_NAMES = {
    0x10: "Shift",   0x11: "Ctrl",    0x12: "Alt",
    0xA0: "LShift",  0xA1: "RShift",
    0xA2: "LCtrl",   0xA3: "RCtrl",
    0xA4: "LAlt",    0xA5: "RAlt",
    0x5B: "LWin",    0x5C: "RWin",
    0x86: "F23",     0x87: "F24",
    0x70: "F1", 0x71: "F2", 0x72: "F3", 0x73: "F4", 0x74: "F5", 0x75: "F6",
    0x76: "F7", 0x77: "F8", 0x78: "F9", 0x79: "F10", 0x7A: "F11", 0x7B: "F12",
}

MODIFIER_VKS = {0x10, 0x11, 0x12, 0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0x5B, 0x5C}


def vk_name(vk: int) -> str:
    return VK_NAMES.get(vk, f"VK_0x{vk:02X}")


# ---- Win32 plumbing ----------------------------------------------------------

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
kernel32.GetModuleHandleW.restype = wintypes.HMODULE


class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


LRESULT = ctypes.c_ssize_t
HOOKPROC = ctypes.WINFUNCTYPE(
    LRESULT, ctypes.c_int, wintypes.WPARAM, ctypes.POINTER(KBDLLHOOKSTRUCT)
)

user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, wintypes.HMODULE, wintypes.DWORD]
user32.SetWindowsHookExW.restype = wintypes.HHOOK
user32.CallNextHookEx.argtypes = [
    wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, ctypes.POINTER(KBDLLHOOKSTRUCT)
]
user32.CallNextHookEx.restype = LRESULT
user32.UnhookWindowsHookEx.argtypes = [wintypes.HHOOK]
user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.GetMessageW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.restype = ctypes.c_int
user32.PeekMessageW.argtypes = [
    ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.UINT, wintypes.UINT
]
user32.PeekMessageW.restype = wintypes.BOOL


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("padding", ctypes.c_ubyte * 32)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUT_UNION)]


user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
user32.SendInput.restype = wintypes.UINT


def send_vk(vk: int, key_up: bool) -> None:
    inp = INPUT()
    inp.type = INPUT_KEYBOARD
    inp.u.ki.wVk = vk
    inp.u.ki.wScan = 0
    inp.u.ki.dwFlags = KEYEVENTF_KEYUP if key_up else 0
    inp.u.ki.time = 0
    inp.u.ki.dwExtraInfo = INJECTED_TAG
    user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))


# ---- detect mode -------------------------------------------------------------

def detect() -> None:
    log.info("Press the Copilot key once. Capturing for ~0.5s after the first event...")
    captured: list[tuple[float, int, int]] = []
    first_t: list[float | None] = [None]

    @HOOKPROC
    def proc(nCode, wParam, lParam):
        if nCode == 0 and wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
            kb = lParam.contents
            if kb.dwExtraInfo != INJECTED_TAG:
                t = time.perf_counter()
                if first_t[0] is None:
                    first_t[0] = t
                captured.append((t, kb.vkCode, kb.scanCode))
        return user32.CallNextHookEx(None, nCode, wParam, lParam)

    hmod = kernel32.GetModuleHandleW(None)
    hook = user32.SetWindowsHookExW(WH_KEYBOARD_LL, proc, hmod, 0)
    if not hook:
        sys.exit(f"SetWindowsHookEx failed (err {ctypes.get_last_error()})")

    msg = ctypes.create_string_buffer(48)
    try:
        deadline = time.perf_counter() + 60  # 60s overall timeout
        while time.perf_counter() < deadline:
            user32.PeekMessageW(msg, None, 0, 0, 1)  # PM_REMOVE — pump LL hook
            time.sleep(0.005)
            if first_t[0] is not None and time.perf_counter() - first_t[0] > 0.5:
                break
    finally:
        user32.UnhookWindowsHookEx(hook)

    if not captured:
        sys.exit("No key events captured. Try again.")

    t0 = captured[0][0]
    seen: dict[int, int] = {}
    for t, vk, sc in captured:
        if t - t0 < 0.2 and vk not in seen:
            seen[vk] = sc

    chord = list(seen.items())
    log.info("Detected chord:")
    for vk, sc in chord:
        log.info("  vk=0x%02X (%s)  scan=0x%02X", vk, vk_name(vk), sc)

    triggers = [vk for vk, _ in chord if vk not in MODIFIER_VKS]
    if not triggers:
        sys.exit("Chord contains only modifier keys — cannot remap.")
    trigger_vk = triggers[-1]
    modifier_chord = sorted(vk for vk, _ in chord if vk != trigger_vk)

    cfg = {
        "trigger_vk": trigger_vk,
        "modifier_vks": modifier_chord,
        "target_vk": DEFAULT_TARGET_VK,
        "_human_chord": [vk_name(vk) for vk, _ in chord],
        "_human_target": vk_name(DEFAULT_TARGET_VK),
    }
    CONFIG.write_text(json.dumps(cfg, indent=2))
    chord_str = " + ".join(cfg["_human_chord"])
    log.info("Saved -> %s", CONFIG)
    log.info("Will remap [%s] -> %s", chord_str, cfg["_human_target"])


# ---- run mode ----------------------------------------------------------------

def show() -> None:
    if not CONFIG.exists():
        sys.exit(f"No config at {CONFIG}. Run detect first.")
    cfg = json.loads(CONFIG.read_text())
    chord_str = " + ".join(cfg.get("_human_chord", []))
    log.info("%s", CONFIG)
    log.info("  chord  : [%s]", chord_str)
    log.info("  target : %s", cfg.get("_human_target", vk_name(cfg["target_vk"])))
    log.info("  raw    : %s", json.dumps(cfg, indent=2))


def run() -> None:
    if not CONFIG.exists():
        sys.exit(f"No config at {CONFIG}. Run with 'detect' first.")
    cfg = json.loads(CONFIG.read_text())
    trigger_vk: int = cfg["trigger_vk"]
    modifier_vks: set[int] = set(cfg["modifier_vks"])
    target_vk: int = cfg.get("target_vk", DEFAULT_TARGET_VK)

    pressed: set[int] = set()
    remap_active = [False]

    @HOOKPROC
    def proc(nCode, wParam, lParam):
        if nCode != 0:
            return user32.CallNextHookEx(None, nCode, wParam, lParam)
        kb = lParam.contents
        if kb.dwExtraInfo == INJECTED_TAG:
            return user32.CallNextHookEx(None, nCode, wParam, lParam)

        vk = kb.vkCode
        is_down = wParam in (WM_KEYDOWN, WM_SYSKEYDOWN)
        is_up = wParam in (WM_KEYUP, WM_SYSKEYUP)

        if is_down:
            pressed.add(vk)
            if vk == trigger_vk and modifier_vks.issubset(pressed):
                # The chord-modifiers (Win+Shift) are spuriously "down" because
                # the driver mashed them in front of F23. Cancel them to the OS,
                # then press the real target key.
                for mvk in modifier_vks:
                    send_vk(mvk, key_up=True)
                send_vk(target_vk, key_up=False)
                remap_active[0] = True
                return 1  # suppress trigger down
        elif is_up:
            pressed.discard(vk)
            if vk == trigger_vk and remap_active[0]:
                send_vk(target_vk, key_up=True)
                remap_active[0] = False
                return 1  # suppress trigger up
            if remap_active[0] and vk in modifier_vks:
                # The driver releases Win/Shift after F23 up, but we already
                # cancelled them at chord-down time — swallow the redundant ups.
                return 1

        return user32.CallNextHookEx(None, nCode, wParam, lParam)

    hmod = kernel32.GetModuleHandleW(None)
    hook = user32.SetWindowsHookExW(WH_KEYBOARD_LL, proc, hmod, 0)
    if not hook:
        sys.exit(f"SetWindowsHookEx failed (err {ctypes.get_last_error()})")

    chord_str = " + ".join(vk_name(v) for v in (*sorted(modifier_vks), trigger_vk))
    log.info("Remapping [%s] -> %s. Running.", chord_str, vk_name(target_vk))

    try:
        msg = ctypes.create_string_buffer(48)
        while True:
            r = user32.GetMessageW(msg, None, 0, 0)
            if r <= 0:
                break
    finally:
        user32.UnhookWindowsHookEx(hook)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mode", choices=["detect", "run", "show"])
    args = ap.parse_args()
    {"detect": detect, "run": run, "show": show}[args.mode]()


if __name__ == "__main__":
    main()
