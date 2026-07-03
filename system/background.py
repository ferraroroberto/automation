import ctypes
import logging
import sys
import winreg as reg

log = logging.getLogger(__name__)


def set_colors(theme: str = "light grey") -> None:
    log.info("Applying '%s' desktop theme", theme)
    if theme.lower() == "black":
        rgb_color = (0, 0, 0)
        accent_color = 0x000000
    elif theme.lower() == "light grey":
        rgb_color = (235, 235, 235)
        accent_color = 0x5A5A5A
    else:
        raise ValueError("Theme must be either 'black' or 'light grey'")

    # Set the background color
    color_ref = (rgb_color[2] << 16) | (rgb_color[1] << 8) | rgb_color[0]
    color_array = (ctypes.c_int * 1)(color_ref)
    element_array = (ctypes.c_int * 1)(1)  # 1 corresponds to COLOR_DESKTOP

    key = reg.OpenKey(reg.HKEY_CURRENT_USER, "Control Panel\\Desktop", 0, reg.KEY_SET_VALUE)
    reg.SetValueEx(key, "Wallpaper", 0, reg.REG_SZ, "")
    reg.CloseKey(key)

    ctypes.windll.user32.SetSysColors(1, element_array, color_array)
    ctypes.windll.user32.SystemParametersInfoW(20, 0, "", 1)

    # Set the taskbar color
    key = reg.OpenKey(reg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 0, reg.KEY_SET_VALUE)
    reg.SetValueEx(key, "AppsUseLightTheme", 0, reg.REG_DWORD, 0 if theme.lower() == "black" else 1)
    reg.SetValueEx(key, "SystemUsesLightTheme", 0, reg.REG_DWORD, 0 if theme.lower() == "black" else 1)
    reg.CloseKey(key)

    key = reg.OpenKey(reg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Accent", 0, reg.KEY_SET_VALUE)
    reg.SetValueEx(key, "AccentColorMenu", 0, reg.REG_DWORD, accent_color)
    reg.CloseKey(key)

    ctypes.windll.user32.SystemParametersInfoW(20, 0, "", 1)
    log.info("Desktop and taskbar colors updated")


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    # Set default theme to "light grey" if no argument is provided
    theme = sys.argv[1] if len(sys.argv) > 1 else "light grey"
    set_colors(theme)


if __name__ == "__main__":
    main()
