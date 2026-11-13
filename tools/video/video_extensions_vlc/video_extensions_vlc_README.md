# Video Extensions VLC Default Setter

This tool helps you set VLC Media Player as the default program for opening various video file formats on Windows.

## Features

- **Automatic Detection**: Checks common video file extensions (.mp4, .avi, .mkv, .mov, etc.)
- **Current Status Display**: Shows which program is currently associated with each extension
- **Selective Changes**: Asks for confirmation before changing each extension (y/n with y as default)
- **VLC Detection**: Automatically finds VLC installation or uses system PATH
- **Registry Management**: Properly updates Windows file associations using registry

## Supported Video Extensions

The script checks the following video file extensions:
- .mp4, .avi, .mkv, .mov, .wmv, .flv, .webm, .m4v
- .3gp, .mpg, .mpeg, .vob, .ogv, .rm, .rmvb, .asf
- .divx, .xvid, .f4v, .mts, .m2ts, .ts, .mxf

## Usage

### Method 1: Batch File (Recommended)
1. Double-click `video_extensions_vlc.bat`
2. The script will automatically attempt to elevate to administrator privileges (you may see a UAC prompt)
3. The script will check each video extension
4. For each extension, it will show current status and ask for confirmation
5. Press Enter or 'y' to set VLC as default, 'n' to skip

### Method 2: PowerShell Script
1. Open PowerShell as Administrator (recommended)
2. Navigate to the script directory
3. Run: `.\video_extensions_vlc.ps1`

## Requirements

- Windows 10/11
- VLC Media Player installed
- PowerShell execution enabled

## How It Works

1. **Detection**: Checks current file associations using Windows `assoc` and `ftype` commands
2. **VLC Location**: Finds VLC installation automatically from common paths or PATH
3. **Association Methods**: Uses multiple methods to set file associations:
   - Basic `assoc/ftype` commands (works without admin rights for many extensions)
   - Registry updates (some require admin rights for UserChoice keys)
4. **Confirmation**: Asks before making changes with y/n prompt (y is default)

## Important Notes

- **Administrator Rights**: The script will automatically attempt to elevate to administrator privileges. Some file associations (especially in Windows 10/11) require admin rights for the UserChoice registry keys. The script will still work for basic associations without admin rights.
- **Partial Success**: Even without admin rights, many extensions will be successfully associated. Extensions that fail will show warnings.
- **Restart Required**: You may need to restart Windows Explorer or log out/in for changes to take effect
- **Backup**: The script doesn't backup existing associations, but you can manually change them back in Windows Settings if needed

## Troubleshooting

- **VLC Not Found**: Install VLC Media Player first
- **Permission Errors**: Run as administrator
- **Changes Not Applied**: Restart Windows Explorer (Ctrl+Shift+Esc → Restart Explorer)

## Files

- `video_extensions_vlc.ps1` - Main PowerShell script
- `video_extensions_vlc.bat` - Batch file launcher
- `video_extensions_vlc_README.md` - This documentation