# Previous Working Configuration: Office 2021 LTSC 32-bit on Wine 11

## Status as of 2026-02-11

**Working but sluggish.** Excel flickering issue resolved by upgrading from
Wine 9.7 to Wine 11 (wine32 multilib build). Performance was noticeably slower
than native Windows.

## System Configuration

- **OS**: CachyOS (Arch-based), kernel 6.18.8-3-cachyos
- **Wine**: wine32 11.1-1 (AUR) - Wine 11 compiled WITHOUT WoW64 (traditional multilib)
- **Architecture**: win32 (32-bit prefix, 32-bit Office binaries)
- **Wine prefix**: `/home/tommy/.wine-msoffice/LTSC`
- **Wine prefix arch**: `#arch=win32` (in system.reg)
- **Windows version**: Windows 7 (per user.reg)
- **NTsync**: Device exists at /dev/ntsync but wine32 does NOT use it
- **Sync mechanism**: Default server-based (no esync/fsync/ntsync enabled)

## Wine Packages Installed

```
wine32 11.1-1              # AUR - traditional multilib Wine 11 (replaced system wine)
wine-mono 10.4.1-1.1       # .NET replacement
lib32-glibc 2.42+          # Plus ~30 other lib32-* dependencies
```

## Office Installation

- **Source**: Pre-built 32-bit archive `msoffice_ltsc.7z` (~1.7GB)
  - Originally from Troplo's script: https://i.troplo.com/i/721f0242a2c0.7z
  - Contains a pre-configured Wine prefix with Office 2021 LTSC already installed
- **Office apps**: Excel, Word, PowerPoint, Outlook, Publisher, Access
- **Install path**: `C:\Program Files\Microsoft Office\root\Office16\`

## Desktop Entry Format (all 6 apps follow this pattern)

```ini
[Desktop Entry]
Type=Application
Name=Microsoft Excel [LTSC]
Icon=/home/tommy/.wine-msoffice/icons/excel_48x1.png
Exec=sh -c 'PATH="/usr/bin:/home/tommy/.local/bin:..." WINEARCH=win32 WINEPREFIX=/home/tommy/.wine-msoffice/LTSC /usr/bin/wine "/home/tommy/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/EXCEL.EXE" "%U"'
Categories=Office;
MimeType=application/msexcel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;
```

### All Desktop Entries

| File | App | Executable | MimeTypes |
|------|-----|-----------|-----------|
| excel-ltsc.desktop | Excel | EXCEL.EXE | msexcel, xlsx |
| word-ltsc.desktop | Word | WINWORD.EXE | msword, docx |
| powerpoint-ltsc.desktop | PowerPoint | POWERPNT.EXE | ms-powerpoint, pptx |
| outlook-ltsc.desktop | Outlook | OUTLOOK.EXE | microsoft-outlook |
| publisher-ltsc.desktop | Publisher | MSPUB.EXE | ms-publisher |
| access-ltsc.desktop | Access | MSACCESS.EXE | msaccess |

## Directory Structure

```
~/.wine-msoffice/
  LTSC/                    # Active Wine prefix (win32)
  LTSC-wine9.7-backup/     # Backup of Wine 9.7 prefix state
  Office2024/              # Abandoned 2024 attempt prefix (win32)
  icons/                   # Application icons (png, 48x48)
  wine/                    # Wine 9.7 binary (original, no longer used)
    usr/bin/wine           # ELF 32-bit
    usr/bin/wine64         # ELF 64-bit
    usr/lib/wine/          # Has i386-unix + x86_64-unix
```

## Proposed Performance Fixes (Never Applied)

These were identified but the session ended before applying them:

### 1. WINEESYNC=1 (Eventfd-based synchronization)

- Proven performance improvement for Wine applications
- System has file descriptor limit of 1,048,576 (well above the ~500K requirement)
- Would replace slow default server-based synchronization
- Add to Exec line: `WINEESYNC=1`

### 2. WINEDEBUG=-all (Suppress debug output)

- Eliminates all Wine debug/fixme/warn messages
- Reduces overhead from string formatting and stderr writes
- Add to Exec line: `WINEDEBUG=-all`

### 3. OpenGL Version Cap

- MaxVersionGL registry key set to 0x3002 (OpenGL 3.0)
- Could be increased if GPU supports higher
- Location: `HKEY_CURRENT_USER\Software\Wine\Direct3D\MaxVersionGL`

### 4. NTsync (Not Available for wine32)

- Kernel module present (/dev/ntsync exists on 6.18.8)
- wine32 package was NOT compiled with NTsync support
- Would require system Wine (WoW64) or a custom wine32 build
- This is a key motivation for moving to 64-bit + system Wine

## Why We're Moving Away From This Setup

1. **Performance**: 32-bit execution with no sync optimization = sluggish
2. **AUR dependency**: wine32 is not officially maintained, could break on updates
3. **No NTsync**: The wine32 build lacks NTsync support, missing a major perf win
4. **Translation overhead**: On a 64-bit system, 32-bit Wine requires lib32-* stack
5. **Better path available**: 64-bit Office + WoW64 Wine = native execution + NTsync
