# Wine + Office Knowledge Base

Lessons learned from getting Microsoft Office running on Linux via Wine.
Reference this before attempting any new installation to avoid repeating mistakes.

---

## Wine Architecture: The Critical Distinction

### Multilib (Traditional) Build
- Contains **both** `i386-unix/` and `x86_64-unix/` directories under `/usr/lib/wine/`
- Can run 32-bit AND 64-bit Windows applications natively
- 32-bit apps use native i386 Linux binaries (fast, no translation)
- Supports `WINEARCH=win32` to create 32-bit-only prefixes
- **Example**: Wine 9.7 bundled with Troplo's script, `wine32` AUR package

### WoW64 (Modern Default) Build
- Contains **only** `x86_64-unix/` under `/usr/lib/wine/`
- Has `i386-windows/` DLLs but NO `i386-unix/` native binaries
- 64-bit Windows apps run natively (fast)
- 32-bit Windows apps run through WoW64 translation layer (PE-level, not native)
- **Cannot** use `WINEARCH=win32` - will error with "not supported in wow64 mode"
- **Example**: Default Arch/CachyOS `wine` package since ~2025

### How to Check Which Build You Have
```bash
# If this directory exists, you have multilib (can run 32-bit natively)
ls /usr/lib/wine/i386-unix/
# If only this exists, you have WoW64 (64-bit native only)
ls /usr/lib/wine/x86_64-unix/
```

### Key Implication for Office
- **32-bit Office** needs either multilib Wine OR accepts WoW64 translation overhead
- **64-bit Office** on WoW64 Wine = native execution = best performance
- The pre-built Troplo archive contains 32-bit Office in a win32 prefix

---

## Wine Versions Tested

### Wine 9.7 (April 2024) - Troplo's Bundled Version
- Traditional multilib build (has i386-unix)
- Bundled as `wine-9.7.zst` (~196MB) in the install script
- **Issues**: Excel flickering when typing (rendering bug)
- **Status**: Works but outdated, flickering makes Excel unusable

### Wine 11.1 - System Package (CachyOS/Arch Default)
- Pure WoW64 build (no i386-unix)
- **Fatal issue with 32-bit**: `wine: could not load kernel32.dll, status c0000135`
  - This happens because there are no i386-unix native libraries
  - `WINEARCH=win32` errors: "not supported in wow64 mode"
  - Cannot run existing 32-bit prefixes AT ALL
- **Good for 64-bit**: Native execution, NTsync support likely compiled in

### wine32 11.1-1 (AUR Package)
- Wine 11 compiled WITHOUT WoW64 flag (traditional multilib)
- Provides i386-unix directory, can create win32 prefixes
- **Resolved**: Excel flickering gone, Office 2021 works
- **Issue**: Sluggish performance, no NTsync support compiled in
- **Conflicts with**: System `wine` package (replaces it)
- After install, run: `systemctl restart systemd-binfmt`

---

## Office Versions

### Office 2021 LTSC (ProPlus2021Volume)
- Support ends: October 2026
- Channel: `PerpetualVL2021`
- Available in 32-bit and 64-bit
- Apps: Word, Excel, PowerPoint, Outlook, Publisher, Access
- **Proven working** on Wine (32-bit via Troplo's script)
- Install path: `C:\Program Files\Microsoft Office\root\Office16\`

### Office 2024 LTSC (ProPlus2024Volume)
- Support ends: October 2029
- Channel: `PerpetualVL2024`
- Available in 32-bit and 64-bit
- Apps: Word, Excel, PowerPoint, Outlook, Access (Publisher REMOVED)
- New Excel functions: TEXTBEFORE, TEXTAFTER, TEXTSPLIT, VSTACK, HSTACK, etc.
- ActiveX disabled by default
- **Not yet tested** on Wine

---

## Office Deployment Tool (ODT)

### What It Is
- Microsoft's official tool for downloading and installing Office offline
- Binary: `setup.exe` (~7.2MB)
- Uses XML configuration files to specify edition, channel, architecture, language

### ODT Commands
```bash
# Download Office source files (RUN ON WINDOWS - fails in Wine)
setup.exe /download config.xml

# Install from local source files (can work in Wine)
setup.exe /configure config.xml
```

### Critical: ODT Download Does NOT Work in Wine
- **Error**: 30088-2056 (3)
- **Cause**: ODT uses BITS (Background Intelligent Transfer Service) to download
- Wine only has BITS stubs:
  ```
  fixme:qmgr:BackgroundCopyJob_AddFileWithRanges ... stub
  fixme:qmgr:BackgroundCopyJob_SetPriority ... stub
  fixme:qmgr:BackgroundCopyJob_SetMaximumDownloadTime (stub)
  ```
- Also has Authenticode certificate validation failures
- **Workaround**: Download on real Windows, transfer files, use `/configure` on Wine

### ODT Configuration Examples

**64-bit Office 2021 LTSC:**
```xml
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
```

**64-bit Office 2024 LTSC:**
```xml
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
```

### App-V (Deprecated)
- Microsoft deprecated App-V packaging in January 2026
- Was never a viable alternative for Wine anyway

---

## Wine Performance Settings

### Synchronization Primitives (Most Important for Performance)

Wine applications spend significant time on thread synchronization. The default
server-based sync is slow. These alternatives exist (in order of preference):

#### NTsync (Best - Kernel-level)
- Requires: Linux 6.14+ kernel, `/dev/ntsync` device node
- Requires: Wine compiled with NTsync support
- **Check availability**: `ls /dev/ntsync` and `WINEFSYNC=0 WINEESYNC=0 wine --version`
- **How to enable**: Usually automatic if Wine is compiled with support
- CachyOS kernel 6.18.8 HAS the device node
- System WoW64 Wine likely compiled with NTsync support
- wine32 AUR package does NOT have NTsync support

#### WINEESYNC (Good - Eventfd-based)
- Requires: High file descriptor limit (500K+, check with `ulimit -Hn`)
- **Enable**: `WINEESYNC=1` environment variable
- CachyOS has limit of 1,048,576 by default (sufficient)
- Proven significant performance improvement
- Works with any Wine build

#### WINEFSYNC (Good - Futex-based)
- Requires: Linux kernel with futex_waitv support (5.16+)
- **Enable**: `WINEFSYNC=1` environment variable
- Similar performance to esync, different mechanism

### Debug Output
- `WINEDEBUG=-all` suppresses all debug/fixme/warn messages
- Reduces overhead from string formatting and stderr I/O
- **Always use in production desktop entries**

### OpenGL Version
- Registry: `HKCU\Software\Wine\Direct3D\MaxVersionGL`
- Value `0x3002` = OpenGL 3.0 (conservative, set by Troplo's script)
- Can increase to match GPU capability (e.g., `0x4006` for OpenGL 4.6)
- Check GPU max: `glxinfo | grep "OpenGL version"`

### Desktop Entry Template (Optimized)
```ini
[Desktop Entry]
Type=Application
Name=Microsoft Excel [LTSC]
Icon=~/.wine-msoffice/icons/excel_48x1.png
Exec=sh -c 'WINEESYNC=1 WINEDEBUG=-all WINEPREFIX=~/.wine-msoffice/LTSC wine "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE" "%%U"'
Categories=Office;
MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;
```

Note: For 64-bit prefix on WoW64 Wine, do NOT set `WINEARCH=win32`.

---

## Common Errors and Solutions

### "wine: could not load kernel32.dll, status c0000135"
- **Cause**: Using WoW64 Wine with a 32-bit (win32) prefix
- **Solutions**:
  1. Install `wine32` from AUR (traditional multilib build)
  2. OR use 64-bit prefix with WoW64 Wine (preferred long-term)

### "wine: WINEARCH is set to 'win32' but this is not supported in wow64 mode"
- **Cause**: Trying to set `WINEARCH=win32` with WoW64-only Wine
- **Solution**: Don't set WINEARCH, or install multilib Wine

### "WINEARCH set to win32 but [prefix] is a 64-bit installation"
- **Cause**: Prefix was created as 64-bit, but WINEARCH=win32 is set
- **Solution**: Delete prefix and recreate: `rm -rf $WINEPREFIX && WINEARCH=win32 wineboot -u`

### ODT Error 30088-2056 (3)
- **Cause**: BITS service is stubbed in Wine
- **Solution**: Download Office source files on real Windows, transfer to Linux

### Wine prefix initialization hangs or crashes
- **Solution**: Run `wineboot -u` first to initialize prefix before any app launch
- For new prefix: `WINEPREFIX=/path/to/prefix wineboot -u`

---

## Distro-Specific Notes

### Arch Linux / CachyOS (2025-2026)
- Default `wine` package is pure WoW64 (64-bit only) since ~2025
- For 32-bit prefix support, need `wine32` from AUR
- `wine32` conflicts with system `wine` (replaces it)
- After installing wine32: `systemctl restart systemd-binfmt`
- CachyOS kernel ships with NTsync support (`/dev/ntsync`)
- lib32-* packages provide 32-bit runtime libraries

### Package Management
```bash
# Install system Wine (WoW64, 64-bit only)
pacman -S wine

# Install multilib Wine (AUR, supports 32-bit)
paru -S wine32

# Check which is installed
pacman -Qs wine
```

---

## File Locations Reference

| Item | Path |
|------|------|
| Wine prefix | `~/.wine-msoffice/LTSC/` |
| Office executables | `~/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/` |
| Desktop entries | `~/.local/share/applications/*-ltsc.desktop` |
| Icons | `~/.wine-msoffice/icons/` |
| Wine system registry | `~/.wine-msoffice/LTSC/system.reg` (check `#arch=` line) |
| Wine user registry | `~/.wine-msoffice/LTSC/user.reg` |
| Wine Direct3D settings | `HKCU\Software\Wine\Direct3D` in user.reg |
