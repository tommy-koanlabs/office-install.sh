# Office on Wine - Implementation Plan

## Project Goal

Run Microsoft Office LTSC (2021 and/or 2024) on Linux using Wine 11, targeting
native 64-bit execution on the distro's default WoW64 Wine build for best
performance.

## Current Status: Distro Hop In Progress

Previous distro was CachyOS. User is switching distros. The plan below is
distro-agnostic but assumes Arch-based (pacman/paru). Adjust package commands
for the new distro.

**What works (proven):** 32-bit Office 2021 LTSC on Wine 11 multilib (wine32).
Flickering fixed, but sluggish. Backup of this setup is in `backup/`.

**Next step:** Fresh install with 64-bit Office on system WoW64 Wine.

## Documentation

Read these before starting work:

- `docs/wine-knowledge-base.md` - **START HERE** - Wine architecture, version
  differences, all errors encountered and solutions, performance settings
- `docs/previous-32bit-setup.md` - Full config of the working 32-bit setup
  (desktop entries, env vars, paths, proposed perf fixes never applied)

## Repository Layout

```
archive/                  # Original Troplo install scripts (tracked in git)
  office-install.sh       # Troplo's original script
  office-install-modified.sh  # Our modified version
  README.md               # Original README
backup/                   # Binary backups from 32-bit install (GITIGNORED)
  msoffice_ltsc.7z        # Pre-built 32-bit Office 2021 LTSC (~1.7GB)
  wine-9.7.zst            # Wine 9.7 multilib binary (~196MB)
  msoffice_icons.7z       # Office application icons (~16KB)
  officedeploymenttool.exe # Microsoft ODT installer (~3.6MB)
  ODT/                    # ODT extracted files
    setup.exe             # ODT setup binary
    configuration-*.xml   # Various ODT configs
    EULA
docs/
  wine-knowledge-base.md  # Comprehensive reference (errors, settings, arch)
  previous-32bit-setup.md # Working 32-bit config documentation
scripts/
  download-office-64bit.bat # Windows batch: ODT download of 64-bit 2021 + 2024
CLAUDE.md                 # This file
```

## Phase 1: Preparation

1. [x] Document the previous working 32-bit configuration and proposed fixes
2. [x] Create Windows batch script to download 64-bit Office 2021 + 2024 via ODT
3. [x] Backup original 32-bit assets to repo with .gitignore
4. [ ] ~~Purge wine32 and reinstall default Wine~~ (skipped - distro hopping instead)
5. [x] Comprehensive knowledge base documenting all Wine settings and lessons learned

## Phase 2: 64-bit Office Installation (Post Distro Hop)

1. [ ] Install distro's default Wine package (should be WoW64 build)
2. [ ] Verify: `ls /usr/lib/wine/x86_64-unix/` exists, `i386-unix/` does NOT
3. [ ] Run `download-office-64bit.bat` on a Windows PC
4. [ ] Transfer OfficeDownloads folder from Windows to Linux
5. [ ] Create 64-bit Wine prefix: `WINEPREFIX=~/.wine-msoffice/LTSC wineboot -u`
   - Do NOT set WINEARCH=win32 (let it default to 64-bit)
6. [ ] Install Office via ODT with local files:
   ```bash
   WINEPREFIX=~/.wine-msoffice/LTSC wine setup.exe /configure config.xml
   ```
7. [ ] Try 2021 first. If it works, also test 2024.
8. [ ] Create desktop entries with performance env vars (see knowledge base)

## Phase 3: Optimization

1. [ ] Check NTsync: `ls /dev/ntsync` and test if Wine build uses it
2. [ ] Enable WINEESYNC=1 if NTsync unavailable (check `ulimit -Hn` >= 500000)
3. [ ] Set WINEDEBUG=-all in all desktop entries
4. [ ] Test OpenGL version cap - increase from 3.0 if GPU supports higher
5. [ ] Benchmark: compare responsiveness with various sync mechanisms

## Key Decisions

- **Why 64-bit**: Native execution on WoW64 Wine (no translation), system Wine
  package (maintained, updated), NTsync support, more CPU registers/RAM
- **Why WoW64 Wine**: Distro default, regular updates, compiled with modern
  features, no AUR dependency
- **Fallback**: 32-bit backup in `backup/` can restore the working (but slow)
  setup using wine32 from AUR - see `docs/previous-32bit-setup.md`

## Quick Reference: Errors to Watch For

| Error | Meaning | See |
|-------|---------|-----|
| `kernel32.dll, status c0000135` | WoW64 Wine + 32-bit prefix | knowledge-base.md |
| `WINEARCH not supported in wow64` | Can't force win32 on WoW64 | knowledge-base.md |
| ODT error `30088-2056` | BITS stubbed, download on Windows | knowledge-base.md |
| Excel flickering | Wine < 11 rendering bug | Use Wine 11+ |
