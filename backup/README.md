# Backup - 32-bit Office Assets

This directory contains backups from the original 32-bit Office 2021 LTSC
installation. All binary files are gitignored.

## Contents (when populated)

| File | Size | Description |
|------|------|-------------|
| msoffice_ltsc.7z | ~1.7 GB | Pre-built 32-bit Office 2021 LTSC Wine prefix |
| wine-9.7.zst | ~196 MB | Wine 9.7 multilib binary (original from Troplo's script) |
| msoffice_icons.7z | ~16 KB | Office application icons (48x48 PNG) |
| officedeploymenttool.exe | ~3.6 MB | Microsoft Office Deployment Tool |
| ODT/setup.exe | ~7.2 MB | ODT setup binary |
| ODT/configuration-*.xml | — | ODT configuration files |
| ODT/EULA | — | ODT license agreement |

These files are kept as a fallback in case the 64-bit migration fails and we
need to restore the working 32-bit configuration.
