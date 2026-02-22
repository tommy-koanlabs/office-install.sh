# Fork notes:
Modified script to use up to date packages and work with locally downloaded wine and office files

# office-install.sh
Microsoft Office 365 & 2021 LTSC installation script for Wine on Arch Linux.

Originally by [Troplo](https://github.com/Troplo/office-install.sh).

Known issues:
- Broken Microsoft login (good thing!)
- Doesn't receive feature updates due to Windows 7 EoL.
- OneNote, and Teams don't work.
- Excel has a tendency to flicker when typing. (**Fixed in Wine 11**)

# Running
```
bash <(curl -s https://raw.githubusercontent.com/Troplo/office-install.sh/main/office-install.sh)
```

If you can't type, wait a bit for the activation popup and restart the Office program.

# Note
These are the original scripts preserved for reference. The active project has
moved to the repo root. See `../CLAUDE.md` for the current plan.
