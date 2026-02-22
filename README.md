# Office on Wine

Microsoft Office LTSC on Linux via Wine 11. Targeting 64-bit native execution.

## Status

Migrating from a working 32-bit Office 2021 LTSC setup to 64-bit for better
performance. See `CLAUDE.md` for the full implementation plan and current progress.

## Quick Start (Post Distro Hop)

1. Install your distro's Wine package (should be WoW64 build)
2. Run `scripts/download-office-64bit.bat` on a Windows PC to get Office files
3. Transfer files to Linux and install via ODT `/configure`
4. See `docs/wine-knowledge-base.md` for all the details

## Documentation

- **[CLAUDE.md](CLAUDE.md)** - Implementation plan and task tracking
- **[docs/wine-knowledge-base.md](docs/wine-knowledge-base.md)** - Wine settings,
  architecture differences, errors and fixes, performance tuning
- **[docs/previous-32bit-setup.md](docs/previous-32bit-setup.md)** - Previous
  working 32-bit configuration (archived)

## Credits

Based on [Troplo's office-install.sh](https://github.com/Troplo/office-install.sh).
Original scripts preserved in `archive/`.
