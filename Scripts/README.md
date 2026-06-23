# Federation Automation

Federation Automation is a Windows PowerShell tool for staging BIM model files, adding selected metadata to IFC files, grouping and federating models into Navisworks NWD outputs, and optionally publishing a final model to Revizto.

## What it supports

- Local and synchronised source folders, including ACC Desktop Connector, OneDrive, and SharePoint locations.
- Optional ProjectWise source acquisition.
- IFC metadata processing.
- Filename-based grouping and Navisworks federation.
- Optional Revizto publishing.
- A GUI editor for JSON configuration, plus command-line automation.

## Quick start

1. Edit the included generic `Scripts/Config.json` in `Federation-Automation.exe`, or edit it in a text editor.
2. Configure at least the source folders, download rows, grouping rows, and output folder.
3. Run `Federation-Automation.exe` for the GUI, or run `006-Main.exe` / `006-Main.ps1` for unattended operation.

`Scripts/Config.xlsx` contains the same generic configuration in the legacy Excel format.

When run with no configuration argument, `006-Main.exe` uses `Config.json` beside the EXE when available; otherwise it falls back to `Config.xlsx`.

## Build

Run `Scripts/000-Gui2Exe.ps1` to build `Federation-Automation.exe` and `Scripts/000-2Exe.ps1` to build `006-Main.exe`.

See [the user manual](Docs/Federation-Automation-User-Manual-Friendly.md) for configuration and operating details.

## Important security note

Do not commit live configuration files, passwords, project paths, Revizto publish codes, logs, outputs, or model files. The included `Config.json` is generic; keep project-specific configuration local.

## Licence and commercial use

This project is offered under the PolyForm Noncommercial License 1.0.0. Noncommercial use, sharing, and modification are permitted under that licence. Commercial use, including selling the software, requires separate permission from the copyright holder.

Donations are voluntary and do not grant commercial-use rights, support entitlement, or a separate licence. See [LICENSE](LICENSE).
