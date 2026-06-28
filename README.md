# Federation Automation

Federation Automation prepares source model files, optionally injects metadata into IFC files, builds Navisworks federation outputs, and can publish the latest valid federation to Revizto when configured.

## Folder Layout

- `Source` contains the PowerShell source scripts and EXE build scripts.
- `Docs` contains the user manual.
- `Templates` contains sanitized generic JSON and Excel configuration templates.
- `Exe_Files` contains the runnable package:
  - `FA_GUI.exe`
  - `FA_Main.exe`
  - `Config.json`
  - `Config.xlsx`
  - `NavisworksOptions.xml`

## Quick Start

1. Open `Exe_Files\FA_GUI.exe`.
2. Open or edit `Exe_Files\Config.json`.
3. Replace the example source rows and folder settings with project-specific values.
4. Validate the configuration.
5. Save, then use `Save and Run`.

The included configuration files are generic templates and intentionally contain no project-specific paths, model codes, credentials, or publish codes.
