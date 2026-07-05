# Federation Automation

Federation Automation prepares BIM model delivery sets by staging source files, optionally extracting IFC object data, optionally injecting metadata into IFC files, grouping models, building Navisworks federation outputs, and optionally publishing the latest valid federation to Revizto.

## Manual

See `Docs/UserManual.md` for the workflow, section descriptions, and selective run options.

Project page:
https://github.com/Gudarzian/FederationAutomation

## Folder Layout

- `Source` contains the PowerShell source scripts and EXE build scripts.
- `Docs` contains the user manual.
- `Templates` contains sanitized generic JSON configuration templates.
- `Exe_Files` contains the runnable package.

## Quick Start

1. Open `Exe_Files\FA_GUI.exe`.
2. Open or edit `Exe_Files\Config.json`.
3. Replace the example source rows and folder settings with project-specific values.
4. Validate the configuration.
5. Save, then use `Save and Run`.

## License

Copyright (c) 2026 Gudarzian

This software is licensed under the PolyForm Noncommercial License 1.0.0. Noncommercial use, copying, modification, and sharing are allowed under the license terms. Commercial use, including selling this software or using it as part of a paid service or product, requires separate permission from the copyright holder.

Full terms:
https://polyformproject.org/licenses/noncommercial/1.0.0
