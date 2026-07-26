# Federation Automation - JSON/CSV Build

Federation Automation prepares BIM model delivery sets by staging source files, optionally extracting IFC object data, optionally injecting metadata into IFC files, grouping models, building Navisworks federation outputs, and optionally publishing the latest valid federation to Revizto.

Manual and source repository:
https://github.com/Gudarzian/FederationAutomation

## JSON/CSV Build Notes

- Configuration format: `Config.json` only.
- Source metadata file: `PWAttributes.csv` under `SourceFolder`.
- Deleted-file report: `Deleted_files.csv`.
- Runtime does not call `Set-ExecutionPolicy`.
- Runtime does not install PowerShell modules.
- Runtime does not use `ImportExcel`, `Export-Excel`, or Excel workbook APIs.

ProjectWise rows can still run if the required ProjectWise PowerShell module is already installed and allowed by the machine policy. This build will not install it automatically.

## Workflow

The stages are connected by folders. Source acquisition writes to `SourceFolder`; IFC data extraction writes CSV files to `IfcDataExtractionFolder`; IFC processing can write metadata-injected files to `ProcessedFolder`; federation reads the selected input folder and writes NWD outputs to `FederationOutputFolder` and, when configured, `DestinationFolder`.

Each stage can be selectively run from Settings:

- `RunDownload`: stage files from configured source rows, or skip and reuse existing files.
- `RunIfcDataExtraction`: export IFC object data to CSV, optionally forcing a refresh.
- `RunProcess`: inject configured metadata into IFC files, optionally forcing reprocessing.
- `RunFederation`: create Navisworks outputs, optionally forcing a rebuild.
- `ReviztoPublish`: publish a valid federation when a publish code is configured.

## Folder Layout for GitHub

- `Source` contains PowerShell source scripts and EXE build scripts.
- `Docs` contains the user manual.
- `Templates` contains sanitized generic JSON configuration templates.
- `Exe_Files` contains the runnable package.

## Run

```powershell
.\FA_GUI.exe
.\FA_Main.exe Config.json
```

For the script-only package, use `Run-GUI.bat` or `Run-Main.bat`. Direct execution or dot-sourcing of unsigned `.ps1` scripts can be blocked by Windows execution policy. The supplied batch launchers use a process-local `-ExecutionPolicy Bypass`, so run them only from the approved package location.

The current PS2EXE-built EXEs are not Authenticode signed and can occasionally be flagged by antivirus or SmartScreen. Do not disable protection or add a personal exclusion. Provide IT/security with the alert, tool version, and EXE SHA-256 hash so they can validate the approved release and apply any centrally managed allow rule.

## License

Copyright (c) 2026 Gudarzian

This software is licensed under the PolyForm Noncommercial License 1.0.0. It may be used, copied, modified, and shared for noncommercial purposes under the license terms. Commercial use, including selling this software or using it as part of a paid service or product, requires separate permission from the copyright holder.

Full terms:
https://polyformproject.org/licenses/noncommercial/1.0.0
