# Federation Automation

Federation Automation prepares BIM model delivery sets by staging source files, extracting IFC data when required, adding configured IFC metadata, building Navisworks federations, and optionally publishing to Revizto.

## Quick start

1. Open `Exe_Files\FA_GUI.exe`.
2. Open or edit `Exe_Files\Config.json`.
3. Configure the source rows and folder settings for the project.
4. Use **Validate**, resolve any errors, then use **Save and Run**.

For a script-only deployment, open `PowerShell_Run_Package\Run-GUI.bat` instead.

## Launching and security

The script-only runners use a process-local PowerShell `-ExecutionPolicy Bypass` because unsigned `.ps1` scripts can be blocked by Windows policy. This does not change the computer's permanent policy; run the batch files only from the approved package location.

The current PS2EXE-built EXEs are not Authenticode signed and may occasionally be flagged by antivirus or SmartScreen. Do not disable protection or add a personal exclusion. Send IT/security the alert, tool version, and EXE SHA-256 hash so they can validate the approved release and apply any centrally managed allow rule.

## Package layout

- `Exe_Files` contains the GUI, pipeline executable, configuration, and Navisworks options file.
- `Templates` contains the sanitized generic configuration template.
- `Docs` contains the concise and full user guides.
- `Source` contains the PowerShell source and build scripts for maintainers.

### Rebuilding executables

`Source\\ThirdParty\\Navisworks\\2027\\Autodesk.Navisworks.Api.dll` is the
compile-time reference for the bundled visual-style add-in. This lets a build
computer create `FA_Main.exe` and `FA_GUI.exe` without Navisworks installed.
It does not make Navisworks portable: Navisworks Manage is still required on a
machine that creates federation outputs.

## Guides

Read `Docs\UserManual.md` for the concise workflow guide and `Docs\Federation-Automation-User-Manual-Friendly.md` for detailed settings and troubleshooting information.

## License

Copyright (c) 2026 Gudarzian. Licensed under the PolyForm Noncommercial License 1.0.0. See `LICENSE` for terms.
