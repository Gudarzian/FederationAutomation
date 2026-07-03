# GitHub Commit Summary

Suggested commit title:

`Prepare generic release library and update configuration templates`

Summary:

- Updated the user manual for the current `FA_GUI.exe` / `FA_Main.exe` package layout.
- Added sanitized generic configuration templates in JSON and Excel formats.
- Created a GitHub library structure with source scripts, docs, templates, and runnable EXE package files.
- Added `Exe_Files` package folder containing `FA_GUI.exe`, `FA_Main.exe`, generic `Config.json`, generic `Config.xlsx`, and `NavisworksOptions.xml`.
- Removed project-specific values from the release configuration templates.

Validation:

- Generic JSON and Excel configs were loaded through the existing configuration adapter.
- Release package was scanned for project-specific template values.
