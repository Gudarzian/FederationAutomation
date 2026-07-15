# Federation Automation User Manual (Friendly JSON/CSV Version)

Version date: 2026-07-04  
Applies to: `FA_GUI.exe`, `FA_Main.exe`, `Config.json`, and the PowerShell scripts in this JSON/CSV build

This manual matches the current `NoExcel_JSON_CSV_Version` scripts. The important point is simple:

- Configuration is JSON only.
- Source metadata is written to CSV.
- Excel configuration files and Excel export are not available in this build.

## 1. What Federation Automation Does

Federation Automation prepares BIM model delivery sets by running a configured pipeline.

In normal use it can:

1. Stage source model files from local folders, synced folders, ACC Desktop Connector folders, UNC paths, or ProjectWise.
2. Write a source metadata CSV, usually `PWAttributes.csv`, into the configured source folder.
3. Optionally extract object-level IFC data to one CSV per IFC file.
4. Optionally inject selected metadata into IFC files and write processed copies.
5. Federate model files in Navisworks using either naming-convention rules or wildcard rules.
6. Optionally copy selected federation outputs to a destination folder.
7. Optionally publish a valid federation output to Revizto.

The tool is controlled by `Config.json`. The GUI edits the same JSON file that the command-line executable runs.

## 2. What Is Different in This Build

This is the JSON/CSV build. It is intentionally lighter than the older Excel-based workflow.

| Area | Current behavior |
| --- | --- |
| Config file | `Config.json` only. `.xlsx` config files are rejected. |
| Source metadata | CSV file under `SourceFolder`, normally `PWAttributes.csv`. |
| Deleted-file report | `Deleted_files.csv`, normally written under the log folder. |
| Excel APIs | Not used by runtime. `ImportExcel`, `Export-Excel`, and workbook APIs are not used. |
| PowerShell execution policy | The runtime does not call `Set-ExecutionPolicy`. |
| PowerShell modules | The runtime does not install modules. Required modules must already be available. |
| GUI run behavior | `FA_GUI.exe` saves JSON, runs `FA_Main.exe`, and shows live progress. |

Some internal function names and a few UI labels still contain old terms such as "Excel", "workbook", or `ExportToXLSX`. In this build, read those as legacy names. The actual files used by the pipeline are JSON and CSV.

## 3. Main Files

| File | Purpose |
| --- | --- |
| `FA_GUI.exe` | Recommended interactive editor and launcher. |
| `FA_Main.exe` | Main automation runner. Can be launched directly from PowerShell. |
| `Config.json` | Runtime configuration file. Keep it beside the EXE or pass its path when launching. |
| `Generic_Config.json` | Starter template for a new project. |
| `NavisworksOptions.xml` | Optional Navisworks options file. Required when saving an older NWD version. |
| `006-Main.ps1` | Source for the main runner. |
| `007-Gui.ps1` | Source for the GUI. |
| `012-SharedFunctions.Ps1` | Shared helpers. |
| `013-ConfigFunctions.ps1` | JSON configuration loader/saver and settings catalog. |
| `021-DownloadFunctions.Ps1` | Source acquisition from local/synced folders and ProjectWise. |
| `031-ProcessFunctions.Ps1` | IFC metadata injection and pass-through file handling. |
| `041-FederationFunctions.Ps1` | Navisworks federation and Revizto publishing helpers. |
| `051-IfcDataExtractionFunctions.ps1` | IFC object data extraction using Python and IfcOpenShell. |

## 4. Working Folders

These folders are controlled in `settings`.

| Setting | Default | Used for |
| --- | --- | --- |
| `LogFolder` | `Logs` | Main run log and support messages. |
| `SourceFolder` | `SourceFiles` | Staged source models and the attributes CSV. |
| `AttributesFile` | `PWAttributes.csv` | Source metadata file stored inside `SourceFolder`. Use a file name only, not a path. |
| `ProcessedFolder` | `ProcessedIFC` | Processed IFC files and pass-through copies for federation. |
| `IfcDataExtractionFolder` | `IFCDataExtraction` | One CSV export per extracted IFC file. |
| `FederationOutputFolder` | `Output` | Grouped Navisworks files, final files, and `federation-summary.json`. |
| `DestinationFolder` | `Destination` | Optional copy location for wildcard outputs marked `CopyToDestination`. |

Relative paths are resolved from the folder where the EXE or script is running.

## 5. Requirements

Minimum:

- Windows.
- `FA_GUI.exe` and `FA_Main.exe` in the same runnable folder.
- A valid `Config.json`.

Needed only for specific stages:

- ProjectWise source rows require ProjectWise PowerShell commands such as `New-PWLogin`, `Get-PWDocumentsBySearch`, `Get-PWDocumentsByGUIDs`, and `CheckOut-PWDocuments`. The script checks for them but does not install them.
- Navisworks federation requires Autodesk Navisworks Manage. If `NavisworksVersion` is blank, the script tries to detect the highest installed version.
- IFC data extraction requires Python and `ifcopenshell`. If Python is not found, the script tries a current-user Python install through `winget`, then creates a user environment under `%LOCALAPPDATA%\Federation-Automation\PythonEnv`.
- Revizto publishing requires `C:\Program Files\Revizto SA\Revizto5\Service\ReviztoConsole.exe` and a valid `ReviztoPublishCode`.

## 6. Quick Start With the GUI

Use the GUI unless you are running the pipeline from a scheduled task or a script.

1. Open `FA_GUI.exe`.
2. Open an existing `Config.json`, or choose `New JSON`.
3. In `Settings`, turn on only the stages you need.
4. Set the main folders: `SourceFolder`, `FederationInputFolder` if needed, `FederationOutputFolder`, and `LogFolder`.
5. Fill the `Download` tab if source acquisition is enabled.
6. Fill the `Attributes` tab if IFC processing is enabled.
7. Fill `Data Extraction` rules if IFC CSV export is enabled.
8. Choose a `Grouping` method and configure the related table.
9. Click `Preflight` and fix any errors.
10. Click `Save and Run`.

The `Run` tab shows the active stage, progress bar, live activity, and a short run summary when the process finishes.

## 7. Running From PowerShell

From the folder containing the EXE:

```powershell
.\FA_Main.exe
.\FA_Main.exe Config.json
.\FA_Main.exe C:\Runs\ProjectA\Config.json
```

Config resolution:

- If no argument is supplied, the runner expects `Config.json` beside `FA_Main.exe`.
- A relative config path is resolved from the EXE folder.
- A full config path is used directly.
- `.xlsx` config files are not supported by this build.

The GUI starts the same runner in the background using:

```powershell
.\FA_Main.exe -ConfigFile "path\to\Config.json"
```

## 8. Stage Controls

The top activation controls in the GUI map to these settings.

| Setting | Values | Meaning |
| --- | --- | --- |
| `RunDownload` | `Yes`, `No` | `Yes` runs source acquisition from the `Download` rows. `No` ignores the `Download` rows and reuses existing files in `SourceFolder`. |
| `RunIfcDataExtraction` | `Yes`, `No`, `Force` | Exports IFC object data to CSV. `Force` ignores the "CSV is current" skip check. |
| `RunProcess` | `Yes`, `No`, `Force` | Injects selected metadata into IFC files. `Force` reprocesses applicable files even when unchanged. |
| `RunFederation` | `Yes`, `No`, `Force` | Builds Navisworks outputs when needed. `Force` rebuilds even if no upstream change is detected. |
| `ReviztoPublish` | `Yes`, `No`, `Force` | Publishes when configured and valid. `Force` can publish an existing valid model. |

Normal `RunFederation = Yes` is change-aware. It runs when source or processed files changed, deleted source files were detected, or the expected federation output is missing. It skips when nothing relevant changed and a valid output already exists.

## 9. How the Stages Connect

The stages pass files through the configured folders.

1. Source acquisition writes model files to `SourceFolder` and writes `AttributesFile` in that same folder.
2. IFC data extraction reads IFC files from `SourceFolder` and writes CSV files to `IfcDataExtractionFolder`.
3. IFC processing reads source IFC files plus the attributes CSV, writes processed IFC files to `ProcessedFolder`, and copies non-IFC federatable files as pass-through files.
4. Federation reads `FederationInputFolder` when set. If it is blank, it reads `ProcessedFolder` when processing is enabled, otherwise `SourceFolder`.
5. Federation writes Navisworks outputs to `FederationOutputFolder`.
6. Wildcard rules can copy selected outputs to `DestinationFolder`.
7. Revizto publishing uses the current valid federation output.

Supported federation input extensions are:

- `.ifc`
- `.dwg`
- `.dgn`
- `.rvt`
- `.nwc`

Generated `.nwc` cache files are filtered out when a matching source model is present, to avoid accidentally federating both the original model and its cache.

## 10. Settings Reference

`Config.json` is organised into these top-level sections:

| JSON section | GUI area |
| --- | --- |
| `settings` | Settings tab |
| `download` | Download tab |
| `pwAttributesList` | Attributes tab |
| `ifcDataExtractionRules` | Data Extraction tab |
| `federation` | Grouping tab when using Naming Convention and Lookups |
| `wildcardSelection` | Grouping tab when using Wildcard Selection |
| `lookups` | Lookups tab |

The GUI can open JSON files where optional sections are missing, then writes the current structure when you save.

### General and Source Settings

| Parameter | What to enter |
| --- | --- |
| `LogFolder` | Folder for run logs. Relative path is allowed. |
| `SourceFolder` | Folder where source files are staged. |
| `AttributesFile` | CSV file name for source metadata, usually `PWAttributes.csv`. It must be a file name only. |
| `RunDownload` | `Yes` or `No`. |
| `SourceAcquisitionMode` | `Auto`, `Local`, or `ProjectWise`. `Auto` allows mixed local and ProjectWise rows. |
| `PWUser` | Optional ProjectWise user name. Leave blank for Bentley IMS sign-in. |
| `PWPass` | Optional ProjectWise password. Avoid storing production passwords in shared JSON. |

### IFC Processing Settings

| Parameter | What to enter |
| --- | --- |
| `RunProcess` | `Yes`, `No`, or `Force`. |
| `ProcessedFolder` | Folder for processed IFC files and pass-through model files. |

### IFC Data Extraction Settings

| Parameter | What to enter |
| --- | --- |
| `RunIfcDataExtraction` | `Yes`, `No`, or `Force`. |
| `IfcDataExtractionFolder` | Folder for exported IFC object CSV files. |
| `IfcDataExtractionMaxFileSizeMB` | Maximum IFC file size to extract. Accepts values like `150`, `500MB`, or `1GB`. |
| `IfcDataExtractionSkipIfCsvIsCurrent` | `Yes` skips extraction when the existing CSV is newer than or equal to the IFC. |

### Federation and Navisworks Settings

| Parameter | What to enter |
| --- | --- |
| `RunFederation` | `Yes`, `No`, or `Force`. |
| `FederationGroupingMethod` | `Naming Convention and Lookups` or `Wildcard Selection`. |
| `IncludeUnmatchedFilesInFederatedModel` | `Yes` adds unmatched naming-convention files into the final federation. |
| `FederationInputFolder` | Optional explicit folder for federation input. Blank lets the script choose. |
| `FederationOutputFolder` | Folder for Navisworks outputs. |
| `DestinationFolder` | Folder for selected wildcard outputs copied after federation. |
| `FederatedFileName` | Final naming-convention output name. Use `.nwf` only when the final output should be NWF; otherwise NWD is used. |
| `NavisworksVersion` | Preferred Navisworks Manage year, such as `2026` or `2027`. Blank enables auto-detect. |
| `NavisworksConfigXML` | Optional Navisworks options XML. Needed for older NWD save versions. |
| Navisworks visual style | Every federation output is saved in `Full Render` style with a graduated background. This is enforced by the bundled Federation Automation Navisworks add-in. |
| `NavisworksSavedNwdVersion` | `Latest`, `2027`, `2026`, or `2016-2025`. |
| `NavisworksViewsImportXML` | Optional saved viewpoints XML, imported into the final naming-convention model only. |
| `NavisworksVisible` | `Yes` shows Navisworks; `No` runs it in the background. |
| `NWDNamingMethod` | `Full`, `OnlyCodes`, `OnlyDesc`, or `Codes-Desc` for grouped NWD names. |

### Revizto Settings

| Parameter | What to enter |
| --- | --- |
| `ReviztoPublish` | `Yes`, `No`, or `Force`. |
| `ReviztoPublishCode` | Revizto scheduler publish code. Required for publishing. |
| `ReviztoMaxAgeMinutes` | Maximum age of a federated model that may be published. Default is `60`. |

## 11. Download Tab

The `Download` tab controls source acquisition.

| Column | Meaning |
| --- | --- |
| `Run` | Whether this row is active. |
| `ReadFolder` | Local/synced folder or ProjectWise folder to search. |
| `FileFilter` | Comma-separated wildcard filters, such as `*.ifc,*.nwc` or `*ARC*.ifc`. Blank means `*`. |
| `Exclude` | Comma-separated terms or wildcard-style text removed after include matching. |
| `SkipIfSame` | `Yes` skips files already staged with the same size. `No` forces copy/download. |
| `CheckDateToo` | When `SkipIfSame` is `Yes`, also compare modified date. |
| `MinState` | Reserved in the current GUI. The download module currently ignores it. |

Local and synced folders are searched only at the folder level, not recursively.

ProjectWise paths can be entered in these forms:

```text
pw://datasource/Documents/Project/Folder
pw:\\datasource\Project\Folder
datasource:\Project\Folder
```

When ProjectWise credentials are blank, the script uses Bentley IMS sign-in. If `PWUser` and `PWPass` are supplied and fail, the script attempts Bentley IMS as a fallback.

The output of this stage is:

- model files in `SourceFolder`
- `PWAttributes.csv` or your configured `AttributesFile`
- optional `Deleted_files.csv` when old staged files are safely moved aside
- the `deleted` folder beside `SourceFolder`

Deleted-file cleanup is skipped when some read-folder searches fail, so the tool does not remove local files based on an incomplete source search.

## 12. Attributes Tab

The `Attributes` tab defines source metadata fields.

| Column | Meaning |
| --- | --- |
| `AttributeName` | Source field name, ProjectWise property path, or local metadata field. |
| `OutputName` | Name written to the attributes CSV and used as the IFC property name. Blank uses `AttributeName`. |
| `ExportToXLSX` | Legacy column name. In this build it means include the value in the exported CSV metadata. |
| `InjectToIFC` | Inject this value into IFC files when `RunProcess` is enabled. |

For local/synced source rows, useful built-in values include:

- `FileName`
- `FileSize`
- `FileUpdateDate`
- `SourcePath`
- `ReadFolder`

For ProjectWise rows, dotted paths such as `DocumentUpdater.Email` and `CustomAttributes.SomeField` can be used when the ProjectWise object exposes those fields.

IFC processing requires the attributes CSV to include `FileName` or `Filename`, because this is how source IFC files are matched to metadata rows.

## 13. Data Extraction Tab

IFC data extraction is a reporting step. It does not modify model files.

| Column | Meaning |
| --- | --- |
| `Run` | Whether the rule is active. |
| `FileInclusions` | Comma-separated wildcard filters for IFC file names. Blank means the rule can include all IFC files. |
| `FileExclusions` | Comma-separated wildcard filters for IFC file names to remove after file inclusions are applied. |
| `TabInclusions` | Comma-separated wildcard filters for IFC property set or quantity set names. Blank includes every tab/source. |
| `TabExclusions` | Comma-separated wildcard filters for IFC property set or quantity set names to remove. |
| `AttributeInclusions` | Comma-separated wildcard filters for IFC attribute names. Blank includes every attribute in the selected tabs. |
| `AttributeExclusions` | Comma-separated wildcard filters for IFC attribute names to remove. |

If there are no active extraction rules, every IFC in `SourceFolder` is eligible and all available attributes are exported. If an active rule is blank, it also includes all IFC files and all attributes. If multiple enabled rules match the same file, the first matching rule is used and later matches are ignored with a warning.

The extractor also adds viewer-style groups so the exported source names are closer to Navisworks/Forma property panels:

| Exported source group | Typical values |
| --- | --- |
| `Item` | name, type, material, source file |
| `Element ID` | object identifier |
| `Element` | IFC class, GUID, object type, tag, size fields, predefined type |
| `Material` | material name |

The fixed `Object` identity columns are always written at the start of the CSV. They are not controlled by the tab filters. Tab filters such as `Material, BaseQuantities, Element` match the viewer-style groups and IFC property/quantity sets exactly; use wildcards such as `*Element*` only when you want a broader match.

The extractor:

- reads IFC files from `SourceFolder`
- skips files larger than `IfcDataExtractionMaxFileSizeMB`
- skips current CSV files when `IfcDataExtractionSkipIfCsvIsCurrent = Yes`, unless the run is forced
- writes one CSV per IFC to `IfcDataExtractionFolder`
- writes `ifc-data-extraction-summary.json`

The CSV has two header rows:

1. source group or property set names
2. attribute names

Then one row is written for each extracted IFC product object.

## 14. IFC Processing

IFC processing is controlled by the `RunProcess` setting and the `Attributes` and `Grouping` tabs.

For IFC files, the process:

- reads the source IFC from `SourceFolder`
- finds the matching row in `AttributesFile`
- sets the `IFCPROJECT` name to the full source file name
- adds selected attributes into a property set named `PWattributes`
- optionally adds filename-derived values into `FileName_sections_Descriptions`
- writes the processed IFC to `ProcessedFolder`

For non-IFC federatable files, such as `.dwg`, `.dgn`, `.rvt`, and `.nwc`, the process copies them through to `ProcessedFolder` so federation can use one clean input folder.

The process is change-aware. It writes:

- `processed-metadata.json`
- `processed-summary.json`
- `processed-attributes.csv`

When the source file, reported metadata, and injected payload have not changed, the file is skipped unless `RunProcess = Force`.

## 15. Grouping Tab: Naming Convention and Lookups

Use `Naming Convention and Lookups` when your model file names follow a dash-separated naming convention.

Example:

```text
PROJECT-ORG-ZONE-STAGE-MOD-DISC-NUMBER.ifc
```

The script splits the file name without extension by `-`. The first part is position `1`, the second is position `2`, and so on. The special value `FileExtension` reads the Windows file extension, such as `.ifc`.

| Federation column | Meaning |
| --- | --- |
| `FieldNames` | Friendly name for this filename part. |
| `Filename-Part` | One-based filename part number, or `FileExtension`. |
| `GroupOrder` | Grouping level. Use `0` when this field is not a grouping level. |
| `Description` | User note. The current federation logic does not rely on this column. |
| `InjectToIFC` | Adds this filename-derived value during IFC processing. |

Files match the naming rules only when every grouping field with `GroupOrder > 0` has a value in the file name. Matched files are grouped into nested Navisworks outputs. Unmatched files can be placed into a separate unmatched-files NWD, and can be included in the final federation when `IncludeUnmatchedFilesInFederatedModel = Yes`.

`NWDNamingMethod` controls grouped output names:

| Value | Output naming |
| --- | --- |
| `Full` | Includes field name and code/description. Safest against duplicate names. |
| `OnlyCodes` | Uses code values only. |
| `OnlyDesc` | Uses lookup descriptions where available. |
| `Codes-Desc` | Uses `code=description` where available. |

Important lookup note: the current script matches `Lookups.Filename-Part` against the `Federation.Filename-Part` value. If the federation row uses `6`, the lookup row should also use `6`. If the federation row uses `FileExtension`, the lookup row should use `FileExtension`.

## 16. Grouping Tab: Wildcard Selection

Use `Wildcard Selection` when you want explicit ordered rules instead of naming-convention grouping.

| Wildcard column | Meaning |
| --- | --- |
| `Run` | Whether the rule is active. |
| `Inclusions` | Comma-separated wildcard filters. Required for an active rule. |
| `Exclusions` | Comma-separated wildcard filters to remove from the matched set. |
| `ExportFileName` | Output Navisworks file name. Use `.nwf` for NWF; otherwise NWD is used. |
| `ReadFromOutputFolder` | When `Yes`, this rule reads previous outputs from `FederationOutputFolder` instead of source models. |
| `CopyToDestination` | When `Yes`, copies this rule's output to `DestinationFolder`. |

Rules run in order. This allows a two-level workflow:

1. First rules create discipline, area, or package outputs from source models.
2. Later rules use `ReadFromOutputFolder = Yes` to combine those earlier outputs.
3. Final selected outputs can be copied to `DestinationFolder`.

In wildcard mode, the rule names drive the outputs. `FederatedFileName` is not the main output controller for this mode. The script records the last created wildcard output in `federation-summary.json`, and that output becomes the publish candidate for Revizto checks.

## 17. Lookups Tab

Lookups translate filename codes into readable descriptions.

| Column | Meaning |
| --- | --- |
| `Filename-Part` | Must match the `Federation.Filename-Part` value used by the current script, such as `4`, `6`, or `FileExtension`. |
| `Code` | Code found in the file name. |
| `Description` | Friendly description for that code. |

Example:

| Filename-Part | Code | Description |
| --- | --- | --- |
| `6` | `ARC` | Architecture |
| `6` | `STR` | Structures |
| `4` | `L01` | Level 1 |

If lookup rows are missing or do not match, the process still runs, but descriptions may show as `N/A` or output names may fall back to codes.

## 18. Preflight, Preview, and Issue Reports

The GUI has several checks that are worth using before a long run.

| Button | What it does |
| --- | --- |
| `Preview Matches` | Checks active download rows and shows local/source matches where possible. |
| `Preview Grouping` | Shows planned naming-convention groups or wildcard rule matches before Navisworks runs. |
| `Preflight` | Saves the JSON, checks folders, runtime files, Navisworks, Python, ProjectWise commands, disk space, and grid validation. |
| `Report Issue` | Creates a support zip under `%LOCALAPPDATA%\Federation-Automation\IssueReports` and asks you to email it to `gudarz@gmail.com`. |
| `Save and Run` | Saves JSON, runs preflight with folder creation, then starts `FA_Main.exe`. |

`Save and Run` will stop before the pipeline starts if preflight finds errors.

Every compiled EXE has a build version in `yyyyMMdd-index` format. The main pipeline writes its build version near the start and end of the log; include that value when reporting an issue.

## 19. Main Outputs

| Output | Location |
| --- | --- |
| Staged source models | `SourceFolder` |
| Source metadata CSV | `SourceFolder\AttributesFile`, usually `SourceFolder\PWAttributes.csv` |
| Deleted file log | `LogFolder\Deleted_files.csv` |
| Moved deleted source files | `deleted` folder beside `SourceFolder` |
| IFC object extraction CSVs | `IfcDataExtractionFolder` |
| IFC extraction summary | `IfcDataExtractionFolder\ifc-data-extraction-summary.json` |
| Processed IFCs and pass-through files | `ProcessedFolder` |
| Processing summary | `ProcessedFolder\processed-summary.json` |
| Processing change metadata | `ProcessedFolder\processed-metadata.json` |
| Grouped and final Navisworks files | `FederationOutputFolder` |
| Federation summary | `FederationOutputFolder\federation-summary.json` |
| Destination copies | `DestinationFolder` |
| Main run log | `LogFolder\PW_Process.log.txt` or timestamped variants |

For cloud-synced log folders such as OneDrive or SharePoint, the logger tries to reuse a stable `PW_Process.log.txt` file so version history can manage retention. If that file is locked or read-only, it falls back to a timestamped log file.

## 20. Common Workflows

### Federate an Existing Local Folder

Use this when files are already available in a local or synced folder.

1. Set `RunDownload = No`.
2. Set `RunProcess = No` unless IFC metadata injection is required.
3. Set `RunFederation = Yes` or `Force`.
4. Set `FederationInputFolder` to the folder containing the model files.
5. Configure `Grouping`.
6. Run `Preview Grouping`, then `Save and Run`.

### Stage Local or ACC Desktop Connector Files

Use this when the source files are in a project folder and should be copied into a controlled local source set.

1. Set `RunDownload = Yes`.
2. Set `SourceAcquisitionMode = Local` or `Auto`.
3. Add active `Download` rows with local or ACC Desktop Connector `ReadFolder` paths.
4. Set `FileFilter`, `Exclude`, `SkipIfSame`, and `CheckDateToo`.
5. Use `Preview Matches`.
6. Enable federation or processing as needed.

### ProjectWise Source Acquisition

Use this when source files should come from ProjectWise.

1. Set `RunDownload = Yes`.
2. Set `SourceAcquisitionMode = ProjectWise` or `Auto`.
3. Enter ProjectWise `ReadFolder` paths.
4. Leave `PWUser` and `PWPass` blank for Bentley IMS, unless the project requires direct credentials.
5. Confirm ProjectWise PowerShell commands are available.
6. Run `Preflight` before running the pipeline.

### IFC Metadata Injection

Use this when selected source metadata must be visible inside downstream IFC consumers.

1. Set `RunDownload = Yes`, or make sure `SourceFolder` already contains both the model files and `PWAttributes.csv`.
2. Set `RunProcess = Yes` or `Force`.
3. In `Attributes`, set `InjectToIFC = Yes` for the fields to inject.
4. In `Grouping`, set `InjectToIFC = Yes` for filename-derived values you want added.
5. Federation can then read `ProcessedFolder` or an explicit `FederationInputFolder`.

### Wildcard Multi-Level Federation

Use this when you want ordered output rules.

1. Set `FederationGroupingMethod = Wildcard Selection`.
2. Add active wildcard rules for first-level outputs.
3. Add a later rule with `ReadFromOutputFolder = Yes` to combine first-level outputs.
4. Mark final outputs with `CopyToDestination = Yes` if required.
5. Run `Preview Grouping`.

## 21. Troubleshooting

| Message or symptom | What to check |
| --- | --- |
| `Expected Config.json` or `Config file not found` | Place `Config.json` beside the EXE or pass the full path. |
| `Unsupported configuration format '.xlsx'` | This build supports JSON only. Open/save as `Config.json`. |
| Excel export unavailable | Expected in this build. Use `Save` to write JSON. |
| No active Download rows | Set at least one `Download.Run` value to `Yes`, or set `RunDownload = No` to reuse existing source files. |
| ProjectWise commands unavailable | Install/register the ProjectWise PowerShell tooling before running ProjectWise rows. |
| ProjectWise login failed | Check datasource name, path format, IMS access, credentials, and whether ProjectWise tools work in the same Windows session. |
| Local read folder not found | Check the path, synced-folder availability, and whether the path is relative to the EXE folder. |
| Download returned zero files | Check `FileFilter`, `Exclude`, folder path, and source mode. Deleted-file cleanup will not run after an unsafe empty or failed search. |
| Attribute file not found | `AttributesFile` must exist under `SourceFolder` when IFC processing runs. Run download first or provide the CSV manually. |
| Attribute CSV has no `FileName` | Add `FileName` or `Filename`; processing cannot match IFC files without it. |
| Python not detected | Install Python manually or allow the user-level `winget` install. Locked-down machines may block this. |
| IFC extraction skips files | Check max file size and whether existing CSV files are already current. Use `Force` if needed. |
| Navisworks not found | Install Navisworks Manage or correct `NavisworksVersion`. |
| Older NWD save version fails | Provide a valid `NavisworksConfigXML`, or set `NavisworksSavedNwdVersion = Latest`. |
| Wildcard rule matched no files | Use `Preview Grouping`; check `Inclusions`, `Exclusions`, and whether the rule reads from source or output folder. |
| Duplicate wildcard output names | Each active wildcard `ExportFileName` must be unique. |
| Naming-convention files are unmatched | Check dash-separated filename parts and `Federation.Filename-Part` positions. |
| Lookup descriptions show `N/A` | Make `Lookups.Filename-Part` match the `Federation.Filename-Part` value used by the script. |
| Revizto publish skipped | Check `ReviztoPublish`, `ReviztoPublishCode`, model freshness, and whether `ReviztoConsole.exe` exists. |

## 22. Build Notes for Maintainers

To build the main executable:

```powershell
.\000-2Exe.ps1
```

To build the GUI executable:

```powershell
.\000-Gui2Exe.ps1
```

Build notes:

- `FA_Main.exe` should be built before `FA_GUI.exe`.
- Build scripts relaunch under Windows PowerShell Desktop when needed.
- `ps2exe` is required for builds.
- `FA_GUI.exe` expects `FA_Main.exe` beside it at runtime.
- The main build bundles the supporting function scripts.

## 23. Known Naming Quirks

These are not blockers, but they explain some wording you may see:

- `PWAttributesList` is still the JSON section name even though local/synced sources are also supported.
- `ExportToXLSX` is still a column name, but this build writes CSV metadata.
- Some internal errors still say "workbook" or "named range"; in this build that normally refers to the CSV/JSON compatibility layer.
- `MinState` is visible in the Download table but is reserved and currently ignored by the download module.

## 24. License

Federation Automation is licensed under the PolyForm Noncommercial License 1.0.0.

You may use, copy, modify, and share the software for noncommercial purposes under the license terms. Commercial use, including selling the software or using it as part of a paid service or product, requires separate permission from the copyright holder.

Full terms:

https://polyformproject.org/licenses/noncommercial/1.0.0
