# Federation Automation Guide (Friendly Version)

Version date: 2026-06-29  
Applies to: `FA_GUI.exe`, `FA_Main.exe`, and the packaged generic configuration templates

## 1. What This Process Does

This process automates the full path from source model files to a published Revizto model. Source files can come from ProjectWise or from normal filesystem folders such as local folders, ACC Desktop Connector, OneDrive, or SharePoint sync locations.

In simple terms, it does four jobs:

1. Build the local source file set.
   The process reads the `Download` range in the config file and stages matching supported model files into `SourceFolder`. Each active `Download` row can point either to a ProjectWise folder or to a local/synced filesystem folder. ProjectWise rows log into the relevant datasource and capture available ProjectWise metadata. Local rows copy files directly from the listed folder without using ProjectWise or the ProjectWise PowerShell module. At the same time the process builds an attributes workbook, usually `PWAttributes.xlsx`, that gives the processing stage a consistent file list and metadata snapshot.
2. Inject selected metadata into IFC files.
   After download, the process reads the attributes workbook together with the `PWAttributesList`, `Federation`, and `Lookups` config ranges. It then creates updated IFC copies in the processed folder, sets the `IfcProject` name to match the full source file name, and adds property sets directly into each IFC so selected metadata and filename-derived values become visible inside downstream tools. This stage does not blindly reprocess everything every time: it keeps metadata records from previous runs and skips IFCs that have not changed in file content or attribute payload, unless processing is forced.
3. Federate supported model files into Navisworks outputs.
   Once source files are ready, choose either **Naming Convention and Lookups** or **Wildcard Selection**. Naming Convention and Lookups uses the `Federation` and `Lookups` rules to build grouped NWDs and a final NWD or NWF from filename parts. Wildcard Selection creates the explicitly named NWD or NWF outputs defined in its ordered rules, including multi-level hierarchies. Federation can run from the processed folder or directly from `SourceFolder`, depending on settings and whether the processing stage ran.
4. Publish the latest valid federated model to Revizto (if enabled).
   If Revizto publishing is enabled in `Settings`, the process checks whether a valid final federated model exists and whether it is new enough to publish. It then uses the configured `ReviztoPublishCode` to run the Revizto console scheduler. This means the script is not just creating files locally; it can also complete the last handoff step and push the newest federated model into the Revizto workflow when the required conditions are met.

## 2. Benefits

Compared with a fully manual federation workflow, this process provides several practical benefits for project teams.

### 2.1 Major time savings

For a large program, manual federation can take several hours every time files need to be prepared, grouped, federated, checked, and published. This process moves most of that repeated work into a configured workflow that can run in a consistent way on a PC or VM.

The initial setup of the config file may take a few hours, and ongoing maintenance is usually much smaller than repeating the full workflow manually each time. Over a project life of 12 to 24 months, the total saving can be significant and may reduce repeat effort by up to about 85%, depending on how often the workflow is run.

### 2.2 Consistency

The other main benefit is consistency. Every run follows the same saved configuration, folder rules, metadata selection, grouping logic, and federation structure. This reduces variation between runs and makes the outputs more predictable for both support staff and end users.

### 2.3 Capturing source metadata

Collecting and saving source metadata manually is repetitive, slow, and easy to do inconsistently. For ProjectWise rows, this process captures the selected ProjectWise attributes and stores them in the attributes workbook. For local or synced folder rows, it records file-level information such as file name, size, update date, source path, and read folder. That gives the team a repeatable snapshot for the staged source file set and removes a large amount of manual checking and copying.

### 2.4 Injecting metadata into IFC files

Adding metadata into IFC files is not a practical manual task for normal project delivery. IFC files are structured text files, and inserting property data correctly requires technical handling that is difficult to do by hand without errors. This process performs that injection automatically so the selected metadata and filename-derived values become available inside the processed IFC files.

### 2.5 Grouping and federating files using the project naming convention

This process can read file-name parts and use them to build consistent file groups and federation hierarchies based on the agreed naming convention. That makes it possible to create smaller grouped outputs as well as a final federated model in a structured and repeatable way. The result is easier navigation, clearer model organisation, and better consistency across the project.

### 2.6 Automated Revizto publishing

When Revizto publishing is enabled, the process can publish the final federated model automatically by using the configured scheduler publish code. This removes another repeated manual step and helps keep Revizto aligned with the latest valid federated output without requiring someone to publish it manually each time.

### 2.7 Logging and traceability

The process creates a detailed log file for each run. That log helps the team understand what happened during the run, which stages ran or were skipped, and where any issue occurred. This makes investigation, support, and future troubleshooting much easier.

### 2.8 Flexible configuration

The process is highly configurable and can be adapted to different project requirements without changing the EXE itself. The JSON configuration controls which stages run, where files are read and written, how outputs are named, and how grouping, processing, and publishing behave. Existing Excel configurations remain supported for import, export, and legacy runs.

For example, any of the main stages can be enabled, disabled, or forced as needed. Folder paths and output file names can be adjusted to suit a specific project, environment, or delivery setup. The EXE can also run different config workbooks by passing the required config file name or full config path as the runtime argument. In practice, this means one EXE can support multiple projects, test setups, or delivery scenarios without needing a separate build for each one.

### 2.9 Faster repeat runs

The process is designed to avoid repeating work when nothing has changed. It can skip files that have already been staged in `SourceFolder`, skip IFCs that have already been processed with the same metadata, and skip federation or publishing when no new valid output is needed.

This makes repeat runs much faster than starting from scratch every time. It is especially useful on long-running projects where the process may be run many times over the life of the job.

### 2.10 Safer handling of deleted and outdated files

The process does not only create new outputs. It also compares the current configured source list against what is already staged in `SourceFolder` and handles removed supported files in a controlled way when it is safe to do so. Old local files can be moved to a `deleted` folder, and deleted-file information can be logged for later reference.

This helps keep the working folders cleaner and reduces the risk of old or obsolete files being left behind and accidentally used in later federation runs.

### 2.11 Better resilience across different environments

The process includes several fallback and compatibility checks to make it more reliable in day-to-day use. For example, it can fall back from direct ProjectWise credentials to Bentley IMS login, and it can fall back to another installed Navisworks version if the requested version is not available.

It also handles common path differences between users and synced environments such as OneDrive or ACC. This makes the workflow easier to reuse across different machines, user profiles, and delivery setups.

## 3. How to Run the EXE

### 3.1 What users need

In the standard packaged setup, the key runtime files are:

- `FA_Main.exe`
- `FA_GUI.exe`
- a valid `Config.json` file (recommended), or a legacy `Config.xlsx` workbook
- `NavisworksOptions.xml` when the configuration references the default Navisworks options file

Place the selected configuration beside the EXE unless you pass a different path at launch. JSON is the primary editable format. If the configuration points to additional external files, such as a custom Navisworks XML file, those files must also be available at the configured paths.

The GitHub/release library keeps runnable files under `Exe_Files`. A normal starter package contains:

- `Exe_Files\FA_GUI.exe`
- `Exe_Files\FA_Main.exe`
- `Exe_Files\Config.json`
- `Exe_Files\Config.xlsx`
- `Exe_Files\NavisworksOptions.xml`

The supplied `Config.json` and `Config.xlsx` are generic templates. They intentionally use relative folders such as `SourceFiles`, `ProcessedIFC`, `Output`, and example source rows such as `ExampleSourceFiles`. Replace those rows with real project folders before enabling source acquisition or federation.

### 3.2 Typical launch examples

The EXE can be started in any of the following ways:

.\FA_Main.exe
.\FA_Main.exe Config.json
.\FA_Main.exe Config.xlsx
.\FA_Main.exe C:\Runs\Project\Config.json

The first example uses the default selection. The next two use a configuration beside the EXE. The final example uses a configuration stored elsewhere, such as a project-specific folder.

### 3.3 How the EXE resolves the config file

The EXE resolves the config input in a simple and predictable way:

- If no argument is passed, it uses `Config.json` when present beside the EXE; otherwise it uses `Config.xlsx`.
- If the argument is a file name or relative path, it is resolved from the folder where the EXE is located.
- If the argument is a full path, that full path is used directly.
- If the config file cannot be found, the run stops immediately.

This allows one EXE to be reused with multiple JSON or Excel configurations while keeping the simple default case of `Config.json` beside the EXE.

### 3.4 What the EXE does at startup

When the EXE starts, it first works from its own folder, then sets the current process execution policy to `RemoteSigned` so required PowerShell modules can load, checks the configuration location, loads its bundled support logic, reads the configuration, and starts logging. The GUI and shared Excel reader also apply the same process-only setting before loading `ImportExcel` for Excel import/export. This does not change the user's machine or profile execution policy. After that, it runs the enabled pipeline stages according to the saved settings.

### 3.5 Using the Federation Automation GUI

`FA_GUI.exe` is the recommended interactive launcher. It opens, edits, validates, saves, and runs the same JSON configuration used by the command-line process.

- **Settings** contains folders, source acquisition, IFC processing, federation, and Revizto options. Section switches enable or disable the related stage.
- **Download**, **Attributes**, and **Grouping** are editable tables. **Lookups** is also available when the Naming Convention and Lookups grouping method is selected.
  - Checkbox columns use compact widths. Hover over column headers for guidance, including wildcard and comma-separated filter examples.
  - Type into the blank row at the bottom to add one row.
  - To add several rows, copy tab-separated rows from Excel, select the first destination cell, and press `Ctrl+V`. The GUI adds required rows automatically.
- **Grouping** contains the grouping-method selector. It shows the filename-part `Federation` table for Naming Convention and Lookups, or the `WildcardSelection` table for Wildcard Selection. Method-specific settings are shown above the active table.
- **Run** saves the configuration, starts the pipeline without showing a separate terminal window, and displays live output. It also shows a progress bar, bold top-level stage, and current detailed activity.

Use **Save** to keep configuration changes without running the process, or **Save and Run** to save and begin a run immediately.

When opening, creating, saving, or exporting configuration files, the GUI starts file dialogs in the folder of the current configuration path. This keeps related JSON, Excel, and XML files together instead of jumping to the last folder used by Windows.

## 4. Introduction to the Config File and How to Set It Up

The behaviour and logic are controlled by a JSON configuration file, typically named `Config.json`. The application can open legacy Excel configurations and export the current JSON settings back to Excel when needed. The current Excel template is `Federation-Automation-Config.xlsx`; `Config.xlsx` is also supported as the default legacy Excel file name when no JSON configuration is present.

For a new project, start from the generic templates:

- `Generic_Config.json`
- `Generic_Config.xlsx`

In the runnable `Exe_Files` package these are copied as `Config.json` and `Config.xlsx` so the EXEs can be tested immediately. The templates contain no project-specific ACC, ProjectWise, Revizto, user-profile, or model-code information.

The configuration contains these collections (Excel uses the same named ranges):

- `Settings`
- `Download`
- `PWAttributesList`
- `Federation`
- `WildcardSelection`
- `Lookups`

Note for legacy Excel configurations: the process does not read arbitrary sheets or tables; it looks for the named ranges/areas listed above. Use the provided current template when maintaining an Excel configuration, especially after schema changes such as `WildcardSelection`, NWC federation, and NWF output support.

### 4.1 `Settings` range (control panel)

This range is mainly used to control the program's behaviour, including:
- Controls stage on/off/force behaviour.
- Defines folders, output file names, and tool options.
- Provides optional credentials and publish controls.

Expected columns:

- `Parameter` (required)
- `Value` (required)
- `Desc` (optional, informational only)

Recognised row keywords (`Parameter`) and accepted values:

Workflow management keywords and assumptions:
- `RunDownload` (legacy alias: `RunPWDownload`)
  - Purpose: Controls whether the source staging/download stage runs.
  - Accepted disable values: `No`, `N`, `False`, `0`, `Ignore`
  - Any other non-empty value -> treated as enabled
  - Missing/blank -> default enabled
- `RunProcess` (aliases: `ForceIfcProcessing`, `ForceIfcProcess`, `ForceProcess`)
  - Purpose: Controls whether IFC attribute processing runs after source staging/download.
  - `Force` -> always process
  - `No`, `N`, `False`, `0`, `Ignore` -> skip process
  - Any other non-empty value -> enable process
  - Missing/blank -> default skip
- `RunFederation` (aliases: `FederationRun`, `RunFederate`, `Federate`, `Federation`)
  - Purpose: Controls whether Navisworks federation is allowed, disabled, or forced.
  - `Force` -> force federation run
  - `No`, `N`, `False`, `0`, `Ignore` -> disable federation
  - Any other non-empty value -> enabled with auto decision rules:
    - Federation runs if at least one IFC changed (processed IFC count > 0, or staged source file count > 0 when processing is skipped and federation input points to `SourceFolder`).
    - Federation runs if deleted source IFCs were detected.
    - Federation runs if the final federated model file is missing.
    - Federation skips only when all of these are true: no changed IFCs, no deleted IFCs, and final model already exists.
  - Missing/blank -> same auto decision rules as above
- `FederationGroupingMethod`
  - `Naming Convention and Lookups` (default) -> uses the `Federation` and `Lookups` collections.
  - `Wildcard Selection` -> uses ordered `WildcardSelection` rules instead.
- `ReviztoPublish` (aliases: `RunRevizto`, `ReviztoRun`, `PublishRevizto`, `Revizto`)
  - Purpose: Controls whether Revizto publish is allowed, disabled, or forced.
  - `Force` -> force publish when valid model exists
  - `No`, `N`, `False`, `0`, `Ignore` -> disable publish
  - Any other non-empty value -> allow publish
  - Missing/blank -> default disabled
- `ReviztoPublishCode`
  - Purpose: Supplies the publish target code required by the Revizto publish command.
  - Required for publish stage to actually run
  - Blank/missing -> publish skipped
- `ReviztoMaxAgeMinutes` / `ReviztoMaxAgeHours` (plus alias variants)
  - Purpose: Sets the freshness window used to decide whether publish should run again.
  - Numeric value > 0 accepted
  - Invalid or <= 0 -> warning and default `60` minutes used

Path and file keywords and assumptions:
- Folder parameters accepted:
  - `LogFolder`: where run logs are written (default `Logs`).
    - If `LogFolder` is a normal local folder, the process creates timestamped log files such as `2026-04-08-2134---PW_Process.log.txt`.
    - If `LogFolder` is inside a OneDrive- or SharePoint-synced folder, the process reuses the same file name `PW_Process.log.txt` so cloud version history can track older runs.
    - In synced folders, the process overwrites `PW_Process.log.txt` at the start of each run.
    - If that synced log file is read-only or locked by another process, the script falls back to a timestamped log file for that run.
  - `SourceFolder`: where staged source model files and the attributes workbook are stored (default `SourceFiles`).
  - `ProcessedFolder`: where processed IFCs and process summary files are written (default `ProcessedIFC`).
  - `FederationInputFolder`: folder used as federation source; if blank, script derives it from run context:
    - uses `ProcessedFolder` when processing is enabled.
    - uses `SourceFolder` when processing is skipped.
- `FederationOutputFolder`: where grouped NWDs and the final federated model are written (default `Output`). Wildcard rules can also read NWDs from this folder to build higher-level outputs.
  - Absolute paths are used as-is
  - Relative paths are resolved from EXE/PS1 folder
  - Known `C:\Users\<other>\...` synced roots are remapped to current user
  - Practical examples:
    - Full absolute path: `C:\Data\Project\ProcessedIFC`
    - Relative path: `ProcessedIFC` or `Output\Models`
    - Absolute synced paths from another user profile are remapped when they use supported roots, for example:
      - `C:\Users\other\OneDrive - Tenant\...` -> current `OneDrive - Tenant\...`
      - `C:\Users\other\OneDrive\...` -> current `OneDrive\...`
      - `C:\Users\other\DC\ACCDocs\...` -> current `DC\ACCDocs\...`
      - `C:\Users\other\<CompanySyncRoot>\...` -> current company sync root derived from `OneDriveCommercial`
- File-name parameters:
  - `AttributesFile`: name of the Excel workbook that stores the staged source metadata snapshot (`Attributes` table) used by the process stage for IFC attribute injection; written under `SourceFolder`.
- `FederatedFileName`: name of the final federated Navisworks model produced by the federation stage; written under `FederationOutputFolder` and used as the publish source for Revizto.
  - If the name ends with `.nwf`, only the final federated model is saved as NWF.
  - If the name ends with `.nwd`, the final federated model is saved as NWD.
  - If no Navisworks extension is supplied, `.nwd` is added automatically.
  - Grouped intermediate models and the unmatched-files model are still saved as NWD.
  - `IncludeUnmatchedFilesInFederatedModel`:
    - Purpose: controls whether the separate unmatched-files NWD is appended into the final federated model.
    - Accepted disable values: `No`, `N`, `False`, `0`, `Ignore`
    - Any other non-empty value -> treated as enabled
    - Missing/blank -> default `No`
    - Important: the separate unmatched-files NWD is still created even when this setting is `No`
  - `NWDNamingMethod`:
    - Purpose: controls how grouped intermediate NWD files are named
    - Accepted values:
      - `OnlyCodes`
      - `OnlyDesc`
      - `Codes-Desc`
      - `Full`
    - Missing/blank -> default `Full`
    - Invalid value -> falls back to `Full`
    - This setting affects grouped intermediate NWD names only
    - It does not rename `FederatedFileName`
  - It does not rename the unmatched-files NWD

`FederatedFileName`, `IncludeUnmatchedFilesInFederatedModel`, and `NWDNamingMethod` apply to **Naming Convention and Lookups** only. Wildcard Selection uses the `ExportFileName` in each rule and does not create an additional automatic final model.

Other important keywords:

- `PWUser`, `PWPass` optional:
  - Used only when at least one active `Download` row points to ProjectWise.
  - If both provided, credentials login is attempted first.
  - If credentials are unsuccessful, Bentley IMS is attempted.
  - If IMS is also unsuccessful, the script stops the source staging/download stage.
- `NavisworksVersion`
  - If missing, highest installed version is auto-detected.
  - If specified but not installed, fallback to highest installed version.
- `NavisworksConfigXML`
  - Defaults to `NavisworksOptions.xml` when blank.
- `NavisworksViewsImportXML`
  - Purpose: optionally imports a Navisworks viewpoints XML file into the final federated NWD.
  - Blank/missing -> ignored
  - If a file name is provided without a path, the script looks for it beside the EXE/PS1
  - If a full path is provided, that full path is used
  - If the value does not end with `.xml`, the script adds `.xml`
  - The viewpoints import runs only when the resolved XML file exists
  - If the XML file cannot be found, the federation still runs and the viewpoints import is skipped with a warning
  - This setting affects the final federated NWD only
- `NavisWorksVisible` / `NavisworksVisible`
  - `Yes` -> GUI mode
  - Any other value -> background/headless mode

### 4.2 `Download` range (source selection)

Purpose:

- Defines which folders/files are read and copied into `SourceFolder`.
- Each active row can point to either ProjectWise or a local/synced filesystem folder.

Expected columns:

- `Run`
- `ReadFolder`
- `FileFilter`
- `Exclude` (optional)
- `SkipIfSame`
- `CheckDateToo`
- `MinState` (currently not used by code)

Mandatory vs optional:

- Required when source staging/download is enabled.
- If no active rows are found, the script stops the source staging/download stage.
- `Exclude` is optional.
- `MinState` is currently ignored.

Accepted values and behaviour:

- `Run`
  - Active only when value is one of: `Yes`, `Y`, `True`, `1`
  - Any other value (including blank) -> row ignored
- `ReadFolder`
  - Required for every active row.
  - For ProjectWise rows, best practice is to copy the folder path directly from ProjectWise Explorer (PWE).
  - ProjectWise forms include:
    - `pw://server:datasource/Documents/Project/Folder`
    - `pw:\server:datasource\Documents\Project\Folder`
    - `datasource:\Project\Folder`
  - For the `pw://` and `pw:\` forms, the script automatically extracts the datasource and removes the leading `Documents\` part before running the search
  - For local/synced rows, use any normal folder path that exists on the machine running the script.
  - Local/synced examples include:
    - `C:\Data\Project\Incoming`
    - `..\Incoming`
    - `C:\Users\<user>\DC\ACCDocs\...`
    - `C:\Users\<user>\OneDrive - Tenant\...`
    - `C:\Users\<user>\<Company SharePoint Sync>\...`
    - `\\server\share\folder`
  - Local/synced folders are read as normal filesystem folders; no ProjectWise login or `PWPS_DAB` module is used for those rows.
  - Rows can be mixed. One row can read ProjectWise, another can read ACC Desktop Connector, another can read OneDrive, and all selected files are copied into the same `SourceFolder`.
  - Local/synced folder reads are not recursive. If a subfolder must be included, add it as a separate `Download` row.
  - If an active local/synced folder does not exist, that row is counted as a failed read-folder search and the script continues with other rows.
- `FileFilter`
  - Wildcards are supported (`*`, `?`)
  - `*` means any number of characters, including none
  - `?` means exactly one character
  - Typical examples:
    - `*.ifc` -> all IFC files
    - `MRP-120-*.ifc` -> all IFC files whose names start with `MRP-120-`
    - `MRP-12?-A.ifc` -> matches names such as `MRP-120-A.ifc` or `MRP-121-A.ifc`, where only one character changes in that position
  - Does not support comma-separated wildcard patterns; if you need different searches, use different rows
- `Exclude`
  - Comma-separated words; any candidate file or ProjectWise document whose name contains any of those words is excluded from the source list
- `SkipIfSame`
  - Controls whether the script should always copy/download the file again or first check whether the staged copy already appears to match the source file
  - `No` (exact) -> force copy/download, even if the staged file already exists
  - Any other value -> use comparison logic before copying/downloading
  - With comparison logic enabled, the script checks the staged file against the source file and skips the copy/download if they are considered the same
  - The exact comparison method depends on `CheckDateToo`:
    - if `CheckDateToo=Yes`, the script compares both file size and file date
    - otherwise, it compares file size only
- `CheckDateToo`
  - `Yes` -> compare date + size
  - Any other value -> compare size only

### 4.3 `PWAttributesList` range (attribute definition map)

Purpose:

- Defines which metadata fields are captured from source files.
- Controls which attributes are written to the attributes Excel workbook.
- Controls which attributes are later injected into the processed IFC files.
- Controls the output label used in the workbook and in the IFC property set.

Expected columns:

- `Row` (optional)
- `AttributeName`
- `OutputName` (optional)
- `ExportToXLSX`
- `InjectToIFC`

Mandatory vs optional:

- Required for meaningful processing when `RunProcess` is enabled.
- If no usable `InjectToIFC` rows are found, the process stage stops.
- `AttributeName` is required for any usable row.
- `OutputName` is optional.
- `Row` is optional.
- Separate `FieldName` and `DataType` columns are not used by the current process for this range.

Accepted values and behaviour:

- `Row`
  - Numeric value used for ordering
  - Blank/invalid -> the script uses the row order in the sheet
- `AttributeName`
  - For ProjectWise rows, this is the actual ProjectWise field name or property path to read during download.
  - For local/synced rows, ProjectWise-only fields are left blank unless they match one of the file-level values the script records, such as `FileName`, `FileSize`, `FileUpdateDate`, `SourceFolder`, `SourcePath`, or `ReadFolder`.
- `OutputName`
  - Optional user-friendly output label
  - When provided, it becomes both the column heading in the attributes Excel workbook and the property name written into the IFC
  - Blank -> defaults to `AttributeName`
- `ExportToXLSX`
  - `Yes`, `Y`, `True`, `1` -> exported to the attributes Excel workbook
  - Any other value -> not exported unless `InjectToIFC` is enabled for that row
- `InjectToIFC`
  - `Yes`, `Y`, `True`, `1` -> included in the process stage and injected into the IFC files
  - Rows enabled for IFC injection are also exported to the attributes Excel workbook because the process stage reads the workbook as its input
  - Any other value -> not injected into IFC
- Attribute value type
  - All values are treated as text by the current process
- Export format
  - The current workflow exports attribute data to the Excel workbook only; it does not use a separate CSV export
- Example
  - If `AttributeName` is `DocumentStatus` and `OutputName` is `Status`, the value is exported to the workbook under `Status`
  - If `InjectToIFC` is also enabled, the same `Status` label is written into the IFC property set


### 4.4 `Federation` range (filename interpretation and grouping)

Purpose:

- Defines how source model filename parts are mapped into group fields and final federation hierarchy.
- Also defines optional filename-part description fields for process enrichment.

Expected columns:

- `FieldNames` (or fallback first column)
- `Filename-Part` (aliases accepted: `FilenamePart`, `Filename_Part`)
- `GroupOrder`
- `InjectToIFC` (used by process-stage filename description injection)
- `Description` (informational for users)

Mandatory vs optional:

- Required when federation stage runs.
- If empty/missing when federation runs -> the script stops the federation stage.
- `Description` is optional (not required by logic).

Accepted values and behaviour:

- `Filename-Part`
  - Normally this is a numeric filename part index, using `1` for the first part, `2` for the second part, and so on
  - For numeric values, the script splits the file name at each `-` and reads the requested position
  - Example: for a file named `MRP-120-C-SME-MOD.ifc`, part `1` is `MRP`, part `2` is `120`, part `3` is `C`, and part `4` is `SME`
  - If `Filename-Part` is blank or not numeric, the script cannot reliably read that code from the file name, so that row does not contribute useful grouping information
  - Special case: `Filename-Part = FileExtension`
  - This tells the script to read the real Windows file extension from the file name, for example `.ifc`, `.rvt`, `.dwg`, or `.dgn`
  - The extension is read from the text after the final `.` in the file name, not from the dash-separated naming parts
  - This special value can be used as a grouping field, an IFC-injection field, or a lookup key just like other federation fields
- `GroupOrder`
  - Controls whether that field is used as a federation grouping level, and in what order the levels are created
  - Any positive number means the field is used for grouping
  - Lower numbers are grouped first and become higher levels in the federation structure
  - Example: if `Discipline` has `GroupOrder = 1`, `Zone` has `GroupOrder = 2`, and `Package` has `GroupOrder = 3`, the process first groups files by `Discipline`, then creates sub-groups by `Zone` inside each discipline, then creates sub-groups by `Package` inside each zone
  - In practical terms, this determines the nested federation hierarchy and the intermediate NWD files that are created
  - `0`, blank, or invalid means that field is not used to split the federation into separate groups
  - Best practice is to use unique positive numbers such as `1`, `2`, `3` so the grouping order is clear
  - Only rows with `GroupOrder > 0` are treated as mandatory for deciding whether a file matches the federation naming pattern
  - Rows with `GroupOrder = 0` can still be used for process-stage filename description injection when `InjectToIFC` is enabled
- `InjectToIFC` (process-stage only)
  - This setting does not control federation grouping itself
  - `Yes`, `Y`, `True`, `1` -> that filename part is also used by the process stage when writing filename-based description properties into IFC files
  - Anything else -> ignored for that IFC injection purpose
- `Description`
  - Optional plain-language note for users
  - Not read by the current script logic

### 4.5 `WildcardSelection` range (explicit wildcard federation)

Use this collection only when `FederationGroupingMethod` is `Wildcard Selection`. The GUI processes its rows strictly from top to bottom, so later rows can build on NWDs created by earlier rows.

Expected columns:

- `Run`
- `Inclusions`
- `Exclusions`
- `ExportFileName`
- `ReadFromOutputFolder`

How each row works:

- `Run` controls whether the wildcard rule is executed. Existing configurations without this column are treated as enabled.
- `Inclusions` is required. Enter one or more comma-separated Windows wildcard patterns. A file is included when it matches **any** pattern. Matching is case-insensitive and includes the file extension; for example, `PRJ*ARC*.ifc`.
- `Exclusions` is optional. Enter comma-separated wildcard patterns. A matching file is excluded if it matches **any** exclusion pattern.
- `ExportFileName` is required and names the Navisworks file created by the rule.
  - If the name ends with `.nwf`, that wildcard row is saved as NWF.
  - If the name ends with `.nwd`, that wildcard row is saved as NWD.
  - If no Navisworks extension is supplied, `.nwd` is added automatically.
  - Each row must have a unique output base name; duplicate names stop the federation to prevent an earlier output from being overwritten.
- `ReadFromOutputFolder` is `No` for source-model rules and `Yes` when the rule should read NWDs already created in `FederationOutputFolder`.

There is no separate final-model checkbox or `FederatedFileName` step in this mode. To create a top-level model, add a later row that reads from the output folder and includes exactly the earlier NWDs you want. The last successfully created wildcard output is treated as the latest top-level result for run-state and optional Revizto publishing.

Example hierarchy:

| Run | Inclusions | Exclusions | ExportFileName | ReadFromOutputFolder |
| --- | --- | --- | --- | --- |
| Yes | `PRJ*ARC*.ifc` |  | `Architecture.nwd` | No |
| Yes | `PRJ*MEP*.ifc` |  | `MEP.nwd` | No |
| Yes | `Architecture.nwd,MEP.nwd` |  | `Project Federated.nwd` | Yes |

Rules with `Run=No` are ignored. Rules with no matches write a warning and are skipped. Federation stops only when no enabled wildcard rule creates an NWD. A rule never includes its own output file.

### 4.6 `Lookups` range (code-to-description mapping)

Purpose:

- Maps filename part codes into readable descriptions.
- Used when translating filename-part codes into human-readable meanings.
- Used by the process stage when injecting filename-part meaning/description values into IFC files.
- Used by the federation stage when building more readable grouping labels.

Expected columns:

- `Filename-Part` (aliases accepted: `FilenamePart`, `Filename_Part`)
- `Code`
- `Description`

Mandatory vs optional:

- Optional overall.
- If missing or invalid, the process still runs but shows warnings.
- If missing or invalid, filename-part description injection into IFC uses fallback text (`N/A`).
- If missing or invalid, federation still runs, but grouping labels use the raw filename code only.

Accepted values and behaviour:

- `Filename-Part`
  - Identifies which filename position the lookup row applies to
  - This must match the same filename-part number used in the `Federation` range
  - Example: if `Filename-Part = 3`, the lookup applies only to the third part of the file name
  - Special case: `Filename-Part = FileExtension` applies the lookup to the real file extension, for example `.ifc`
- `Code`
  - The exact code expected in that filename position
  - Matching is done against the text read from the file name
  - Example: if the third filename part is `C`, the script looks for a row where `Filename-Part = 3` and `Code = C`
- `Description`
  - The readable meaning for that code
  - When a match is found, the process stage writes the IFC value as `Code = Description`
  - Example: a lookup of `3 / C / Civil` becomes `C = Civil` in the IFC property value
  - In federation, the description can also be used when building grouped NWD file names, depending on `NWDNamingMethod`
- Row usage
  - A lookup row is used only when `Filename-Part`, `Code`, and `Description` are all non-blank
  - Blank in any required field -> row ignored
- Matching rule
  - The lookup key is the combination of `Filename-Part` and `Code`
  - This means the same code can have different meanings in different filename positions
- Duplicate keys
  - If multiple rows use the same `Filename-Part` + `Code` combination, the first matching row in the table is used
  - Later duplicates are ignored
- Fallback behaviour
  - If no matching lookup is found during processing, the IFC description value falls back to `N/A`
  - If no matching lookup is found during federation:
    - grouping still uses the raw code
    - `OnlyDesc` falls back to the raw code for naming that segment
    - `Codes-Desc` falls back to the raw code for naming that segment

## 5. Process Run Flow and Key Rules

### 5.1 Global Rules Used Throughout the Run

These rules apply across the whole pipeline, not just one stage.

- False-like values are treated as disabled in many settings:
  - `No`, `N`, `False`, `0`, `Ignore`
- True-like values in row flags are:
  - `Yes`, `Y`, `True`, `1`
- `Force` has special meaning for selected stage settings.
- Blank values can trigger defaults; the default may be enabled or disabled depending on the setting.
- Several settings also accept legacy alias names for backward compatibility. Common examples:
  - Run process: `RunProcess`, `ForceIfcProcessing`, `ForceIfcProcess`, `ForceProcess`
  - Run federation: `RunFederation`, `FederationRun`, `RunFederate`, `Federate`, `Federation`
  - Revizto publish switch: `ReviztoPublish`, `RunRevizto`, `ReviztoRun`, `PublishRevizto`, `Revizto`
  - Revizto publish code: `ReviztoPublishCode`
  - Navisworks config XML: `NavisworksConfigXML`
- Main orchestration defaults:
  - `RunDownload` is effectively enabled unless set false-like; when disabled, source acquisition is skipped and existing files are used.
  - `RunProcess` is disabled when missing.
  - `RunFederation` is conditionally enabled unless explicitly disabled.
  - `ReviztoPublish` is disabled unless enabled.
- If source staging/download is unavailable:
  - Process runs only if `RunProcess=Force`.
  - Federation runs only if `RunFederation=Force`.
- If federation did not run:
  - Revizto can still run only when `ReviztoPublish=Force` and a final model already exists.
- Folder settings are normalised by replacing `/` with `\` and trimming spaces.
- Known synced roots under `C:\Users\<username>\...` are remapped to the current user profile for:
  - OneDrive commercial
  - OneDrive personal
  - ACC docs (`DC\ACCDocs`)
  - Company SharePoint sync root derived from `OneDriveCommercial`
- Tokenized path suffixes such as `@OneDrive`, `@SPSync`, and `@ACC` are not used by the current script logic.

Run-wide output:

- `LogFolder\...\PW_Process.log.txt`
- In local non-synced log folders, the actual file name is usually timestamped per run.
- In OneDrive/SharePoint-synced log folders, the process normally writes to the fixed file name `PW_Process.log.txt` and relies on cloud version history.

### 5.2 Stage A: Source Staging / Download

Purpose: build the local source file set and an attributes workbook snapshot.

How it works:

- Runs only `Download` rows where `Run` is true-like (`Yes`, `Y`, `True`, `1`).
- Reads each active row's `ReadFolder` value and decides whether that row is ProjectWise or local/synced filesystem.
- ProjectWise auth flow, used only for ProjectWise rows:
  - If `PWUser` and `PWPass` are provided, try credentials first.
  - If credentials are unsuccessful or not provided, use Bentley IMS.
  - If both are unsuccessful, the script stops the run.
- Local/synced rows do not use ProjectWise login and do not require the `PWPS_DAB` module.
- Copies or downloads the files that match the configured `ReadFolder` and filter rules into `SourceFolder`.
- Creates or updates the `Attributes` table in `AttributesFile`.
- Handles removed supported source files safely by moving old local copies to `deleted`, logging them, and cleaning related downstream leftovers when it is safe to do so.

Key rules for this stage:

- `Download.Run` controls active rows; inactive rows are ignored.
- `ReadFolder` is required for every active row.
- ProjectWise `ReadFolder` formats include:
  - `pw://server:datasource/Documents/...`
  - `pw:\server:datasource\Documents\...`
  - `datasource:\...`
- Local/synced `ReadFolder` values can be absolute paths, relative paths, UNC paths, ACC Desktop Connector paths, OneDrive paths, or SharePoint sync paths.
- Local/synced folders are not read recursively. Add subfolders as separate rows if needed.
- Different rows can use different source types and different ProjectWise datasources.
- `FileFilter` uses wildcard matching (`*`, `?`) and can contain multiple comma-separated patterns, for example `*ARC*.ifc,*CEW*.nwc`.
- `Exclude` is comma-delimited and removes matches by name.
- `CheckDateToo` influences date-based comparison when `SkipIfSame` is used.
- Source staging/download safety rules:
  - If all searches are unsuccessful, the stage is treated as unavailable.
  - If partial searches are unsuccessful, cleanup deletion logic is skipped to avoid false deletions.
- `MinState` currently exists in the workbook but is not used by the current script logic.

Typical outputs from this stage:

- `SourceFolder\*.<supported model extension>`
- `SourceFolder\AttributesFile` (for example `PWAttributes.xlsx`)
- `deleted\*.<supported model extension>` and `Deleted_files.xlsx` when source files were removed

### 5.3 Stage B: Process IFC Attributes

Purpose: enrich IFC files with selected metadata while carrying forward other supported federation files unchanged.

How it works:

- Uses file-name matching to map each IFC to a row in `Attributes` (column `FileName`, with `Filename` accepted for compatibility).
- Only IFC files are modified in this stage.
- Supported non-IFC federation files such as `dwg`, `dgn`, and `rvt` are copied through unchanged when processing is enabled, so federation can still use the processed folder as its source.
- Updates `IfcProject.Name` to the full source file name, including the extension.
- Performs the IFC rewrite in a single streaming pass so top-level renaming and property-set insertion happen together.
- Injects into IFC property sets:
  - `PWattributes`
  - `FileName_sections_Descriptions`
- Uses metadata fingerprinting (`processed-metadata.json`) to skip unchanged IFCs.
- Writes run summary (`processed-summary.json`) with processed/skipped/error counts.

Key rules for this stage:

- `ExportToXLSX` true-like rows in `PWAttributesList` are exported to the attributes workbook.
- `InjectToIFC` true-like rows in `PWAttributesList` are selected for IFC injection.
- `InjectToIFC` rows are also exported so the processing stage can read them from the workbook.
- `OutputName` blank means the exported column name and IFC property name both default to `AttributeName`.
- Separate `FieldName` and `DataType` columns are not used by the current process for `PWAttributesList`.
- There is no separate CSV fallback in the current workflow; attribute export goes to the Excel workbook.
- All injected attribute values are treated as text by the current process logic.
- If no usable attribute definitions are found, the script stops the processing stage.
- `IfcProject` is the only top-level IFC object renamed by the current process logic.
- The current process does not rename `IfcSite`.
- Match logic:
  - tries exact file name first
  - then tries no-extension matching
- If a file does not satisfy the federation naming pattern, processing still continues
  - selected source metadata is still injected into the IFC
  - filename-derived description properties are skipped only when the required federation grouping parts are missing
- If `Lookups` is missing or invalid:
  - a loud warning is shown
  - filename-part descriptions become `N/A`
  - processing continues

Typical outputs from this stage:

- `ProcessedFolder\*.ifc`
- `ProcessedFolder\processed-metadata.json`
- `ProcessedFolder\processed-summary.json`

### 5.4 Stage C: Federation

Purpose: generate Navisworks federation outputs using the selected federation grouping method.

How it works:

- Uses supported federation source files from the federation input folder, currently `.ifc`, `.dwg`, `.dgn`, `.rvt`, and `.nwc`.
- NWC files are accepted, but generated cache files are skipped when the matching source model is present in the same folder. For example, `Model.nwc` or `Model.ifc.nwc` is skipped when `Model.ifc` is also present.
- **Naming Convention and Lookups** uses `Federation` filename-part rules, optional `Lookups`, bottom-up grouping, `NWDNamingMethod`, and the optional unmatched-files NWD.
- **Wildcard Selection** evaluates the `WildcardSelection` rules in row order. Source rules read supported model files, including standalone NWC files after generated-cache filtering. Output-folder rules read prior Navisworks outputs and can form higher-level models.
- A wildcard row with no matches is warned and skipped. Duplicate wildcard `ExportFileName` base names stop the stage rather than overwrite an output.
- Wildcard Selection creates only the outputs specified by its rules; create the top-level federation explicitly as a final output-folder rule. A wildcard output is NWF only when that row's `ExportFileName` ends with `.nwf`; otherwise it defaults to NWD.

Key rules for this stage:

- If `FederationInputFolder` is blank, the source folder is derived from run context.
- `GroupOrder > 0` defines grouping hierarchy.
- Normal dash-separated part extraction uses numeric `Filename-Part` values.
- Special case: `Filename-Part = FileExtension` reads the real Windows file extension, including the leading dot.
- Only rows with `GroupOrder > 0` are used to decide whether a file matches the federation naming pattern.
- `FileExtension` can participate in grouping when it has a positive `GroupOrder`, but it does not change the matched/unmatched naming-pattern test.
- In other words, a file must first match the normal dash-separated naming convention before `FileExtension` can be used as an extra hierarchy level.
- IFC and other supported model files that do not match the federation naming pattern are still federated into a separate unmatched-files NWD.
- If all candidates are unmatched, the stage can still create the unmatched-files NWD.
- Current `Lookups` use in this stage is limited to:
  - `Filename-Part`
  - `Code`
  - `Description`
- `NWDNamingMethod` affects grouped intermediate NWD names only:
  - `OnlyCodes` -> uses raw codes only
  - `OnlyDesc` -> uses descriptions only, with fallback to code when description is missing
  - `Codes-Desc` -> uses `Code=Description`, with fallback to code when description is missing
  - `Full` -> uses the legacy full format such as `Field_Code (Description)`
- If `NWDNamingMethod` produces duplicate grouped NWD names, the federation stage stops with a naming-collision error so outputs are not overwritten.
- The unmatched-files NWD name is built automatically as:
  - `<FederatedFileName without extension> - Unmatched Files.nwd`
- The final federated model includes the unmatched-files NWD only when `IncludeUnmatchedFilesInFederatedModel` is enabled.
- The final federated model is saved as NWF only when `FederatedFileName` explicitly ends with `.nwf`; otherwise it is saved as NWD.
- Revizto publishing uses the final/top-level model path detected by the run. If the final model is NWF, the user must configure the Revizto scheduler task to publish that NWF.
- If `NavisworksViewsImportXML` resolves to an existing XML file, the final federated model command adds:
  - `-ExecuteAddInPlugin NativeExportPluginAdaptor_XmlViewpointsImportPlugin_Import.Navisworks <xml file>`
- The viewpoints import is applied to the final federated model only.
- Grouped intermediate NWDs and the unmatched-files NWD are not affected by `NavisworksViewsImportXML`.
- Navisworks handling:
  - If the requested version is not installed, the script tries the highest installed version.
  - If no Navisworks installation is found, the federation stage is skipped.
- Output safety checks:
  - no supported federation files found -> stage stops
  - grouped NWD naming collision -> stage stops
  - no grouped NWDs and no unmatched-files NWD produced -> stage stops
  - final federated model missing -> stage stops, except when only unmatched files exist and `IncludeUnmatchedFilesInFederatedModel=No`

Typical outputs from this stage:

- `FederationOutputFolder\*.nwd` (grouped NWDs and unmatched-files NWD when needed)
- the final federated file as `.nwd` or `.nwf`, based on `FederatedFileName`
- `FederationOutputFolder\federation-summary.json`

### 5.5 Stage D: Revizto Publish

Purpose: publish only when publish is enabled and a valid, sufficiently recent federated model is available.

How it works:

- Requires `ReviztoPublishCode`.
- Uses model age gates (`ReviztoMaxAgeMinutes/Hours`, default 60 min).
- In normal operation, publish runs only when a new federated model was created in the current run.
- Skips publish when disabled, when the federated model is missing, or when the model is older than the allowed freshness window.
- Force-publish is possible, but only when an existing final federated model is available.

Key rules for this stage:

- Revizto executable path is fixed in the script:
  - `C:\Program Files\Revizto SA\Revizto5\Service\ReviztoConsole.exe`
- If the publish code is missing or the Revizto executable is missing, publish is skipped.
- Default max model age is 60 minutes when no valid age setting is provided.
- Invalid age values fall back to the default with warning.

## 6. Common Messages (Plain Meaning)

- `Config file not found`: wrong config argument or missing file.
- `Every active Download row must have a ReadFolder value`: at least one active row is missing its source folder.
- `No active rows found in the Download named range`: no `Download` rows have `Run` set to a true-like value.
- `ProjectWise login ... and Bentley IMS fallback`: both auth methods were unsuccessful for a ProjectWise row; script stops the source staging/download stage.
- `LOOKUPS TABLE WARNING`: process continues, but lookup descriptions default to `N/A`.
- `IFCPROJECT not found`: the IFC does not contain the expected root project entity, so the processing stage stops for that file.
- `Federation input folder not found`: path setting or run context issue.
- `filename-derived attributes skipped; name does not match Federation pattern`: the file was still processed, but filename-based IFC description properties were skipped because the required grouping parts were not present.
- `Federation stopped: final federated model was not created`: Navisworks did not produce expected output, so the script stops the federation stage.
- `NWD naming collision detected`: the selected `NWDNamingMethod` produced duplicate grouped NWD names; use `Full` or adjust descriptions so names are unique.
- `Final federated model was not created because only unmatched files were available and IncludeUnmatchedFilesInFederatedModel is disabled`: the separate unmatched-files NWD was created, but it was intentionally not appended into the final project NWD.
- `Revizto publish skipped`: disabled setting, missing code/exe/model, or model older than allowed age.

## 7. Good Operating Habits

- Keep `Config.json` under controlled backup/versioning. Export an Excel copy only when a spreadsheet review or bulk edit is useful.
- Change one settings area at a time, then run and verify logs.
- Keep file naming conventions aligned with `Federation` definition.
- Use `Force` only when you understand what normal gating is preventing.
- Treat `Desc` column text as guidance only; script logic is defined by code, not by description text.

## 8. Screenshots

Use:

- `Docs/Federation-Automation-Screenshot-Checklist.md`

Recommended placement:

- Add one screenshot at the start of each stage section.
- Add one screenshot in each config-tab explanation under Section 4 (`Settings`, `Download`, `PWAttributesList`, `Federation`, `Lookups`).
