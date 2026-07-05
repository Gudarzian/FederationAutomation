# Federation Automation User Manual

Federation Automation prepares BIM model delivery sets by running a configurable pipeline. It can stage source models, extract IFC object data to CSV, inject metadata into IFC files, group models, build Navisworks NWD outputs, and optionally publish the latest valid federation to Revizto.

The public project page is available at:
https://github.com/Gudarzian/FederationAutomation

## How the Stages Connect

The stages are connected by the folders in Settings.

1. Source acquisition copies local files or retrieves ProjectWise files into `SourceFolder`.
2. IFC data extraction can read IFC files from the source set and write CSV files into `IfcDataExtractionFolder`.
3. IFC processing can read IFC files and write metadata-injected copies into `ProcessedFolder`.
4. Federation reads either `ProcessedFolder`, `SourceFolder`, or an explicit `FederationInputFolder`, then writes grouped and final Navisworks outputs into `FederationOutputFolder`.
5. Destination copying and Revizto publishing use the federation outputs when those options are enabled.

This means you can run the whole pipeline, or disable stages that are not needed for a particular update.

## Selective Running

Use the activation controls in the Settings tab to decide which stages run.

- `RunDownload`: `Yes` stages files from the Download rows; `No` reuses files already in the source folder.
- `RunIfcDataExtraction`: `Yes` exports IFC object data to CSV; `No` skips reporting; `Force` extracts even if current CSV files already exist.
- `RunProcess`: `Yes` updates IFC metadata when required; `No` skips IFC modification; `Force` rewrites applicable IFC files.
- `RunFederation`: `Yes` builds Navisworks outputs when needed; `No` skips federation; `Force` rebuilds even when no upstream changes are detected.
- `ReviztoPublish`: `Yes` publishes when configured and valid; `No` skips publishing; `Force` publishes an available valid model.

## Sections

### General Settings

General settings define shared folders, Navisworks defaults, and the stage activation controls. Configure these first because later sections use these folders to pass results between stages.

### Source Acquisition

Source acquisition uses the Download rows to copy local files or retrieve ProjectWise files. Filters, exclusions, minimum workflow state, and skip-same checks decide which files are staged. Disable this section when the models are already in the source folder.

### Attributes

The Attributes tab defines which source metadata fields are exported and which fields are injected into IFC files during processing. The same definitions also make the processing output easier to audit.

### IFC Data Extraction

Data Extraction exports object-level IFC data to CSV. It is a reporting step and does not modify model files. Enabled rules filter which IFC files are extracted; if no enabled rules exist, every IFC in the configured source folder can be extracted.

### IFC Processing

IFC processing injects selected metadata into IFC models and writes processed copies. Federation can use those processed files, or the processing stage can be disabled so federation uses the original staged source files.

### Grouping and Federation

Grouping defines how model files become grouped NWDs and a final federated model. `Naming Convention and Lookups` uses filename parts and lookup descriptions. `Wildcard Selection` uses explicit include/exclude rules and can also copy selected outputs to the destination folder.

### Lookups

Lookups translate codes found in filenames into readable descriptions. They are used by grouping and naming rules when the selected federation method needs descriptive output names.

### Revizto Publishing

Revizto publishing is optional and depends on a valid federated model and publish code. Leave it disabled when only local NWD outputs are required.

## License

Federation Automation is licensed under the PolyForm Noncommercial License 1.0.0.

You may use, copy, modify, and share the software for noncommercial purposes under the license terms. Commercial use, including selling the software or using it as part of a paid service or product, requires separate permission from the copyright holder.

Full terms:
https://polyformproject.org/licenses/noncommercial/1.0.0
