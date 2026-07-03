<#
Configuration adapters for Federation Automation.
JSON is the editable, application-friendly format.  Existing Excel named-range
configuration remains supported so current projects do not need to migrate at once.
#>

function ConvertTo-PipelineSettingsRows {
    [CmdletBinding()]
    param($Settings)

    if ($null -eq $Settings) { return @() }
    if ($Settings -is [System.Collections.IEnumerable] -and -not ($Settings -is [string]) -and
        -not ($Settings -is [System.Collections.IDictionary]) -and
        ($Settings | Select-Object -First 1).PSObject.Properties.Name -contains 'Parameter') {
        return @($Settings)
    }

    $rows = @()
    foreach ($property in $Settings.PSObject.Properties) {
        $value = $property.Value
        if ($value -is [pscustomobject] -and ($value.PSObject.Properties.Name -contains 'Value')) {
            $rows += [pscustomobject]@{
                Parameter = $property.Name
                Value     = $value.Value
                Desc      = $value.Desc
            }
        }
        else {
            $rows += [pscustomobject]@{ Parameter = $property.Name; Value = $value; Desc = '' }
        }
    }
    return $rows
}

function Get-FEDAUTOSettingsCatalog {
    <# Canonical settings shown by the JSON-first GUI, including pipeline defaults. #>
    $items = @(
        @('Working folders & metadata','LogFolder','Logs','Folder for run logs. Relative paths are based on the application folder.'),
        @('Working folders & metadata','SourceFolder','SourceFiles','Folder where downloaded or copied source model files are staged.'),
        @('Working folders & metadata','AttributesFile','PWAttributes.xlsx','Excel workbook under SourceFolder that stores captured source metadata.'),
        @('Source acquisition','RunDownload','Yes','Retrieve or stage source model files before running the remaining stages.'),
        @('Source acquisition','SourceAcquisitionMode','Auto','Select Local, ProjectWise, or Auto. Auto supports a mixture of configured source rows.'),
        @('Source acquisition','PWUser','','Optional ProjectWise user name. Leave blank to use the normal Bentley/IMS sign-in.'),
        @('Source acquisition','PWPass','','Optional ProjectWise password. Prefer Windows Credential Manager rather than storing a password in JSON.'),
        @('IFC processing','RunProcess','No','Yes runs IFC processing when needed; No skips it; Force reprocesses all applicable IFC files.'),
        @('IFC processing','ProcessedFolder','ProcessedIFC','Folder for processed IFC files and process summaries.'),
        @('Federation & Navisworks','RunFederation','Yes','Yes uses change-based federation; No disables it; Force always rebuilds the federation.'),
        @('Federation & Navisworks','FederationGroupingMethod','Naming Convention and Lookups','Choose Naming Convention and Lookups or Wildcard Selection.'),
        @('Federation & Navisworks','IncludeUnmatchedFilesInFederatedModel','No','Adds models that do not match federation naming rules into the final federated Navisworks model.'),
        @('Federation & Navisworks','FederationInputFolder','','Leave blank to use ProcessedFolder when processing runs, otherwise SourceFolder.'),
        @('Federation & Navisworks','FederationOutputFolder','Output','Folder where grouped Navisworks files and the final federated model are created.'),
        @('Federation & Navisworks','DestinationFolder','Destination','Folder where selected wildcard outputs are copied after federation finishes.'),
        @('Federation & Navisworks','FederatedFileName','Project Federated.nwd','Final federated model file name. Use .nwf to save only the final model as NWF; otherwise .nwd is added when omitted.'),
        @('Federation & Navisworks','NavisworksVersion','','Preferred installed Navisworks version. Leave blank for automatic detection.'),
        @('Federation & Navisworks','NavisworksConfigXML','NavisworksOptions.xml','Optional Navisworks XML options file.'),
        @('Federation & Navisworks','NavisworksSavedNwdVersion','Latest','NWD file version to write. Latest uses the running Navisworks version.'),
        @('Federation & Navisworks','NavisworksViewsImportXML','','Optional XML file of saved views to import.'),
        @('Federation & Navisworks','NavisworksVisible','No','Yes shows Navisworks while federation runs; No runs it in the background.'),
        @('Federation & Navisworks','NWDNamingMethod','Full','Naming for grouped NWD files: Full, OnlyCodes, OnlyDesc, or Codes-Desc.'),
        @('Revizto publishing','ReviztoPublish','No','Yes allows publish; No disables it; Force publishes when a valid model is available.'),
        @('Revizto publishing','ReviztoPublishCode','','Revizto scheduler publish code, required only when publishing is enabled.'),
        @('Revizto publishing','ReviztoMaxAgeMinutes','60','Maximum age, in minutes, of a federated model that may be published.')
    )
    return @($items | ForEach-Object { [pscustomobject]@{ Section=$_[0]; Parameter=$_[1]; DefaultValue=$_[2]; Desc=$_[3] } })
}

function Merge-FEDAUTOSettingsWithCatalog {
    [CmdletBinding()]
    param([array]$Settings)
    $existing = @{}
    foreach ($row in $Settings) { if ($row -and $row.Parameter) { $existing[$row.Parameter.ToString().ToLowerInvariant()] = $row } }
    $merged = @()
    foreach ($definition in Get-FEDAUTOSettingsCatalog) {
        $key = $definition.Parameter.ToLowerInvariant()
        $row = $existing[$key]
        if (-not $row -and $key -eq 'rundownload') { $row = $existing['runpwdownload'] }
        $value = if ($row -and $null -ne $row.Value -and -not [string]::IsNullOrWhiteSpace($row.Value.ToString())) { $row.Value } else { $definition.DefaultValue }
        $isDefault = -not $row -or $null -eq $row.Value -or [string]::IsNullOrWhiteSpace($row.Value.ToString())
        $merged += [pscustomobject]@{ Section=$definition.Section; Parameter=$definition.Parameter; Value=$value; Desc=$definition.Desc; DefaultValue=$definition.DefaultValue; IsDefault=$isDefault }
    }
    foreach ($row in $Settings) {
        if ($row -and $row.Parameter -and $row.Parameter -ne 'RunPWDownload' -and -not ($merged.Parameter -contains $row.Parameter)) {
            $merged += [pscustomobject]@{ Section='Other settings'; Parameter=$row.Parameter; Value=$row.Value; Desc=$row.Desc; DefaultValue=''; IsDefault=$false }
        }
    }
    return $merged
}

function ConvertTo-PipelineRows {
    [CmdletBinding()]
    param($Rows)
    if ($null -eq $Rows) { return @() }
    return @($Rows | ForEach-Object { [pscustomobject]$_ })
}

function Get-PipelineConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ConfigPath,
        [string]$BasePath
    )

    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        throw "Configuration file not found: $ConfigPath"
    }

    $extension = [System.IO.Path]::GetExtension($ConfigPath).ToLowerInvariant()
    if ($extension -eq '.json') {
        try { $raw = Get-Content -LiteralPath $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
        catch { throw "Configuration JSON is invalid in '$ConfigPath'. $($_.Exception.Message)" }

        return [pscustomobject]@{
            Format           = 'Json'
            Settings         = @(ConvertTo-PipelineSettingsRows $raw.settings)
            Download         = @(ConvertTo-PipelineRows $raw.download)
            PWAttributesList = @(ConvertTo-PipelineRows $raw.pwAttributesList)
            Federation       = @(ConvertTo-PipelineRows $raw.federation)
            WildcardSelection = @(ConvertTo-PipelineRows $raw.wildcardSelection)
            Lookups          = @(ConvertTo-PipelineRows $raw.lookups)
        }
    }

    if ($extension -notin @('.xlsx', '.xlsm')) {
        throw "Unsupported configuration format '$extension'. Use .json, .xlsx, or .xlsm."
    }
    if (-not (Get-Command Get-ExcelDataSafe -ErrorAction SilentlyContinue)) {
        throw 'Excel configuration support is unavailable because Get-ExcelDataSafe was not loaded.'
    }
    # WildcardSelection was introduced after the legacy Excel format.  Its
    # absence in existing workbooks is valid and simply means no rules exist.
    $wildcardSelectionRows = @()
    try { $wildcardSelectionRows = @(Get-ExcelDataSafe -Path $ConfigPath -NamedRange 'WildcardSelection') }
    catch { $wildcardSelectionRows = @() }
    return [pscustomobject]@{
        Format           = 'Excel'
        Settings         = @(Get-SettingsSafe -ConfigPath $ConfigPath -BasePath $BasePath)
        Download         = @(Get-ExcelDataSafe -Path $ConfigPath -NamedRange 'Download')
        PWAttributesList = @(Get-ExcelDataSafe -Path $ConfigPath -NamedRange 'PWAttributesList')
        Federation       = @(Get-ExcelDataSafe -Path $ConfigPath -NamedRange 'Federation')
        WildcardSelection = $wildcardSelectionRows
        Lookups          = @(Get-ExcelDataSafe -Path $ConfigPath -NamedRange 'Lookups')
    }
}

function Save-PipelineJsonConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [array]$Settings,
        [array]$Download,
        [array]$PWAttributesList,
        [array]$Federation,
        [array]$WildcardSelection,
        [array]$Lookups
    )

    $settingsObject = [ordered]@{}
    foreach ($row in $Settings) {
        if ($null -eq $row -or [string]::IsNullOrWhiteSpace($row.Parameter)) { continue }
        $settingsObject[$row.Parameter.ToString()] = $row.Value
    }
    if (@($Settings | Where-Object { $null -ne $_ }).Count -gt 0 -and @($settingsObject.Keys).Count -eq 0) {
        throw 'Settings rows were supplied but none contained a Parameter name. Configuration was not written.'
    }
    $document = [ordered]@{
        schemaVersion    = 1
        settings         = $settingsObject
        download         = @($Download | Where-Object { $null -ne $_ })
        pwAttributesList = @($PWAttributesList | Where-Object { $null -ne $_ })
        federation       = @($Federation | Where-Object { $null -ne $_ })
        wildcardSelection = @($WildcardSelection | Where-Object { $null -ne $_ })
        lookups          = @($Lookups | Where-Object { $null -ne $_ })
    }
    $json = $document | ConvertTo-Json -Depth 12
    $directory = Split-Path -Parent $Path
    if ([string]::IsNullOrWhiteSpace($directory)) { $directory = (Get-Location).Path }
    if (-not (Test-Path -LiteralPath $directory)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }

    # Write and validate a sibling temporary file first.  The original config is
    # only replaced once a complete, parseable JSON document exists on disk.
    $temporaryPath = Join-Path $directory ('.{0}.{1}.tmp' -f ([IO.Path]::GetFileName($Path)), [guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($temporaryPath, $json, [Text.UTF8Encoding]::new($false))
        $null = Get-Content -LiteralPath $temporaryPath -Raw | ConvertFrom-Json -ErrorAction Stop
        Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue }
    }
}
