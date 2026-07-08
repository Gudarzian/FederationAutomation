<#
Main pipeline entry point.
- Locates supporting scripts/resources (even when packaged as resources).
- Orchestrates download, processing, federation, and publish steps.
- Reads config-driven settings to honor skip/force flags.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$ConfigFile
)

function Enable-FEDAUTOProcessModuleLoading {
    # JSON/CSV build: intentionally does not change PowerShell execution policy.
    return
}

Enable-FEDAUTOProcessModuleLoading

# Resolve the folder that this script, or the compiled executable, lives in.
function Get-FEDAUTOApplicationBasePath {
    $hostProcessNames = @('powershell', 'pwsh', 'powershell_ise')
    $candidatePaths = New-Object System.Collections.Generic.List[string]
    try {
        $assembly = [System.Reflection.Assembly]::GetEntryAssembly()
        if ($assembly -and -not [string]::IsNullOrWhiteSpace($assembly.Location)) { [void]$candidatePaths.Add($assembly.Location) }
    }
    catch { }
    try {
        $processPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if (-not [string]::IsNullOrWhiteSpace($processPath)) { [void]$candidatePaths.Add($processPath) }
    }
    catch { }

    foreach ($candidatePath in $candidatePaths) {
        try {
            if ([System.IO.Path]::GetExtension($candidatePath) -ieq '.exe') {
                $processName = [System.IO.Path]::GetFileNameWithoutExtension($candidatePath)
                if ($hostProcessNames -notcontains $processName) { return (Split-Path -Parent $candidatePath) }
            }
        }
        catch { }
    }

    if ($PSCommandPath) { return (Split-Path -Parent $PSCommandPath) }
    if ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
    return (Get-Location).ProviderPath
}
$basePath = Get-FEDAUTOApplicationBasePath

function Get-FEDAUTOExecutableBuildInfo {
    param([string]$DefaultName = 'FA_Main')

    $nameVariable = Get-Variable -Name FEDAUTOExecutableName -Scope Script -ErrorAction SilentlyContinue
    $versionVariable = Get-Variable -Name FEDAUTOExecutableBuildVersion -Scope Script -ErrorAction SilentlyContinue
    $timeVariable = Get-Variable -Name FEDAUTOExecutableBuildTime -Scope Script -ErrorAction SilentlyContinue
    $name = if ($nameVariable -and -not [string]::IsNullOrWhiteSpace($nameVariable.Value)) { $nameVariable.Value } else { $DefaultName }
    $version = if ($versionVariable -and -not [string]::IsNullOrWhiteSpace($versionVariable.Value)) { $versionVariable.Value } else { 'source' }
    $builtAt = if ($timeVariable -and -not [string]::IsNullOrWhiteSpace($timeVariable.Value)) { $timeVariable.Value } else { '' }
    return [pscustomobject]@{ Name = $name; Version = $version; BuiltAt = $builtAt }
}

function Format-FEDAUTOExecutableBuildInfo {
    param($BuildInfo)
    if ($BuildInfo -and -not [string]::IsNullOrWhiteSpace($BuildInfo.BuiltAt)) {
        return ("{0} build {1} ({2})" -f $BuildInfo.Name, $BuildInfo.Version, $BuildInfo.BuiltAt)
    }
    return ("{0} build {1}" -f $BuildInfo.Name, $BuildInfo.Version)
}

$script:FEDAUTOCurrentBuildInfo = Get-FEDAUTOExecutableBuildInfo -DefaultName 'FA_Main'

# Ensure relative paths resolve from the script/EXE folder.
$locationPushed = $false
try {
    Push-Location -Path $basePath
    $locationPushed = $true
}
catch {
    Write-Warning "Unable to set working directory to '$basePath'. Relative paths will use the current location."
}

function Resolve-ConfigFilePath {
    param(
        [string]$InputPath,
        [string]$RootPath
    )
    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        foreach ($defaultName in 'Config.json') {
            $defaultPath = Join-Path $RootPath $defaultName
            if (Test-Path -LiteralPath $defaultPath -PathType Leaf) { return $defaultPath }
        }
        throw "No default configuration was found in '$RootPath'. Expected Config.json."
    }
    $candidate = $InputPath.Trim()
    if ([System.IO.Path]::IsPathRooted($candidate)) { return $candidate }
    return (Join-Path $RootPath $candidate)
}

# Runtime CLI now provides only the first positional argument as the config file.
$ConfigFile = Resolve-ConfigFilePath -InputPath $ConfigFile -RootPath $basePath
if (-not (Test-Path $ConfigFile -PathType Leaf)) {
    throw "Config file not found: $ConfigFile"
}

# Load split support files; packaged EXE runs load embedded resources when disk files are unavailable.
# If local script files exist, always reload them for PS1 runs so updates take effect.
$functionFiles = @(
    "012-SharedFunctions.Ps1",
    "013-ConfigFunctions.ps1",
    "021-DownloadFunctions.Ps1",
    "031-ProcessFunctions.Ps1",
    "051-IfcDataExtractionFunctions.ps1",
    "041-FederationFunctions.Ps1"
)
$localPaths = $functionFiles | ForEach-Object { Join-Path $basePath $_ }
$missingLocal = @($localPaths | Where-Object { -not (Test-Path $_) })
$needsFunctions = -not (Get-Command Resolve-SupportFile -ErrorAction SilentlyContinue) -or
    -not (Get-Command Update-IfcProjectProperties -ErrorAction SilentlyContinue) -or
    -not (Get-Command Invoke-IfcDataExtraction -ErrorAction SilentlyContinue)
$forceReload = $PSCommandPath -and ($missingLocal.Count -eq 0)
if ($needsFunctions -or $forceReload) {
    if ($missingLocal.Count -eq 0) {
        foreach ($path in $localPaths) {
            # Local dev/editor runs: load from disk so debugging in VS Code stays normal.
            . $path
        }
    }
    else {
        # Embedded EXE run: load functions directly from resources without extraction.
        $asm = [System.Reflection.Assembly]::GetEntryAssembly()
        if (-not $asm) {
            throw ("Support files not found beside the script: {0}" -f ($functionFiles -join ', '))
        }

        $resourceNames = $asm.GetManifestResourceNames()
        function Import-EmbeddedScript {
            param(
                $Assembly,
                [string]$ResourceName
            )
            $stream = $Assembly.GetManifestResourceStream($ResourceName)
            try {
                $reader = New-Object System.IO.StreamReader($stream)
                $scriptText = $reader.ReadToEnd()
            }
            finally {
                if ($reader) { $reader.Dispose() }
                if ($stream) { $stream.Dispose() }
            }
            $scriptBlock = [ScriptBlock]::Create($scriptText)
            . $scriptBlock
        }

        foreach ($file in $functionFiles) {
            $resourceName = $resourceNames | Where-Object { $_ -like "*$file" } | Select-Object -First 1
            if (-not $resourceName) {
                throw "Embedded support file not found: $file"
            }
            Import-EmbeddedScript -Assembly $asm -ResourceName $resourceName
        }
    }
}

function Get-SettingValue {
    param(
        [array]$Settings,
        [string[]]$Names
    )
    if (-not $Settings -or -not $Names -or $Names.Count -eq 0) { return $null }
    $lookup = $Names | ForEach-Object { $_.ToLowerInvariant() }
    foreach ($row in $Settings) {
        if (-not $row) { continue }
        $paramRaw = $row.Parameter
        $paramText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $paramRaw
        }
        else {
            if ($null -ne $paramRaw) { $paramRaw.ToString().Trim() } else { $null }
        }
        if ([string]::IsNullOrWhiteSpace($paramText)) { continue }
        if ($lookup -contains $paramText.ToLowerInvariant()) {
            return $row.Value
        }
    }
    return $null
}

# Pre-resolve supporting assets so downstream functions can use fixed paths.
$supporting = @(
    "NavisworksOptions.xml"
) | ForEach-Object { Resolve-SupportFile -FileName $_ -BasePath $basePath } | Out-Null

# Internal control flags are settings-driven (no CLI overrides).
$ForceProcess = $false
$SkipDownload = $false
$SkipProcess = $false
$SkipFederate = $false
$LogsFolder = $null

$global:MainTranscriptActive = $false
$global:MainLogInfo = $null
$mainTranscriptStarted = $false
$pipelineTimer = [System.Diagnostics.Stopwatch]::StartNew()
$stepResults = New-Object System.Collections.Generic.List[object]
$mainLogInfo = $null
$resolvedConfig = $ConfigFile
$settingsCache = @()
$settingsMissing = $true

# Load Settings and required configuration data once in main through the JSON adapter.
if (-not (Get-Command Get-PipelineConfiguration -ErrorAction SilentlyContinue)) {
    throw 'Configuration adapter was not loaded.'
}
$pipelineConfiguration = Get-PipelineConfiguration -ConfigPath $resolvedConfig -BasePath $basePath
$settingsCache = @($pipelineConfiguration.Settings)
if ($settingsCache.Count -eq 1 -and $null -eq $settingsCache[0]) { $settingsCache = @() }
$settingsMissing = -not $settingsCache -or $settingsCache.Count -eq 0
$downloadRangeCache = @($pipelineConfiguration.Download)
$pwAttributesListCache = @($pipelineConfiguration.PWAttributesList)
$federationRangeCache = @($pipelineConfiguration.Federation)
$wildcardSelectionCache = @($pipelineConfiguration.WildcardSelection)
$ifcDataExtractionRulesCache = @($pipelineConfiguration.IfcDataExtractionRules)
$lookupsRangeCache = @($pipelineConfiguration.Lookups)

# Resolve log file path for the main transcript.
$logFolderValue = $LogsFolder
if (-not $logFolderValue -and $settingsCache -and $settingsCache.Count -gt 0) {
    $logFolderValue = $settingsCache | Where-Object { $_.Parameter -eq "LogFolder" } | Select-Object -ExpandProperty Value
}
if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
    $logFolderValue = CureFolderPath $logFolderValue
}
if (-not $logFolderValue) { $logFolderValue = Join-Path $basePath 'Logs' }
try {
    if (-not (Test-Path $logFolderValue)) { New-Item -Path $logFolderValue -ItemType Directory -Force | Out-Null }
}
catch {
    Write-Warning "Could not create or access log folder '$logFolderValue'. Falling back to TEMP. Error: $_"
    $logFolderValue = Join-Path $env:TEMP 'PWLogs'
    if (-not (Test-Path $logFolderValue)) { New-Item -Path $logFolderValue -ItemType Directory -Force | Out-Null }
}
if (Get-Command Get-LogFilePath -ErrorAction SilentlyContinue) {
    $mainLogInfo = Get-LogFilePath -LogsFolder $logFolderValue -BaseName 'PW_Process.log.txt'
}
else {
    $mainLogInfo = [pscustomobject]@{
        Path   = (Join-Path $logFolderValue 'PW_Process.log.txt')
        Append = $true
    }
}
$mainLogInfo = [pscustomobject]@{
    Path      = $mainLogInfo.Path
    Append    = $mainLogInfo.Append
    BaseName  = 'PW_Process.log.txt'
    LogsFolder = $logFolderValue
    Reset     = $(if ($mainLogInfo.PSObject.Properties.Name -contains 'Reset') { $mainLogInfo.Reset } else { $false })
}
$global:MainLogInfo = $mainLogInfo
try {
    if (Get-Command Initialize-LogFileTarget -ErrorAction SilentlyContinue) {
        Initialize-LogFileTarget -LogInfo $mainLogInfo
    }
    Start-Transcript -Path $mainLogInfo.Path -Append:$($mainLogInfo.Append) -ErrorAction Stop | Out-Null
    $mainTranscriptStarted = $true
    $global:MainTranscriptActive = $true
}
catch {
    Write-Warning "Transcript could not be started. Logging will be console-only. Error: $_"
}
Write-Host "Logging to $($mainLogInfo.Path)"
Write-Host ("Executable build version: {0}" -f (Format-FEDAUTOExecutableBuildInfo $script:FEDAUTOCurrentBuildInfo))

# Log settings resolution (defaults vs Settings) for traceability.
Write-Host ""
Write-Host "=== Settings Resolution ===" -ForegroundColor Cyan
if ($settingsMissing) {
    Write-Host "Settings section missing; defaults will be used where applicable." -ForegroundColor Yellow
}
if (Test-Path $resolvedConfig) {
    Write-Host ("ConfigFile: {0}" -f $resolvedConfig)
}
function Write-SettingsLine {
    param(
        [string]$Name,
        [string]$Value,
        [string]$Source,
        [switch]$IsMissing
    )
    $displayValue = if ($null -eq $Value -or [string]::IsNullOrWhiteSpace($Value)) { "[missing]" } else { $Value }
    Write-Host ("{0}: " -f $Name) -NoNewline -ForegroundColor Green
    Write-Host $displayValue -NoNewline -ForegroundColor Red
    Write-Host (" [{0}]" -f $Source) -ForegroundColor Yellow
}
function Resolve-SettingEntry {
    param(
        [array]$Settings,
        [string[]]$Names,
        [string]$DefaultValue,
        [string]$DisplayName,
        [switch]$Sensitive
    )
    $nameLabel = if ($DisplayName) { $DisplayName } else { ($Names -join "/") }
    $value = Get-SettingValue -Settings $Settings -Names $Names
    $source = "Settings"
    $defaultLabel = if (-not $Settings -or $Settings.Count -eq 0) { "Default (Settings missing)" } else { "Default" }
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace($value.ToString())) {
        $value = $DefaultValue
        $source = $defaultLabel
    }
    $displayValue = if ($Sensitive) { if ($value) { "[set]" } else { "[missing]" } } else { $value }
    $isMissing = ($displayValue -eq "[missing]")
    Write-SettingsLine -Name $nameLabel -Value $displayValue -Source $source -IsMissing:$isMissing
}

$sourceFolderDefault = "SourceFiles"
$processedDefault = "ProcessedIFC"
$attributesDefault = "PWAttributes.csv"
$outputDefault = "Output"
$logDefault = "Logs"
$navisworksOptionsDefault = "NavisworksOptions.xml"
$federatedDefault = "Project Federated.nwd"

Resolve-SettingEntry -Settings $settingsCache -Names @('PWUser') -DefaultValue '' -DisplayName 'PWUser'
Resolve-SettingEntry -Settings $settingsCache -Names @('PWPass') -DefaultValue '' -DisplayName 'PWPass' -Sensitive
Resolve-SettingEntry -Settings $settingsCache -Names @('SourceFolder') -DefaultValue $sourceFolderDefault -DisplayName 'SourceFolder'
Resolve-SettingEntry -Settings $settingsCache -Names @('ProcessedFolder') -DefaultValue $processedDefault -DisplayName 'ProcessedFolder'
Resolve-SettingEntry -Settings $settingsCache -Names @('AttributesFile') -DefaultValue $attributesDefault -DisplayName 'AttributesFile'
Resolve-SettingEntry -Settings $settingsCache -Names @('LogFolder') -DefaultValue $logDefault -DisplayName 'LogFolder'
Resolve-SettingEntry -Settings $settingsCache -Names @('IfcDataExtractionFolder') -DefaultValue 'IFCDataExtraction' -DisplayName 'IfcDataExtractionFolder'
Resolve-SettingEntry -Settings $settingsCache -Names @('IfcDataExtractionMaxFileSizeMB') -DefaultValue '150' -DisplayName 'IfcDataExtractionMaxFileSizeMB'
Resolve-SettingEntry -Settings $settingsCache -Names @('IfcDataExtractionSkipIfCsvIsCurrent') -DefaultValue 'Yes' -DisplayName 'IfcDataExtractionSkipIfCsvIsCurrent'
Resolve-SettingEntry -Settings $settingsCache -Names @('FederationOutputFolder') -DefaultValue $outputDefault -DisplayName 'FederationOutputFolder'
Resolve-SettingEntry -Settings $settingsCache -Names @('DestinationFolder') -DefaultValue 'Destination' -DisplayName 'DestinationFolder'

$runDownloadValue = Get-SettingValue -Settings $settingsCache -Names @('RunDownload','RunPWDownload')
$runDownloadEnabled = $true
$runDownloadSource = if ($runDownloadValue) { "Settings" } else { if ($settingsMissing) { "Default (Settings missing)" } else { "Default" } }
if ($runDownloadValue) {
    $normalizedDownloadRun = $runDownloadValue.ToString().Trim().ToLowerInvariant()
    if ($normalizedDownloadRun -in @('no','n','false','0','ignore')) { $runDownloadEnabled = $false }
}
$runDownloadDisplay = if ($runDownloadEnabled) { "Yes" } else { "No" }
Write-SettingsLine -Name "RunDownload" -Value $runDownloadDisplay -Source $runDownloadSource

$federatedFileSetting = Get-SettingValue -Settings $settingsCache -Names @('FederatedFileName')
$federatedFileSource = if ($federatedFileSetting) { "Settings" } else { if ($settingsMissing) { "Default (Settings missing)" } else { "Default" } }
if ([string]::IsNullOrWhiteSpace($federatedFileSetting)) { $federatedFileSetting = $federatedDefault }
$finalModelIsNwf = if (Get-Command Test-FinalNavisworksFileNameIsNwf -ErrorAction SilentlyContinue) {
    Test-FinalNavisworksFileNameIsNwf -Name $federatedFileSetting
}
else {
    $federatedFileSetting.EndsWith('.nwf', [System.StringComparison]::OrdinalIgnoreCase)
}
$finalOutputExtension = if ($finalModelIsNwf) { '.nwf' } else { '.nwd' }
if (Get-Command Resolve-NavisworksOutputBaseName -ErrorAction SilentlyContinue) {
    $federatedFileSetting = Resolve-NavisworksOutputBaseName -Name $federatedFileSetting
}
elseif ($federatedFileSetting.EndsWith('.nwd', [System.StringComparison]::OrdinalIgnoreCase) -or $federatedFileSetting.EndsWith('.nwf', [System.StringComparison]::OrdinalIgnoreCase)) {
    $federatedFileSetting = $federatedFileSetting.Substring(0, $federatedFileSetting.Length - 4)
}
$federatedFileSetting = "{0}{1}" -f $federatedFileSetting, $finalOutputExtension
Write-SettingsLine -Name "FederatedFileName" -Value $federatedFileSetting -Source $federatedFileSource

$includeUnmatchedSettingValue = Get-SettingValue -Settings $settingsCache -Names @('IncludeUnmatchedFilesInFederatedModel')
$includeUnmatchedSettingSource = if ($includeUnmatchedSettingValue) { "Settings" } else { if ($settingsMissing) { "Default (Settings missing)" } else { "Default" } }
$includeUnmatchedFilesInFederatedModel = $false
if ($includeUnmatchedSettingValue) {
    $normalizedIncludeUnmatched = $includeUnmatchedSettingValue.ToString().Trim().ToLowerInvariant()
    if ($normalizedIncludeUnmatched -notin @('no','n','false','0','ignore')) {
        $includeUnmatchedFilesInFederatedModel = $true
    }
}
$includeUnmatchedDisplay = if ($includeUnmatchedFilesInFederatedModel) { "Yes" } else { "No" }
Write-SettingsLine -Name "IncludeUnmatchedFilesInFederatedModel" -Value $includeUnmatchedDisplay -Source $includeUnmatchedSettingSource

$nwdNamingMethodSetting = Get-SettingValue -Settings $settingsCache -Names @('NWDNamingMethod')
$nwdNamingMethodSource = if ($nwdNamingMethodSetting) { "Settings" } else { if ($settingsMissing) { "Default (Settings missing)" } else { "Default" } }
$nwdNamingMethodValue = 'Full'
if ($nwdNamingMethodSetting) {
    $normalizedNwdNamingMethod = $nwdNamingMethodSetting.ToString().Trim().ToLowerInvariant()
    if ($normalizedNwdNamingMethod -in @('onlycodes', 'onlydesc', 'codes-desc', 'full')) {
        $nwdNamingMethodValue = $nwdNamingMethodSetting
    }
    else {
        $nwdNamingMethodSource = 'Default (invalid setting)'
    }
}
Write-SettingsLine -Name "NWDNamingMethod" -Value $nwdNamingMethodValue -Source $nwdNamingMethodSource

$runProcessValue = Get-SettingValue -Settings $settingsCache -Names @('RunProcess','ForceIfcProcessing','ForceIfcProcess','ForceProcess')
$runProcessEnabled = $false
$runProcessSource = if ($runProcessValue) { "Settings" } else { if ($settingsMissing) { "Default (Settings missing)" } else { "Default" } }
if ($runProcessValue) {
    $normalizedRun = $runProcessValue.ToString().Trim().ToLowerInvariant()
    if ($normalizedRun -eq 'force') {
        $runProcessEnabled = $true
    }
    elseif ($normalizedRun -in @('no','n','false','0','ignore')) { $runProcessEnabled = $false }
    else { $runProcessEnabled = $true }
}
$runProcessDisplay = if ($runProcessEnabled) { "Yes" } else { "No" }
Write-SettingsLine -Name "RunProcess" -Value $runProcessDisplay -Source $runProcessSource
if (-not $runProcessEnabled) {
    $SkipProcess = $true
}

$runIfcDataExtractionValue = Get-SettingValue -Settings $settingsCache -Names @('RunIfcDataExtraction')
$runIfcDataExtractionEnabled = $false
$runIfcDataExtractionSource = if ($runIfcDataExtractionValue) { "Settings" } else { if ($settingsMissing) { "Default (Settings missing)" } else { "Default" } }
if ($runIfcDataExtractionValue) {
    $normalizedIfcExtractionRun = $runIfcDataExtractionValue.ToString().Trim().ToLowerInvariant()
    if ($normalizedIfcExtractionRun -notin @('no','n','false','0','ignore','')) {
        $runIfcDataExtractionEnabled = $true
    }
}
$runIfcDataExtractionDisplay = if ($runIfcDataExtractionEnabled) { if ($normalizedIfcExtractionRun -eq 'force') { "Force" } else { "Yes" } } else { "No" }
Write-SettingsLine -Name "RunIfcDataExtraction" -Value $runIfcDataExtractionDisplay -Source $runIfcDataExtractionSource

$fedInputSetting = Get-SettingValue -Settings $settingsCache -Names @('FederationInputFolder')
$fedInputSource = if ($fedInputSetting) {
    "Settings"
} else {
    if ($settingsMissing) { "Derived (RunProcess, Settings missing)" } else { "Derived (RunProcess)" }
}
$fedInputValue = if ($fedInputSetting) { $fedInputSetting } else { if ($runProcessEnabled) { $processedDefault } else { $sourceFolderDefault } }
Write-SettingsLine -Name "FederationInputFolder" -Value $fedInputValue -Source $fedInputSource

$navisworksConfigSetting = Get-SettingValue -Settings $settingsCache -Names @('NavisworksConfigXML')
$navisworksConfigSource = if ($navisworksConfigSetting -and -not [string]::IsNullOrWhiteSpace($navisworksConfigSetting.ToString())) {
    "Settings"
} else {
    if ($settingsMissing) { "Default (Settings missing)" } else { "Default" }
}
$navisworksConfig = if ($navisworksConfigSetting -and -not [string]::IsNullOrWhiteSpace($navisworksConfigSetting.ToString())) { $navisworksConfigSetting } else { $navisworksOptionsDefault }
Write-SettingsLine -Name "NavisworksConfigXML" -Value $navisworksConfig -Source $navisworksConfigSource

$navisworksSavedNwdVersionSetting = Get-SettingValue -Settings $settingsCache -Names @('NavisworksSavedNwdVersion')
$navisworksSavedNwdVersionSource = if ($navisworksSavedNwdVersionSetting -and -not [string]::IsNullOrWhiteSpace($navisworksSavedNwdVersionSetting.ToString())) {
    "Settings"
} else {
    if ($settingsMissing) { "Default (Settings missing)" } else { "Default" }
}
$navisworksSavedNwdVersion = if ($navisworksSavedNwdVersionSetting -and -not [string]::IsNullOrWhiteSpace($navisworksSavedNwdVersionSetting.ToString())) { $navisworksSavedNwdVersionSetting } else { "Latest" }
Write-SettingsLine -Name "NavisworksSavedNwdVersion" -Value $navisworksSavedNwdVersion -Source $navisworksSavedNwdVersionSource

$navisworksViewsImportSetting = Get-SettingValue -Settings $settingsCache -Names @('NavisworksViewsImportXML')
if ($navisworksViewsImportSetting -and -not [string]::IsNullOrWhiteSpace($navisworksViewsImportSetting.ToString())) {
    $navisworksViewsImportInfo = Resolve-OptionalXmlSettingPath -Value $navisworksViewsImportSetting -BasePath $basePath
    $navisworksViewsImportSource = 'Settings'
    if ($navisworksViewsImportInfo.Exists) {
        $navisworksViewsImportDisplay = $navisworksViewsImportInfo.CandidatePath
    }
    else {
        $navisworksViewsImportDisplay = "{0} [not found, ignored]" -f $navisworksViewsImportInfo.CandidatePath
    }
}
else {
    $navisworksViewsImportDisplay = '[blank, ignored]'
    $navisworksViewsImportSource = if ($settingsMissing) { 'Default (Settings missing)' } else { 'Blank/ignored' }
}
Write-SettingsLine -Name "NavisworksViewsImportXML" -Value $navisworksViewsImportDisplay -Source $navisworksViewsImportSource

$navisworksVisibleSetting = Get-SettingValue -Settings $settingsCache -Names @('NavisWorksVisible','NavisworksVisible')
$navisworksVisibleSource = if ($navisworksVisibleSetting -and -not [string]::IsNullOrWhiteSpace($navisworksVisibleSetting.ToString())) {
    "Settings"
} else {
    if ($settingsMissing) { "Default (Settings missing)" } else { "Default" }
}
$navisworksVisible = if ($navisworksVisibleSetting -and -not [string]::IsNullOrWhiteSpace($navisworksVisibleSetting.ToString())) { $navisworksVisibleSetting } else { "No" }
Write-SettingsLine -Name "NavisworksVisible" -Value $navisworksVisible -Source $navisworksVisibleSource

$navisworksVersionSetting = Get-SettingValue -Settings $settingsCache -Names @('NavisworksVersion')
$navisworksVersionValue = $navisworksVersionSetting
$navisworksVersionSource = if ($navisworksVersionSetting) { "Settings" } else { if ($settingsMissing) { "Auto (Settings missing)" } else { "Auto" } }
if (-not [string]::IsNullOrWhiteSpace($navisworksVersionValue) -and (Get-Command Resolve-NavisworksInstallPath -ErrorAction SilentlyContinue)) {
    if (-not (Resolve-NavisworksInstallPath -Version $navisworksVersionValue)) {
        if (Get-Command Resolve-NavisworksVersion -ErrorAction SilentlyContinue) {
            $fallbackVersion = Resolve-NavisworksVersion
            if ($fallbackVersion) {
                $navisworksVersionValue = $fallbackVersion
                $navisworksVersionSource = "Fallback (Settings version not installed)"
            }
        }
    }
}
if ([string]::IsNullOrWhiteSpace($navisworksVersionValue) -and (Get-Command Resolve-NavisworksVersion -ErrorAction SilentlyContinue)) {
    $navisworksVersionValue = Resolve-NavisworksVersion
    if ($navisworksVersionValue) {
        $navisworksVersionSource = if ($settingsMissing) { "Auto (Settings missing)" } else { "Auto" }
    }
}
if ([string]::IsNullOrWhiteSpace($navisworksVersionValue)) { $navisworksVersionValue = "Not found" }
Write-SettingsLine -Name "NavisworksVersion" -Value $navisworksVersionValue -Source $navisworksVersionSource -IsMissing:($navisworksVersionValue -eq "Not found")

$runFederationValue = Get-SettingValue -Settings $settingsCache -Names @('RunFederation','FederationRun','RunFederate','Federate','Federation')
$runFederationForcedBySetting = $false
$runFederationDisabledBySetting = $false
if ($runFederationValue) {
    $runFederationNormalized = $runFederationValue.ToString().Trim().ToLowerInvariant()
    if ($runFederationNormalized -eq 'force') { $runFederationForcedBySetting = $true }
    elseif ($runFederationNormalized -in @('no','n','false','0','ignore')) { $runFederationDisabledBySetting = $true }
}
$runFederationSource = if ($runFederationValue) { "Settings" } else { if ($settingsMissing) { "Default (Settings missing)" } else { "Default" } }
$runFederationDisplay = if ($runFederationValue) { $runFederationValue } else { "Yes" }
Write-SettingsLine -Name "RunFederation" -Value $runFederationDisplay -Source $runFederationSource

$reviztoCode = Get-SettingValue -Settings $settingsCache -Names @('ReviztoPublishCode')
$reviztoCodeSource = if ($reviztoCode) { "Settings" } else { if ($settingsMissing) { "Default (Settings missing)" } else { "Default" } }
$reviztoCodeDisplay = if ($reviztoCode) { "[set]" } else { "[missing]" }
Write-SettingsLine -Name "ReviztoPublishCode" -Value $reviztoCodeDisplay -Source $reviztoCodeSource -IsMissing:($reviztoCodeDisplay -eq "[missing]")

$reviztoPublishSetting = Get-SettingValue -Settings $settingsCache -Names @('ReviztoPublish','RunRevizto','ReviztoRun','ReviztoPublish','PublishRevizto','Revizto')
$reviztoPublishForcedBySetting = $false
if ($reviztoPublishSetting) {
    $normalizedReviztoPublish = $reviztoPublishSetting.ToString().Trim().ToLowerInvariant()
    if ($normalizedReviztoPublish -eq 'force') {
        $reviztoPublishForcedBySetting = $true
    }
}
$reviztoPublishSource = if ($reviztoPublishSetting) { "Settings" } else { if ($settingsMissing) { "Default (Settings missing)" } else { "Default" } }
$reviztoPublishDisplay = if ($reviztoPublishSetting) { $reviztoPublishSetting } else { "No" }
Write-SettingsLine -Name "ReviztoPublish" -Value $reviztoPublishDisplay -Source $reviztoPublishSource

$reviztoPublishEnabled = $false
if ($reviztoPublishSetting) {
    $reviztoPublishEnabled = $reviztoPublishSetting.ToString().Trim().ToLowerInvariant() -notin @('no','n','false','0','ignore','')
}
$pipelineStagePlan = [ordered]@{
    Download          = $runDownloadEnabled
    IfcDataExtraction = $runIfcDataExtractionEnabled
    Process           = $runProcessEnabled
    Federation        = -not $runFederationDisabledBySetting
    Revizto           = $reviztoPublishEnabled
}
Write-Host ("FEDAUTO_STAGE_PLAN|Download={0}|IfcDataExtraction={1}|Process={2}|Federation={3}|Revizto={4}" -f $pipelineStagePlan.Download, $pipelineStagePlan.IfcDataExtraction, $pipelineStagePlan.Process, $pipelineStagePlan.Federation, $pipelineStagePlan.Revizto)

$reviztoAgeHours = Get-SettingValue -Settings $settingsCache -Names @('ReviztoMaxAgeHours','ReviztoPublishMaxAgeHours','ReviztoAgeHours','ReviztoMaxAgeHrs','ReviztoPublishAgeHours')
$reviztoAgeMinutes = Get-SettingValue -Settings $settingsCache -Names @('ReviztoMaxAgeMinutes','ReviztoPublishMaxAgeMinutes','ReviztoAgeMinutes','ReviztoMaxAgeMins','ReviztoPublishAgeMinutes')
if ($reviztoAgeHours) {
    Write-SettingsLine -Name "ReviztoMaxAgeHours" -Value $reviztoAgeHours -Source "Settings"
}
elseif ($reviztoAgeMinutes) {
    Write-SettingsLine -Name "ReviztoMaxAgeMinutes" -Value $reviztoAgeMinutes -Source "Settings"
}
else {
    $ageSource = if ($settingsMissing) { "Default (Settings missing)" } else { "Default" }
    Write-SettingsLine -Name "ReviztoMaxAgeMinutes" -Value "60" -Source $ageSource
}

function Resolve-FinalFederatedModelPath {
    param(
        [array]$Settings,
        [string]$RootPath
    )
    $outputFolder = Get-SettingValue -Settings $Settings -Names @('FederationOutputFolder')
    if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
        $outputFolder = CureFolderPath $outputFolder
    }
    if ([string]::IsNullOrWhiteSpace($outputFolder)) { $outputFolder = 'Output' }

    $federatedName = Get-SettingValue -Settings $Settings -Names @('FederatedFileName')
    if ([string]::IsNullOrWhiteSpace($federatedName)) { $federatedName = 'Project Federated' }
    $outputExtension = if ($federatedName.EndsWith('.nwf', [System.StringComparison]::OrdinalIgnoreCase)) { '.nwf' } else { '.nwd' }
    if (Get-Command Resolve-NavisworksOutputBaseName -ErrorAction SilentlyContinue) {
        $federatedName = Resolve-NavisworksOutputBaseName -Name $federatedName
    }
    elseif ($federatedName.EndsWith('.nwd', [System.StringComparison]::OrdinalIgnoreCase) -or $federatedName.EndsWith('.nwf', [System.StringComparison]::OrdinalIgnoreCase)) {
        $federatedName = $federatedName.Substring(0, $federatedName.Length - 4)
    }
    $federatedName = "{0}{1}" -f $federatedName, $outputExtension
    if (-not [System.IO.Path]::IsPathRooted($outputFolder)) {
        $outputFolder = Join-Path $RootPath $outputFolder
    }
    return (Join-Path $outputFolder $federatedName)
}

function Resolve-FEDAUTOSummaryFilePath {
    param(
        [string]$StoredPath,
        [string]$RelativePath,
        [string]$RootFolder
    )

    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($RelativePath) -and -not [string]::IsNullOrWhiteSpace($RootFolder)) {
        $relativeCandidate = if ([System.IO.Path]::IsPathRooted($RelativePath)) {
            if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) { CureFolderPath $RelativePath } else { $RelativePath }
        }
        else {
            Join-Path $RootFolder $RelativePath
        }
        [void]$candidates.Add($relativeCandidate)
    }
    if (-not [string]::IsNullOrWhiteSpace($StoredPath)) {
        $cleanPath = if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) { CureFolderPath $StoredPath } else { $StoredPath }
        if ([System.IO.Path]::IsPathRooted($cleanPath)) {
            [void]$candidates.Add($cleanPath)
            if (-not [string]::IsNullOrWhiteSpace($RootFolder)) {
                [void]$candidates.Add((Join-Path $RootFolder ([System.IO.Path]::GetFileName($cleanPath))))
            }
        }
        elseif (-not [string]::IsNullOrWhiteSpace($RootFolder)) {
            [void]$candidates.Add((Join-Path $RootFolder $cleanPath))
        }
    }

    $seen = @{}
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $key = $candidate.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $(if ($candidates.Count -gt 0) { $candidates[0] } else { $null })
}

function Resolve-ProcessInputs {
    param(
        [array]$Settings,
        [string]$RootPath
    )

    $downloadFolder = Get-SettingValue -Settings $Settings -Names @('SourceFolder')
    if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
        $downloadFolder = CureFolderPath $downloadFolder
    }
    if ([string]::IsNullOrWhiteSpace($downloadFolder)) { $downloadFolder = 'SourceFiles' }
    if (-not [System.IO.Path]::IsPathRooted($downloadFolder)) {
        $downloadFolder = Join-Path $RootPath $downloadFolder
    }

    $attributesFile = Get-SettingValue -Settings $Settings -Names @('AttributesFile')
    if ([string]::IsNullOrWhiteSpace($attributesFile)) { $attributesFile = 'PWAttributes.csv' }
    $attributesPath = if ([System.IO.Path]::IsPathRooted($attributesFile)) { $attributesFile } else { Join-Path $downloadFolder $attributesFile }

    $attributeRows = @()
    $attributesTimestamp = ""
    if (Test-Path $attributesPath -PathType Leaf) {
        try {
            $attributeRows = @(Get-ExcelDataSafe -Path $attributesPath -NamedRange 'Attributes')
        }
        catch {
            Write-Warning "Unable to read attributes CSV from '$attributesPath'. Processing will handle this as a stage error. Error: $_"
            $attributeRows = @()
        }
        try {
            $attributesTimestamp = (Get-Item $attributesPath).LastWriteTimeUtc.ToString("o")
        }
        catch {
            $attributesTimestamp = ""
        }
    }

    return [pscustomobject]@{
        SourceFolder        = $downloadFolder
        AttributesPath      = $attributesPath
        AttributeRows       = $attributeRows
        AttributesTimestamp = $attributesTimestamp
    }
}

function New-FEDAUTOStageSummary {
    param(
        [string]$Name,
        [bool]$Enabled,
        [string]$Reason
    )
    [pscustomobject]@{
        Name     = $Name
        Enabled  = $Enabled
        Status   = $(if ($Enabled) { 'Pending' } else { 'Off' })
        Reason   = $Reason
        Details  = New-Object System.Collections.Generic.List[string]
    }
}

function Set-FEDAUTOStageSummary {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$Status,
        [string]$Reason,
        [string[]]$Details
    )
    if (-not $script:FEDAUTOStageSummaries -or -not $script:FEDAUTOStageSummaries.Contains($Name)) { return }
    $summary = $script:FEDAUTOStageSummaries[$Name]
    if (-not [string]::IsNullOrWhiteSpace($Status)) { $summary.Status = $Status }
    if (-not [string]::IsNullOrWhiteSpace($Reason)) { $summary.Reason = $Reason }
    foreach ($detail in @($Details)) {
        if (-not [string]::IsNullOrWhiteSpace($detail)) { $summary.Details.Add($detail) | Out-Null }
    }
}

function Get-FEDAUTOStageSummaryPath {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('IfcDataExtraction','Process','Federation')][string]$Stage
    )
    switch ($Stage) {
        'IfcDataExtraction' {
            $folder = Get-SettingValue -Settings $settingsCache -Names @('IfcDataExtractionFolder')
            if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) { $folder = CureFolderPath $folder }
            if ([string]::IsNullOrWhiteSpace($folder)) { $folder = 'IFCDataExtraction' }
            if (-not [System.IO.Path]::IsPathRooted($folder)) { $folder = Join-Path $basePath $folder }
            return (Join-Path $folder 'ifc-data-extraction-summary.json')
        }
        'Process' {
            $folder = Get-SettingValue -Settings $settingsCache -Names @('ProcessedFolder')
            if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) { $folder = CureFolderPath $folder }
            if ([string]::IsNullOrWhiteSpace($folder)) { $folder = 'ProcessedIFC' }
            if (-not [System.IO.Path]::IsPathRooted($folder)) { $folder = Join-Path $basePath $folder }
            return (Join-Path $folder 'processed-summary.json')
        }
        'Federation' {
            $folder = Get-SettingValue -Settings $settingsCache -Names @('FederationOutputFolder')
            if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) { $folder = CureFolderPath $folder }
            if ([string]::IsNullOrWhiteSpace($folder)) { $folder = 'Output' }
            if (-not [System.IO.Path]::IsPathRooted($folder)) { $folder = Join-Path $basePath $folder }
            return (Join-Path $folder 'federation-summary.json')
        }
    }
}

function Get-FEDAUTOJsonSummary {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try { return (Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -ErrorAction Stop) }
    catch { return $null }
}

function Complete-FEDAUTOStageFromSummary {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('IfcDataExtraction','Process','Federation')][string]$Stage
    )
    $summaryPath = Get-FEDAUTOStageSummaryPath -Stage $Stage
    $summary = Get-FEDAUTOJsonSummary -Path $summaryPath
    switch ($Stage) {
        'IfcDataExtraction' {
            if ($summary) {
                Set-FEDAUTOStageSummary -Name $Stage -Status 'Successful' -Reason 'IFC data extraction completed.' -Details @(
                    ("Exported: {0}; Skipped: {1}; Failed: {2}; Rule-skipped inputs: {3}" -f $summary.Exported, $summary.Skipped, $summary.Failed, $summary.RuleSkippedInputCount),
                    ("Output folder: {0}" -f $summary.ExportFolderRelativePath)
                )
            }
            else {
                Set-FEDAUTOStageSummary -Name $Stage -Status 'Successful' -Reason 'IFC data extraction completed, but summary JSON was not found.'
            }
        }
        'Process' {
            if ($summary) {
                Set-FEDAUTOStageSummary -Name $Stage -Status 'Successful' -Reason 'IFC processing completed.' -Details @(
                    ("Processed: {0}; Skipped/current: {1}; Failed: {2}" -f $summary.Processed, $summary.Skipped, $summary.Failed),
                    ("Summary file: {0}" -f $summaryPath)
                )
            }
            else {
                Set-FEDAUTOStageSummary -Name $Stage -Status 'Successful' -Reason 'IFC processing completed, but summary JSON was not found.'
            }
        }
        'Federation' {
            if ($summary) {
                if ($summary.GroupingMethod -eq 'Wildcard Selection') {
                    $rules = @($summary.WildcardRules)
                    $createdRules = @($rules | Where-Object { $_.Status -eq 'Created' }).Count
                    $skippedRules = @($rules | Where-Object { $_.Status -like 'Skipped*' }).Count
                    $matchedFiles = 0
                    foreach ($rule in $rules) {
                        if ($rule -and ($rule.PSObject.Properties.Name -contains 'MatchedFiles')) {
                            $matchedFiles += @($rule.MatchedFiles).Count
                        }
                    }
                    Set-FEDAUTOStageSummary -Name $Stage -Status 'Successful' -Reason 'Wildcard federation completed.' -Details @(
                        ("Rules created: {0}; Rules skipped/no matches: {1}; Matched input references: {2}" -f $createdRules, $skippedRules, $matchedFiles),
                        ("Last created model: {0}" -f $summary.LastCreatedModelRelativePath),
                        ("Output folder: {0}" -f $summary.OutputFolderRelativePath)
                    )
                    return
                }
                Set-FEDAUTOStageSummary -Name $Stage -Status 'Successful' -Reason 'Federation completed.' -Details @(
                    ("Grouping: {0}; Matched inputs: {1}; Unmatched inputs: {2}; Final created: {3}" -f $summary.GroupingMethod, $summary.MatchedInputCount, $summary.UnmatchedInputCount, $summary.FinalModelCreated),
                    ("Output folder: {0}" -f $summary.OutputFolderRelativePath)
                )
            }
            else {
                Set-FEDAUTOStageSummary -Name $Stage -Status 'Successful' -Reason 'Federation completed, but summary JSON was not found.'
            }
        }
    }
}

$script:FEDAUTOStageSummaries = [ordered]@{
    Download          = New-FEDAUTOStageSummary -Name 'Download' -Enabled $pipelineStagePlan.Download -Reason $(if ($pipelineStagePlan.Download) { 'Enabled by RunDownload.' } else { 'Off: RunDownload is disabled.' })
    IfcDataExtraction = New-FEDAUTOStageSummary -Name 'IfcDataExtraction' -Enabled $pipelineStagePlan.IfcDataExtraction -Reason $(if ($pipelineStagePlan.IfcDataExtraction) { 'Enabled by RunIfcDataExtraction.' } else { 'Off: RunIfcDataExtraction is disabled.' })
    Process           = New-FEDAUTOStageSummary -Name 'Process' -Enabled $pipelineStagePlan.Process -Reason $(if ($pipelineStagePlan.Process) { 'Enabled by RunProcess.' } else { 'Off: RunProcess is disabled.' })
    Federation        = New-FEDAUTOStageSummary -Name 'Federation' -Enabled $pipelineStagePlan.Federation -Reason $(if ($pipelineStagePlan.Federation) { 'Enabled by RunFederation.' } else { 'Off: RunFederation is disabled.' })
    Revizto           = New-FEDAUTOStageSummary -Name 'Revizto' -Enabled $pipelineStagePlan.Revizto -Reason $(if ($pipelineStagePlan.Revizto) { 'Enabled by ReviztoPublish.' } else { 'Off: ReviztoPublish is disabled.' })
}

try {
    # Determine federation behavior from Settings.
    $forceFederate = $runFederationForcedBySetting
    $downloadSkipForceOnlyMode = $false
    $downloadIntentionallyDisabled = $false
    $downloadSkipReason = $null
    $federationRan = $false
    $finalModelPath = $null
    $settings = $null
    if ($runFederationDisabledBySetting) {
        $SkipFederate = $true
        $skipMessage = "Federation disabled by Settings parameter 'RunFederation'."
        Write-Log -Message $skipMessage -Color 'Yellow' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
        Set-FEDAUTOStageSummary -Name 'Federation' -Status 'Off' -Reason $skipMessage
    }

    if ($SkipDownload -or -not $runDownloadEnabled) {
        $downloadIntentionallyDisabled = $true
        $downloadSkipReason = "source acquisition disabled by Settings parameter 'RunDownload'. Existing files will be used and Download rows are ignored."
        Write-Log -Message ("Download skipped: {0}" -f $downloadSkipReason) -Color 'Yellow' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
        $global:PWDownloadedCount = 0
        $global:PWDeletedCount = 0
        Set-FEDAUTOStageSummary -Name 'Download' -Status 'Skipped' -Reason $downloadSkipReason -Details @('Existing source files were used.')
    }
    else {
        # Pull latest IFCs and metadata from ProjectWise/ACC based on config.
        $pwUserValue = $null
        $pwPassValue = $null
        if ($settingsCache -and $settingsCache.Count -gt 0) {
            $pwUserValue = $settingsCache | Where-Object { $_.Parameter -eq "PWUser" } | Select-Object -ExpandProperty Value
            $pwPassValue = $settingsCache | Where-Object { $_.Parameter -eq "PWPass" } | Select-Object -ExpandProperty Value
            if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
                $pwUserValue = Normalize-ExcelText -Text $pwUserValue
                $pwPassValue = Normalize-ExcelText -Text $pwPassValue
            }
            else {
                if ($null -ne $pwUserValue) { $pwUserValue = $pwUserValue.ToString().Trim() }
                if ($null -ne $pwPassValue) { $pwPassValue = $pwPassValue.ToString().Trim() }
            }
        }
        try {
            $global:PWDownloadStatus = 'NotRun'
            $stepResult = Invoke-Step -Name "Download" -Action {
                Invoke-PWDownload -InputExcelFile $ConfigFile -Settings $settingsCache -DownloadRows $downloadRangeCache -PWAttributeListRows $pwAttributesListCache -LogsFolder $LogsFolder -PWUser $pwUserValue -PWPass $pwPassValue
            }
            if ($stepResult) { $stepResults.Add($stepResult) | Out-Null }
            $downloadDetails = @(
                ("Download status: {0}" -f $global:PWDownloadStatus),
                ("Downloaded/copied: {0}; Deleted/relocated: {1}" -f $global:PWDownloadedCount, $global:PWDeletedCount)
            )
            Set-FEDAUTOStageSummary -Name 'Download' -Status 'Successful' -Reason 'Source acquisition completed.' -Details $downloadDetails
        }
        catch {
            $downloadError = $_.ToString()
            if ($downloadError -match 'ProjectWise login failed|did not create an active session|session dropped or was not established|Not logged into a ProjectWise datasource') {
                $downloadSkipForceOnlyMode = $true
                $downloadSkipReason = "ProjectWise login/session failed during download."
                Write-Log -Message ("ERROR: {0} {1}" -f $downloadSkipReason, $downloadError) -Color 'Red' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
                $global:PWDownloadedCount = 0
                $global:PWDeletedCount = 0
                Set-FEDAUTOStageSummary -Name 'Download' -Status 'Failed' -Reason $downloadSkipReason -Details @($downloadError)
            }
            else {
                Set-FEDAUTOStageSummary -Name 'Download' -Status 'Failed' -Reason 'Source acquisition failed.' -Details @($downloadError)
                throw
            }
        }
        if (-not $downloadSkipForceOnlyMode -and $global:PWDownloadStatus -in @('Empty','AllSearchFailed')) {
            $downloadSkipForceOnlyMode = $true
            $downloadSkipReason = if ($global:PWDownloadStatus -eq 'AllSearchFailed') { "All read-folder searches failed." } else { "Download returned zero files." }
            Write-Log -Message ("ERROR: {0} Treating as skipped download." -f $downloadSkipReason) -Color 'Red' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
            $global:PWDownloadedCount = 0
            $global:PWDeletedCount = 0
            Set-FEDAUTOStageSummary -Name 'Download' -Status 'Failed' -Reason $downloadSkipReason
        }
    }
    if ($runIfcDataExtractionEnabled) {
        if ($downloadSkipForceOnlyMode) {
            Write-Log -Message ("IFC data extraction is running using existing source files because download is unavailable ({0})." -f $downloadSkipReason) -Color 'Yellow' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
        }
        $stepResult = Invoke-Step -Name "IFC Data Extraction" -Action {
            Invoke-IfcDataExtraction -Settings $settingsCache -IfcDataExtractionRules $ifcDataExtractionRulesCache -LogsFolder $LogsFolder
        }
        if ($stepResult) { $stepResults.Add($stepResult) | Out-Null }
        Complete-FEDAUTOStageFromSummary -Stage 'IfcDataExtraction'
    }
    $processedCount = $null
    if (-not $SkipProcess) {
        # Decide whether processing is forced or disabled by Settings.
        if ($settingsCache -and $settingsCache.Count -gt 0) {
            $settingsForProcess = $settingsCache
            # RunProcess values: Yes=process if new/updated files, No=skip, Force=always process.
            $runProcessKeys = @('runprocess','forceifcprocessing','forceifcprocess','forceprocess')
            $runProcessMatches = $settingsForProcess | Where-Object {
                $paramRaw = $_.Parameter
                $paramText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
                    Normalize-ExcelText -Text $paramRaw
                }
                else {
                    if ($null -ne $paramRaw) { $paramRaw.ToString().Trim() } else { $null }
                }
                if ([string]::IsNullOrWhiteSpace($paramText)) { return $false }
                $runProcessKeys -contains $paramText.ToLowerInvariant()
            }
            $runProcessRow = $null
            foreach ($candidate in $runProcessMatches) {
                $valueText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
                    Normalize-ExcelText -Text $candidate.Value
                }
                else {
                    if ($null -ne $candidate.Value) { $candidate.Value.ToString().Trim() } else { $null }
                }
                if (-not [string]::IsNullOrWhiteSpace($valueText)) {
                    $runProcessRow = $candidate
                    break
                }
            }
            if (-not $runProcessRow -and $runProcessMatches) {
                $runProcessRow = $runProcessMatches | Select-Object -First 1
            }
            if (-not $runProcessRow -or $null -eq $runProcessRow.Value -or [string]::IsNullOrWhiteSpace($runProcessRow.Value.ToString())) {
                $SkipProcess = $true
                $message = "Processing disabled by default (RunProcess missing)."
                Set-FEDAUTOStageSummary -Name 'Process' -Status 'Off' -Reason $message
                if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                    Write-Log -Message $message -Color 'Yellow' -Settings $settingsForProcess -LogsFolderOverride $LogsFolder -BasePath $basePath
                }
                else {
                    Write-Host $message -ForegroundColor Yellow
                }
            }
            else {
                $valueText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
                    Normalize-ExcelText -Text $runProcessRow.Value
                }
                else {
                    $runProcessRow.Value.ToString().Trim()
                }
                $normalized = $valueText.ToLowerInvariant()
                if ($normalized -eq 'force') {
                    $ForceProcess = $true
                    $SkipProcess = $false
                    $message = ("Processing forced by Settings parameter '{0}'." -f $runProcessRow.Parameter)
                    if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                        Write-Log -Message $message -Color 'Yellow' -Settings $settingsForProcess -LogsFolderOverride $LogsFolder -BasePath $basePath
                    }
                    else {
                        Write-Host $message -ForegroundColor Yellow
                    }
                }
                elseif ($normalized -in @('no','n','false','0','ignore')) {
                    $SkipProcess = $true
                    $message = ("Processing disabled by Settings parameter '{0}'." -f $runProcessRow.Parameter)
                    Set-FEDAUTOStageSummary -Name 'Process' -Status 'Off' -Reason $message
                    if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
                        Write-Log -Message $message -Color 'Yellow' -Settings $settingsForProcess -LogsFolderOverride $LogsFolder -BasePath $basePath
                    }
                    else {
                        Write-Host $message -ForegroundColor Yellow
                    }
                }
                else {
                    $SkipProcess = $false
                }
            }
        }
        if (-not $SkipProcess) {
            if ($downloadSkipForceOnlyMode -and -not $ForceProcess) {
                $SkipProcess = $true
                $processSkipMessage = ("Processing skipped because download is unavailable ({0}). RunProcess must be Force to continue." -f $downloadSkipReason)
                Write-Log -Message $processSkipMessage -Color 'Yellow' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
                Set-FEDAUTOStageSummary -Name 'Process' -Status 'Skipped' -Reason $processSkipMessage
            }
            elseif ($downloadSkipForceOnlyMode -and $ForceProcess) {
                Write-Log -Message ("Processing is running in Force mode while download is unavailable ({0})." -f $downloadSkipReason) -Color 'Yellow' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
            }
        }
        if (-not $SkipProcess) {
            $processInputInfo = Resolve-ProcessInputs -Settings $settingsCache -RootPath $basePath
            $processAttributesPath = $processInputInfo.AttributesPath
            $processAttributesTimestamp = $processInputInfo.AttributesTimestamp
            $attributeRowsForProcess = @($processInputInfo.AttributeRows)
            # Inject IFC attributes and capture the processed count for federation decisions.
            $stepResult = Invoke-Step -Name "Process IFC Attributes" -Action {
                Invoke-ProcessIfcAttributes -ConfigFile $ConfigFile -Settings $settingsCache -AttributeDefinitionRows $pwAttributesListCache -FederationRows $federationRangeCache -LookupRows $lookupsRangeCache -AttributeRows $attributeRowsForProcess -AttributesWorkbookPath $processAttributesPath -AttributesWorkbookTimestamp $processAttributesTimestamp -Force:$ForceProcess -LogsFolder $LogsFolder
            }
            if ($stepResult) { $stepResults.Add($stepResult) | Out-Null }
            Complete-FEDAUTOStageFromSummary -Stage 'Process'
            $summaryPath = $null
            $settings = $settingsCache
            $processedFolderValue = $settings | Where-Object { $_.Parameter -eq "ProcessedFolder" } | Select-Object -ExpandProperty Value
            if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
                $processedFolderValue = CureFolderPath $processedFolderValue
            }
            if (-not $processedFolderValue) { $processedFolderValue = "ProcessedIFC" }
            if ($processedFolderValue) {
                $summaryPath = Join-Path $processedFolderValue "processed-summary.json"
            }
            if (-not $summaryPath) {
                $summaryPath = Join-Path $basePath "processed-summary.json"
            }
            if (Test-Path $summaryPath -PathType Leaf) {
                try {
                    $summary = Get-Content -Path $summaryPath -Raw | ConvertFrom-Json
                    $processedCount = $summary.Processed
                }
                catch {
                    Write-Warning "Unable to read processing summary at '$summaryPath'. Proceeding with federation."
                    $processedCount = $null
                }
            }
        }
    }
    if ($downloadSkipForceOnlyMode -and -not $forceFederate) {
        $SkipFederate = $true
        $federationSkipMessage = ("Federation skipped because download is unavailable ({0}). RunFederation must be Force to continue." -f $downloadSkipReason)
        Write-Log -Message $federationSkipMessage -Color 'Yellow' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
        Set-FEDAUTOStageSummary -Name 'Federation' -Status 'Skipped' -Reason $federationSkipMessage
    }
    elseif ($downloadSkipForceOnlyMode -and $forceFederate) {
        Write-Log -Message ("Federation is running in Force mode while download is unavailable ({0})." -f $downloadSkipReason) -Color 'Yellow' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
    }
    if (-not $SkipFederate) {
        # Evaluate existing federated model state so we can skip or force federation intelligently.
        $finalModelMissing = $false
        $finalModelPath = $null
        $preModelTimestamp = $null
        $settings = $settingsCache
        $groupingMethodValue = Get-SettingValue -Settings $settings -Names @('FederationGroupingMethod')
        $isWildcardSelection = $groupingMethodValue -and $groupingMethodValue.ToString().Trim().ToLowerInvariant() -eq 'wildcard selection'
        $outputFolderValue = $settings | Where-Object { $_.Parameter -eq "FederationOutputFolder" } | Select-Object -ExpandProperty Value
        if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
            $outputFolderValue = CureFolderPath $outputFolderValue
        }
        if (-not $outputFolderValue) { $outputFolderValue = "Output" }
        if ($outputFolderValue) {
            $federatedFileName = $settings | Where-Object { $_.Parameter -eq "FederatedFileName" } | Select-Object -ExpandProperty Value
            if ([string]::IsNullOrWhiteSpace($federatedFileName)) { $federatedFileName = 'Project Federated' }
            $finalExtension = if ($federatedFileName.EndsWith('.nwf', [System.StringComparison]::OrdinalIgnoreCase)) { '.nwf' } else { '.nwd' }
            if (Get-Command Resolve-NavisworksOutputBaseName -ErrorAction SilentlyContinue) {
                $federatedFileName = Resolve-NavisworksOutputBaseName -Name $federatedFileName
            }
            elseif ($federatedFileName.EndsWith('.nwd', [System.StringComparison]::OrdinalIgnoreCase) -or $federatedFileName.EndsWith('.nwf', [System.StringComparison]::OrdinalIgnoreCase)) {
                $federatedFileName = $federatedFileName.Substring(0, $federatedFileName.Length - 4)
            }
            $federatedFileName = "{0}{1}" -f $federatedFileName, $finalExtension
            $finalModelPath = Join-Path $outputFolderValue $federatedFileName
            if (-not (Test-Path $finalModelPath -PathType Leaf)) {
                $finalModelMissing = $true
            }
            else {
                $preModelTimestamp = (Get-Item $finalModelPath).LastWriteTimeUtc
            }
        }
        if ($isWildcardSelection -and $outputFolderValue) {
            $wildcardSummaryPath = Join-Path $outputFolderValue 'federation-summary.json'
            $finalModelPath = $null
            $preModelTimestamp = $null
            $finalModelMissing = $true
            if (Test-Path $wildcardSummaryPath -PathType Leaf) {
                try {
                    $wildcardSummary = Get-Content -Path $wildcardSummaryPath -Raw | ConvertFrom-Json
                    $lastWildcardModel = Resolve-FEDAUTOSummaryFilePath -StoredPath $wildcardSummary.LastCreatedModelPath -RelativePath $wildcardSummary.LastCreatedModelRelativePath -RootFolder $outputFolderValue
                    if ($wildcardSummary.GroupingMethod -eq 'Wildcard Selection' -and $lastWildcardModel -and (Test-Path -LiteralPath $lastWildcardModel -PathType Leaf)) {
                        $finalModelPath = $lastWildcardModel
                        $preModelTimestamp = (Get-Item $finalModelPath).LastWriteTimeUtc
                        $finalModelMissing = $false
                    }
                }
                catch { Write-Warning "Unable to read wildcard federation summary at '$wildcardSummaryPath'. Error: $_" }
            }
        }
        $includeUnmatchedSettingText = Get-SettingValue -Settings $settings -Names @('IncludeUnmatchedFilesInFederatedModel')
        $includeUnmatchedInFinal = $false
        if ($includeUnmatchedSettingText) {
            $normalizedIncludeText = $includeUnmatchedSettingText.ToString().Trim().ToLowerInvariant()
            if ($normalizedIncludeText -notin @('no','n','false','0','ignore')) {
                $includeUnmatchedInFinal = $true
            }
        }
        if ($finalModelMissing -and -not $includeUnmatchedInFinal -and $outputFolderValue) {
            $federationSummaryPath = Join-Path $outputFolderValue 'federation-summary.json'
            if (Test-Path $federationSummaryPath -PathType Leaf) {
                try {
                    $federationSummary = Get-Content -Path $federationSummaryPath -Raw | ConvertFrom-Json
                    $summaryMatchedCount = 0
                    if ($null -ne $federationSummary.MatchedInputCount) {
                        [void][int]::TryParse($federationSummary.MatchedInputCount.ToString(), [ref]$summaryMatchedCount)
                    }
                    $summaryUnmatchedCount = 0
                    if ($null -ne $federationSummary.UnmatchedInputCount) {
                        [void][int]::TryParse($federationSummary.UnmatchedInputCount.ToString(), [ref]$summaryUnmatchedCount)
                    }
                    $summaryFinalCreated = $false
                    if ($null -ne $federationSummary.FinalModelCreated) {
                        $summaryFinalCreated = $federationSummary.FinalModelCreated.ToString().Trim().ToLowerInvariant() -in @('true','1')
                    }
                    $summaryIncludeUnmatched = $false
                    if ($null -ne $federationSummary.IncludeUnmatchedFilesInFederatedModel) {
                        $summaryIncludeUnmatched = $federationSummary.IncludeUnmatchedFilesInFederatedModel.ToString().Trim().ToLowerInvariant() -in @('true','1','yes','y')
                    }
                    $summaryUnmatchedModelPath = Resolve-FEDAUTOSummaryFilePath -StoredPath $federationSummary.UnmatchedModelPath -RelativePath $federationSummary.UnmatchedModelRelativePath -RootFolder $outputFolderValue
                    if (-not $summaryIncludeUnmatched -and -not $summaryFinalCreated -and $summaryMatchedCount -le 0 -and $summaryUnmatchedCount -gt 0 -and
                        -not [string]::IsNullOrWhiteSpace($summaryUnmatchedModelPath) -and (Test-Path -LiteralPath $summaryUnmatchedModelPath -PathType Leaf)) {
                        $finalModelMissing = $false
                    }
                }
                catch {
                    Write-Warning "Unable to read federation summary at '$federationSummaryPath'. Error: $_"
                }
            }
        }
        $deletedCountForRun = 0
        if (Get-Variable -Name PWDeletedCount -Scope Global -ErrorAction SilentlyContinue) {
            $deletedCountForRun = [int]$global:PWDeletedCount
        }
        $downloadedCountForRun = 0
        if (Get-Variable -Name PWDownloadedCount -Scope Global -ErrorAction SilentlyContinue) {
            $downloadedCountForRun = [int]$global:PWDownloadedCount
        }

        function Normalize-PathValue {
            param([string]$PathValue, [string]$Root)
            if ([string]::IsNullOrWhiteSpace($PathValue)) { return $null }
            $normalized = ($PathValue -replace '/', '\').Trim()
            if (-not [System.IO.Path]::IsPathRooted($normalized)) {
                $normalized = Join-Path $Root $normalized
            }
            try {
                $normalized = [System.IO.Path]::GetFullPath($normalized)
            }
            catch {
                # ignore
            }
            return $normalized.TrimEnd('\')
        }

        $downloadFolderValue = $settings | Where-Object { $_.Parameter -eq "SourceFolder" } | Select-Object -ExpandProperty Value
        if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
            $downloadFolderValue = CureFolderPath $downloadFolderValue
        }
        if (-not $downloadFolderValue) { $downloadFolderValue = "SourceFiles" }

        $processedFolderValueForFed = $settings | Where-Object { $_.Parameter -eq "ProcessedFolder" } | Select-Object -ExpandProperty Value
        if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
            $processedFolderValueForFed = CureFolderPath $processedFolderValueForFed
        }
        if (-not $processedFolderValueForFed) { $processedFolderValueForFed = "ProcessedIFC" }

        $federationInputValue = $settings | Where-Object { $_.Parameter -eq "FederationInputFolder" } | Select-Object -ExpandProperty Value
        if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
            $federationInputValue = CureFolderPath $federationInputValue
        }
        if (-not $federationInputValue) {
            $federationInputValue = if ($SkipProcess) { $downloadFolderValue } else { $processedFolderValueForFed }
        }

        $normalizedDownload = Normalize-PathValue -PathValue $downloadFolderValue -Root $basePath
        $normalizedFederationInput = Normalize-PathValue -PathValue $federationInputValue -Root $basePath
        $useDownloadInput = $SkipProcess -and $normalizedDownload -and $normalizedFederationInput -and
            ($normalizedDownload.Equals($normalizedFederationInput, [System.StringComparison]::OrdinalIgnoreCase))

        $changeCount = if ($useDownloadInput) { $downloadedCountForRun } else { $processedCount }
        $changeLabel = if ($useDownloadInput) { "downloaded" } else { "processed" }
        $federationRan = $false
        if (-not $forceFederate -and $null -ne $changeCount -and [int]$changeCount -le 0 -and -not $finalModelMissing -and $deletedCountForRun -le 0) {
            $federatedDisplay = if ($finalModelPath) {
                Split-Path -Leaf $finalModelPath
            }
            elseif ($federatedFileName) {
                $federatedFileName
            }
            else {
                "Federated model"
            }
            Write-Host $federatedDisplay -NoNewline -ForegroundColor Green
            Write-Host (" Skipping federation >>>> no federation source files were {0}" -f $changeLabel) -ForegroundColor Red
            Set-FEDAUTOStageSummary -Name 'Federation' -Status 'Skipped' -Reason ("No federation source files were {0}, no deletions were detected, and the final model already exists." -f $changeLabel) -Details @("Final model: $finalModelPath")
        }
        else {
            $reasons = @()
            if ($forceFederate) { $reasons += "forced by RunFederation" }
            if ($deletedCountForRun -gt 0) { $reasons += ("deleted source files detected ({0})" -f $deletedCountForRun) }
            if ($finalModelMissing) { $reasons += "final federated model missing" }
            if ($null -ne $changeCount -and [int]$changeCount -gt 0) { $reasons += ("{0} federation source files: {1}" -f $changeLabel, [int]$changeCount) }
            if ($reasons.Count -gt 0) {
                Write-Host ("Federation triggered: {0}" -f ($reasons -join "; ")) -ForegroundColor Yellow
            }
            $stepResult = Invoke-Step -Name "Group/Federate Files" -Action {
                Invoke-GroupFilesForFedProcess -ConfigFile $ConfigFile -Settings $settingsCache -FederationRows $federationRangeCache -WildcardSelectionRows $wildcardSelectionCache -LookupRows $lookupsRangeCache -LogsFolder $LogsFolder -UseFederationInputFolder:$SkipProcess
            }
            if ($stepResult) { $stepResults.Add($stepResult) | Out-Null }
            $federationRan = $true
            Complete-FEDAUTOStageFromSummary -Stage 'Federation'
        }

        if ($isWildcardSelection -and $federationRan -and $outputFolderValue) {
            $wildcardSummaryPath = Join-Path $outputFolderValue 'federation-summary.json'
            if (Test-Path $wildcardSummaryPath -PathType Leaf) {
                try {
                    $wildcardSummary = Get-Content -Path $wildcardSummaryPath -Raw | ConvertFrom-Json
                    $lastWildcardModel = Resolve-FEDAUTOSummaryFilePath -StoredPath $wildcardSummary.LastCreatedModelPath -RelativePath $wildcardSummary.LastCreatedModelRelativePath -RootFolder $outputFolderValue
                    if ($lastWildcardModel -and (Test-Path -LiteralPath $lastWildcardModel -PathType Leaf)) {
                        $finalModelPath = $lastWildcardModel
                    }
                }
                catch { Write-Warning "Unable to read wildcard federation summary at '$wildcardSummaryPath'. Error: $_" }
            }
        }

        # Only publish to Revizto when a new federated model is produced.
        $newFederatedModel = $false
        if ($federationRan -and $finalModelPath -and (Test-Path $finalModelPath -PathType Leaf)) {
            $postModelTimestamp = (Get-Item $finalModelPath).LastWriteTimeUtc
            if ($null -eq $preModelTimestamp -or $postModelTimestamp -gt $preModelTimestamp) {
                $newFederatedModel = $true
            }
        }
        if ($federationRan) {
            $preText = if ($preModelTimestamp) { $preModelTimestamp.ToString("o") } else { "null" }
            $postText = if ($postModelTimestamp) { $postModelTimestamp.ToString("o") } else { "null" }
            Write-Host ("Revizto check: final model path='{0}', pre='{1}', post='{2}'" -f $finalModelPath, $preText, $postText) -ForegroundColor Yellow
        }

        if ($federationRan) {
            $reviztoAllowed = $false
            $reviztoForced = $false
            $reviztoSettingName = "ReviztoPublish (default)"
            if ($settings) {
                $reviztoSettingRow = $settings | Where-Object {
                    $paramRaw = $_.Parameter
                    $paramText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
                        Normalize-ExcelText -Text $paramRaw
                    }
                    else {
                        if ($null -ne $paramRaw) { $paramRaw.ToString().Trim() } else { $null }
                    }
                    if ([string]::IsNullOrWhiteSpace($paramText)) { return $false }
                    $paramText.ToLowerInvariant() -in @('runrevizto','reviztorun','reviztopublish','publishrevizto','revizto')
                } | Select-Object -First 1
                if ($reviztoSettingRow -and $null -ne $reviztoSettingRow.Value) {
                    $valueText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
                        Normalize-ExcelText -Text $reviztoSettingRow.Value
                    }
                    else {
                        $reviztoSettingRow.Value.ToString().Trim()
                    }
                    if (-not [string]::IsNullOrWhiteSpace($valueText)) {
                        $normalized = $valueText.ToLowerInvariant()
                        if ($normalized -eq 'force') {
                            $reviztoAllowed = $true
                            $reviztoForced = $true
                            $reviztoSettingName = $reviztoSettingRow.Parameter
                        }
                        elseif ($normalized -in @('no','n','false','0','ignore')) {
                            $reviztoAllowed = $false
                            $reviztoSettingName = $reviztoSettingRow.Parameter
                        }
                        else {
                            $reviztoAllowed = $true
                            $reviztoSettingName = $reviztoSettingRow.Parameter
                        }
                    }
                }
            }

            $reviztoMaxAgeMinutes = 60.0
            $ageMinutesRaw = Get-SettingValue -Settings $settings -Names @(
                'ReviztoMaxAgeMinutes','ReviztoPublishMaxAgeMinutes','ReviztoAgeMinutes','ReviztoMaxAgeMins','ReviztoPublishAgeMinutes'
            )
            $ageHoursRaw = Get-SettingValue -Settings $settings -Names @(
                'ReviztoMaxAgeHours','ReviztoPublishMaxAgeHours','ReviztoAgeHours','ReviztoMaxAgeHrs','ReviztoPublishAgeHours'
            )
            $parsedAge = $null
            if ($null -ne $ageHoursRaw -and -not [string]::IsNullOrWhiteSpace($ageHoursRaw.ToString())) {
                $hours = 0.0
                if ([double]::TryParse($ageHoursRaw.ToString(), [ref]$hours)) {
                    $parsedAge = $hours * 60.0
                }
                else {
                    Write-Log -Message "Revizto publish warning: invalid Revizto max age hours value '$ageHoursRaw'. Using default $reviztoMaxAgeMinutes minutes." -Color 'Yellow' -Settings $settings -LogsFolderOverride $LogsFolder -BasePath $basePath
                }
            }
            elseif ($null -ne $ageMinutesRaw -and -not [string]::IsNullOrWhiteSpace($ageMinutesRaw.ToString())) {
                $minutes = 0.0
                if ([double]::TryParse($ageMinutesRaw.ToString(), [ref]$minutes)) {
                    $parsedAge = $minutes
                }
                else {
                    Write-Log -Message "Revizto publish warning: invalid Revizto max age minutes value '$ageMinutesRaw'. Using default $reviztoMaxAgeMinutes minutes." -Color 'Yellow' -Settings $settings -LogsFolderOverride $LogsFolder -BasePath $basePath
                }
            }
            if ($null -ne $parsedAge) {
                if ($parsedAge -gt 0) {
                    $reviztoMaxAgeMinutes = $parsedAge
                }
                else {
                    Write-Log -Message "Revizto publish warning: Revizto max age must be > 0. Using default $reviztoMaxAgeMinutes minutes." -Color 'Yellow' -Settings $settings -LogsFolderOverride $LogsFolder -BasePath $basePath
                }
            }

            $publishBlockers = New-Object System.Collections.Generic.List[string]
            if (-not $reviztoAllowed) {
                $settingLabel = if ($reviztoSettingName) { $reviztoSettingName } else { 'Revizto publish setting' }
                $publishBlockers.Add(("Revizto publish skipped: disabled by Settings parameter '{0}'." -f $settingLabel)) | Out-Null
            }
            if (-not $finalModelPath -or -not (Test-Path $finalModelPath -PathType Leaf)) {
                $publishBlockers.Add(("Revizto publish skipped: federated model not found at '{0}'." -f $finalModelPath)) | Out-Null
            }
            else {
                $modelTimestamp = if ($postModelTimestamp) { $postModelTimestamp } else { (Get-Item $finalModelPath).LastWriteTimeUtc }
                $ageMinutes = (New-TimeSpan -Start $modelTimestamp -End (Get-Date).ToUniversalTime()).TotalMinutes
                if ($ageMinutes -gt $reviztoMaxAgeMinutes) {
                    $publishBlockers.Add(("Revizto publish skipped: federated model is {0:N1} minutes old; max allowed is {1:N1} minutes." -f $ageMinutes, $reviztoMaxAgeMinutes)) | Out-Null
                }
            }

            if ($publishBlockers.Count -gt 0) {
                foreach ($warn in $publishBlockers) {
                    Write-Log -Message $warn -Color 'Yellow' -Settings $settings -LogsFolderOverride $LogsFolder -BasePath $basePath
                }
                Set-FEDAUTOStageSummary -Name 'Revizto' -Status 'Skipped' -Reason 'Revizto publish was blocked.' -Details @($publishBlockers.ToArray())
            }
            elseif ($newFederatedModel -or $reviztoForced) {
                $stepResult = Invoke-Step -Name "Publish Revizto" -Action { Invoke-PublishRevizto -ConfigFile $ConfigFile -Settings $settingsCache -LogsFolder $LogsFolder }
                if ($stepResult) { $stepResults.Add($stepResult) | Out-Null }
                Set-FEDAUTOStageSummary -Name 'Revizto' -Status 'Successful' -Reason 'Revizto publish step completed.' -Details @("Model: $finalModelPath")
            }
            else {
                Write-Log -Message "Revizto publish skipped: no new federated model was created." -Color 'Yellow' -Settings $settings -LogsFolderOverride $LogsFolder -BasePath $basePath
                Set-FEDAUTOStageSummary -Name 'Revizto' -Status 'Skipped' -Reason 'No new federated model was created.'
            }
        }
    }
    if (-not $federationRan -and $reviztoPublishForcedBySetting) {
        $forcedModelPath = if ($finalModelPath) { $finalModelPath } else { Resolve-FinalFederatedModelPath -Settings $settingsCache -RootPath $basePath }
        if (-not [string]::IsNullOrWhiteSpace($forcedModelPath) -and -not [System.IO.Path]::IsPathRooted($forcedModelPath)) {
            $forcedModelPath = Join-Path $basePath $forcedModelPath
        }
        if ($forcedModelPath -and (Test-Path $forcedModelPath -PathType Leaf)) {
            Write-Log -Message ("Revizto publish forced by Settings. Using existing federated model: '{0}'." -f $forcedModelPath) -Color 'Yellow' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
            $stepResult = Invoke-Step -Name "Publish Revizto" -Action { Invoke-PublishRevizto -ConfigFile $ConfigFile -Settings $settingsCache -LogsFolder $LogsFolder }
            if ($stepResult) { $stepResults.Add($stepResult) | Out-Null }
            Set-FEDAUTOStageSummary -Name 'Revizto' -Status 'Successful' -Reason 'Revizto publish was forced and the publish step completed.' -Details @("Model: $forcedModelPath")
        }
        else {
            Write-Log -Message ("Revizto publish force requested, but federated model was not found at '{0}'. Skipping publish." -f $forcedModelPath) -Color 'Yellow' -Settings $settingsCache -LogsFolderOverride $LogsFolder -BasePath $basePath
            Set-FEDAUTOStageSummary -Name 'Revizto' -Status 'Skipped' -Reason ("Revizto publish force requested, but federated model was not found at '{0}'." -f $forcedModelPath)
        }
    }
    Write-Host ""
    Write-Host "Pipeline finished." -ForegroundColor Green
}
catch {
    $pipelineError = if ($_.Exception -and -not [string]::IsNullOrWhiteSpace($_.Exception.Message)) {
        $_.Exception.Message
    }
    else {
        $_.ToString()
    }
    $errorMessage = "ERROR: Pipeline halted: $pipelineError"
    $script:FEDAUTOFatalError = $pipelineError
    if ($pipelineError -match "Step '([^']+)' failed") {
        $failedStepName = $Matches[1]
        $failedStageName = switch ($failedStepName) {
            'Download' { 'Download' }
            'IFC Data Extraction' { 'IfcDataExtraction' }
            'Process IFC Attributes' { 'Process' }
            'Group/Federate Files' { 'Federation' }
            'Publish Revizto' { 'Revizto' }
            default { $null }
        }
        if ($failedStageName) {
            Set-FEDAUTOStageSummary -Name $failedStageName -Status 'Failed' -Reason $pipelineError
        }
    }
    Write-Host $errorMessage -ForegroundColor Red
    Write-Error $errorMessage
    throw $errorMessage
}
finally {
    $pipelineTimer.Stop()
    Write-Host ""
    Write-Host "=== Run Summary ===" -ForegroundColor Cyan
    $stageOrder = @('Download','IfcDataExtraction','Process','Federation','Revizto')
    foreach ($stageName in $stageOrder) {
        if (-not $script:FEDAUTOStageSummaries.Contains($stageName)) { continue }
        $summary = $script:FEDAUTOStageSummaries[$stageName]
        if ($summary.Enabled -and $summary.Status -eq 'Pending') {
            if ($script:FEDAUTOFatalError) {
                $summary.Status = 'Not completed'
                $summary.Reason = 'Pipeline halted before this stage completed.'
            }
            else {
                $summary.Status = 'Not run'
                if ([string]::IsNullOrWhiteSpace($summary.Reason)) { $summary.Reason = 'No run condition triggered this enabled stage.' }
            }
        }
        $enabledText = if ($summary.Enabled) { 'On' } else { 'Off' }
        $reasonText = if ([string]::IsNullOrWhiteSpace($summary.Reason)) { '' } else { " - $($summary.Reason)" }
        Write-Host ("{0}: {1}; {2}{3}" -f $summary.Name, $enabledText, $summary.Status, $reasonText)
        foreach ($detail in @($summary.Details | Select-Object -First 6)) {
            if (-not [string]::IsNullOrWhiteSpace($detail)) {
                Write-Host ("  - {0}" -f $detail)
            }
        }
        if ($summary.Details.Count -gt 6) {
            Write-Host ("  - ... {0} more detail(s)" -f ($summary.Details.Count - 6))
        }
    }
    Write-Host ""
    Write-Host "=== Pipeline Totals ===" -ForegroundColor Cyan
    foreach ($result in $stepResults) {
        if (-not $result) { continue }
        Write-Host ("{0} duration: {1}" -f $result.Name, (Format-Duration -Duration $result.Duration))
    }
    Write-Host ("Grand total: {0}" -f (Format-Duration -Duration $pipelineTimer.Elapsed)) -ForegroundColor Green
    Write-Host ("Executable build version: {0}" -f (Format-FEDAUTOExecutableBuildInfo $script:FEDAUTOCurrentBuildInfo))
    # Ensure any active transcripts are closed (including unexpected nested starts).
    $stoppedCount = 0
    while ($true) {
        try {
            Stop-Transcript -ErrorAction Stop | Out-Null
            $stoppedCount++
        }
        catch {
            break
        }
    }
    $global:MainTranscriptActive = $false
    $global:MainLogInfo = $null
    if ($locationPushed) {
        Pop-Location -ErrorAction SilentlyContinue | Out-Null
    }
}



