<#
Shared utilities for Federation Automation.
- Excel helpers, path normalization, logging, and support utilities.
#>

function Ensure-ModuleAvailable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$InstallName = $Name
    )
    if (Get-Module -ListAvailable -Name $Name) {
        Import-Module -Name $Name -ErrorAction Stop
        return
    }
    throw "Required PowerShell module '$Name' is not installed. The JSON/CSV build does not install modules or change execution policy at runtime."
}
function Ensure-FEDAUTOProcessExecutionPolicy {
<#
JSON/CSV build: intentionally does not change PowerShell execution policy.
#>
    [CmdletBinding()]
    param()
    return
}
function Resolve-RelativePath {
<#
Resolves a relative path against a root path; returns rooted paths unchanged.
#>
    [CmdletBinding()]
    param(
        [string]$PathValue,
        [string]$Root
    )
    if ([string]::IsNullOrWhiteSpace($PathValue)) { return $PathValue }
    if ([System.IO.Path]::IsPathRooted($PathValue)) { return $PathValue }
    return (Join-Path $Root $PathValue)
}

function CureFolderPath {
<#
Normalizes folder/path text from configuration values without touching the file
system. Kept dependency-free so compiled GUI and pipeline builds can call it
early during startup.
#>
    [CmdletBinding()]
    param([string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return '' }
    $text = $FolderPath.Trim()
    $text = $text.Trim('"', "'")
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    return [Environment]::ExpandEnvironmentVariables($text)
}

function Get-FederatableModelExtensions {
<#
Returns the model file extensions supported by the federation workflow.
#>
    return @('.ifc', '.dwg', '.dgn', '.rvt', '.nwc')
}

function Test-GeneratedNwcForSourceNames {
<#
Returns true when an NWC file name looks like a Navisworks cache generated from
one of the supplied source model names. Navisworks commonly creates both
Model.nwc and Model.ifc.nwc style cache names.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$NwcName,
        [Parameter(Mandatory = $true)][array]$SourceNames,
        [string[]]$SourceExtensions = @('.ifc', '.dwg', '.dgn', '.rvt')
    )

    if ([string]::IsNullOrWhiteSpace($NwcName)) { return $false }
    if ([System.IO.Path]::GetExtension($NwcName).ToLowerInvariant() -ne '.nwc') { return $false }

    $normalizedNwcName = [System.IO.Path]::GetFileName($NwcName).ToLowerInvariant()
    $normalizedSourceExtensions = @($SourceExtensions | ForEach-Object { $_.ToLowerInvariant() })

    foreach ($sourceName in @($SourceNames)) {
        if ([string]::IsNullOrWhiteSpace($sourceName)) { continue }
        $sourceFileName = [System.IO.Path]::GetFileName($sourceName.ToString())
        $sourceExtension = [System.IO.Path]::GetExtension($sourceFileName).ToLowerInvariant()
        if ($normalizedSourceExtensions -notcontains $sourceExtension) { continue }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sourceFileName)
        if ($normalizedNwcName -eq ("{0}.nwc" -f $baseName).ToLowerInvariant()) { return $true }
        if ($normalizedNwcName -eq ("{0}.nwc" -f $sourceFileName).ToLowerInvariant()) { return $true }
    }

    return $false
}

function Get-ExcelDataSafe {
<#
JSON/CSV build compatibility shim.
Reads CSV or JSON tabular files and returns PowerShell objects. Spreadsheet workbooks are intentionally unsupported in this build.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$NamedRange
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Tabular data file not found: $Path"
    }

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -eq '.csv') {
        return @(Import-Csv -LiteralPath $Path)
    }
    if ($extension -eq '.json') {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($raw.PSObject.Properties.Name -contains $NamedRange) { return @($raw.$NamedRange) }
        return @($raw)
    }

    throw "Spreadsheet files are not supported in the JSON/CSV build: $Path. Use .json configuration files and .csv attribute files."
}
function Get-SettingsSafe {
<# JSON/CSV build: settings are read from Config.json only. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ConfigPath,
        [string]$BasePath
    )
    throw 'Spreadsheet configuration files are not supported in the JSON/CSV build. Use Config.json.'
}
function Ensure-ExcelTable {
<# JSON/CSV build: CSV files do not need Excel tables. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$WorksheetName,
        [Parameter(Mandatory = $true)][string]$TableName
    )
    return
}
function Get-LogFilePath {
<#
Returns the log file path and append mode.
- Reuses a stable filename for OneDrive/SharePoint-synced folders so version history can do the retention work.
- Falls back to a timestamped filename only when the stable cloud file is read-only or locked.
- Uses a timestamped filename for other folders to avoid local run collisions.
#>
    param(
        [Parameter(Mandatory = $true)][string]$LogsFolder,
        [Parameter(Mandatory = $true)][string]$BaseName
    )
    if ($global:MainLogInfo -and $global:MainLogInfo.Path -and
        ($null -eq $BaseName -or $global:MainLogInfo.BaseName -eq $BaseName)) {
        return $global:MainLogInfo
    }

    $normalizedLogsFolder = CureFolderPath $LogsFolder
    $cloudRoots = @()
    if (-not [string]::IsNullOrWhiteSpace($Env:OneDriveCommercial)) {
        $cloudRoots += (CureFolderPath $Env:OneDriveCommercial)
    }
    if (-not [string]::IsNullOrWhiteSpace($Env:OneDrive)) {
        $cloudRoots += (CureFolderPath $Env:OneDrive)
    }
    if (-not [string]::IsNullOrWhiteSpace($Env:OneDriveCommercial)) {
        $sharePointRoot = CureFolderPath ($Env:OneDriveCommercial -replace 'OneDrive - ', '')
        if (-not [string]::IsNullOrWhiteSpace($sharePointRoot)) {
            $cloudRoots += $sharePointRoot
        }
    }

    foreach ($root in ($cloudRoots | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        if ($normalizedLogsFolder.Equals($root, [System.StringComparison]::OrdinalIgnoreCase) -or
            $normalizedLogsFolder.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            $stablePath = Join-Path $normalizedLogsFolder $BaseName
            $useFallback = $false

            if (Test-Path $stablePath) {
                try {
                    $item = Get-Item -LiteralPath $stablePath -ErrorAction Stop
                    if (($item.Attributes -band [System.IO.FileAttributes]::ReadOnly) -ne 0) {
                        $useFallback = $true
                    }
                    else {
                        $lockProbe = [System.IO.File]::Open($stablePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
                        $lockProbe.Close()
                    }
                }
                catch {
                    $useFallback = $true
                }
            }

            if (-not $useFallback) {
                return [pscustomobject]@{
                    Path   = $stablePath
                    Append = $false
                    Reset  = (Test-Path $stablePath)
                }
            }
        }
    }

    # Use a timestamp prefix for non-cloud folders, or when the stable cloud file cannot be overwritten.
    $prefix = (Get-Date).ToString('yyyy-MM-dd-HHmm---')
    return [pscustomobject]@{
        Path   = (Join-Path $normalizedLogsFolder ($prefix + $BaseName))
        Append = $false
        Reset  = $false
    }
}

function Initialize-LogFileTarget {
<#
Prepares a log file path for the current run.
- Clears the existing file once when Reset is requested.
- Leaves timestamped/new files untouched.
#>
    param(
        [Parameter(Mandatory = $true)][psobject]$LogInfo
    )

    if (-not $LogInfo.Path) { return }
    if (-not $global:PreparedLogFiles) {
        $global:PreparedLogFiles = @{}
    }
    if ($global:PreparedLogFiles.ContainsKey($LogInfo.Path)) { return }

    if ($LogInfo.PSObject.Properties.Name -contains 'Reset' -and $LogInfo.Reset -and (Test-Path $LogInfo.Path)) {
        Set-Content -Path $LogInfo.Path -Value $null
    }

    $global:PreparedLogFiles[$LogInfo.Path] = $true
}

function New-TempWorkspace {
    <#
        Creates a temporary working folder and returns its path.
        - Prefers D:\temp if D: exists and has >= 100 GB free.
        - Falls back to C:\temp; warns if C: has < 100 GB free.
    #>
    [CmdletBinding()]
    param()

    $requiredFreeBytes = 100GB

    function Get-DriveInfo([string]$letter) {
        try { return [System.IO.DriveInfo]::new($letter) }
        catch { return $null }
    }

    $targetDrive = $null
    $driveD = Get-DriveInfo -letter 'D'
    if ($driveD -and $driveD.IsReady -and $driveD.TotalFreeSpace -ge $requiredFreeBytes) {
        $targetDrive = 'D'
    }

    if (-not $targetDrive) {
        $targetDrive = 'C'
        $driveC = Get-DriveInfo -letter 'C'
        if ($driveC -and $driveC.TotalFreeSpace -lt $requiredFreeBytes) {
            $freeGB = [math]::Round($driveC.TotalFreeSpace / 1GB, 2)
            Write-Warning ("Only {0} GB free on drive C:. Proceeding anyway." -f $freeGB)
        }
    }

    $rootPath = Join-Path -Path ("{0}:" -f $targetDrive) -ChildPath 'temp'
    if (-not (Test-Path $rootPath)) {
        New-Item -Path $rootPath -ItemType Directory -Force | Out-Null
    }

    $folderName = "tmp_{0}" -f ([guid]::NewGuid().ToString('N'))
    $workspacePath = Join-Path -Path $rootPath -ChildPath $folderName
    New-Item -Path $workspacePath -ItemType Directory -Force | Out-Null
    return $workspacePath
}

function Remove-TempWorkspace {
    <#
        Deletes all contents inside a temp workspace and removes the folder.
        - Accepts the folder path returned by New-TempWorkspace.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        Write-Verbose "Temp folder '$Path' not found."
        return
    }

    try {
        Get-ChildItem -Path $Path -Recurse -Force -ErrorAction Stop |
            Remove-Item -Recurse -Force -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to delete all items in '$Path': $_"
    }

    try {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to remove folder '$Path': $_"
    }
}

function Resolve-SupportFile {
<#
Resolves a support file path, extracting embedded resources when needed.
#>
    param(
        [Parameter(Mandatory = $true)][string]$FileName,
        [string]$BasePath
    )

    $rootPath = if ($BasePath) { $BasePath } elseif ($PSCommandPath) {
        Split-Path -Parent $PSCommandPath
    }
    elseif ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    else {
        Add-Type -AssemblyName System.Reflection
        Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location)
    }

    $candidate = Join-Path $rootPath $FileName
    if (Test-Path $candidate) { return $candidate }

    # Try to extract from embedded resources (when running as packaged EXE)
    $asm = [System.Reflection.Assembly]::GetEntryAssembly()
    if (-not $asm) { return $candidate }
    $resourceName = $asm.GetManifestResourceNames() |
        Where-Object { $_ -like "*$FileName" } |
        Select-Object -First 1
    if (-not $resourceName) { return $candidate }

    $tempDir = Join-Path $env:APPDATA "Pythonx"
    if (-not (Test-Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    }
    $targetPath = Join-Path $tempDir $FileName
    if (Test-Path $targetPath) { return $targetPath }

    $stream = $asm.GetManifestResourceStream($resourceName)
    $outStream = [System.IO.File]::Open($targetPath, 'Create', 'Write', 'ReadWrite')
    $stream.CopyTo($outStream)
    $outStream.Dispose()
    $stream.Dispose()

    return $targetPath
}

function Resolve-OptionalXmlSettingPath {
<#
Normalizes an optional XML file setting.
- Blank values are ignored.
- Adds a .xml suffix when missing.
- Relative paths are resolved from the runtime base folder.
- Returns the candidate path plus whether it exists.
#>
    [CmdletBinding()]
    param(
        [string]$Value,
        [string]$BasePath
    )

    $rootPath = if ($BasePath) { $BasePath } elseif ($PSCommandPath) {
        Split-Path -Parent $PSCommandPath
    }
    elseif ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    else {
        Add-Type -AssemblyName System.Reflection
        Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location)
    }

    $text = Normalize-ExcelText -Text $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{
            RequestedValue = $null
            CandidatePath  = $null
            Exists         = $false
        }
    }

    $text = $text -replace '/', '\'
    if (-not $text.EndsWith('.xml', [System.StringComparison]::OrdinalIgnoreCase)) {
        $text = "{0}.xml" -f $text
    }

    $candidatePath = if ([System.IO.Path]::IsPathRooted($text)) {
        $text
    }
    else {
        Join-Path $rootPath $text
    }

    return [pscustomobject]@{
        RequestedValue = $text
        CandidatePath  = $candidatePath
        Exists         = (Test-Path $candidatePath -PathType Leaf)
    }
}

function Invoke-Step {
<#
Wraps a pipeline step with a standard header and execution.
#>
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    $stepStart = Get-Date
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host ""
    Write-Host "=== $Name ===" -ForegroundColor Cyan
    try {
        & $Action
    }
    catch {
        $message = "ERROR: Step '$Name' failed: $_"
        Write-Host $message -ForegroundColor Red
        throw
    }
    finally {
        $timer.Stop()
        $stepEnd = Get-Date
        $durationText = Format-Duration -Duration $timer.Elapsed
        Write-Host ("=== END {0} (Duration: {1}) ===" -f $Name, $durationText) -ForegroundColor Cyan
    }

    return [pscustomobject]@{
        Name     = $Name
        Start    = $stepStart
        End      = $stepEnd
        Duration = $timer.Elapsed
    }
}

function Format-Duration {
<#
Formats a TimeSpan for logging.
#>
    param($Duration)
    if ($null -eq $Duration) { return "00:00:00" }
    if (-not ($Duration -is [TimeSpan])) {
        $parsed = [TimeSpan]::Zero
        if (-not [TimeSpan]::TryParse($Duration.ToString(), [ref]$parsed)) {
            return "00:00:00"
        }
        $Duration = $parsed
    }
    $hhmmss = "{0:hh\:mm\:ss}" -f $Duration
    return ("{0} ({1:N2} min)" -f $hhmmss, $Duration.TotalMinutes)
}

function Write-Log {
<#
Writes a message to console and the main pipeline log (single entry point).
#>
    param(
        [string]$Message,
        [string]$Color = 'Yellow',
        [array]$Settings,
        [string]$LogsFolderOverride,
        [string]$BasePath,
        [switch]$NoConsole
    )

    if (-not $NoConsole) {
        if (-not [string]::IsNullOrWhiteSpace($Color)) {
            Write-Host $Message -ForegroundColor $Color
        }
        else {
            Write-Host $Message
        }
    }

    if ($global:MainTranscriptActive) {
        return
    }

    $logFolderValue = if ($LogsFolderOverride) { $LogsFolderOverride } elseif ($Settings) {
        $Settings | Where-Object { $_.Parameter -eq "LogFolder" } | Select-Object -ExpandProperty Value
    } else { $null }
    if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
        $logFolderValue = CureFolderPath $logFolderValue
    }
    if (-not $logFolderValue) {
        $rootPath = if ($BasePath) { $BasePath } elseif ($PSCommandPath) {
            Split-Path -Parent $PSCommandPath
        }
        elseif ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') {
            Split-Path -Parent $MyInvocation.MyCommand.Path
        }
        else {
            Add-Type -AssemblyName System.Reflection
            Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location)
        }
        $logFolderValue = Join-Path $rootPath 'Logs'
    }
    try {
        if (-not (Test-Path $logFolderValue)) { New-Item -Path $logFolderValue -ItemType Directory -Force | Out-Null }
    }
    catch {
        Write-Warning "Could not create or access log folder '$logFolderValue'."
        return
    }

    $logLine = ("{0} {1}" -f (Get-Date).ToString("s"), $Message)
    if (Get-Command Get-LogFilePath -ErrorAction SilentlyContinue) {
        $logInfo = Get-LogFilePath -LogsFolder $logFolderValue -BaseName 'PW_Process.log.txt'
        Initialize-LogFileTarget -LogInfo $logInfo
        Add-Content -Path $logInfo.Path -Value $logLine
    }
    else {
        Add-Content -Path (Join-Path $logFolderValue 'PW_Process.log.txt') -Value $logLine
    }
}

function Write-MainLog {
<#
Backward-compatible wrapper for Write-Log.
#>
    param(
        [string]$Message,
        [string]$Color = 'Yellow',
        [array]$Settings,
        [string]$LogsFolderOverride,
        [string]$BasePath,
        [switch]$NoConsole
    )
    Write-Log -Message $Message -Color $Color -Settings $Settings -LogsFolderOverride $LogsFolderOverride -BasePath $BasePath -NoConsole:$NoConsole
}

function Get-RowValue {
<#
Returns the first non-blank value from a row using candidate column names.
#>
    param($Row, [string[]]$Candidates)
    foreach ($name in $Candidates) {
        if ($Row.PSObject.Properties.Name -contains $name) {
            $val = $Row.$name
            if (-not ([string]::IsNullOrWhiteSpace($val))) { return [string]$val }
        }
    }
    return $null
}

function Normalize-ExcelText {
<#
Trims Excel text and removes non-breaking/zero-width spaces.
#>
    param([string]$Text)
    if ($null -eq $Text) { return $null }
    $clean = [regex]::Replace($Text, "[\u00a0\u2000-\u200b]", "")
    return $clean.Trim()
}

function Get-FederationPartDefinitions {
<#
Parses Federation rows into filename-part definitions used by processing and federation.
#>
    [CmdletBinding()]
    param(
        [array]$FederationRows
    )

    $definitions = @()
    $rowOrder = 0
    foreach ($row in @($FederationRows)) {
        if (-not $row) { continue }
        $rowOrder++

        $fieldName = Get-RowValue -Row $row -Candidates @('FieldNames', 'FieldName')
        if ([string]::IsNullOrWhiteSpace($fieldName) -and $row.PSObject.Properties.Count -gt 0) {
            $fieldName = $row.PSObject.Properties[0].Value
        }
        $fieldName = Normalize-ExcelText -Text $fieldName
        if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }

        $fileNamePart = Get-RowValue -Row $row -Candidates @('Filename-Part', 'FilenamePart', 'Filename_Part')
        $fileNamePart = Normalize-ExcelText -Text $fileNamePart
        if ([string]::IsNullOrWhiteSpace($fileNamePart)) { continue }

        $isFileExtension = $fileNamePart.ToLowerInvariant() -eq 'fileextension'
        $position = 0
        if (-not $isFileExtension) {
            if (-not [int]::TryParse($fileNamePart, [ref]$position) -or $position -le 0) { continue }
        }

        $groupOrder = 0
        $groupOrderText = Get-RowValue -Row $row -Candidates @('GroupOrder')
        if (-not [string]::IsNullOrWhiteSpace($groupOrderText)) {
            [void][int]::TryParse($groupOrderText, [ref]$groupOrder)
        }

        $injectText = Normalize-ExcelText -Text (Get-RowValue -Row $row -Candidates @('InjectToIFC', 'InjectToIfc', 'Inject2IFC', 'Inject2Ifc'))
        $injectEnabled = $false
        if (-not [string]::IsNullOrWhiteSpace($injectText)) {
            $injectEnabled = $injectText.ToLowerInvariant() -in @('yes', 'y', 'true', '1')
        }

        $definitions += [pscustomobject]@{
            Name         = $fieldName
            GroupOrder   = $groupOrder
            FileNamePart = $fileNamePart
            Position     = $position
            SortPosition = $(if ($isFileExtension) { [int]::MaxValue } else { $position })
            IsFileExtension = $isFileExtension
            InjectToIFC  = $injectEnabled
            RowOrder     = $rowOrder
        }
    }

    return @($definitions | Sort-Object SortPosition, RowOrder)
}

function Split-ModelFileNameParts {
<#
Splits a model file name (without extension) into normalized dash-separated parts.
#>
    [CmdletBinding()]
    param(
        [string]$FileName
    )

    if ([string]::IsNullOrWhiteSpace($FileName)) { return @() }
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    if ([string]::IsNullOrWhiteSpace($baseName)) { return @() }
    return @($baseName -split '-' | ForEach-Object { Normalize-ExcelText -Text $_ })
}

function Get-ModelFileExtensionPart {
<#
Returns the normalized Windows file extension, including the leading dot.
#>
    [CmdletBinding()]
    param(
        [string]$FileName
    )

    if ([string]::IsNullOrWhiteSpace($FileName)) { return $null }
    $extension = [System.IO.Path]::GetExtension($FileName)
    if ([string]::IsNullOrWhiteSpace($extension)) { return $null }
    return Normalize-ExcelText -Text $extension
}

function Get-FederationPartValue {
<#
Returns the value for a federation part definition from a file name.
- Numeric parts come from dash-separated filename sections.
- The special FileExtension part comes from the Windows file extension.
#>
    [CmdletBinding()]
    param(
        [string]$FileName,
        [psobject]$Definition,
        [array]$NameParts
    )

    if (-not $Definition) { return $null }
    if ($Definition.PSObject.Properties.Name -contains 'IsFileExtension' -and $Definition.IsFileExtension) {
        return Get-ModelFileExtensionPart -FileName $FileName
    }

    $parts = if ($null -ne $NameParts) { @($NameParts) } else { @(Split-ModelFileNameParts -FileName $FileName) }
    $position = 0
    if (-not [int]::TryParse([string]$Definition.Position, [ref]$position) -or $position -le 0) { return $null }
    $index = $position - 1
    if ($index -lt 0 -or $index -ge $parts.Count) { return $null }
    return Normalize-ExcelText -Text $parts[$index]
}

function Test-FederationFileNameMatch {
<#
Returns true when all configured Federation filename parts exist and are non-blank in the file name.
#>
    [CmdletBinding()]
    param(
        [string]$FileName,
        [array]$PartDefinitions
    )

    $definitions = @($PartDefinitions | Where-Object {
        $_ -and
        -not ($_.PSObject.Properties.Name -contains 'IsFileExtension' -and $_.IsFileExtension) -and
        $_.Position -gt 0
    })
    if ($definitions.Count -eq 0) { return $true }

    $parts = @(Split-ModelFileNameParts -FileName $FileName)
    foreach ($definition in $definitions) {
        $index = [int]$definition.Position - 1
        if ($index -lt 0 -or $index -ge $parts.Count) { return $false }
        if ([string]::IsNullOrWhiteSpace($parts[$index])) { return $false }
    }

    return $true
}






