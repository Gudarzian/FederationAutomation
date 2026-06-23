<#
Federation-related functions for Federation Automation.
- IFC grouping, Navisworks federation, and publish helpers.
#>

function BuildGroupHierarchy {
<#
Builds a hierarchical grouping tree from flat file rows.
- Groups by the provided keys in order, creating nested nodes.
- Leaf nodes return an array of full file paths (from FileName).
#>
# === RECURSIVE GROUP BUILDER ===
    param(
        [array]$Data,
        [array]$GroupKeys = @(),
        [string]$ParentName = "",
        [string]$NamingMethod = "Full"
    )
    if (-not $GroupKeys -or $GroupKeys.Count -eq 0) {
        # Leaf case: return an array of full file paths for federation.
        return @($Data.FileName)
    }
    $resolvedNamingMethod = Resolve-NwdNamingMethod -Value $NamingMethod
    $nameSeparator = if ($resolvedNamingMethod -eq 'full') { '---' } else { '-' }
    $currentKey    = $GroupKeys[0]
    $remainingKeys = if ($GroupKeys.Count -gt 1) { $GroupKeys | Select-Object -Skip 1 } else { @() }
    # Split input rows by the current grouping key (e.g., Disc, Zone, Level).
    $groups = $Data | Group-Object $currentKey | Sort-Object Name
    $result = @()
    foreach ($g in $groups) {
        # Build a stable group name chain for logging and output naming.
        $sampleRow = $g.Group | Select-Object -First 1
        $descPropertyName = "__Desc_{0}" -f $currentKey
        $description = $null
        if ($sampleRow -and $sampleRow.PSObject.Properties.Name -contains $descPropertyName) {
            $description = $sampleRow.$descPropertyName
        }
        $currentNamePart = Get-NwdGroupNamePart -Key $currentKey -Code $g.Name -Description $description -NamingMethod $resolvedNamingMethod
        $fullName = if ($ParentName) { "$ParentName$nameSeparator$currentNamePart" } else { $currentNamePart }
        $result += [pscustomobject]@{
            Name     = $g.Name
            Key      = $currentKey
            FullName = $fullName
            Children = BuildGroupHierarchy -Data $g.Group -GroupKeys $remainingKeys -ParentName $fullName -NamingMethod $resolvedNamingMethod
        }
    }
    return $result
}

function Resolve-NwdNamingMethod {
<#
Normalizes NWD naming mode values and falls back to Full.
#>
    [CmdletBinding()]
    param(
        [string]$Value
    )

    $text = Normalize-ExcelText -Text $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return 'full' }

    switch ($text.ToLowerInvariant()) {
        'onlycodes' { return 'onlycodes' }
        'onlydesc' { return 'onlydesc' }
        'codes-desc' { return 'codes-desc' }
        'full' { return 'full' }
        default { return 'full' }
    }
}

function Get-NwdGroupNamePart {
<#
Builds one NWD name segment for a grouped field based on the configured naming method.
#>
    [CmdletBinding()]
    param(
        [string]$Key,
        [string]$Code,
        [string]$Description,
        [string]$NamingMethod = 'Full'
    )

    $resolvedMethod = Resolve-NwdNamingMethod -Value $NamingMethod
    $codeText = Normalize-ExcelText -Text $Code
    $descText = Normalize-ExcelText -Text $Description
    if ([string]::IsNullOrWhiteSpace($codeText)) { $codeText = 'Unknown' }

    switch ($resolvedMethod) {
        'onlycodes' {
            return $codeText
        }
        'onlydesc' {
            if (-not [string]::IsNullOrWhiteSpace($descText)) { return $descText }
            return $codeText
        }
        'codes-desc' {
            if (-not [string]::IsNullOrWhiteSpace($descText)) { return ("{0}={1}" -f $codeText, $descText) }
            return $codeText
        }
        default {
            $valueText = if (-not [string]::IsNullOrWhiteSpace($descText)) { "{0} ({1})" -f $codeText, $descText } else { $codeText }
            return ("{0}_{1}" -f $Key, $valueText)
        }
    }
}

function Get-GroupHierarchyNodeNames {
<#
Returns the FullName values for all non-leaf grouping nodes.
#>
    [CmdletBinding()]
    param(
        [array]$Groups
    )

    $names = @()
    foreach ($group in @($Groups)) {
        if (-not $group) { continue }
        if ($group.PSObject.Properties.Name -contains 'FullName') {
            $names += $group.FullName
        }
        if ($group.PSObject.Properties.Name -contains 'Children' -and $group.Children -and
            $group.Children[0] -and $group.Children[0].PSObject.Properties.Name -contains 'FullName') {
            $names += Get-GroupHierarchyNodeNames -Groups $group.Children
        }
    }
    return @($names)
}

function Get-SafeFileName {
<#
Returns a filesystem-safe file name derived from a label.
- Replaces invalid characters with '_', trims trailing dots/spaces.
#>
    param(
        [Parameter(Mandatory = $true)][string]$Name
    )
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder($Name)
    foreach ($ch in $invalid) { [void]$sb.Replace($ch, '_') }
    ($sb.ToString()).Trim().Trim('.')
}

function Resolve-NavisworksInstallPath {
<#
Resolves the install folder for a specific Navisworks Manage version.
#>
    param([string]$Version)
    if ([string]::IsNullOrWhiteSpace($Version)) { return $null }
    $roots = @(
        'C:\Program Files\Autodesk',
        'C:\Program Files (x86)\Autodesk'
    )
    foreach ($root in $roots) {
        $candidate = Join-Path $root ("Navisworks Manage {0}" -f $Version)
        $roamer = Join-Path $candidate 'Roamer.exe'
        if (Test-Path $roamer -PathType Leaf) {
            return $candidate
        }
    }
    return $null
}

function Resolve-NavisworksVersion {
<#
Finds the highest installed Navisworks Manage version (year) from Program Files.
#>
    $roots = @(
        'C:\Program Files\Autodesk',
        'C:\Program Files (x86)\Autodesk'
    )
    $versions = @()
    foreach ($root in $roots) {
        if (-not (Test-Path $root -PathType Container)) { continue }
        $dirs = Get-ChildItem -Path $root -Directory -Filter 'Navisworks Manage *' -ErrorAction SilentlyContinue
        foreach ($dir in $dirs) {
            $match = [regex]::Match($dir.Name, 'Navisworks Manage\s+(?<ver>\d{4})')
            if ($match.Success) {
                $verText = $match.Groups['ver'].Value
                $verNum = 0
                if ([int]::TryParse($verText, [ref]$verNum)) {
                    $roamer = Join-Path $dir.FullName 'Roamer.exe'
                    if (Test-Path $roamer -PathType Leaf) {
                        $versions += $verNum
                    }
                }
            }
        }
    }
    if ($versions.Count -gt 0) {
        return ($versions | Sort-Object -Descending | Select-Object -First 1)
    }
    return $null
}

if (-not (Get-Variable -Name PWAttributeDefinitions -Scope Script -ErrorAction SilentlyContinue)) {
    $script:PWAttributeDefinitions = @()
}

function CreateNwd {
<#
Creates a Navisworks NWD from a group node.
- Leaf groups: Children are IFC paths; parents: Children are child NWDs.
- Uses the group's FullName for output name (sanitized in caller where used).
#>
    param(
        $InputFiles,
        $OutputPath = $OutputPath,
        $NavisworksVersion = $NavisworksVersion,
        $NavisworksOptions = $NavisworksOptions,
        $NavisworksOptionsPath = $NavisworksOptionsPath,
        $NavisworksViewsImportXmlPath = $NavisworksViewsImportXmlPath,
        $BasePath = $BasePath
    )
    # Basic runtime info for console/trace logs.
    Write-Host "Processing $($InputFiles.FullName)... Do not close the window."
    $start = Get-Date

    # Output name is derived from the grouping hierarchy and sanitized for filesystem.
    $safeName      = Get-SafeFileName -Name $InputFiles.FullName
    $outputNwd     = $OutputPath + "\" + $safeName + ".nwd"
    $navisworksDir = Resolve-NavisworksInstallPath -Version $NavisworksVersion
    $roamerExe     = if ($navisworksDir) { Join-Path $navisworksDir 'Roamer.exe' } else { $null }

    # Build arguments for Roamer.exe to append all group inputs.
    $arg = @()
    foreach ($Inputfile in $InputFiles.Children) {
        $arg += '-AppendFile', $Inputfile
    }
    $arg += '-SaveFile', $outputNwd
    $arg += $GNWUI

    # Resolve options: Settings override > fallback to NavisworksOptions.xml.
    $optionsPath = $null
    if ($NavisworksOptionsPath) {
        $optionsPath = $NavisworksOptionsPath
    }
    elseif ($NavisworksOptions) {
        $optionsPath = $NavisworksOptions
    }
    if ($optionsPath) {
        if (-not [System.IO.Path]::IsPathRooted($optionsPath)) {
            $root = if ($BasePath) { $BasePath } else { $PSScriptRoot }
            $optionsPath = Join-Path $root $optionsPath
        }
        if (-not (Test-Path $optionsPath)) {
            $root = if ($BasePath) { $BasePath } else { $PSScriptRoot }
            $resolved = Resolve-SupportFile -FileName $optionsPath -BasePath $root
            if (Test-Path $resolved) { $optionsPath = $resolved } else { $optionsPath = $null }
        }
    }
    if (-not $optionsPath) {
        $root = if ($BasePath) { $BasePath } else { $PSScriptRoot }
        $fallback = Resolve-SupportFile -FileName 'NavisworksOptions.xml' -BasePath $root
        if (Test-Path $fallback) { $optionsPath = $fallback }
    }
    if ($optionsPath) { $arg += '-options', $optionsPath }
    if ($NavisworksViewsImportXmlPath) {
        $arg += '-ExecuteAddInPlugin', 'NativeExportPluginAdaptor_XmlViewpointsImportPlugin_Import.Navisworks', $NavisworksViewsImportXmlPath
    }
    $arg += '-Exit'

    # Build command string for logging and reproducibility.
    $argString = ($arg | ForEach-Object {
        if ($_ -match '\s') { '"' + $_ + '"' } else { $_ }
    }) -join ' '

    if ($InputFiles.Children.Count -gt 0) {
        # Ensure output folder exists, then run Navisworks in batch mode.
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        $p = Start-Process -FilePath $roamerExe -WorkingDirectory $navisworksDir -ArgumentList $argString -PassThru
        $p.WaitForExit()
    } else {
        Write-Host "No matching files found for provided wildcards."
    }

    $end     = Get-Date
    $minutes = [math]::Round((New-TimeSpan -Start $start -End $end).TotalMinutes, 2)

    $logTxt = @"
NWD processed for $outputNwd
Number of model(s) federated: $($InputFiles.Children.Count)
Navisworks Version: $NavisworksVersion
End Time: $end
Time spent: $minutes minutes
NWD output path: $outputNwd
Model(s) federated: $([string]::Join(', ', $InputFiles.Children))
NW Arguments: $argString
"@
    Write-Host $logTxt

    # Return a summary object for further use
    [pscustomobject]@{
        Model             = $outputNwd
        FederatedFiles    = $InputFiles.Children
        OutputPath        = $OutputPath
        NavisworksVersion = $NavisworksVersion
        StartTime         = $start
        EndTime           = $end
        TimeMinutes       = $minutes
        Command           = $argString
        NWArguments       = ($InputFiles.Children.Count -gt 0)
    }
}

function FederateGroup {
    param(
        [Parameter(Mandatory = $true)]
        $Group,
        [int]$Level = 0,
        [string]$NavisworksOptionsPath,
        [string]$NavisworksOptions,
        [string]$BasePath
    )
    # Print the tree and federate bottom-up so child NWDs exist for parents.
    $indent = ' ' * ($Level * 2)
    Write-Host "$indent$($Group.FullName)" -ForegroundColor Cyan

    $hasNested = ($Group.Children -and $Group.Children[0].PSObject.Properties['Children'])
    if ($hasNested) {
        $childOutputs = @()
        foreach ($child in $Group.Children) {
            $childOutputs += FederateGroup -Group $child -Level ($Level + 1)
        }
        $groupForCreate = [pscustomobject]@{
            FullName = $Group.FullName
            Children = $childOutputs
        }
        $res = CreateNwd -InputFiles $groupForCreate -NavisworksOptionsPath $NavisworksOptionsPath -NavisworksOptions $NavisworksOptions -BasePath $BasePath
        return $res.Model
    }
    else {
        foreach ($file in $Group.Children) { Write-Host "$indent  $file" }
        $res = CreateNwd -InputFiles $Group -NavisworksOptionsPath $NavisworksOptionsPath -NavisworksOptions $NavisworksOptions -BasePath $BasePath
        return $res.Model
    }
}

function Invoke-GroupFilesForFedProcess {
<#
Groups IFCs and federates them into NWD outputs.
#>
    [CmdletBinding()]
    param(
        [string]$ConfigFile = "Config.xlsx",
        [array]$Settings = @(),
        [Parameter(Mandatory = $true)][array]$FederationRows,
        [array]$WildcardSelectionRows = @(),
        [array]$LookupRows,
        [string]$LogsFolder,
        [switch]$UseFederationInputFolder
    )

    # Resolve script root to find Config.xlsx and related assets.
    $basePath = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
    elseif ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    else {
        Add-Type -AssemblyName System.Reflection
        Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location)
    }
    $ConfigFilePath = if ([System.IO.Path]::IsPathRooted($ConfigFile)) {
        $ConfigFile
    }
    else {
        Join-Path $basePath $ConfigFile
    }
    $ConfigSettings = @($Settings)
    if (-not $ConfigSettings -or $ConfigSettings.Count -eq 0) {
        throw "Settings were not supplied to Invoke-GroupFilesForFedProcess."
    }
    $runFederationKeys = @('runfederation','federationrun','runfederate','federate','federation')
    $runFederationRow = $ConfigSettings | Where-Object {
        $paramRaw = $_.Parameter
        $paramText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $paramRaw
        }
        else {
            if ($null -ne $paramRaw) { $paramRaw.ToString().Trim() } else { $null }
        }
        if ([string]::IsNullOrWhiteSpace($paramText)) { return $false }
        $runFederationKeys -contains $paramText.ToLowerInvariant()
    } | Select-Object -First 1
    if ($runFederationRow -and $null -ne $runFederationRow.Value) {
        $valueText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $runFederationRow.Value
        }
        else {
            $runFederationRow.Value.ToString().Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($valueText)) {
            $normalized = $valueText.ToLowerInvariant()
            if ($normalized -in @('no','n','false','0','ignore')) {
                Write-Host ("Federation ignored by Settings parameter '{0}'." -f $runFederationRow.Parameter) -ForegroundColor Yellow
                return
            }
        }
    }
    # Resolve log path (including OneDrive/ACC tokens) and start transcript logging.
    $LogsFolder = if ($LogsFolder) { $LogsFolder } else { CureFolderPath($ConfigSettings | Where-Object { $_.Parameter -eq "LogFolder" } | Select-Object -ExpandProperty Value) }
    if (-not $LogsFolder) { $LogsFolder = Join-Path $basePath 'Logs' }
    try {
        if (-not (Test-Path $LogsFolder)) { New-Item -Path $LogsFolder -ItemType Directory -Force | Out-Null }
    }
    catch {
        Write-Warning "Could not create or access log folder '$LogsFolder'. Falling back to TEMP. Error: $_"
        $LogsFolder = Join-Path $env:TEMP 'PWLogs'
        if (-not (Test-Path $LogsFolder)) { New-Item -Path $LogsFolder -ItemType Directory -Force | Out-Null }
    }
    $logInfo = Get-LogFilePath -LogsFolder $LogsFolder -BaseName 'PW_Process.log.txt'
    $LogFilePath = $logInfo.Path
    $TranscriptStarted = $false
    if (-not $global:MainTranscriptActive) {
        try {
            Initialize-LogFileTarget -LogInfo $logInfo
            Start-Transcript -Path $LogFilePath -Append:$($logInfo.Append) -ErrorAction Stop | Out-Null
            $TranscriptStarted = $true
            $global:MainTranscriptActive = $true
        }
        catch {
            Write-Warning "Transcript could not be started. Logging will be console-only. Error: $_"
        }
    }
    $ProcessStartTime = Get-Date
    Write-Host "Logging to $LogFilePath"
    Write-Host "=== START 050-GroupFilesForFedProcess ==="
    Write-Host "Process start time: $ProcessStartTime"
    Write-Host "Config file: $ConfigFilePath"

    # Load federation inputs/outputs and Navisworks preferences from Settings.
    $federationInputPath = CureFolderPath ($ConfigSettings | Where-Object { $_.Parameter -eq "FederationInputFolder" }  | Select-Object -ExpandProperty Value)
    $processedFolderPath = CureFolderPath ($ConfigSettings | Where-Object { $_.Parameter -eq "ProcessedFolder" }  | Select-Object -ExpandProperty Value)
    $sourceFolderPath = CureFolderPath ($ConfigSettings | Where-Object { $_.Parameter -eq "SourceFolder" }  | Select-Object -ExpandProperty Value)
    if ([string]::IsNullOrWhiteSpace($processedFolderPath)) { $processedFolderPath = "ProcessedIFC" }
    if ([string]::IsNullOrWhiteSpace($sourceFolderPath)) { $sourceFolderPath = "SourceFiles" }
    if ([string]::IsNullOrWhiteSpace($federationInputPath)) {
        $runProcessSetting = $ConfigSettings | Where-Object {
            $paramRaw = $_.Parameter
            $paramText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
                Normalize-ExcelText -Text $paramRaw
            }
            else {
                if ($null -ne $paramRaw) { $paramRaw.ToString().Trim() } else { $null }
            }
            if ([string]::IsNullOrWhiteSpace($paramText)) { return $false }
            $paramText.ToLowerInvariant() -in @('runprocess','forceifcprocessing','forceifcprocess','forceprocess')
        } | Select-Object -First 1
        $runProcessEnabled = $false
        if ($runProcessSetting -and $null -ne $runProcessSetting.Value) {
            $valueText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
                Normalize-ExcelText -Text $runProcessSetting.Value
            }
            else {
                $runProcessSetting.Value.ToString().Trim()
            }
            if (-not [string]::IsNullOrWhiteSpace($valueText)) {
                $normalized = $valueText.ToLowerInvariant()
                if ($normalized -in @('no','n','false','0','ignore')) {
                    $runProcessEnabled = $false
                }
                else {
                    $runProcessEnabled = $true
                }
            }
        }
        $federationInputPath = if ($runProcessEnabled) { $processedFolderPath } else { $sourceFolderPath }
    }
    if ($UseFederationInputFolder) {
        $InputFolderPath = if ($federationInputPath) { $federationInputPath } else { $processedFolderPath }
    }
    else {
        $InputFolderPath = if ($processedFolderPath) { $processedFolderPath } else { $federationInputPath }
    }
    $OutputPath          = CureFolderPath ($ConfigSettings | Where-Object { $_.Parameter -eq "FederationOutputFolder" }  | Select-Object -ExpandProperty Value)
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = "Output" }
    $InputFolderPath = Resolve-RelativePath -PathValue $InputFolderPath -Root $basePath
    $OutputPath = Resolve-RelativePath -PathValue $OutputPath -Root $basePath
    if (-not (Test-Path $InputFolderPath -PathType Container)) {
        $message = "Federation input folder not found: '$InputFolderPath'."
        Write-Log -Message $message -Color 'Red' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
        throw $message
    }
    if (-not (Test-Path $OutputPath -PathType Container)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    $NavisWorksVisible   = $ConfigSettings | Where-Object { $_.Parameter -eq "NavisWorksVisible" }  | Select-Object -ExpandProperty Value
    if ($NavisWorksVisible -eq "Yes") { $GNWUI = "-ShowGui" } else { $GNWUI = "-NoGui" }
    $NavisworksVersion   = $ConfigSettings | Where-Object { $_.Parameter -eq "NavisworksVersion" }  | Select-Object -ExpandProperty Value
    $navisworksProvided = -not [string]::IsNullOrWhiteSpace($NavisworksVersion)
    if (-not $navisworksProvided) {
        $resolvedVersion = Resolve-NavisworksVersion
        if ($resolvedVersion) {
            $NavisworksVersion = $resolvedVersion
        }
    }
    $NavisworksOptions   = $ConfigSettings | Where-Object { $_.Parameter -eq "NavisworksConfigXML" }  | Select-Object -ExpandProperty Value
    if ([string]::IsNullOrWhiteSpace($NavisworksOptions)) {
        $NavisworksOptions = "NavisworksOptions.xml"
    }
    $NavisworksViewsImportXmlPath = $null
    $navisworksViewsImportSetting = $ConfigSettings | Where-Object {
        $paramRaw = $_.Parameter
        $paramText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $paramRaw
        }
        else {
            if ($null -ne $paramRaw) { $paramRaw.ToString().Trim() } else { $null }
        }
        if ([string]::IsNullOrWhiteSpace($paramText)) { return $false }
        $paramText.ToLowerInvariant() -eq 'navisworksviewsimportxml'
    } | Select-Object -First 1
    if ($navisworksViewsImportSetting -and $null -ne $navisworksViewsImportSetting.Value) {
        $viewsImportInfo = Resolve-OptionalXmlSettingPath -Value $navisworksViewsImportSetting.Value -BasePath $basePath
        if ($viewsImportInfo.Exists) {
            $NavisworksViewsImportXmlPath = $viewsImportInfo.CandidatePath
        }
        elseif (-not [string]::IsNullOrWhiteSpace($viewsImportInfo.RequestedValue)) {
            $warning = "NavisworksViewsImportXML file not found: '$($viewsImportInfo.CandidatePath)'. Viewpoints import will be skipped."
            Write-Host $warning -ForegroundColor Yellow
            Write-Log -Message $warning -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
        }
    }
    $FederatedFileName = $ConfigSettings | Where-Object { $_.Parameter -eq "FederatedFileName" } | Select-Object -ExpandProperty Value
    if ([string]::IsNullOrWhiteSpace($FederatedFileName)) { $FederatedFileName = 'Project Federated' }
    if ($FederatedFileName.EndsWith('.nwd', [System.StringComparison]::OrdinalIgnoreCase)) {
        $FederatedFileName = $FederatedFileName.Substring(0, $FederatedFileName.Length - 4)
    }
    $nwdNamingMethod = 'full'
    $nwdNamingMethodRow = $ConfigSettings | Where-Object {
        $paramRaw = $_.Parameter
        $paramText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $paramRaw
        }
        else {
            if ($null -ne $paramRaw) { $paramRaw.ToString().Trim() } else { $null }
        }
        if ([string]::IsNullOrWhiteSpace($paramText)) { return $false }
        $paramText.ToLowerInvariant() -eq 'nwdnamingmethod'
    } | Select-Object -First 1
    if ($nwdNamingMethodRow -and $null -ne $nwdNamingMethodRow.Value) {
        $requestedNwdNamingMethod = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $nwdNamingMethodRow.Value
        }
        else {
            $nwdNamingMethodRow.Value.ToString().Trim()
        }
        $resolvedNwdNamingMethod = Resolve-NwdNamingMethod -Value $requestedNwdNamingMethod
        if ($resolvedNwdNamingMethod -eq 'full' -and -not [string]::IsNullOrWhiteSpace($requestedNwdNamingMethod) -and
            $requestedNwdNamingMethod.ToLowerInvariant() -ne 'full') {
            $warning = "Invalid NWDNamingMethod '$requestedNwdNamingMethod'. Falling back to Full."
            Write-Host $warning -ForegroundColor Yellow
            Write-Log -Message $warning -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
        }
        $nwdNamingMethod = $resolvedNwdNamingMethod
    }
    $includeUnmatchedSetting = $ConfigSettings | Where-Object {
        $paramRaw = $_.Parameter
        $paramText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $paramRaw
        }
        else {
            if ($null -ne $paramRaw) { $paramRaw.ToString().Trim() } else { $null }
        }
        if ([string]::IsNullOrWhiteSpace($paramText)) { return $false }
        $paramText.ToLowerInvariant() -eq 'includeunmatchedfilesinfederatedmodel'
    } | Select-Object -First 1
    $includeUnmatchedFilesInFederatedModel = $false
    if ($includeUnmatchedSetting -and $null -ne $includeUnmatchedSetting.Value) {
        $includeValue = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $includeUnmatchedSetting.Value
        }
        else {
            $includeUnmatchedSetting.Value.ToString().Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($includeValue)) {
            $normalized = $includeValue.ToLowerInvariant()
            if ($normalized -in @('no', 'n', 'false', '0', 'ignore')) {
                $includeUnmatchedFilesInFederatedModel = $false
            }
            else {
                $includeUnmatchedFilesInFederatedModel = $true
            }
        }
    }
    $OutputSheet = "GroupedResults"
    Write-Host "Input folder: $InputFolderPath"
    Write-Host "Output folder: $OutputPath"
    Write-Host "Navisworks version: $NavisworksVersion"
    $navisworksDir = Resolve-NavisworksInstallPath -Version $NavisworksVersion
    if ($navisworksProvided -and -not $navisworksDir) {
        $fallbackVersion = Resolve-NavisworksVersion
        if ($fallbackVersion) {
            $warning = "Navisworks Manage $NavisworksVersion not found. Falling back to Navisworks Manage $fallbackVersion."
            Write-Host $warning -ForegroundColor Yellow
            Write-Log -Message $warning -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
            $NavisworksVersion = $fallbackVersion
            $navisworksDir = Resolve-NavisworksInstallPath -Version $NavisworksVersion
        }
    }
    if (-not $navisworksDir) {
        $reason = if ($navisworksProvided) {
            "Navisworks Manage $NavisworksVersion not found and no other versions are installed."
        }
        else {
            "No Navisworks Manage installation found on this machine."
        }
        Write-Host ("Federation skipped: {0}" -f $reason) -ForegroundColor Yellow
        Write-Log -Message ("Federation skipped: {0}" -f $reason) -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
        if ($TranscriptStarted) {
            Stop-Transcript | Out-Null
            $global:MainTranscriptActive = $false
        }
        return
    }
    Write-Host "Navisworks options: $NavisworksOptions"
    if ($NavisworksViewsImportXmlPath) {
        Write-Host "Navisworks viewpoints import XML (final model only): $NavisworksViewsImportXmlPath"
    }
    else {
        Write-Host "Navisworks viewpoints import XML (final model only): [not used]"
    }
    Write-Host "GUI mode: $GNWUI"
    Write-Host "Federated file name: $FederatedFileName"
    Write-Host "NWD naming method: $nwdNamingMethod"
    Write-Host "Include unmatched in final federated model: $(if ($includeUnmatchedFilesInFederatedModel) { 'Yes' } else { 'No' })"

    $groupingMethod = $ConfigSettings | Where-Object { $_.Parameter -eq 'FederationGroupingMethod' } | Select-Object -ExpandProperty Value -First 1
    $useWildcardSelection = -not [string]::IsNullOrWhiteSpace($groupingMethod) -and $groupingMethod.ToString().Trim().ToLowerInvariant() -eq 'wildcard selection'
    if ($useWildcardSelection) {
        $wildcardRows = @($WildcardSelectionRows | Where-Object { $_ })
        if ($wildcardRows.Count -eq 0) { throw 'Wildcard Selection is active but no wildcard rules were supplied.' }
        Write-Host 'Federation grouping method: Wildcard Selection' -ForegroundColor Cyan
        $ruleSummary = @()
        $createdOutputNames = @{}
        $lastCreatedModelPath = $null
        $federatableExtensions = if (Get-Command Get-FederatableModelExtensions -ErrorAction SilentlyContinue) { Get-FederatableModelExtensions } else { @('.ifc', '.dwg', '.dgn', '.rvt') }
        for ($ruleIndex = 0; $ruleIndex -lt $wildcardRows.Count; $ruleIndex++) {
            $rule = $wildcardRows[$ruleIndex]
            $inclusionText = [string](Get-RowValue -Row $rule -Candidates @('Inclusions','Inclusion'))
            $exclusionText = [string](Get-RowValue -Row $rule -Candidates @('Exclusions','Exclusion'))
            $inclusions = @($inclusionText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            $exclusions = @($exclusionText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            $outputName = ([string](Get-RowValue -Row $rule -Candidates @('ExportFileName','Export File Name'))).Trim()
            $readFromOutput = ([string](Get-RowValue -Row $rule -Candidates @('ReadFromOutputFolder','Read From Output Folder'))).Trim().ToLowerInvariant() -in @('yes','y','true','1')
            if ($inclusions.Count -eq 0 -or [string]::IsNullOrWhiteSpace($outputName)) {
                $warning = "Wildcard rule $($ruleIndex + 1) skipped: Inclusions and ExportFileName are required."
                Write-Warning $warning; Write-Log -Message $warning -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
                continue
            }
            if ($outputName.EndsWith('.nwd', [StringComparison]::OrdinalIgnoreCase)) { $outputName = $outputName.Substring(0, $outputName.Length - 4) }
            if ([string]::IsNullOrWhiteSpace($outputName) -or $outputName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) { throw "Wildcard rule $($ruleIndex + 1) has an invalid ExportFileName." }
            if ($createdOutputNames.ContainsKey($outputName.ToLowerInvariant())) { throw "Wildcard rules cannot use duplicate ExportFileName values: '$outputName'." }
            $createdOutputNames[$outputName.ToLowerInvariant()] = $true
            $sourcePath = if ($readFromOutput) { $OutputPath } else { $InputFolderPath }
            $targetPath = Join-Path $OutputPath ("$outputName.nwd")
            $candidates = @(Get-ChildItem -LiteralPath $sourcePath -File -ErrorAction Stop | Where-Object {
                ($readFromOutput -or $federatableExtensions -contains $_.Extension.ToLowerInvariant()) -and
                $_.FullName -ne $targetPath
            })
            $matches = @($candidates | Where-Object {
                $name = $_.Name
                ($inclusions | Where-Object { $name -like $_ }).Count -gt 0 -and
                ($exclusions | Where-Object { $name -like $_ }).Count -eq 0
            } | Select-Object -ExpandProperty FullName)
            if ($matches.Count -eq 0) {
                $warning = "Wildcard rule $($ruleIndex + 1) '$outputName' matched no files in '$sourcePath'; skipped."
                Write-Warning $warning; Write-Log -Message $warning -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
                $ruleSummary += [pscustomobject]@{ Rule=$ruleIndex + 1; ExportFileName="$outputName.nwd"; SourceFolder=$sourcePath; MatchedFiles=@(); Status='Skipped - no matches' }
                continue
            }
            $result = CreateNwd -InputFiles ([pscustomobject]@{ FullName=$outputName; Children=$matches }) -OutputPath $OutputPath -NavisworksVersion $NavisworksVersion -NavisworksOptionsPath $NavisworksOptions -NavisworksOptions $NavisworksOptions -BasePath $basePath
            if (-not $result.Model -or -not (Test-Path -LiteralPath $result.Model -PathType Leaf)) { throw "Wildcard rule $($ruleIndex + 1) failed to create '$outputName.nwd'." }
            $lastCreatedModelPath = $result.Model
            $ruleSummary += [pscustomobject]@{ Rule=$ruleIndex + 1; ExportFileName="$outputName.nwd"; SourceFolder=$sourcePath; MatchedFiles=$matches; Status='Created' }
        }
        if (-not $lastCreatedModelPath) { throw 'Wildcard Selection created no NWD outputs.' }
        [ordered]@{ Timestamp=(Get-Date).ToUniversalTime().ToString('o'); GroupingMethod='Wildcard Selection'; InputFolderPath=$InputFolderPath; OutputFolderPath=$OutputPath; WildcardRules=$ruleSummary; LastCreatedModelPath=$lastCreatedModelPath } | ConvertTo-Json -Depth 10 | Set-Content -Path (Join-Path $OutputPath 'federation-summary.json') -Encoding UTF8
        if ($TranscriptStarted) { Stop-Transcript | Out-Null; $global:MainTranscriptActive = $false }
        return
    }

    <#
    Normalizes lookup values by trimming and removing non-breaking spaces.
    #>
    function Normalize-LookupText {
        param([string]$Text)
        if ($null -eq $Text) { return $null }
        $clean = $Text -replace '\u00a0', ''
        return $clean.Trim()
    }

    # Optional lookup mapping: filename part + code -> descriptive label.
    # Only Filename-Part, Code, and Description are read from the Lookups range.
    $lookupSheet = @($LookupRows)
    $lookupMap = @{}
    if ($lookupSheet) {
        $hasPart = ($lookupSheet | Get-Member -MemberType NoteProperty | Where-Object Name -eq 'Filename-Part')
        $hasCode = ($lookupSheet | Get-Member -MemberType NoteProperty | Where-Object Name -eq 'Code')
        $hasDesc = ($lookupSheet | Get-Member -MemberType NoteProperty | Where-Object Name -eq 'Description')
        if ($hasPart -and $hasCode -and $hasDesc) {
            foreach ($row in $lookupSheet) {
                $partKey = Normalize-LookupText -Text (Get-RowValue -Row $row -Candidates @('Filename-Part','FilenamePart','Filename_Part'))
                $code = Normalize-LookupText -Text $row.Code
                $desc = Normalize-LookupText -Text $row.Description
                if ([string]::IsNullOrWhiteSpace($partKey) -or [string]::IsNullOrWhiteSpace($code) -or [string]::IsNullOrWhiteSpace($desc)) { continue }
                $mapKey = ("{0}||{1}" -f $partKey, $code)
                if (-not $lookupMap.ContainsKey($mapKey)) { $lookupMap[$mapKey] = $desc }
            }
        }
    }

    # Federation sheet defines which filename parts form grouping keys and order.
    # A Description column may also exist for user reference, but it is not used by current logic.
    $rawConfig = @($FederationRows)
    if (-not $rawConfig -or $rawConfig.Count -eq 0) {
        throw "Federation named range data was not supplied to Invoke-GroupFilesForFedProcess."
    }
    $partDefinitionsRaw = @(Get-FederationPartDefinitions -FederationRows $rawConfig)
    if ($partDefinitionsRaw.Count -eq 0) {
        throw "Federation range does not contain usable filename-part definitions."
    }
    $filePartsByName = [ordered]@{}
    foreach ($definition in $partDefinitionsRaw) {
        $filePartsByName[$definition.Name] = [pscustomobject]@{
            Name      = $definition.Name
            Position  = [int]$definition.Position
            SortPosition = [int]$definition.SortPosition
            PartKey   = $definition.FileNamePart
            GroupOrder = [int]$definition.GroupOrder
            IsFileExtension = [bool]$definition.IsFileExtension
            RowOrder  = [int]$definition.RowOrder
        }
    }
    $partDefinitions = @($filePartsByName.GetEnumerator() | ForEach-Object { $_.Value } | Sort-Object SortPosition, RowOrder)
    $matchDefinitions = @($partDefinitions | Where-Object { $_.GroupOrder -gt 0 -and -not $_.IsFileExtension })
    $GroupBy = @(
        $partDefinitions |
            Where-Object { $_.GroupOrder -gt 0 } |
            Sort-Object GroupOrder, SortPosition, RowOrder |
            Select-Object -ExpandProperty Name -Unique
    )
    if ($GroupBy.Count -eq 0) {
        throw "Federation range does not contain any grouping fields with GroupOrder > 0."
    }
    Write-Host "Grouping fields:" ($GroupBy -join ", ") -ForegroundColor Cyan
    # Scan supported model files and build grouping rows with decoded filename parts.
    $federatableExtensions = if (Get-Command Get-FederatableModelExtensions -ErrorAction SilentlyContinue) {
        Get-FederatableModelExtensions
    }
    else {
        @('.ifc', '.dwg', '.dgn', '.rvt')
    }
    $modelCandidates = @(Get-ChildItem -Path $InputFolderPath -File -ErrorAction Stop | Where-Object {
        $federatableExtensions -contains $_.Extension.ToLowerInvariant()
    })
    if ($modelCandidates.Count -eq 0) {
        $message = "No supported federation files found in input folder '$InputFolderPath'."
        Write-Log -Message $message -Color 'Red' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
        throw $message
    }
    $rows = @()
    $unmatchedFiles = @()
    foreach ($f in $modelCandidates) {
        if (-not (Test-FederationFileNameMatch -FileName $f.Name -PartDefinitions $matchDefinitions)) {
            $unmatchedFiles += $f.FullName
            continue
        }

        $parts = @(Split-ModelFileNameParts -FileName $f.Name)
        $ht = [ordered]@{}
        $ht['FileName'] = $f.FullName
        foreach ($definition in $partDefinitions) {
            $code = Get-FederationPartValue -FileName $f.Name -Definition $definition -NameParts $parts
            $mapped = $null
            if (-not [string]::IsNullOrWhiteSpace($code) -and $definition.PartKey) {
                $lookupKey = ("{0}||{1}" -f $definition.PartKey, $code)
                if ($lookupMap.ContainsKey($lookupKey)) {
                    $mapped = $lookupMap[$lookupKey]
                }
            }
            $ht[$definition.Name] = if ([string]::IsNullOrWhiteSpace($code)) { '' } else { $code }
            $ht["__Desc_$($definition.Name)"] = $mapped
        }
        $rows += [pscustomobject]$ht
    }
    Write-Host ("Files matched Federation naming rules: {0}" -f $rows.Count) -ForegroundColor Green
    Write-Host ("Files not matched and moved to separate NWD: {0}" -f $unmatchedFiles.Count) -ForegroundColor Yellow

    $TopOutputs = @()
    if ($rows.Count -gt 0) {
        $GroupedStructure = BuildGroupHierarchy -Data $rows -GroupKeys $GroupBy -NamingMethod $nwdNamingMethod
        $groupNodeNames = @(Get-GroupHierarchyNodeNames -Groups $GroupedStructure)
        $duplicateNodeNames = @($groupNodeNames | Group-Object | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name)
        if ($duplicateNodeNames.Count -gt 0) {
            $sampleNames = @($duplicateNodeNames | Select-Object -First 10)
            $message = ("NWD naming collision detected for NWDNamingMethod '{0}'. Duplicate grouped NWD name(s): {1}. Use 'Full' or adjust lookup descriptions to make names unique." -f $nwdNamingMethod, ($sampleNames -join '; '))
            Write-Log -Message $message -Color 'Red' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
            throw $message
        }
        $TopOutputs = foreach ($top in $GroupedStructure) {
            FederateGroup -Group $top -NavisworksOptionsPath $NavisworksOptions -NavisworksOptions $NavisworksOptions -BasePath $basePath
        }
        Write-Host "Top-level federated outputs:" -ForegroundColor Green
        $TopOutputs | ForEach-Object { Write-Host "  $_" }
    }
    else {
        $message = ("Federation naming rules matched none of the {0} supported file(s) in '{1}'. Building only the unmatched-files NWD." -f $modelCandidates.Count, $InputFolderPath)
        Write-Host $message -ForegroundColor Yellow
        Write-Log -Message $message -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
    }

    $topNwds = @($TopOutputs | Where-Object { $_ -and (Test-Path $_) })
    $unmatchedModelPath = $null
    if ($unmatchedFiles.Count -gt 0) {
        $unmatchedGroup = [pscustomobject]@{
            FullName = ("{0} - Unmatched Files" -f $FederatedFileName)
            Children = $unmatchedFiles
        }
        $unmatchedRes = CreateNwd -InputFiles $unmatchedGroup -NavisworksOptionsPath $NavisworksOptions -NavisworksOptions $NavisworksOptions -BasePath $basePath
        $unmatchedModelPath = if ($unmatchedRes) { $unmatchedRes.Model } else { $null }
        if ([string]::IsNullOrWhiteSpace($unmatchedModelPath) -or -not (Test-Path $unmatchedModelPath -PathType Leaf)) {
            $message = "Federation failed: unmatched-files NWD was not created."
            Write-Log -Message $message -Color 'Red' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
            throw $message
        }
        Write-Host "Unmatched files model:" -ForegroundColor Green
        Write-Host "  $unmatchedModelPath"
    }

    $finalInputModels = @()
    if ($topNwds.Count -gt 0) { $finalInputModels += $topNwds }
    if ($includeUnmatchedFilesInFederatedModel -and $unmatchedModelPath) {
        $finalInputModels += $unmatchedModelPath
    }

    $finalModelPath = $null
    if ($finalInputModels.Count -gt 0) {
        $finalGroup = [pscustomobject]@{
            FullName = $FederatedFileName
            Children = $finalInputModels
        }
        $finalRes = CreateNwd -InputFiles $finalGroup -NavisworksOptionsPath $NavisworksOptions -NavisworksOptions $NavisworksOptions -NavisworksViewsImportXmlPath $NavisworksViewsImportXmlPath -BasePath $basePath
        $finalModelPath = if ($finalRes) { $finalRes.Model } else { $null }
        if ([string]::IsNullOrWhiteSpace($finalModelPath) -or -not (Test-Path $finalModelPath -PathType Leaf)) {
            $message = ("Federation failed: final federated model was not created. Expected output path: '{0}'." -f $finalModelPath)
            Write-Log -Message $message -Color 'Red' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
            throw $message
        }
        Write-Host "Final federated model:" -ForegroundColor Green
        Write-Host "  $finalModelPath"
    }
    elseif ($unmatchedModelPath) {
        $message = "Final federated model was not created because only unmatched files were available and IncludeUnmatchedFilesInFederatedModel is disabled."
        Write-Host $message -ForegroundColor Yellow
        Write-Log -Message $message -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
    }
    else {
        $message = ("Federation failed: no top-level NWD outputs were created from input folder '{0}'." -f $InputFolderPath)
        Write-Log -Message $message -Color 'Red' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
        throw $message
    }

    $summary = [ordered]@{
        Timestamp                              = (Get-Date).ToUniversalTime().ToString("o")
        InputFolderPath                        = $InputFolderPath
        OutputFolderPath                       = $OutputPath
        MatchedInputCount                      = $rows.Count
        MatchedInputFiles                      = @($rows | Select-Object -ExpandProperty FileName)
        UnmatchedInputCount                    = $unmatchedFiles.Count
        UnmatchedInputFiles                    = @($unmatchedFiles)
        UnmatchedModelPath                     = $unmatchedModelPath
        IncludeUnmatchedFilesInFederatedModel  = $includeUnmatchedFilesInFederatedModel
        FinalModelPath                         = $finalModelPath
        FinalModelCreated                      = -not [string]::IsNullOrWhiteSpace($finalModelPath)
    }
    $summaryPath = Join-Path $OutputPath 'federation-summary.json'
    $summary | ConvertTo-Json -Depth 10 | Set-Content -Path $summaryPath -Encoding UTF8
    $ProcessEndTime = Get-Date
    $TotalDuration = New-TimeSpan -Start $ProcessStartTime -End $ProcessEndTime
    Write-Host "Process finish time: $ProcessEndTime"
    Write-Host "Total duration: $([math]::Round($TotalDuration.TotalMinutes,2)) minutes"
    Write-Host "=== END 050-GroupFilesForFedProcess ==="
    if ($TranscriptStarted) {
        Stop-Transcript | Out-Null
        $global:MainTranscriptActive = $false
    }
}

function Invoke-PublishRevizto {
<#
Runs the Revizto publish schedule based on settings.
#>
    [CmdletBinding()]
    param(
        [string]$ConfigFile = "Config.xlsx",
        [array]$Settings = @(),
        [string]$LogsFolder
    )

    $basePath = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
    elseif ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    else {
        Add-Type -AssemblyName System.Reflection
        Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location)
    }

    # Pull publish code and log folder from Settings.
    $settings = @($Settings)
    if (-not $settings -or $settings.Count -eq 0) {
        throw "Settings were not supplied to Invoke-PublishRevizto."
    }
    $reviztoCode = $settings | Where-Object { $_.Parameter -eq "ReviztoPublishCode" } | Select-Object -ExpandProperty Value -First 1
    $reviztoCode = if ($null -ne $reviztoCode) { $reviztoCode.ToString().Trim() } else { $null }
    $logFolderValue = if ($LogsFolder) { $LogsFolder } else {
        $settings | Where-Object { $_.Parameter -eq "LogFolder" } | Select-Object -ExpandProperty Value
    }
    if (Get-Command CureFolderPath -ErrorAction SilentlyContinue) {
        $logFolderValue = CureFolderPath $logFolderValue
    }
    if (-not $logFolderValue) { $logFolderValue = Join-Path $basePath 'Logs' }
    try {
        if (-not (Test-Path $logFolderValue)) { New-Item -Path $logFolderValue -ItemType Directory -Force | Out-Null }
    }
    catch {
        Write-Warning "Could not create or access log folder '$logFolderValue'."
    }

    $logInfo = Get-LogFilePath -LogsFolder $logFolderValue -BaseName 'PW_Process.log.txt'
    $LogFilePath = $logInfo.Path
    $TranscriptStarted = $false
    if (-not $global:MainTranscriptActive) {
        try {
            Initialize-LogFileTarget -LogInfo $logInfo
            Start-Transcript -Path $LogFilePath -Append:$($logInfo.Append) -ErrorAction Stop | Out-Null
            $TranscriptStarted = $true
            $global:MainTranscriptActive = $true
        }
        catch {
            Write-Warning "Transcript could not be started. Logging will be console-only. Error: $_"
        }
    }
    Write-Host "Logging to $LogFilePath"
    Write-Host "=== START 060-PublishRevizto ==="

    $logLine = $null
    if ([string]::IsNullOrWhiteSpace($reviztoCode)) {
        $logLine = "Revizto publish skipped: ReviztoPublishCode is empty."
        Write-Host $logLine -ForegroundColor Yellow
    }
    else {
        $reviztoExe = "C:\\Program Files\\Revizto SA\\Revizto5\\Service\\ReviztoConsole.exe"
        if (-not (Test-Path $reviztoExe -PathType Leaf)) {
            $logLine = "Revizto publish skipped: ReviztoConsole.exe not found at '$reviztoExe'."
            Write-Warning $logLine
        }
        else {
            Write-Host "Running Revizto publish schedule..." -ForegroundColor Cyan
            & $reviztoExe scheduler run $reviztoCode | Out-Null
            $exitCode = $LASTEXITCODE
            if ($exitCode -ne 0) {
                $logLine = "Revizto publish failed (exit code $exitCode)."
                Write-Warning $logLine
            }
            else {
                $logLine = "Revizto publish completed successfully."
                Write-Host $logLine -ForegroundColor Green
            }
        }
    }

    Write-Host "=== END 060-PublishRevizto ==="
    if ($TranscriptStarted) {
        Stop-Transcript | Out-Null
        $global:MainTranscriptActive = $false
    }

    if ($logLine) {
        Write-Log -Message $logLine -Color '' -Settings $settings -LogsFolderOverride $LogsFolder -BasePath $basePath -NoConsole
    }
}

