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

function Resolve-NavisworksOutputFormat {
<#
Normalizes the requested Navisworks output format.
#>
    [CmdletBinding()]
    param([string]$Value)

    $text = Normalize-ExcelText -Text $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return 'NWD' }
    switch ($text.ToLowerInvariant()) {
        'nwf' { return 'NWF' }
        'nwd' { return 'NWD' }
        default { return 'NWD' }
    }
}

function Get-NavisworksOutputExtension {
    [CmdletBinding()]
    param([string]$OutputFormat)

    if ((Resolve-NavisworksOutputFormat -Value $OutputFormat) -eq 'NWF') { return '.nwf' }
    return '.nwd'
}

function Resolve-NavisworksOutputBaseName {
    [CmdletBinding()]
    param([string]$Name)

    $baseName = Normalize-ExcelText -Text $Name
    if ([string]::IsNullOrWhiteSpace($baseName)) { return $baseName }
    if ($baseName.EndsWith('.nwd', [System.StringComparison]::OrdinalIgnoreCase) -or
        $baseName.EndsWith('.nwf', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $baseName.Substring(0, $baseName.Length - 4)
    }
    return $baseName
}

function Test-FinalNavisworksFileNameIsNwf {
<#
Only the final federated model may switch to NWF, and only when the configured
final filename explicitly ends with .nwf.
#>
    [CmdletBinding()]
    param([string]$Name)

    $text = Normalize-ExcelText -Text $Name
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    return $text.EndsWith('.nwf', [System.StringComparison]::OrdinalIgnoreCase)
}

function Resolve-NavisworksOutputFormatFromFileName {
<#
Returns NWF only when a configured output filename explicitly ends with .nwf.
All other values default to NWD, including blank names and names without an
extension.
#>
    [CmdletBinding()]
    param([string]$Name)

    $text = Normalize-ExcelText -Text $Name
    if (-not [string]::IsNullOrWhiteSpace($text) -and $text.EndsWith('.nwf', [System.StringComparison]::OrdinalIgnoreCase)) {
        return 'NWF'
    }
    return 'NWD'
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

function Get-FederationCandidateFiles {
<#
Returns supported federation source files and skips Navisworks cache files that
were generated from source model files already present in the same folder.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string[]]$Extensions = @('.ifc', '.dwg', '.dgn', '.rvt', '.nwc'),
        [switch]$KeepGeneratedNwc
    )

    $allFiles = @(Get-ChildItem -LiteralPath $Path -File -ErrorAction Stop)
    $normalizedExtensions = @($Extensions | ForEach-Object { $_.ToLowerInvariant() })
    $candidates = @($allFiles | Where-Object { $normalizedExtensions -contains $_.Extension.ToLowerInvariant() })
    if ($KeepGeneratedNwc) { return $candidates }

    $sourceExtensions = @('.ifc', '.dwg', '.dgn', '.rvt')
    $generatedNwcNames = @{}
    foreach ($source in @($allFiles | Where-Object { $sourceExtensions -contains $_.Extension.ToLowerInvariant() })) {
        $generatedNwcNames["$($source.BaseName).nwc".ToLowerInvariant()] = $true
        $generatedNwcNames["$($source.Name).nwc".ToLowerInvariant()] = $true
    }

    $filtered = @()
    foreach ($candidate in $candidates) {
        if ($candidate.Extension.ToLowerInvariant() -eq '.nwc' -and $generatedNwcNames.ContainsKey($candidate.Name.ToLowerInvariant())) {
            continue
        }
        $filtered += $candidate
    }
    return $filtered
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

function Resolve-NavisworksOptionsPath {
    [CmdletBinding()]
    param(
        [string]$Value,
        [string]$BasePath
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $candidate = $Value
    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        $root = if ($BasePath) { $BasePath } else { $PSScriptRoot }
        $candidate = Join-Path $root $candidate
    }
    if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }

    $root = if ($BasePath) { $BasePath } else { $PSScriptRoot }
    $resolved = Resolve-SupportFile -FileName $Value -BasePath $root
    if ($resolved -and (Test-Path -LiteralPath $resolved -PathType Leaf)) { return $resolved }
    return $null
}

function Resolve-NavisworksSavedNwdVersion {
<#
Maps the GUI save-version choice to Navisworks' hidden NWD XML enum.
For Navisworks 2027 this is: 0=2027, 1=2026, 2=2016-2025.
#>
    [CmdletBinding()]
    param(
        [string]$Value,
        [string]$NavisworksVersion
    )

    $requested = Normalize-ExcelText -Text $Value
    if ([string]::IsNullOrWhiteSpace($requested)) { $requested = 'Latest' }
    $navVersionNumber = 0
    [void][int]::TryParse((Normalize-ExcelText -Text $NavisworksVersion), [ref]$navVersionNumber)
    $latestLabel = if ($navVersionNumber -gt 0) { $navVersionNumber.ToString() } else { 'Latest' }

    if ($requested -eq 'Latest') {
        return [pscustomobject]@{ Requested='Latest'; DisplayValue=$latestLabel; InternalValue=0; IsLatest=$true }
    }
    if ($requested -eq '2016-2025') {
        $bucketValue = if ($navVersionNumber -ge 2027) { 2 } else { 1 }
        return [pscustomobject]@{ Requested=$requested; DisplayValue='2016-2025'; InternalValue=$bucketValue; IsLatest=$false }
    }

    $requestedNumber = 0
    if (-not [int]::TryParse($requested, [ref]$requestedNumber)) {
        throw "Invalid NavisworksSavedNwdVersion '$requested'. Use Latest, 2027, 2026, or 2016-2025."
    }
    if ($navVersionNumber -gt 0 -and $requestedNumber -gt $navVersionNumber) {
        throw "NavisworksSavedNwdVersion '$requested' is newer than Navisworks Manage $NavisworksVersion and cannot be saved by this installation."
    }
    $enumValue = 0
    if ($navVersionNumber -gt 0) {
        $enumValue = $navVersionNumber - $requestedNumber
        if ($enumValue -lt 0 -or $enumValue -gt 1) {
            throw "NavisworksSavedNwdVersion '$requested' is not available from Navisworks Manage $NavisworksVersion. Use Latest, $NavisworksVersion, $($navVersionNumber - 1), or 2016-2025."
        }
    }
    return [pscustomobject]@{ Requested=$requested; DisplayValue=$requested; InternalValue=$enumValue; IsLatest=$false }
}

function New-NavisworksTemporaryOptionsPath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        throw "Navisworks options XML not found: '$SourcePath'. Choose 'Latest' or provide an XML file so the hidden NWD save-version option can be patched."
    }

    $tempRoot = Join-Path $env:TEMP 'FederationAutomation'
    if (-not (Test-Path -LiteralPath $tempRoot -PathType Container)) {
        New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null
    }
    $tempPath = Join-Path $tempRoot ("NavisworksOptions-{0}.xml" -f ([guid]::NewGuid().ToString('N')))
    Copy-Item -LiteralPath $SourcePath -Destination $tempPath -Force
    return $tempPath
}

function Set-NavisworksOptionsNwdVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OptionsPath,
        [Parameter(Mandatory = $true)]$VersionOption
    )

    if (-not (Test-Path -LiteralPath $OptionsPath -PathType Leaf)) {
        throw "Navisworks options XML not found: '$OptionsPath'. Choose 'Latest' or provide an XML file so the hidden NWD save-version option can be patched."
    }

    try { [xml]$doc = Get-Content -LiteralPath $OptionsPath -Raw -ErrorAction Stop }
    catch { throw "Navisworks options XML could not be loaded: '$OptionsPath'. $($_.Exception.Message)" }

    $node = $doc.SelectSingleNode('//optionset[@name="nwd"]/option[@name="default_format"]/data/name')
    if (-not $node) {
        throw "Navisworks options XML '$OptionsPath' does not contain the hidden NWD default_format option."
    }

    $node.SetAttribute('internal', ("lcop_nwd_version:{0}" -f $VersionOption.InternalValue))
    $node.InnerText = $VersionOption.DisplayValue
    $doc.Save($OptionsPath)
    return $OptionsPath
}

function Remove-NavisworksTemporaryOptionsPath {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    try {
        $tempRoot = [System.IO.Path]::GetFullPath((Join-Path $env:TEMP 'FederationAutomation')).TrimEnd('\')
        $target = [System.IO.Path]::GetFullPath($Path)
        if ($target.StartsWith($tempRoot + '\', [System.StringComparison]::OrdinalIgnoreCase) -and
            (Test-Path -LiteralPath $target -PathType Leaf)) {
            Remove-Item -LiteralPath $target -Force -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Could not remove temporary Navisworks options XML '$Path'. $_"
    }
}

function Copy-WildcardOutputsToDestination {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][array]$CopyRequests,
        [Parameter(Mandatory = $true)][string]$DestinationFolder,
        [Parameter(Mandatory = $true)][array]$Settings,
        [string]$LogsFolder,
        [string]$BasePath
    )

    if ($CopyRequests.Count -eq 0) { return @() }
    if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
        throw "DestinationFolder is required because at least one wildcard rule has CopyToDestination enabled."
    }

    try {
        if (-not (Test-Path -LiteralPath $DestinationFolder -PathType Container)) {
            New-Item -Path $DestinationFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        $probePath = Join-Path $DestinationFolder ('.federation-destination-write-test-{0}.tmp' -f ([guid]::NewGuid().ToString('N')))
        [IO.File]::WriteAllText($probePath, 'test')
        Remove-Item -LiteralPath $probePath -Force -ErrorAction Stop
    }
    catch {
        $message = "Destination folder '$DestinationFolder' is not writable or could not be created. It may be locked or unavailable. $($_.Exception.Message)"
        Write-Log -Message $message -Color 'Red' -Settings $Settings -LogsFolderOverride $LogsFolder -BasePath $BasePath
        throw $message
    }

    $copyResults = @()
    foreach ($request in $CopyRequests) {
        $sourcePath = $request.SourcePath
        $destinationPath = Join-Path $DestinationFolder ([IO.Path]::GetFileName($sourcePath))
        try {
            Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force -ErrorAction Stop
            $message = "Copied wildcard output to destination: '$destinationPath'"
            Write-Host $message -ForegroundColor Green
            Write-Log -Message $message -Color 'Green' -Settings $Settings -LogsFolderOverride $LogsFolder -BasePath $BasePath -NoConsole
            $copyResults += [pscustomobject]@{ Rule=$request.Rule; SourcePath=$sourcePath; DestinationPath=$destinationPath; Status='Copied' }
        }
        catch {
            $message = "Failed to copy wildcard output '$sourcePath' to destination '$destinationPath'. The destination file or folder may be locked. $($_.Exception.Message)"
            Write-Log -Message $message -Color 'Red' -Settings $Settings -LogsFolderOverride $LogsFolder -BasePath $BasePath
            throw $message
        }
    }
    return @($copyResults)
}

if (-not (Get-Variable -Name PWAttributeDefinitions -Scope Script -ErrorAction SilentlyContinue)) {
    $script:PWAttributeDefinitions = @()
}

function Write-FederationNwdProgress {
    param(
        [string]$Status,
        [string]$Name
    )
    $total = if ($script:FEDAUTONwdProgressTotal) { [int]$script:FEDAUTONwdProgressTotal } else { 0 }
    if ($total -le 0) { return }
    $current = if ($script:FEDAUTONwdProgressCurrent) { [int]$script:FEDAUTONwdProgressCurrent } else { 0 }
    $safeStatus = if ([string]::IsNullOrWhiteSpace($Status)) { 'Processing' } else { $Status }
    $safeName = if ($Name) { $Name -replace '\|', ' ' } else { '' }
    Write-Host ("FEDERATION_NWD_PROGRESS|Current={0}|Total={1}|Status={2}|Name={3}" -f $current, $total, $safeStatus, $safeName)
}

function Complete-FederationNwdProgress {
    param(
        [string]$Status,
        [string]$Name
    )
    $total = if ($script:FEDAUTONwdProgressTotal) { [int]$script:FEDAUTONwdProgressTotal } else { 0 }
    if ($total -le 0) { return }
    $current = if ($script:FEDAUTONwdProgressCurrent) { [int]$script:FEDAUTONwdProgressCurrent } else { 0 }
    $script:FEDAUTONwdProgressCurrent = [Math]::Min($total, $current + 1)
    Write-FederationNwdProgress -Status $Status -Name $Name
}

function CreateNwd {
<#
Creates a Navisworks model from a group node.
- Leaf groups: Children are source model paths; parents: Children are child Navisworks outputs.
- Uses the group's FullName for output name (sanitized in caller where used).
#>
    param(
        $InputFiles,
        $OutputPath = $OutputPath,
        $NavisworksVersion = $NavisworksVersion,
        $NavisworksOptions = $NavisworksOptions,
        $NavisworksOptionsPath = $NavisworksOptionsPath,
        $NavisworksViewsImportXmlPath = $NavisworksViewsImportXmlPath,
        $OutputFormat = 'NWD',
        $BasePath = $BasePath
    )
    # Basic runtime info for console/trace logs.
    Write-Host "Processing $($InputFiles.FullName)... Do not close the window."
    Write-FederationNwdProgress -Status 'Creating' -Name $InputFiles.FullName
    $start = Get-Date

    # Output name is derived from the grouping hierarchy and sanitized for filesystem.
    $safeName      = Get-SafeFileName -Name $InputFiles.FullName
    $resolvedOutputFormat = Resolve-NavisworksOutputFormat -Value $OutputFormat
    $outputExtension = Get-NavisworksOutputExtension -OutputFormat $resolvedOutputFormat
    $outputModel   = $OutputPath + "\" + $safeName + $outputExtension
    $navisworksDir = Resolve-NavisworksInstallPath -Version $NavisworksVersion
    $roamerExe     = if ($navisworksDir) { Join-Path $navisworksDir 'Roamer.exe' } else { $null }

    # Build arguments for Roamer.exe to append all group inputs.
    $arg = @()
    foreach ($Inputfile in $InputFiles.Children) {
        $arg += '-AppendFile', $Inputfile
    }
    $arg += '-SaveFile', $outputModel
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
        $optionsPath = Resolve-NavisworksOptionsPath -Value $optionsPath -BasePath $BasePath
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
Navisworks $resolvedOutputFormat processed for $outputModel
Number of model(s) federated: $($InputFiles.Children.Count)
Navisworks Version: $NavisworksVersion
End Time: $end
Time spent: $minutes minutes
Navisworks output path: $outputModel
Model(s) federated: $([string]::Join(', ', $InputFiles.Children))
NW Arguments: $argString
"@
    Write-Host $logTxt
    $nwdProgressStatus = if ($InputFiles.Children.Count -gt 0) { 'Created' } else { 'Skipped' }
    Complete-FederationNwdProgress -Status $nwdProgressStatus -Name $InputFiles.FullName

    # Return a summary object for further use
    [pscustomobject]@{
        Model             = $outputModel
        FederatedFiles    = $InputFiles.Children
        OutputPath        = $OutputPath
        OutputFormat      = $resolvedOutputFormat
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
        [string]$OutputFormat = 'NWD',
        [string]$BasePath
    )
    # Print the tree and federate bottom-up so child outputs exist for parents.
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
        $res = CreateNwd -InputFiles $groupForCreate -NavisworksOptionsPath $NavisworksOptionsPath -NavisworksOptions $NavisworksOptions -OutputFormat $OutputFormat -BasePath $BasePath
        return $res.Model
    }
    else {
        foreach ($file in $Group.Children) { Write-Host "$indent  $file" }
        $res = CreateNwd -InputFiles $Group -NavisworksOptionsPath $NavisworksOptionsPath -NavisworksOptions $NavisworksOptions -OutputFormat $OutputFormat -BasePath $BasePath
        return $res.Model
    }
}

function Invoke-GroupFilesForFedProcess {
<#
Groups IFCs and federates them into Navisworks outputs.
#>
    [CmdletBinding()]
    param(
        [string]$ConfigFile = "Config.json",
        [array]$Settings = @(),
        [Parameter(Mandatory = $true)][array]$FederationRows,
        [array]$WildcardSelectionRows = @(),
        [array]$LookupRows,
        [string]$LogsFolder,
        [switch]$UseFederationInputFolder
    )

    # Resolve script root to find Config.json and related assets.
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
    $TemporaryNavisworksOptionsPath = $null
    trap {
        Remove-NavisworksTemporaryOptionsPath -Path $TemporaryNavisworksOptionsPath
        throw
    }
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
    $DestinationFolder   = CureFolderPath ($ConfigSettings | Where-Object { $_.Parameter -eq "DestinationFolder" } | Select-Object -ExpandProperty Value)
    if ([string]::IsNullOrWhiteSpace($DestinationFolder)) { $DestinationFolder = "Destination" }
    $InputFolderPath = Resolve-RelativePath -PathValue $InputFolderPath -Root $basePath
    $OutputPath = Resolve-RelativePath -PathValue $OutputPath -Root $basePath
    $DestinationFolder = Resolve-RelativePath -PathValue $DestinationFolder -Root $basePath
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
    $NavisworksSavedNwdVersion = $ConfigSettings | Where-Object { $_.Parameter -eq "NavisworksSavedNwdVersion" } | Select-Object -ExpandProperty Value -First 1
    if ([string]::IsNullOrWhiteSpace($NavisworksSavedNwdVersion)) {
        $NavisworksSavedNwdVersion = "Latest"
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
    $finalOutputFormat = if (Test-FinalNavisworksFileNameIsNwf -Name $FederatedFileName) { 'NWF' } else { 'NWD' }
    $finalOutputExtension = Get-NavisworksOutputExtension -OutputFormat $finalOutputFormat
    $FederatedFileName = Resolve-NavisworksOutputBaseName -Name $FederatedFileName
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
    Write-Host "Destination folder: $DestinationFolder"
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
    $nwdSaveVersionOption = Resolve-NavisworksSavedNwdVersion -Value $NavisworksSavedNwdVersion -NavisworksVersion $NavisworksVersion
    $resolvedOptionsPath = Resolve-NavisworksOptionsPath -Value $NavisworksOptions -BasePath $basePath
    if (-not $resolvedOptionsPath -and -not $nwdSaveVersionOption.IsLatest) {
        $reason = "NavisworksSavedNwdVersion '$($nwdSaveVersionOption.Requested)' requires a Navisworks options XML, but '$NavisworksOptions' was not found."
        Write-Host ("Federation skipped: {0}" -f $reason) -ForegroundColor Yellow
        Write-Log -Message ("Federation skipped: {0}" -f $reason) -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
        if ($TranscriptStarted) {
            Stop-Transcript | Out-Null
            $global:MainTranscriptActive = $false
        }
        return
    }
    if ($resolvedOptionsPath) {
        $TemporaryNavisworksOptionsPath = New-NavisworksTemporaryOptionsPath -SourcePath $resolvedOptionsPath
        $NavisworksOptions = Set-NavisworksOptionsNwdVersion -OptionsPath $TemporaryNavisworksOptionsPath -VersionOption $nwdSaveVersionOption
    }
    Write-Host "Navisworks options: $NavisworksOptions"
    if ($resolvedOptionsPath -and $TemporaryNavisworksOptionsPath) {
        Write-Host "Original Navisworks options XML preserved: $resolvedOptionsPath"
        Write-Host "Temporary Navisworks options XML: $TemporaryNavisworksOptionsPath"
    }
    Write-Host "Saved NWD version: $($nwdSaveVersionOption.Requested) ($($nwdSaveVersionOption.DisplayValue))"
    if ($NavisworksViewsImportXmlPath) {
        Write-Host "Navisworks viewpoints import XML (final model only): $NavisworksViewsImportXmlPath"
    }
    else {
        Write-Host "Navisworks viewpoints import XML (final model only): [not used]"
    }
    Write-Host "GUI mode: $GNWUI"
    Write-Host "Final Navisworks output format: $finalOutputFormat"
    Write-Host "Federated file name: $FederatedFileName$finalOutputExtension"
    Write-Host "NWD naming method: $nwdNamingMethod"
    Write-Host "Include unmatched in final federated model: $(if ($includeUnmatchedFilesInFederatedModel) { 'Yes' } else { 'No' })"

    $groupingMethod = $ConfigSettings | Where-Object { $_.Parameter -eq 'FederationGroupingMethod' } | Select-Object -ExpandProperty Value -First 1
    $useWildcardSelection = -not [string]::IsNullOrWhiteSpace($groupingMethod) -and $groupingMethod.ToString().Trim().ToLowerInvariant() -eq 'wildcard selection'
    if ($useWildcardSelection) {
        function Test-WildcardSelectionRuleEnabled {
            param($Rule)
            if ($null -eq $Rule) { return $false }
            if (-not ($Rule.PSObject.Properties.Name -contains 'Run')) { return $true }
            $runText = [string](Get-RowValue -Row $Rule -Candidates @('Run','Enabled'))
            if ([string]::IsNullOrWhiteSpace($runText)) { return $false }
            return $runText.Trim().ToLowerInvariant() -in @('yes','y','true','1')
        }

        $allWildcardRows = @($WildcardSelectionRows | Where-Object { $_ })
        if ($allWildcardRows.Count -eq 0) { throw 'Wildcard Selection is active but no wildcard rules were supplied.' }
        $wildcardRows = @($allWildcardRows | Where-Object { Test-WildcardSelectionRuleEnabled -Rule $_ })
        if ($wildcardRows.Count -eq 0) { throw 'Wildcard Selection is active but no enabled wildcard rules were supplied. Set Run to Yes for at least one row.' }
        $disabledWildcardCount = $allWildcardRows.Count - $wildcardRows.Count
        Write-Host 'Federation grouping method: Wildcard Selection' -ForegroundColor Cyan
        if ($disabledWildcardCount -gt 0) {
            Write-Host ("Wildcard Selection: {0} disabled rule(s) skipped." -f $disabledWildcardCount) -ForegroundColor Yellow
        }
        $ruleSummary = @()
        $destinationCopyRequests = @()
        $createdOutputNames = @{}
        $lastCreatedModelPath = $null
        $script:FEDAUTONwdProgressTotal = $wildcardRows.Count
        $script:FEDAUTONwdProgressCurrent = 0
        Write-FederationNwdProgress -Status 'Queued' -Name 'Wildcard Selection'
        $federatableExtensions = if (Get-Command Get-FederatableModelExtensions -ErrorAction SilentlyContinue) { Get-FederatableModelExtensions } else { @('.ifc', '.dwg', '.dgn', '.rvt', '.nwc') }
        for ($ruleIndex = 0; $ruleIndex -lt $wildcardRows.Count; $ruleIndex++) {
            $rule = $wildcardRows[$ruleIndex]
            $inclusionText = [string](Get-RowValue -Row $rule -Candidates @('Inclusions','Inclusion'))
            $exclusionText = [string](Get-RowValue -Row $rule -Candidates @('Exclusions','Exclusion'))
            $inclusions = @($inclusionText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            $exclusions = @($exclusionText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            $outputName = ([string](Get-RowValue -Row $rule -Candidates @('ExportFileName','Export File Name'))).Trim()
            $readFromOutput = ([string](Get-RowValue -Row $rule -Candidates @('ReadFromOutputFolder','Read From Output Folder'))).Trim().ToLowerInvariant() -in @('yes','y','true','1')
            $copyToDestination = ([string](Get-RowValue -Row $rule -Candidates @('CopyToDestination','Copy To Destination'))).Trim().ToLowerInvariant() -in @('yes','y','true','1')
            if ($inclusions.Count -eq 0 -or [string]::IsNullOrWhiteSpace($outputName)) {
                $warning = "Wildcard rule $($ruleIndex + 1) skipped: Inclusions and ExportFileName are required."
                Write-Warning $warning; Write-Log -Message $warning -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
                Complete-FederationNwdProgress -Status 'Skipped' -Name ("Wildcard rule {0}" -f ($ruleIndex + 1))
                continue
            }
            $ruleOutputFormat = Resolve-NavisworksOutputFormatFromFileName -Name $outputName
            $ruleOutputExtension = Get-NavisworksOutputExtension -OutputFormat $ruleOutputFormat
            $outputName = Resolve-NavisworksOutputBaseName -Name $outputName
            if ([string]::IsNullOrWhiteSpace($outputName) -or $outputName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) { throw "Wildcard rule $($ruleIndex + 1) has an invalid ExportFileName." }
            if ($createdOutputNames.ContainsKey($outputName.ToLowerInvariant())) { throw "Wildcard rules cannot use duplicate ExportFileName values: '$outputName'." }
            $createdOutputNames[$outputName.ToLowerInvariant()] = $true
            $sourcePath = if ($readFromOutput) { $OutputPath } else { $InputFolderPath }
            $targetPath = Join-Path $OutputPath ("$outputName$ruleOutputExtension")
            $candidates = if ($readFromOutput) {
                @(Get-ChildItem -LiteralPath $sourcePath -File -ErrorAction Stop | Where-Object { $_.FullName -ne $targetPath })
            }
            else {
                @(Get-FederationCandidateFiles -Path $sourcePath -Extensions $federatableExtensions | Where-Object { $_.FullName -ne $targetPath })
            }
            $matches = @($candidates | Where-Object {
                $name = $_.Name
                ($inclusions | Where-Object { $name -like $_ }).Count -gt 0 -and
                ($exclusions | Where-Object { $name -like $_ }).Count -eq 0
            } | Select-Object -ExpandProperty FullName)
            if ($matches.Count -eq 0) {
                $warning = "Wildcard rule $($ruleIndex + 1) '$outputName' matched no files in '$sourcePath'; skipped."
                Write-Warning $warning; Write-Log -Message $warning -Color 'Yellow' -Settings $ConfigSettings -LogsFolderOverride $LogsFolder -BasePath $basePath
                $ruleSummary += [pscustomobject]@{ Rule=$ruleIndex + 1; ExportFileName="$outputName$ruleOutputExtension"; OutputFormat=$ruleOutputFormat; SourceFolder=$sourcePath; MatchedFiles=@(); CopyToDestination=$copyToDestination; Status='Skipped - no matches' }
                Complete-FederationNwdProgress -Status 'Skipped' -Name $outputName
                continue
            }
            $result = CreateNwd -InputFiles ([pscustomobject]@{ FullName=$outputName; Children=$matches }) -OutputPath $OutputPath -NavisworksVersion $NavisworksVersion -NavisworksOptionsPath $NavisworksOptions -NavisworksOptions $NavisworksOptions -OutputFormat $ruleOutputFormat -BasePath $basePath
            if (-not $result.Model -or -not (Test-Path -LiteralPath $result.Model -PathType Leaf)) { throw "Wildcard rule $($ruleIndex + 1) failed to create '$outputName$ruleOutputExtension'." }
            $lastCreatedModelPath = $result.Model
            if ($copyToDestination) {
                $destinationCopyRequests += [pscustomobject]@{ Rule=$ruleIndex + 1; SourcePath=$result.Model }
            }
            $ruleSummary += [pscustomobject]@{ Rule=$ruleIndex + 1; ExportFileName="$outputName$ruleOutputExtension"; OutputFormat=$ruleOutputFormat; SourceFolder=$sourcePath; MatchedFiles=$matches; CopyToDestination=$copyToDestination; Status='Created' }
        }
        if (-not $lastCreatedModelPath) { throw "Wildcard Selection created no Navisworks outputs." }
        $destinationCopies = @(Copy-WildcardOutputsToDestination -CopyRequests $destinationCopyRequests -DestinationFolder $DestinationFolder -Settings $ConfigSettings -LogsFolder $LogsFolder -BasePath $basePath)
        [ordered]@{ Timestamp=(Get-Date).ToUniversalTime().ToString('o'); GroupingMethod='Wildcard Selection'; InputFolderPath=$InputFolderPath; OutputFolderPath=$OutputPath; DestinationFolder=$DestinationFolder; DestinationCopies=$destinationCopies; WildcardRules=$ruleSummary; LastCreatedModelPath=$lastCreatedModelPath } | ConvertTo-Json -Depth 10 | Set-Content -Path (Join-Path $OutputPath 'federation-summary.json') -Encoding UTF8
        Remove-NavisworksTemporaryOptionsPath -Path $TemporaryNavisworksOptionsPath
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
        @('.ifc', '.dwg', '.dgn', '.rvt', '.nwc')
    }
    $modelCandidates = @(Get-FederationCandidateFiles -Path $InputFolderPath -Extensions $federatableExtensions)
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
        $script:FEDAUTONwdProgressTotal = $groupNodeNames.Count
        if ($unmatchedFiles.Count -gt 0) { $script:FEDAUTONwdProgressTotal++ }
        $script:FEDAUTONwdProgressTotal++
        $script:FEDAUTONwdProgressCurrent = 0
        Write-FederationNwdProgress -Status 'Queued' -Name 'Grouped federation'
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
        $script:FEDAUTONwdProgressTotal = 0
        if ($unmatchedFiles.Count -gt 0) { $script:FEDAUTONwdProgressTotal++ }
        if ($unmatchedFiles.Count -gt 0 -and $includeUnmatchedFilesInFederatedModel) { $script:FEDAUTONwdProgressTotal++ }
        $script:FEDAUTONwdProgressCurrent = 0
        Write-FederationNwdProgress -Status 'Queued' -Name 'Unmatched federation'
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
        $finalRes = CreateNwd -InputFiles $finalGroup -NavisworksOptionsPath $NavisworksOptions -NavisworksOptions $NavisworksOptions -NavisworksViewsImportXmlPath $NavisworksViewsImportXmlPath -OutputFormat $finalOutputFormat -BasePath $basePath
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
        FinalOutputFormat                      = $finalOutputFormat
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
    Remove-NavisworksTemporaryOptionsPath -Path $TemporaryNavisworksOptionsPath
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
        [string]$ConfigFile = "Config.json",
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


