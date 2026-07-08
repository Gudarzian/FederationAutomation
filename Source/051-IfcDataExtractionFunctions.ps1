<#
IFC data extraction functions for Federation Automation.
- Exports object-level IFC property sets and quantity sets to CSV.
#>

function ConvertTo-FEDAUTOIfcMaxFileSizeBytes {
    param($Value)
    $defaultMb = 150.0
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace($Value.ToString())) {
        return [long]($defaultMb * 1MB)
    }

    $text = $Value.ToString().Trim()
    $numberText = $text
    $unit = 'MB'
    if ($text -match '^\s*(?<number>\d+(?:\.\d+)?)\s*(?<unit>KB|K|MB|M|GB|G|B|BYTES?)?\s*$') {
        $numberText = $Matches['number']
        if ($Matches['unit']) { $unit = $Matches['unit'].ToUpperInvariant() }
    }

    $number = 0.0
    if (-not [double]::TryParse($numberText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and
        -not [double]::TryParse($numberText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$number)) {
        Write-Warning ("Invalid IfcDataExtractionMaxFileSizeMB value '{0}'. Using default {1} MB." -f $Value, $defaultMb)
        return [long]($defaultMb * 1MB)
    }
    if ($number -le 0) {
        Write-Warning ("IfcDataExtractionMaxFileSizeMB must be greater than zero. Using default {0} MB." -f $defaultMb)
        return [long]($defaultMb * 1MB)
    }

    switch ($unit) {
        { $_ -in @('B','BYTE','BYTES') } { return [long]$number }
        { $_ -in @('KB','K') } { return [long]($number * 1KB) }
        { $_ -in @('GB','G') } { return [long]($number * 1GB) }
        default { return [long]($number * 1MB) }
    }
}

function Get-FEDAUTOIfcPythonExtractorScript {
    return @'
import argparse
import csv
import fnmatch
import json
import os
import sys

import ifcopenshell
import ifcopenshell.util.element


SKIP_CLASSES = {
    "IfcProject",
    "IfcSite",
    "IfcBuilding",
    "IfcBuildingStorey",
    "IfcSpace",
    "IfcOpeningElement",
}


def scalar(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (str, int, float)):
        return value
    if isinstance(value, (list, tuple)):
        return "; ".join(str(scalar(item)) for item in value if scalar(item) != "")
    return str(value)


def split_terms(value):
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


def wildcard_match(value, patterns):
    text = str(value or "").lower()
    return any(fnmatch.fnmatchcase(text, pattern.lower()) for pattern in patterns)


def filter_allows(value, inclusions, exclusions):
    include_match = not inclusions or wildcard_match(value, inclusions)
    if not include_match:
        return False
    return not wildcard_match(value, exclusions)


def entity_name(entity):
    if entity is None:
        return ""
    return getattr(entity, "Name", None) or getattr(entity, "LongName", None) or ""


def type_name(product):
    try:
        type_entity = ifcopenshell.util.element.get_type(product)
    except Exception:
        type_entity = None
    return entity_name(type_entity)


def predefined_type(product):
    try:
        return ifcopenshell.util.element.get_predefined_type(product) or ""
    except Exception:
        return getattr(product, "PredefinedType", "") or ""


def collect_material_names(value, seen=None):
    if seen is None:
        seen = set()
    names = []
    if value is None:
        return names
    marker = id(value)
    if marker in seen:
        return names
    seen.add(marker)
    if isinstance(value, (list, tuple, set)):
        for item in value:
            names.extend(collect_material_names(item, seen))
        return names
    name = getattr(value, "Name", None)
    if name:
        names.append(str(name))
    for attribute in (
        "ForLayerSet",
        "MaterialLayers",
        "MaterialProfiles",
        "Materials",
        "Material",
        "RelatingMaterial",
    ):
        child = getattr(value, attribute, None)
        if child is not None:
            names.extend(collect_material_names(child, seen))
    unique = []
    for item in names:
        if item and item not in unique:
            unique.append(item)
    return unique


def material_name(product):
    try:
        material = ifcopenshell.util.element.get_material(product, should_skip_usage=True)
    except Exception:
        material = None
    return "; ".join(collect_material_names(material))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--ifc", required=True)
    parser.add_argument("--csv", required=True)
    parser.add_argument("--tab-inclusions", default="")
    parser.add_argument("--tab-exclusions", default="")
    parser.add_argument("--attribute-inclusions", default="")
    parser.add_argument("--attribute-exclusions", default="")
    args = parser.parse_args()

    tab_inclusions = split_terms(args.tab_inclusions)
    tab_exclusions = split_terms(args.tab_exclusions)
    attribute_inclusions = split_terms(args.attribute_inclusions)
    attribute_exclusions = split_terms(args.attribute_exclusions)

    model = ifcopenshell.open(args.ifc)
    products = [
        product
        for product in model.by_type("IfcProduct")
        if getattr(product, "GlobalId", None) and not product.is_a() in SKIP_CLASSES
    ]

    columns = [
        ("Object", "EntityId", "Object||EntityId"),
        ("Object", "GlobalId", "Object||GlobalId"),
        ("Object", "Name", "Object||Name"),
        ("Object", "Class", "Object||Class"),
    ]
    column_keys = {column[2] for column in columns}
    rows = []

    def add_filtered(row, source_name, attribute_name, value):
        if not filter_allows(source_name, tab_inclusions, tab_exclusions):
            return
        if not filter_allows(attribute_name, attribute_inclusions, attribute_exclusions):
            return
        key = source_name + "||" + attribute_name
        if key not in column_keys:
            columns.append((source_name, attribute_name, key))
            column_keys.add(key)
        row[key] = scalar(value)

    for product in products:
        global_id = getattr(product, "GlobalId", "") or ""
        product_name = getattr(product, "Name", "") or ""
        product_type_name = type_name(product)
        material_text = material_name(product)
        row = {
            "Object||EntityId": "#" + str(product.id()),
            "Object||GlobalId": global_id,
            "Object||Name": product_name,
            "Object||Class": product.is_a(),
        }

        add_filtered(row, "Item", "Name", product_name)
        add_filtered(row, "Item", "Type", product_type_name)
        add_filtered(row, "Item", "Material", material_text)
        add_filtered(row, "Item", "Source File", os.path.basename(args.ifc))

        add_filtered(row, "Element ID", "Name", global_id)

        add_filtered(row, "Element", "IfcClass", product.is_a())
        add_filtered(row, "Element", "IfcGUID", global_id)
        add_filtered(row, "Element", "GlobalId", global_id)
        add_filtered(row, "Element", "Name", product_name)
        for attribute_name in (
            "ObjectType",
            "Tag",
            "Description",
            "OverallHeight",
            "OverallWidth",
            "OverallDepth",
        ):
            if hasattr(product, attribute_name):
                add_filtered(row, "Element", attribute_name, getattr(product, attribute_name, ""))
        add_filtered(row, "Element", "PredefinedType", predefined_type(product))

        add_filtered(row, "Material", "Name", material_text)

        psets = ifcopenshell.util.element.get_psets(product, psets_only=False, qtos_only=False)
        for source_name, attributes in psets.items():
            if not isinstance(attributes, dict):
                continue
            for attribute_name, value in attributes.items():
                if attribute_name == "id" or isinstance(value, dict):
                    continue
                add_filtered(row, source_name, attribute_name, value)
        rows.append(row)

    out_dir = os.path.dirname(args.csv)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with open(args.csv, "w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow([column[0] for column in columns])
        writer.writerow([column[1] for column in columns])
        for row in rows:
            writer.writerow([row.get(column[2], "") for column in columns])

    print(json.dumps({
        "IfcPath": args.ifc,
        "CsvPath": args.csv,
        "ObjectCount": len(rows),
        "AttributeCount": max(0, len(columns) - 4),
    }))


if __name__ == "__main__":
    main()
'@
}

function Invoke-FEDAUTOExternalCommand {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$ArgumentList = @(),
        [string]$WorkingDirectory
    )

    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $FilePath
    $escapedArguments = @($ArgumentList | ForEach-Object {
        $argument = if ($null -eq $_) { '' } else { $_.ToString() }
        if ($argument -notmatch '[\s"]') { $argument }
        else { '"' + ($argument -replace '"', '\"') + '"' }
    })
    $processInfo.Arguments = $escapedArguments -join ' '
    if ($WorkingDirectory) { $processInfo.WorkingDirectory = $WorkingDirectory }
    $processInfo.UseShellExecute = $false
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $process = [System.Diagnostics.Process]::Start($processInfo)
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

function Test-FEDAUTOIfcPythonPath {
    param([string]$PythonPath)
    if ([string]::IsNullOrWhiteSpace($PythonPath) -or -not (Test-Path $PythonPath -PathType Leaf)) { return $false }
    $result = Invoke-FEDAUTOExternalCommand -FilePath $PythonPath -ArgumentList @('-c','import sys; print(sys.executable)')
    return $result.ExitCode -eq 0
}

function Get-FEDAUTOIfcPythonPath {
    $candidatePaths = New-Object System.Collections.Generic.List[string]
    foreach ($commandName in @('python','python3')) {
        $command = Get-Command $commandName -ErrorAction SilentlyContinue
        if ($command -and $command.Source) { $candidatePaths.Add($command.Source) | Out-Null }
    }
    $pyLauncher = Get-Command 'py' -ErrorAction SilentlyContinue
    if ($pyLauncher -and $pyLauncher.Source) {
        $result = Invoke-FEDAUTOExternalCommand -FilePath $pyLauncher.Source -ArgumentList @('-3','-c','import sys; print(sys.executable)')
        if ($result.ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($result.StdOut)) {
            $candidatePaths.Add($result.StdOut.Trim()) | Out-Null
        }
    }
    $userPythonRoot = Join-Path $env:LOCALAPPDATA 'Programs\Python'
    if (Test-Path $userPythonRoot -PathType Container) {
        Get-ChildItem -Path $userPythonRoot -Filter python.exe -Recurse -File -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            ForEach-Object { $candidatePaths.Add($_.FullName) | Out-Null }
    }

    foreach ($path in @($candidatePaths | Select-Object -Unique)) {
        if (Test-FEDAUTOIfcPythonPath -PythonPath $path) { return $path }
    }
    return $null
}

function Install-FEDAUTOUserPython {
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget -or -not $winget.Source) {
        throw 'Python is not installed and winget was not found. Install Python for the current user, then rerun IFC data extraction.'
    }
    Write-Host 'Python was not found. Installing Python for the current user with winget...' -ForegroundColor Yellow
    $result = Invoke-FEDAUTOExternalCommand -FilePath $winget.Source -ArgumentList @(
        'install',
        '--exact',
        '--id', 'Python.Python.3.12',
        '--scope', 'user',
        '--accept-package-agreements',
        '--accept-source-agreements'
    )
    if ($result.ExitCode -ne 0) {
        throw ("Python user install failed. {0} {1}" -f $result.StdOut, $result.StdErr)
    }
    $pythonPath = Get-FEDAUTOIfcPythonPath
    if (-not $pythonPath) {
        throw 'Python installation completed, but python.exe could not be located. Restart the app or install Python manually for the current user.'
    }
    return $pythonPath
}

function Ensure-FEDAUTOIfcPythonEnvironment {
    $pythonPath = Get-FEDAUTOIfcPythonPath
    if (-not $pythonPath) { $pythonPath = Install-FEDAUTOUserPython }

    $envRoot = Join-Path $env:LOCALAPPDATA 'Federation-Automation\PythonEnv'
    $venvPython = Join-Path $envRoot 'Scripts\python.exe'
    if (-not (Test-Path $venvPython -PathType Leaf)) {
        Write-Host ("Creating Python environment: {0}" -f $envRoot) -ForegroundColor Yellow
        $parent = Split-Path -Parent $envRoot
        if (-not (Test-Path $parent -PathType Container)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $result = Invoke-FEDAUTOExternalCommand -FilePath $pythonPath -ArgumentList @('-m','venv',$envRoot)
        if ($result.ExitCode -ne 0 -or -not (Test-Path $venvPython -PathType Leaf)) {
            throw ("Failed to create Python environment. {0} {1}" -f $result.StdOut, $result.StdErr)
        }
    }

    foreach ($package in @('ifcopenshell')) {
        $show = Invoke-FEDAUTOExternalCommand -FilePath $venvPython -ArgumentList @('-m','pip','show',$package)
        if ($show.ExitCode -ne 0) {
            Write-Host ("Installing Python package '{0}' in the user environment..." -f $package) -ForegroundColor Yellow
            $pip = Invoke-FEDAUTOExternalCommand -FilePath $venvPython -ArgumentList @('-m','pip','install','--upgrade','pip')
            if ($pip.ExitCode -ne 0) { throw ("Failed to upgrade pip. {0} {1}" -f $pip.StdOut, $pip.StdErr) }
            $install = Invoke-FEDAUTOExternalCommand -FilePath $venvPython -ArgumentList @('-m','pip','install',$package)
            if ($install.ExitCode -ne 0) { throw ("Failed to install Python package '{0}'. {1} {2}" -f $package, $install.StdOut, $install.StdErr) }
        }
    }
    return $venvPython
}

function Export-IfcObjectAttributesCsvPython {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$IfcPath,
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][string]$PythonPath,
        [string]$TabInclusions = '',
        [string]$TabExclusions = '',
        [string]$AttributeInclusions = '',
        [string]$AttributeExclusions = ''
    )

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("FEDAUTO-IfcExtract-{0}" -f ([guid]::NewGuid().ToString('N')))
    $scriptPath = Join-Path $tempRoot 'extract_ifc_attributes.py'
    try {
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        [System.IO.File]::WriteAllText($scriptPath, (Get-FEDAUTOIfcPythonExtractorScript), [System.Text.UTF8Encoding]::new($false))
        $arguments = @($scriptPath, '--ifc', $IfcPath, '--csv', $CsvPath)
        if (-not [string]::IsNullOrWhiteSpace($TabInclusions)) { $arguments += @('--tab-inclusions', $TabInclusions) }
        if (-not [string]::IsNullOrWhiteSpace($TabExclusions)) { $arguments += @('--tab-exclusions', $TabExclusions) }
        if (-not [string]::IsNullOrWhiteSpace($AttributeInclusions)) { $arguments += @('--attribute-inclusions', $AttributeInclusions) }
        if (-not [string]::IsNullOrWhiteSpace($AttributeExclusions)) { $arguments += @('--attribute-exclusions', $AttributeExclusions) }
        $result = Invoke-FEDAUTOExternalCommand -FilePath $PythonPath -ArgumentList $arguments
        if ($result.ExitCode -ne 0) {
            throw ("Python IFC extraction failed. {0} {1}" -f $result.StdOut, $result.StdErr)
        }
        $jsonText = $result.StdOut.Trim()
        if ([string]::IsNullOrWhiteSpace($jsonText)) {
            throw 'Python IFC extraction did not return a summary.'
        }
        return ($jsonText | ConvertFrom-Json)
    }
    finally {
        if (Test-Path $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Write-FEDAUTOIfcExtractionProgress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$FileName,
        [string]$Status
    )
    $percent = if ($Total -gt 0) { [Math]::Round(($Current / [double]$Total) * 100, 2) } else { 100 }
    Write-Host ("IFC_DATA_EXTRACTION_PROGRESS|Current={0}|Total={1}|Percent={2}|Status={3}|File={4}" -f $Current, $Total, $percent, $Status, $FileName)
}

function Write-FEDAUTOIfcExtractionFileResult {
    param(
        [int]$Index,
        [int]$Total,
        [string]$Status,
        [string]$FileName,
        [double]$Seconds,
        [int]$ObjectCount = 0,
        [int]$AttributeCount = 0
    )
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Write-Host ("IFC_DATA_EXTRACTION_FILE|Time={0}|Index={1}|Total={2}|Status={3}|Seconds={4}|Objects={5}|Attributes={6}|File={7}" -f $timestamp, $Index, $Total, $Status, ([Math]::Round($Seconds, 2)), $ObjectCount, $AttributeCount, $FileName)
}

function Get-FEDAUTOIfcExtractionRuleTerms {
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value.ToString() -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Get-FEDAUTOIfcExtractionRuleValue {
    param(
        $Rule,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ($null -eq $Rule) { return '' }
    if ($Rule.PSObject.Properties.Name -contains $Name) {
        $value = $Rule.$Name
        if ($null -ne $value) { return $value.ToString() }
    }
    return ''
}

function Test-FEDAUTOIfcExtractionPatternMatch {
    param(
        [string]$Value,
        [string[]]$Patterns
    )
    if (-not $Patterns -or $Patterns.Count -eq 0) { return $false }
    $text = if ($null -eq $Value) { '' } else { $Value }
    return @(($Patterns | Where-Object { $text -like $_ })).Count -gt 0
}

function Test-FEDAUTOIfcExtractionRuleEnabled {
    param($Rule)
    if ($null -eq $Rule) { return $false }
    if (-not ($Rule.PSObject.Properties.Name -contains 'Run')) { return $true }
    $runText = if ($null -ne $Rule.Run) { $Rule.Run.ToString().Trim().ToLowerInvariant() } else { '' }
    return $runText -in @('yes','y','true','1')
}

function Test-FEDAUTOIfcExtractionRuleMatchesFile {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [Parameter(Mandatory = $true)]$Rule
    )
    $inclusionText = Get-FEDAUTOIfcExtractionRuleValue -Rule $Rule -Name 'FileInclusions'
    $exclusionText = Get-FEDAUTOIfcExtractionRuleValue -Rule $Rule -Name 'FileExclusions'
    $inclusions = @(Get-FEDAUTOIfcExtractionRuleTerms -Value $inclusionText)
    $exclusions = @(Get-FEDAUTOIfcExtractionRuleTerms -Value $exclusionText)
    $includeMatch = ($inclusions.Count -eq 0) -or (Test-FEDAUTOIfcExtractionPatternMatch -Value $File.Name -Patterns $inclusions)
    if (-not $includeMatch) { return $false }
    return -not (Test-FEDAUTOIfcExtractionPatternMatch -Value $File.Name -Patterns $exclusions)
}

function Get-FEDAUTOIfcExtractionRuleFilterHash {
    param($Rule)
    if ($null -eq $Rule) { return 'NoActiveRules:v1' }
    $payload = [ordered]@{
        FileInclusions      = Get-FEDAUTOIfcExtractionRuleValue -Rule $Rule -Name 'FileInclusions'
        FileExclusions      = Get-FEDAUTOIfcExtractionRuleValue -Rule $Rule -Name 'FileExclusions'
        TabInclusions       = Get-FEDAUTOIfcExtractionRuleValue -Rule $Rule -Name 'TabInclusions'
        TabExclusions       = Get-FEDAUTOIfcExtractionRuleValue -Rule $Rule -Name 'TabExclusions'
        AttributeInclusions = Get-FEDAUTOIfcExtractionRuleValue -Rule $Rule -Name 'AttributeInclusions'
        AttributeExclusions = Get-FEDAUTOIfcExtractionRuleValue -Rule $Rule -Name 'AttributeExclusions'
    }
    $json = $payload | ConvertTo-Json -Compress -Depth 4
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $hashBytes = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return ([System.BitConverter]::ToString($hashBytes) -replace '-', '').ToLowerInvariant()
}

function Get-FEDAUTOIfcExtractionSelectedRule {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [array]$Rules = @()
    )

    $activeRules = @($Rules | Where-Object { Test-FEDAUTOIfcExtractionRuleEnabled -Rule $_ })
    if ($activeRules.Count -eq 0) {
        return [pscustomobject]@{
            Selected = $true
            Rule = $null
            RuleIndex = $null
            IgnoredRuleIndexes = @()
            RuleFilterHash = Get-FEDAUTOIfcExtractionRuleFilterHash -Rule $null
            ColumnFiltersActive = $false
            TabInclusions = ''
            TabExclusions = ''
            AttributeInclusions = ''
            AttributeExclusions = ''
        }
    }

    $matches = New-Object System.Collections.Generic.List[object]
    for ($index = 0; $index -lt $Rules.Count; $index++) {
        $rule = $Rules[$index]
        if (-not (Test-FEDAUTOIfcExtractionRuleEnabled -Rule $rule)) { continue }
        if (Test-FEDAUTOIfcExtractionRuleMatchesFile -File $File -Rule $rule) {
            $matches.Add([pscustomobject]@{ Rule=$rule; RuleIndex=($index + 1) }) | Out-Null
        }
    }
    if ($matches.Count -eq 0) {
        return [pscustomobject]@{ Selected = $false; Rule = $null; RuleIndex = $null; IgnoredRuleIndexes = @() }
    }

    $selected = $matches[0]
    $tabInclusions = Get-FEDAUTOIfcExtractionRuleValue -Rule $selected.Rule -Name 'TabInclusions'
    $tabExclusions = Get-FEDAUTOIfcExtractionRuleValue -Rule $selected.Rule -Name 'TabExclusions'
    $attributeInclusions = Get-FEDAUTOIfcExtractionRuleValue -Rule $selected.Rule -Name 'AttributeInclusions'
    $attributeExclusions = Get-FEDAUTOIfcExtractionRuleValue -Rule $selected.Rule -Name 'AttributeExclusions'
    return [pscustomobject]@{
        Selected = $true
        Rule = $selected.Rule
        RuleIndex = $selected.RuleIndex
        IgnoredRuleIndexes = @($matches | Select-Object -Skip 1 | Select-Object -ExpandProperty RuleIndex)
        RuleFilterHash = Get-FEDAUTOIfcExtractionRuleFilterHash -Rule $selected.Rule
        ColumnFiltersActive = -not (
            [string]::IsNullOrWhiteSpace($tabInclusions) -and
            [string]::IsNullOrWhiteSpace($tabExclusions) -and
            [string]::IsNullOrWhiteSpace($attributeInclusions) -and
            [string]::IsNullOrWhiteSpace($attributeExclusions)
        )
        TabInclusions = $tabInclusions
        TabExclusions = $tabExclusions
        AttributeInclusions = $attributeInclusions
        AttributeExclusions = $attributeExclusions
    }
}

function Test-FEDAUTOIfcExtractionFileSelected {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [array]$Rules = @()
    )
    return (Get-FEDAUTOIfcExtractionSelectedRule -File $File -Rules $Rules).Selected
}

function ConvertTo-FEDAUTOIfcRelativePath {
    param(
        [string]$Path,
        [string]$RootFolder
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($RootFolder)) { return $Path }
    try {
        $fullPath = [System.IO.Path]::GetFullPath($Path)
        $fullRoot = [System.IO.Path]::GetFullPath($RootFolder).TrimEnd('\')
        if ($fullPath.Equals($fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) { return '' }
        if ($fullPath.StartsWith($fullRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            return $fullPath.Substring($fullRoot.Length + 1)
        }
    }
    catch { }
    return $Path
}

function Invoke-IfcDataExtraction {
    [CmdletBinding()]
    param(
        [array]$Settings = @(),
        [array]$IfcDataExtractionRules = @(),
        [string]$LogsFolder
    )

    Write-Host "=== START IFC Data Extraction ===" -ForegroundColor Cyan
    $extractionTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $basePath = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
    elseif ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    else {
        Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location)
    }

    $sourceFolder = CureFolderPath (Get-ConfigValue -Settings $Settings -Name 'SourceFolder' -Default 'SourceFiles')
    if ([string]::IsNullOrWhiteSpace($sourceFolder)) { $sourceFolder = 'SourceFiles' }
    if (-not [System.IO.Path]::IsPathRooted($sourceFolder)) { $sourceFolder = Join-Path $basePath $sourceFolder }

    $exportFolder = CureFolderPath (Get-ConfigValue -Settings $Settings -Name 'IfcDataExtractionFolder' -Default 'IFCDataExtraction')
    if ([string]::IsNullOrWhiteSpace($exportFolder)) { $exportFolder = 'IFCDataExtraction' }
    if (-not [System.IO.Path]::IsPathRooted($exportFolder)) { $exportFolder = Join-Path $basePath $exportFolder }
    $engineDisplay = 'Python / IfcOpenShell'
    $maxFileSizeValue = Get-ConfigValue -Settings $Settings -Name 'IfcDataExtractionMaxFileSizeMB' -Default '150'
    $maxFileSizeBytes = ConvertTo-FEDAUTOIfcMaxFileSizeBytes -Value $maxFileSizeValue
    $maxFileSizeDisplayMb = [Math]::Round($maxFileSizeBytes / 1MB, 2)
    $runValue = Get-ConfigValue -Settings $Settings -Name 'RunIfcDataExtraction' -Default 'No'
    $forceExtraction = ($null -ne $runValue -and $runValue.ToString().Trim().ToLowerInvariant() -eq 'force')
    $skipIfCurrentValue = Get-ConfigValue -Settings $Settings -Name 'IfcDataExtractionSkipIfCsvIsCurrent' -Default 'Yes'
    $skipIfCsvIsCurrent = $true
    if ($null -ne $skipIfCurrentValue -and $skipIfCurrentValue.ToString().Trim().ToLowerInvariant() -in @('no','n','false','0','ignore')) {
        $skipIfCsvIsCurrent = $false
    }

    if (-not (Test-Path $sourceFolder -PathType Container)) {
        throw "IFC data extraction source folder not found: $sourceFolder"
    }
    if (-not (Test-Path $exportFolder -PathType Container)) {
        New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null
    }
    $summaryPath = Join-Path $exportFolder 'ifc-data-extraction-summary.json'
    $previousExtractionRecords = @{}
    if (Test-Path -LiteralPath $summaryPath -PathType Leaf) {
        try {
            $previousSummary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json -ErrorAction Stop
            foreach ($record in @($previousSummary.Files)) {
                if (-not $record) { continue }
                $key = ''
                if ($record.PSObject.Properties.Name -contains 'IfcRelativePath' -and -not [string]::IsNullOrWhiteSpace($record.IfcRelativePath)) {
                    $key = $record.IfcRelativePath.ToString().ToLowerInvariant()
                }
                elseif ($record.PSObject.Properties.Name -contains 'IfcPath' -and -not [string]::IsNullOrWhiteSpace($record.IfcPath)) {
                    $key = [System.IO.Path]::GetFileName($record.IfcPath).ToLowerInvariant()
                }
                if (-not [string]::IsNullOrWhiteSpace($key)) { $previousExtractionRecords[$key] = $record }
            }
        }
        catch {
            Write-Warning ("Unable to read previous IFC data extraction summary for filter cache checks. Existing CSV files will be treated as stale when filters are active. Error: {0}" -f $_.Exception.Message)
        }
    }

    Write-Host "Source folder: $sourceFolder"
    Write-Host "Export folder: $exportFolder"
    Write-Host "Extraction engine: $engineDisplay"
    Write-Host ("Maximum IFC file size: {0} MB" -f $maxFileSizeDisplayMb)
    Write-Host ("Skip extraction when CSV is current: {0}" -f $(if ($skipIfCsvIsCurrent -and -not $forceExtraction) { 'Yes' } elseif ($forceExtraction) { 'No (forced)' } else { 'No' }))
    $activeRuleCount = @($IfcDataExtractionRules | Where-Object { Test-FEDAUTOIfcExtractionRuleEnabled -Rule $_ }).Count
    Write-Host ("IFC data extraction rules: {0} active rule(s)" -f $activeRuleCount)
    if ($activeRuleCount -eq 0) {
        Write-Host "No active data extraction rules were supplied; every IFC file in the source/download output folder is eligible."
    }

    $allIfcFiles = @(Get-ChildItem -Path $sourceFolder -Filter '*.ifc' -File)
    $selectedIfcFiles = New-Object System.Collections.Generic.List[object]
    $ruleSkippedCount = 0
    $ruleConflictCount = 0
    foreach ($candidateIfc in $allIfcFiles) {
        $selection = Get-FEDAUTOIfcExtractionSelectedRule -File $candidateIfc -Rules $IfcDataExtractionRules
        if (-not $selection.Selected) {
            $ruleSkippedCount++
            continue
        }
        if ($selection.IgnoredRuleIndexes -and $selection.IgnoredRuleIndexes.Count -gt 0) {
            $ruleConflictCount++
            Write-Warning ("IFC data extraction rule conflict for '{0}': rule {1} matched first; later matching rule(s) {2} ignored for this file." -f $candidateIfc.Name, $selection.RuleIndex, ($selection.IgnoredRuleIndexes -join ', '))
        }
        $selectedIfcFiles.Add([pscustomobject]@{ File=$candidateIfc; Selection=$selection }) | Out-Null
    }
    if ($ruleSkippedCount -gt 0) {
        Write-Host ("IFC data extraction rules excluded {0} IFC file(s)." -f $ruleSkippedCount) -ForegroundColor Yellow
    }
    $exports = New-Object System.Collections.Generic.List[object]
    $skipped = New-Object System.Collections.Generic.List[object]
    $failures = New-Object System.Collections.Generic.List[object]
    $eligibleFileCount = @($selectedIfcFiles | Where-Object { $_.File.Length -le $maxFileSizeBytes }).Count
    $pythonPath = $null
    if ($eligibleFileCount -gt 0) { $pythonPath = Ensure-FEDAUTOIfcPythonEnvironment }
    $processedFileIndex = 0
    foreach ($ifcEntry in $selectedIfcFiles) {
        $processedFileIndex++
        $ifc = $ifcEntry.File
        $selection = $ifcEntry.Selection
        $ifcRelativePath = ConvertTo-FEDAUTOIfcRelativePath -Path $ifc.FullName -RootFolder $sourceFolder
        if ($ifc.Length -gt $maxFileSizeBytes) {
            $message = ("Skipping IFC data extraction for '{0}' because file size {1} MB exceeds the configured limit of {2} MB." -f $ifc.Name, ([Math]::Round($ifc.Length / 1MB, 2)), $maxFileSizeDisplayMb)
            Write-Warning $message
            Write-FEDAUTOIfcExtractionProgress -Current $processedFileIndex -Total $selectedIfcFiles.Count -Status 'Skipped' -FileName $ifc.Name
            Write-FEDAUTOIfcExtractionFileResult -Index $processedFileIndex -Total $selectedIfcFiles.Count -Status 'Skipped' -FileName $ifc.Name -Seconds 0
            $skipped.Add([pscustomobject]@{
                IfcPath       = $ifc.FullName
                IfcRelativePath = $ifcRelativePath
                SizeBytes     = $ifc.Length
                MaxSizeBytes  = $maxFileSizeBytes
                RuleIndex     = $selection.RuleIndex
                RuleFilterHash = $selection.RuleFilterHash
                Reason        = $message
            }) | Out-Null
            continue
        }
        $csvPath = Join-Path $exportFolder ([System.IO.Path]::ChangeExtension($ifc.Name, '.csv'))
        if ($skipIfCsvIsCurrent -and -not $forceExtraction -and (Test-Path $csvPath -PathType Leaf)) {
            $csvItem = Get-Item -LiteralPath $csvPath
            if ($csvItem.LastWriteTimeUtc -ge $ifc.LastWriteTimeUtc) {
                $previousRecord = $previousExtractionRecords[$ifcRelativePath.ToLowerInvariant()]
                if (-not $previousRecord) { $previousRecord = $previousExtractionRecords[$ifc.Name.ToLowerInvariant()] }
                $previousFilterHash = ''
                if ($previousRecord -and $previousRecord.PSObject.Properties.Name -contains 'RuleFilterHash') {
                    $previousFilterHash = if ($null -ne $previousRecord.RuleFilterHash) { $previousRecord.RuleFilterHash.ToString() } else { '' }
                }
                $filtersAreCurrent = ($previousFilterHash -eq $selection.RuleFilterHash)
                if (-not $filtersAreCurrent -and [string]::IsNullOrWhiteSpace($previousFilterHash) -and -not $selection.ColumnFiltersActive) {
                    $filtersAreCurrent = $true
                }
                if ($filtersAreCurrent) {
                    $message = ("Skipping IFC data extraction for '{0}' because existing CSV is current." -f $ifc.Name)
                    Write-Host $message -ForegroundColor Yellow
                    Write-FEDAUTOIfcExtractionProgress -Current $processedFileIndex -Total $selectedIfcFiles.Count -Status 'Skipped' -FileName $ifc.Name
                    Write-FEDAUTOIfcExtractionFileResult -Index $processedFileIndex -Total $selectedIfcFiles.Count -Status 'SkippedCurrent' -FileName $ifc.Name -Seconds 0
                    $skipped.Add([pscustomobject]@{
                        IfcPath   = $ifc.FullName
                        IfcRelativePath = $ifcRelativePath
                        CsvPath   = $csvPath
                        CsvRelativePath = (ConvertTo-FEDAUTOIfcRelativePath -Path $csvPath -RootFolder $exportFolder)
                        RuleIndex = $selection.RuleIndex
                        RuleFilterHash = $selection.RuleFilterHash
                        Reason    = $message
                    }) | Out-Null
                    continue
                }
                Write-Host ("Existing CSV for '{0}' is current by timestamp, but the selected data extraction filters changed; extracting again." -f $ifc.Name) -ForegroundColor Yellow
            }
        }
        try {
            Write-FEDAUTOIfcExtractionProgress -Current ($processedFileIndex - 1) -Total $selectedIfcFiles.Count -Status 'Starting' -FileName $ifc.Name
            Write-Host ("Extracting IFC data: {0}" -f $ifc.Name)
            $fileTimer = [System.Diagnostics.Stopwatch]::StartNew()
            $result = Export-IfcObjectAttributesCsvPython -IfcPath $ifc.FullName -CsvPath $csvPath -PythonPath $pythonPath -TabInclusions $selection.TabInclusions -TabExclusions $selection.TabExclusions -AttributeInclusions $selection.AttributeInclusions -AttributeExclusions $selection.AttributeExclusions
            $fileTimer.Stop()
            $result | Add-Member -NotePropertyName Seconds -NotePropertyValue ([Math]::Round($fileTimer.Elapsed.TotalSeconds, 2)) -Force
            $result | Add-Member -NotePropertyName IfcRelativePath -NotePropertyValue $ifcRelativePath -Force
            $result | Add-Member -NotePropertyName CsvRelativePath -NotePropertyValue (ConvertTo-FEDAUTOIfcRelativePath -Path $csvPath -RootFolder $exportFolder) -Force
            $result | Add-Member -NotePropertyName RuleIndex -NotePropertyValue $selection.RuleIndex -Force
            $result | Add-Member -NotePropertyName RuleFilterHash -NotePropertyValue $selection.RuleFilterHash -Force
            $result | Add-Member -NotePropertyName TabInclusions -NotePropertyValue $selection.TabInclusions -Force
            $result | Add-Member -NotePropertyName TabExclusions -NotePropertyValue $selection.TabExclusions -Force
            $result | Add-Member -NotePropertyName AttributeInclusions -NotePropertyValue $selection.AttributeInclusions -Force
            $result | Add-Member -NotePropertyName AttributeExclusions -NotePropertyValue $selection.AttributeExclusions -Force
            $exports.Add($result) | Out-Null
            Write-FEDAUTOIfcExtractionProgress -Current $processedFileIndex -Total $selectedIfcFiles.Count -Status 'Exported' -FileName $ifc.Name
            Write-FEDAUTOIfcExtractionFileResult -Index $processedFileIndex -Total $selectedIfcFiles.Count -Status 'Exported' -FileName $ifc.Name -Seconds $fileTimer.Elapsed.TotalSeconds -ObjectCount $result.ObjectCount -AttributeCount $result.AttributeCount
            Write-Host ("  Exported {0} object row(s), {1} attribute column(s) in {2}s: {3}" -f $result.ObjectCount, $result.AttributeCount, ([Math]::Round($fileTimer.Elapsed.TotalSeconds, 2)), $csvPath)
        }
        catch {
            $failures.Add([pscustomobject]@{ IfcPath=$ifc.FullName; IfcRelativePath=$ifcRelativePath; RuleIndex=$selection.RuleIndex; RuleFilterHash=$selection.RuleFilterHash; Error=$_.ToString() }) | Out-Null
            Write-FEDAUTOIfcExtractionProgress -Current $processedFileIndex -Total $selectedIfcFiles.Count -Status 'Failed' -FileName $ifc.Name
            Write-Warning ("Failed to extract IFC data from '{0}': {1}" -f $ifc.Name, $_)
        }
    }

    $extractionTimer.Stop()
    $durationSeconds = [Math]::Round($extractionTimer.Elapsed.TotalSeconds, 2)
    $summary = [ordered]@{
        SourceFolder  = $sourceFolder
        SourceFolderRelativePath = (ConvertTo-FEDAUTOIfcRelativePath -Path $sourceFolder -RootFolder $basePath)
        ExportFolder  = $exportFolder
        ExportFolderRelativePath = (ConvertTo-FEDAUTOIfcRelativePath -Path $exportFolder -RootFolder $basePath)
        Engine        = $engineDisplay
        MaxFileSizeMB = $maxFileSizeDisplayMb
        SkipIfCsvIsCurrent = $skipIfCsvIsCurrent
        ForceExtraction = $forceExtraction
        RuleSkippedInputCount = $ruleSkippedCount
        RuleConflictCount = $ruleConflictCount
        DurationSeconds = $durationSeconds
        Exported      = $exports.Count
        Skipped       = $skipped.Count
        Failed        = $failures.Count
        Timestamp     = (Get-Date).ToUniversalTime().ToString('o')
        Files         = @($exports.ToArray())
        SkippedFiles  = @($skipped.ToArray())
        FailedFiles   = @($failures.ToArray())
    }
    $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $summaryPath -Encoding UTF8
    Write-Host ("IFC data extraction exported {0} file(s), skipped {1}, failed {2}. Duration: {3}s." -f $exports.Count, $skipped.Count, $failures.Count, $durationSeconds)
    Write-Host "=== END IFC Data Extraction ==="
    return [pscustomobject]$summary
}
