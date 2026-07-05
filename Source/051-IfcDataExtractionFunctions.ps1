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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--ifc", required=True)
    parser.add_argument("--csv", required=True)
    args = parser.parse_args()

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

    for product in products:
        row = {
            "Object||EntityId": "#" + str(product.id()),
            "Object||GlobalId": getattr(product, "GlobalId", "") or "",
            "Object||Name": getattr(product, "Name", "") or "",
            "Object||Class": product.is_a(),
        }
        psets = ifcopenshell.util.element.get_psets(product, psets_only=False, qtos_only=False)
        for source_name, attributes in psets.items():
            if not isinstance(attributes, dict):
                continue
            for attribute_name, value in attributes.items():
                if attribute_name == "id" or isinstance(value, dict):
                    continue
                key = source_name + "||" + attribute_name
                if key not in column_keys:
                    columns.append((source_name, attribute_name, key))
                    column_keys.add(key)
                row[key] = scalar(value)
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
        [Parameter(Mandatory = $true)][string]$PythonPath
    )

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("FEDAUTO-IfcExtract-{0}" -f ([guid]::NewGuid().ToString('N')))
    $scriptPath = Join-Path $tempRoot 'extract_ifc_attributes.py'
    try {
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        [System.IO.File]::WriteAllText($scriptPath, (Get-FEDAUTOIfcPythonExtractorScript), [System.Text.UTF8Encoding]::new($false))
        $result = Invoke-FEDAUTOExternalCommand -FilePath $PythonPath -ArgumentList @($scriptPath,'--ifc',$IfcPath,'--csv',$CsvPath)
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

function Test-FEDAUTOIfcExtractionRuleEnabled {
    param($Rule)
    if ($null -eq $Rule) { return $false }
    if (-not ($Rule.PSObject.Properties.Name -contains 'Run')) { return $true }
    $runText = if ($null -ne $Rule.Run) { $Rule.Run.ToString().Trim().ToLowerInvariant() } else { '' }
    return $runText -in @('yes','y','true','1')
}

function Test-FEDAUTOIfcExtractionFileSelected {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [array]$Rules = @()
    )
    $activeRules = @($Rules | Where-Object { Test-FEDAUTOIfcExtractionRuleEnabled -Rule $_ })
    if ($activeRules.Count -eq 0) { return $true }
    foreach ($rule in $activeRules) {
        $inclusionText = if ($rule.PSObject.Properties.Name -contains 'Inclusions') { $rule.Inclusions } elseif ($rule.PSObject.Properties.Name -contains 'Inclusion') { $rule.Inclusion } else { $null }
        $exclusionText = if ($rule.PSObject.Properties.Name -contains 'Exclusions') { $rule.Exclusions } elseif ($rule.PSObject.Properties.Name -contains 'Exclusion') { $rule.Exclusion } else { $null }
        $inclusions = @(Get-FEDAUTOIfcExtractionRuleTerms -Value $inclusionText)
        $exclusions = @(Get-FEDAUTOIfcExtractionRuleTerms -Value $exclusionText)
        $includeMatch = ($inclusions.Count -eq 0) -or (($inclusions | Where-Object { $File.Name -like $_ }).Count -gt 0)
        if (-not $includeMatch) { continue }
        $excludeMatch = ($exclusions | Where-Object { $File.Name -like $_ }).Count -gt 0
        if (-not $excludeMatch) { return $true }
    }
    return $false
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
    $ifcFiles = @($allIfcFiles | Where-Object { Test-FEDAUTOIfcExtractionFileSelected -File $_ -Rules $IfcDataExtractionRules })
    $ruleSkippedCount = $allIfcFiles.Count - $ifcFiles.Count
    if ($ruleSkippedCount -gt 0) {
        Write-Host ("IFC data extraction rules excluded {0} IFC file(s)." -f $ruleSkippedCount) -ForegroundColor Yellow
    }
    $exports = New-Object System.Collections.Generic.List[object]
    $skipped = New-Object System.Collections.Generic.List[object]
    $failures = New-Object System.Collections.Generic.List[object]
    $eligibleFileCount = @($ifcFiles | Where-Object { $_.Length -le $maxFileSizeBytes }).Count
    $pythonPath = $null
    if ($eligibleFileCount -gt 0) { $pythonPath = Ensure-FEDAUTOIfcPythonEnvironment }
    $processedFileIndex = 0
    foreach ($ifc in $ifcFiles) {
        $processedFileIndex++
        if ($ifc.Length -gt $maxFileSizeBytes) {
            $message = ("Skipping IFC data extraction for '{0}' because file size {1} MB exceeds the configured limit of {2} MB." -f $ifc.Name, ([Math]::Round($ifc.Length / 1MB, 2)), $maxFileSizeDisplayMb)
            Write-Warning $message
            Write-FEDAUTOIfcExtractionProgress -Current $processedFileIndex -Total $ifcFiles.Count -Status 'Skipped' -FileName $ifc.Name
            Write-FEDAUTOIfcExtractionFileResult -Index $processedFileIndex -Total $ifcFiles.Count -Status 'Skipped' -FileName $ifc.Name -Seconds 0
            $skipped.Add([pscustomobject]@{
                IfcPath       = $ifc.FullName
                SizeBytes     = $ifc.Length
                MaxSizeBytes  = $maxFileSizeBytes
                Reason        = $message
            }) | Out-Null
            continue
        }
        $csvPath = Join-Path $exportFolder ([System.IO.Path]::ChangeExtension($ifc.Name, '.csv'))
        if ($skipIfCsvIsCurrent -and -not $forceExtraction -and (Test-Path $csvPath -PathType Leaf)) {
            $csvItem = Get-Item -LiteralPath $csvPath
            if ($csvItem.LastWriteTimeUtc -ge $ifc.LastWriteTimeUtc) {
                $message = ("Skipping IFC data extraction for '{0}' because existing CSV is current." -f $ifc.Name)
                Write-Host $message -ForegroundColor Yellow
                Write-FEDAUTOIfcExtractionProgress -Current $processedFileIndex -Total $ifcFiles.Count -Status 'Skipped' -FileName $ifc.Name
                Write-FEDAUTOIfcExtractionFileResult -Index $processedFileIndex -Total $ifcFiles.Count -Status 'SkippedCurrent' -FileName $ifc.Name -Seconds 0
                $skipped.Add([pscustomobject]@{
                    IfcPath   = $ifc.FullName
                    CsvPath   = $csvPath
                    Reason    = $message
                }) | Out-Null
                continue
            }
        }
        try {
            Write-FEDAUTOIfcExtractionProgress -Current ($processedFileIndex - 1) -Total $ifcFiles.Count -Status 'Starting' -FileName $ifc.Name
            Write-Host ("Extracting IFC data: {0}" -f $ifc.Name)
            $fileTimer = [System.Diagnostics.Stopwatch]::StartNew()
            $result = Export-IfcObjectAttributesCsvPython -IfcPath $ifc.FullName -CsvPath $csvPath -PythonPath $pythonPath
            $fileTimer.Stop()
            $result | Add-Member -NotePropertyName Seconds -NotePropertyValue ([Math]::Round($fileTimer.Elapsed.TotalSeconds, 2)) -Force
            $exports.Add($result) | Out-Null
            Write-FEDAUTOIfcExtractionProgress -Current $processedFileIndex -Total $ifcFiles.Count -Status 'Exported' -FileName $ifc.Name
            Write-FEDAUTOIfcExtractionFileResult -Index $processedFileIndex -Total $ifcFiles.Count -Status 'Exported' -FileName $ifc.Name -Seconds $fileTimer.Elapsed.TotalSeconds -ObjectCount $result.ObjectCount -AttributeCount $result.AttributeCount
            Write-Host ("  Exported {0} object row(s), {1} attribute column(s) in {2}s: {3}" -f $result.ObjectCount, $result.AttributeCount, ([Math]::Round($fileTimer.Elapsed.TotalSeconds, 2)), $csvPath)
        }
        catch {
            $failures.Add([pscustomobject]@{ IfcPath=$ifc.FullName; Error=$_.ToString() }) | Out-Null
            Write-FEDAUTOIfcExtractionProgress -Current $processedFileIndex -Total $ifcFiles.Count -Status 'Failed' -FileName $ifc.Name
            Write-Warning ("Failed to extract IFC data from '{0}': {1}" -f $ifc.Name, $_)
        }
    }

    $extractionTimer.Stop()
    $durationSeconds = [Math]::Round($extractionTimer.Elapsed.TotalSeconds, 2)
    $summary = [ordered]@{
        SourceFolder  = $sourceFolder
        ExportFolder  = $exportFolder
        Engine        = $engineDisplay
        MaxFileSizeMB = $maxFileSizeDisplayMb
        SkipIfCsvIsCurrent = $skipIfCsvIsCurrent
        ForceExtraction = $forceExtraction
        RuleSkippedInputCount = $ruleSkippedCount
        DurationSeconds = $durationSeconds
        Exported      = $exports.Count
        Skipped       = $skipped.Count
        Failed        = $failures.Count
        Timestamp     = (Get-Date).ToUniversalTime().ToString('o')
        Files         = @($exports.ToArray())
        SkippedFiles  = @($skipped.ToArray())
        FailedFiles   = @($failures.ToArray())
    }
    $summaryPath = Join-Path $exportFolder 'ifc-data-extraction-summary.json'
    $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $summaryPath -Encoding UTF8
    Write-Host ("IFC data extraction exported {0} file(s), skipped {1}, failed {2}. Duration: {3}s." -f $exports.Count, $skipped.Count, $failures.Count, $durationSeconds)
    Write-Host "=== END IFC Data Extraction ==="
    return [pscustomobject]$summary
}
