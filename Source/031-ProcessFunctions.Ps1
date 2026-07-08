<#
Processing-related functions for Federation Automation.
- IFC attribute injection and supporting helpers.
#>

function Get-ConfigValue {
<#
Returns a config setting value by parameter name with default fallback.
#>
    param([array]$Settings, [string]$Name, $Default = $null)

    $match = $Settings |
        Where-Object { $_.Parameter -eq $Name } |
        Select-Object -First 1 |
        ForEach-Object { $_.Value }

    if ($null -eq $match) { return $Default }
    return $match
}

function Normalize-ComparisonKey {
<#
Normalizes text for loose column name matching.
#>
    param([string]$Text)
    $normalized = Normalize-ExcelText -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $null }
    $normalized = $normalized.ToLowerInvariant()
    # ignore common separators so "File_UpdateDate" and "File Update Date" both map to the same key
    return [regex]::Replace($normalized, "[\s_\-]", "")
}

function Initialize-RowFieldMap {
<#
Builds a normalized column name map for fast row lookups.
#>
    param($Row)
    if (-not $Row) { return }
    if ($Row.PSObject.Properties.Name -contains '__NormalizedFieldMap') { return }
    $map = @{}
    foreach ($prop in $Row.PSObject.Properties) {
        if ($prop.Name -eq '__NormalizedFieldMap') { continue }
        $key = Normalize-ComparisonKey -Text $prop.Name
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if (-not $map.ContainsKey($key)) { $map[$key] = $prop.Value }
    }
    $Row | Add-Member -NotePropertyName '__NormalizedFieldMap' -NotePropertyValue $map -Force
}

function Get-RowFieldValue {
<#
Looks up a row value using normalized column name matching.
#>
    param($Row, [string]$ColumnName)
    if (-not $Row) { return $null }
    $target = Normalize-ComparisonKey -Text $ColumnName
    if ([string]::IsNullOrWhiteSpace($target)) { return $null }
    $mapProp = $Row.PSObject.Properties['__NormalizedFieldMap']
    if ($mapProp) {
        $map = $mapProp.Value
        if ($map -and $map.ContainsKey($target)) { return $map[$target] }
        return $null
    }
    foreach ($prop in $Row.PSObject.Properties) {
        $candidate = Normalize-ComparisonKey -Text $prop.Name
        if ($candidate -eq $target) { return $prop.Value }
    }
    return $null
}

function Get-FirstRowFieldValue {
<#
Returns the first matching value across candidate column names.
#>
    param($Row, [string[]]$ColumnNames)
    foreach ($name in $ColumnNames) {
        $value = Get-RowFieldValue -Row $Row -ColumnName $name
        if ($null -ne $value) { return $value }
    }
    return $null
}

function ConvertTo-Hashtable {
<#
Converts objects into a hashtable structure for JSON metadata.
#>
    # Some environments don't support ConvertFrom-Json -AsHashtable; this recreates that behavior manually
    param($InputObject)

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $hash = @{}
        foreach ($key in $InputObject.Keys) {
            $hash[$key] = ConvertTo-Hashtable -InputObject $InputObject[$key]
        }
        return $hash
    }
    if ($InputObject -is [pscustomobject]) {
        $hash = @{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            $hash[$prop.Name] = ConvertTo-Hashtable -InputObject $prop.Value
        }
        return $hash
    }
    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        $list = @()
        foreach ($item in $InputObject) {
            $list += ,(ConvertTo-Hashtable -InputObject $item)
        }
        return $list
    }
    return $InputObject
}

function Get-PropertyValue {
<#
Gets a property value from dictionaries or PSCustomObjects.
#>
    param($Object, [string]$Name)
    if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($Name)) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    if ($Object.PSObject -and $Object.PSObject.Properties.Name -contains $Name) {
        return $Object.$Name
    }
    return $null
}

function Set-PropertyValue {
<#
Sets a property value on dictionaries or PSCustomObjects.
#>
    param($Object, [string]$Name, $Value)
    if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($Name)) { return }
    if ($Object -is [System.Collections.IDictionary]) {
        $Object[$Name] = $Value
        return
    }
    # PSCustomObject fallback
    if ($Object.PSObject -and $Object.PSObject.Properties.Name -contains $Name) {
        $Object.$Name = $Value
    } else {
        Add-Member -InputObject $Object -NotePropertyName $Name -NotePropertyValue $Value -Force
    }
}

function Resolve-FEDAUTOStateFilePath {
<#
Resolves a state-file path from portable metadata first, then legacy absolute
metadata. This prevents shared OneDrive/SharePoint/ACC runs from reprocessing
only because a previous user's C:\Users\...\ path is stored in JSON.
#>
    param(
        [string]$StoredPath,
        [string]$RootFolder,
        [string]$RelativePath
    )

    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($RelativePath) -and -not [string]::IsNullOrWhiteSpace($RootFolder)) {
        [void]$candidates.Add((Join-Path $RootFolder $RelativePath))
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

function ConvertTo-FEDAUTORelativePath {
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

function Get-StringHash {
<#
Computes a SHA256 hash for a string.
#>
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join ""
    } finally {
        $sha.Dispose()
    }
}

function Convert-FileSizeValue {
<#
Normalizes file size values to numeric when possible.
#>
    param($Value)
    if ($null -eq $Value) { return $null }
    $text = $Value.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $parsed = 0
    if ([long]::TryParse($text, [ref]$parsed)) { return $parsed }
    return $text
}

function Convert-DateValue {
<#
Normalizes date-like values to a consistent string format.
#>
    param($Value)
    # Normalize anything that looks like a date into yyyy/MM/dd HH:mm so viewers show consistent text
    if ($null -eq $Value) { return $null }
    $targetFormat = "yyyy/MM/dd HH:mm"
    if ($Value -is [datetime]) {
        return $Value.ToString($targetFormat)
    }
    if ($Value -is [double] -or $Value -is [single] -or
        $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        try {
            $date = [datetime]::FromOADate([double]$Value)
            return $date.ToString($targetFormat)
        } catch {
            # fall back to string parsing
        }
    }
    $text = $Value.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    # Handle JSON-style epoch markers (/Date(123456789)/)
    $dateMatch = [regex]::Match($text, "^/Date\((?<ticks>-?\d+)\)/$")
    if ($dateMatch.Success) {
        $ticks = [long]::Parse($dateMatch.Groups['ticks'].Value)
        $epoch = [DateTimeOffset]::FromUnixTimeMilliseconds($ticks).DateTime
        return $epoch.ToString($targetFormat)
    }

    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($text, [System.Globalization.CultureInfo]::InvariantCulture,[System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
        return $parsed.ToString($targetFormat)
    }
    if ([datetime]::TryParse($text, [System.Globalization.CultureInfo]::CurrentCulture,[System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
        return $parsed.ToString($targetFormat)
    }
    return $text
}

function Should-FormatAsDate {
<#
Determines whether an attribute should be formatted as a date.
#>
    param($Definition)
    # Treat any attribute mentioning date/time (in either the workbook column or the target name) as a date
    if ($null -eq $Definition) { return $false }
    $candidates = @($Definition.Name, $Definition.WorkbookColumn)
    foreach ($item in $candidates) {
        $text = Normalize-ExcelText -Text $item
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        if ($text.ToLowerInvariant() -match 'date|time') { return $true }
    }
    return $false
}

function Convert-FileSizeToBits {
<#
Converts file size values from bytes to a bits string.
#>
    param($Value)
    # ProjectWise reports bytes; multiply by 8 and append "bits" for clarity in the IFC viewer
    if ($null -eq $Value) { return $null }
    $number = 0.0
    $text = $Value.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number) -or
        [double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::CurrentCulture, [ref]$number)) {
        $bits = [math]::Round($number * 8, 0)
        return ("{0} bits" -f [long]$bits)
    }
    return ("{0} bits" -f $text)
}

function Should-FormatAsFileSize {
<#
Determines whether an attribute should be treated as file size.
#>
    param($Definition)
    # Simple heuristic: any attribute with "size" in its name is treated as a file-size entry
    if ($null -eq $Definition) { return $false }
    $candidates = @($Definition.Name, $Definition.WorkbookColumn)
    foreach ($item in $candidates) {
        $text = Normalize-ExcelText -Text $item
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        if ($text.ToLowerInvariant() -match 'size') { return $true }
    }
    return $false
}

function Resolve-PythonExecutable {
<#
Resolves a Python executable path or command.
#>
    param([string]$ExecutableHint)

    if ([string]::IsNullOrWhiteSpace($ExecutableHint)) {
        $ExecutableHint = "python"
    }

    if (Test-Path $ExecutableHint -PathType Leaf) {
        return (Resolve-Path $ExecutableHint).Path
    }

    try {
        $commandInfo = Get-Command $ExecutableHint -ErrorAction Stop
        if ($commandInfo.Source) { return $commandInfo.Source }
        if ($commandInfo.Path) { return $commandInfo.Path }
        if ($commandInfo.Definition) { return $commandInfo.Definition }
    }
    catch {
        throw "Unable to locate Python executable '$ExecutableHint'. Install Python or update the Config.json setting 'PythonExecutable'."
    }

    throw "Unable to resolve Python executable '$ExecutableHint'."
}

function Ensure-PythonPackages {
<#
Ensures required Python packages are installed in the environment.
#>
    param(
        [string]$PythonPath,
        [string[]]$Packages,
        [switch]$UpgradePip
    )

    if (-not (Test-Path $PythonPath -PathType Leaf)) {
        throw "Python interpreter not found at '$PythonPath'."
    }

    if ($UpgradePip) {
        & $PythonPath -m pip install --upgrade pip | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to upgrade pip in the Python environment at '$PythonPath'."
        }
    }

    foreach ($package in $Packages) {
        if ([string]::IsNullOrWhiteSpace($package)) { continue }
        & $PythonPath -m pip show $package | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Installing Python package '$package'..." -ForegroundColor Yellow
            & $PythonPath -m pip install --upgrade $package
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to install required Python package '$package'."
            }
        }
    }
}

function Ensure-PythonEnvironment {
<#
Ensures a Python virtual environment exists and is ready.
#>
    param(
        [string]$BootstrapPython,
        [string]$EnvironmentPath,
        [string[]]$RequiredPackages
    )

    if ([string]::IsNullOrWhiteSpace($EnvironmentPath)) {
        throw "Python environment path is not configured."
    }

    $venvPython = Join-Path $EnvironmentPath "Scripts\\python.exe"
    $created = $false

    if (-not (Test-Path $venvPython -PathType Leaf)) {
        Write-Host "Python environment not found at '$EnvironmentPath'; creating a new virtual environment..." -ForegroundColor Yellow
        $parentDir = Split-Path -Parent $EnvironmentPath
        if (-not (Test-Path $parentDir -PathType Container)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }
        & $BootstrapPython -m venv $EnvironmentPath
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $venvPython -PathType Leaf)) {
            throw "Failed to create Python virtual environment at '$EnvironmentPath'."
        }
        $created = $true
    }

    Ensure-PythonPackages -PythonPath $venvPython -Packages $RequiredPackages -UpgradePip:$created
    return $venvPython
}

function New-IfcGuid {
    $guid = [Guid]::NewGuid()
    $hex = $guid.ToString("N")
    $bytes = New-Object byte[] 16
    for ($i = 0; $i -lt 16; $i++) {
        $bytes[$i] = [Convert]::ToByte($hex.Substring($i * 2, 2), 16)
    }

    $b64 = [Convert]::ToBase64String($bytes).TrimEnd("=")
    $stdAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    $ifcAlphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$"
    $chars = $b64.ToCharArray() | ForEach-Object { $ifcAlphabet[$stdAlphabet.IndexOf($_)] }
    return -join $chars
}

function Escape-IfcString {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    return $Value -replace "'", "''"
}

function Read-IfcStringLiteral {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $text = $Value.Trim()
    if ($text.StartsWith("'") -and $text.EndsWith("'") -and $text.Length -ge 2) {
        $inner = $text.Substring(1, $text.Length - 2)
        return $inner -replace "''", "'"
    }
    return $text
}

function Split-IfcArguments {
    param([string]$Text)
    $args = @()
    $sb = New-Object System.Text.StringBuilder
    $depth = 0
    $inString = $false
    $i = 0
    while ($i -lt $Text.Length) {
        $ch = $Text[$i]
        if ($inString) {
            if ($ch -eq "'") {
                if ($i + 1 -lt $Text.Length -and $Text[$i + 1] -eq "'") {
                    [void]$sb.Append("''")
                    $i += 2
                    continue
                }
                $inString = $false
            }
            [void]$sb.Append($ch)
            $i++
            continue
        }

        if ($ch -eq "'") {
            $inString = $true
            [void]$sb.Append($ch)
            $i++
            continue
        }
        elseif ($ch -eq "(") {
            $depth++
            [void]$sb.Append($ch)
            $i++
            continue
        }
        elseif ($ch -eq ")") {
            if ($depth -gt 0) { $depth-- }
            [void]$sb.Append($ch)
            $i++
            continue
        }
        elseif ($ch -eq "," -and $depth -eq 0) {
            $args += $sb.ToString().Trim()
            [void]$sb.Clear()
            $i++
            continue
        }
        [void]$sb.Append($ch)
        $i++
    }
    if ($sb.Length -gt 0) {
        $args += $sb.ToString().Trim()
    }
    return $args
}

function Parse-IfcEntityLine {
    param([string]$Line)
    if ($Line -match '^\s*#(\d+)\s*=\s*([A-Za-z0-9_]+)\s*\((.*)\)\s*;\s*$') {
        return [pscustomobject]@{
            Id       = "#$($Matches[1])"
            Entity   = $Matches[2].ToUpperInvariant()
            ArgsText = $Matches[3]
        }
    }
    return $null
}

function Parse-IfcIdList {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text) -or $Text.Trim() -eq '$') { return @() }
    $trimmed = $Text.Trim()
    if ($trimmed.StartsWith("(") -and $trimmed.EndsWith(")")) {
        $trimmed = $trimmed.Substring(1, $trimmed.Length - 2)
    }
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return @() }
    return ($trimmed -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function Convert-ToIfcNominalValue {
    param(
        [object]$Value,
        [string]$DataType
    )
    if ($null -eq $Value) { return '$' }
    $text = $Value.ToString()
    if ([string]::IsNullOrWhiteSpace($text)) { return '$' }

    $dtype = if ($DataType) { $DataType.Trim().ToLowerInvariant() } else { 'text' }
    switch ($dtype) {
        { $_ -in @('number', 'float', 'double', 'real') } {
            $number = 0.0
            if (-not [double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
                throw "Value '$text' is not a valid number."
            }
            return ("IFCREAL({0})" -f $number.ToString([System.Globalization.CultureInfo]::InvariantCulture))
        }
        { $_ -in @('integer', 'int') } {
            $number = 0
            if (-not [int]::TryParse($text, [System.Globalization.NumberStyles]::Integer, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
                throw "Value '$text' is not a valid integer."
            }
            return ("IFCINTEGER({0})" -f $number.ToString([System.Globalization.CultureInfo]::InvariantCulture))
        }
        { $_ -in @('boolean', 'bool') } {
            $truthy = $false
            if ($Value -is [bool]) {
                $truthy = [bool]$Value
            }
            else {
                $truthy = $text.Trim().ToLowerInvariant() -in @('1','true','t','yes','y')
            }
            return ("IFCBOOLEAN({0})" -f (if ($truthy) { '.T.' } else { '.F.' }))
        }
        { $_ -in @('label', 'text', 'string') } {
            $escaped = Escape-IfcString $text
            return "IFCLABEL('$escaped')"
        }
        default {
            $escaped = Escape-IfcString $text
            return "IFCTEXT('$escaped')"
        }
    }
}

function Build-IfcEntityLine {
    param(
        [string]$Id,
        [string]$Entity,
        [array]$EntityArgs
    )
    return ("{0} = {1}({2});" -f $Id, $Entity, ($EntityArgs -join ', '))
}

function Build-IfcPropertySingleValueLine {
    param(
        [string]$Id,
        [string]$Name,
        [string]$ValueLiteral
    )
    $escapedName = Escape-IfcString $Name
    $args = @(
        "'$escapedName'",
        '$',
        $ValueLiteral,
        '$'
    )
    return Build-IfcEntityLine -Id $Id -Entity 'IFCPROPERTYSINGLEVALUE' -EntityArgs $args
}

function Build-IfcPropertySetLine {
    param(
        [string]$Id,
        [string]$OwnerHistoryId,
        [string]$Name,
        [string[]]$PropertyIds
    )
    $guid = New-IfcGuid
    $escapedName = Escape-IfcString $Name
    $propList = if ($PropertyIds -and $PropertyIds.Count -gt 0) {
        "({0})" -f ($PropertyIds -join ', ')
    } else {
        '$'
    }
    $args = @(
        "'$guid'",
        $OwnerHistoryId,
        "'$escapedName'",
        '$',
        $propList
    )
    return Build-IfcEntityLine -Id $Id -Entity 'IFCPROPERTYSET' -EntityArgs $args
}

function Build-IfcRelDefinesLine {
    param(
        [string]$Id,
        [string]$OwnerHistoryId,
        [string]$Name,
        [string]$ProjectId,
        [string]$PropertySetId
    )
    $guid = New-IfcGuid
    $escapedName = Escape-IfcString $Name
    $related = "($ProjectId)"
    $args = @(
        "'$guid'",
        $OwnerHistoryId,
        "'$escapedName'",
        '$',
        $related,
        $PropertySetId
    )
    return Build-IfcEntityLine -Id $Id -Entity 'IFCRELDEFINESBYPROPERTIES' -EntityArgs $args
}

function Set-IfcEntityNameLine {
<#
Updates the Name attribute (3rd IFC argument) for a single IFC entity line.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Line,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $parsed = Parse-IfcEntityLine -Line $Line
    if (-not $parsed) { return $Line }
    $escapedName = Escape-IfcString $Name
    $pattern = '^(?<prefix>\s*#\d+\s*=\s*[A-Za-z0-9_]+\(\s*''(?:(?:'''')|[^''])*''\s*,\s*(?:#\d+|\$)\s*,\s*)(?<current>\$|''(?:(?:'''')|[^''])*'')(?<suffix>\s*,.*\)\s*;\s*)$'
    if ($Line -match $pattern) {
        return ("{0}'{1}'{2}" -f $Matches['prefix'], $escapedName, $Matches['suffix'])
    }
    return $Line
}

function Update-IfcProjectProperties {
<#
Injects property sets and attributes into an IFC file (IfcProject scope).
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [Parameter(Mandatory = $true)][array]$PropertySets,
        [string]$TopLevelName
    )

    if (-not (Test-Path $SourcePath -PathType Leaf)) {
        throw "Input IFC not found: $SourcePath"
    }

    $destDir = Split-Path -Parent $DestinationPath
    if (-not (Test-Path $destDir -PathType Container)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    $tempName = "{0}.{1}.tmp" -f ([System.IO.Path]::GetFileName($DestinationPath)), ([System.Guid]::NewGuid().ToString("N"))
    $tempPath = Join-Path $destDir $tempName

    # One-pass stream rewrite:
    # - rename IFCPROJECT as soon as it is found
    # - track max entity id cheaply
    # - append new IFC entities immediately before ENDSEC of the DATA section
    $maxId = [long]0
    $projectId = $null
    $ownerHistoryId = $null
    $inDataSection = $false
    $dataEndsecFound = $false
    $projectNameUpdated = $false

    $reader = New-Object System.IO.StreamReader($SourcePath)
    $writer = New-Object System.IO.StreamWriter($tempPath, $false, [System.Text.Encoding]::ASCII)
    try {
        while (($line = $reader.ReadLine()) -ne $null) {
            $trimmedLine = $line.Trim()

            if (-not $inDataSection -and $trimmedLine -eq 'DATA;') {
                $inDataSection = $true
            }

            $idMatch = [regex]::Match($line, '^\s*#(?<id>\d+)\s*=')
            if ($idMatch.Success) {
                $idValue = [long]$idMatch.Groups['id'].Value
                if ($idValue -gt $maxId) { $maxId = $idValue }
            }

            if (-not $projectId -and $line.IndexOf('IFCPROJECT(', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $projectMatch = [regex]::Match($line, '^\s*#(?<id>\d+)\s*=\s*IFCPROJECT\(')
                if ($projectMatch.Success) {
                    $projectId = "#$($projectMatch.Groups['id'].Value)"
                    $ownerMatch = [regex]::Match($line, "^\s*#\d+\s*=\s*IFCPROJECT\(\s*'(?:(?:'')|[^'])*'\s*,\s*(?<owner>#\d+|\$)\s*,")
                    if ($ownerMatch.Success -and $ownerMatch.Groups['owner'].Value -match '^#\d+$') {
                        $ownerHistoryId = $ownerMatch.Groups['owner'].Value
                    }
                }
            }

            if (-not $projectNameUpdated -and -not [string]::IsNullOrWhiteSpace($TopLevelName) -and $projectId) {
                $projectLinePattern = '^\s*' + [regex]::Escape($projectId) + '\s*=\s*IFCPROJECT\('
                if ($line -match $projectLinePattern) {
                    $line = Set-IfcEntityNameLine -Line $line -Name $TopLevelName
                    $projectNameUpdated = $true
                }
            }

            if ($inDataSection -and $trimmedLine -eq 'ENDSEC;') {
                if (-not $projectId) {
                    throw "IFCPROJECT not found in: $SourcePath"
                }
                if (-not $ownerHistoryId) { $ownerHistoryId = "#2" }

                $insertLines = New-Object System.Collections.Generic.List[string]
                $nextId = $maxId + 1
                foreach ($psetInput in $PropertySets) {
                    if (-not $psetInput) { continue }
                    $psetName = $psetInput.name
                    $attributes = $psetInput.attributes
                    if ([string]::IsNullOrWhiteSpace($psetName) -or -not $attributes -or $attributes.Count -eq 0) { continue }

                    $propIds = New-Object System.Collections.Generic.List[string]
                    foreach ($attr in $attributes) {
                        if (-not $attr) { continue }
                        $propName = $attr.name
                        if ([string]::IsNullOrWhiteSpace($propName)) { continue }
                        if ($propName.Trim().ToLowerInvariant() -eq 'description') {
                            $propName = 'PWDescription'
                        }

                        $valueLiteral = $null
                        try {
                            $valueLiteral = Convert-ToIfcNominalValue -Value $attr.value -DataType $attr.data_type
                        }
                        catch {
                            Write-Warning "Skipping attribute '$propName' in '$psetName': $_"
                            continue
                        }
                        if ([string]::IsNullOrWhiteSpace($valueLiteral)) { continue }

                        $propId = "#$nextId"
                        $nextId++
                        $insertLines.Add((Build-IfcPropertySingleValueLine -Id $propId -Name $propName -ValueLiteral $valueLiteral))
                        $propIds.Add($propId)
                    }

                    if ($propIds.Count -eq 0) { continue }

                    $psetId = "#$nextId"
                    $nextId++
                    $insertLines.Add((Build-IfcPropertySetLine -Id $psetId -OwnerHistoryId $ownerHistoryId -Name $psetName -PropertyIds $propIds.ToArray()))

                    $relId = "#$nextId"
                    $nextId++
                    $relName = "{0}_Assignment" -f $psetName
                    $insertLines.Add((Build-IfcRelDefinesLine -Id $relId -OwnerHistoryId $ownerHistoryId -Name $relName -ProjectId $projectId -PropertySetId $psetId))
                }

                foreach ($insertLine in $insertLines) {
                    $writer.WriteLine($insertLine)
                }
                $writer.WriteLine($line)
                $dataEndsecFound = $true
                $inDataSection = $false
                continue
            }

            $writer.WriteLine($line)
        }
    }
    finally {
        if ($reader) { $reader.Dispose() }
        if ($writer) { $writer.Dispose() }
    }

    if (-not $projectId) {
        throw "IFCPROJECT not found in: $SourcePath"
    }
    if (-not $dataEndsecFound) {
        throw "ENDSEC for DATA section not found in: $SourcePath"
    }

    Move-Item -Path $tempPath -Destination $DestinationPath -Force
}

function Invoke-ProcessIfcAttributes {
<#
Processes IFC files and injects attributes based on configuration.
#>
    [CmdletBinding()]
    param(
        [string]$ConfigFile = "Config.json",
        [array]$Settings = @(),
        [Parameter(Mandatory = $true)][array]$AttributeDefinitionRows,
        [Parameter(Mandatory = $true)][array]$FederationRows,
        [array]$LookupRows,
        [Parameter(Mandatory = $true)][array]$AttributeRows,
        [Parameter(Mandatory = $true)][string]$AttributesWorkbookPath,
        [Parameter(Mandatory = $true)][string]$AttributesWorkbookTimestamp,
        [switch]$Force,
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

    $LogsFolder = if ($LogsFolder) { $LogsFolder } else { $logsFolder } # capture pre-set value (case-insensitive)

    # --- Load settings + resolve paths ---
    $ConfigFilePath = if ([System.IO.Path]::IsPathRooted($ConfigFile)) {
        $ConfigFile
    }
    else {
        Join-Path $basePath $ConfigFile
    }
    $Settings = @($Settings)
    $LogsFolder = if ($LogsFolder) { $LogsFolder } else { CureFolderPath( Get-ConfigValue -Settings $Settings -Name "LogFolder" -Default "Logs") }
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
    Write-Host "=== START 030-ProcessIfcAttributes ==="
    Write-Host "Process start time: $ProcessStartTime"
    Write-Host "Config file: $ConfigFilePath"

    # Honor RunProcess=Force at the function level in case the caller didn't enforce it.
    $runProcessKeys = @('runprocess','forceifcprocessing','forceifcprocess','forceprocess')
    $runProcessRow = $Settings | Where-Object {
        $paramRaw = $_.Parameter
        $paramText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $paramRaw
        }
        else {
            if ($null -ne $paramRaw) { $paramRaw.ToString().Trim() } else { $null }
        }
        if ([string]::IsNullOrWhiteSpace($paramText)) { return $false }
        $runProcessKeys -contains $paramText.ToLowerInvariant()
    } | Select-Object -First 1
    if (-not $Force -and (-not $runProcessRow -or $null -eq $runProcessRow.Value -or [string]::IsNullOrWhiteSpace($runProcessRow.Value.ToString()))) {
        Write-Host "Processing ignored by default (RunProcess missing)." -ForegroundColor Yellow
        return
    }
    if ($runProcessRow -and $null -ne $runProcessRow.Value) {
        $valueText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $runProcessRow.Value
        }
        else {
            $runProcessRow.Value.ToString().Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($valueText)) {
            $normalized = $valueText.ToLowerInvariant()
            if ($normalized -eq 'force') {
                $Force = $true
                Write-Host ("Processing forced by Settings parameter '{0}'." -f $runProcessRow.Parameter) -ForegroundColor Yellow
            }
            elseif ($normalized -in @('no','n','false','0','ignore')) {
                Write-Host ("Processing ignored by Settings parameter '{0}'." -f $runProcessRow.Parameter) -ForegroundColor Yellow
                return
            }
        }
    }

    $localFolder     = CureFolderPath( Get-ConfigValue -Settings $Settings -Name "SourceFolder" -Default "SourceFiles")
    $processedFolder = CureFolderPath (Get-ConfigValue -Settings $Settings -Name "ProcessedFolder" -Default "ProcessedIFC")
    $attributesFile  = if ($AttributesWorkbookPath) { Split-Path -Leaf $AttributesWorkbookPath } else { $null }

    Write-Host "Source folder: $localFolder"
    Write-Host "Processed folder: $processedFolder"
    Write-Host "Attributes file: $attributesFile"
    Write-Host "IFC injection: PowerShell (append-only)"

    function Format-ByteCount {
        param($Bytes)
        if ($null -eq $Bytes) { return "n/a" }
        $number = 0.0
        if (-not [double]::TryParse($Bytes.ToString(), [ref]$number)) { return $Bytes.ToString() }
        $mb = [math]::Round($number / 1MB, 2)
        return ("{0} bytes ({1} MB)" -f [long]$number, $mb)
    }

    foreach ($value in @($localFolder, $processedFolder, $AttributesWorkbookPath)) {
        if ([string]::IsNullOrWhiteSpace($value)) {
            throw "Missing required processing inputs (SourceFolder / ProcessedFolder / AttributesWorkbookPath)."
        }
    }

    if (-not (Test-Path $localFolder -PathType Container)) {
        throw "Source folder not found: $localFolder"
    }
    $localFolder = (Resolve-Path $localFolder).Path

    if (-not (Test-Path $ProcessedFolder)) {
        New-Item -ItemType Directory -Path $ProcessedFolder -Force | Out-Null
    }
    $ProcessedFolder = (Resolve-Path $ProcessedFolder).Path

    if (-not [System.IO.Path]::IsPathRooted($AttributesWorkbookPath)) {
        $AttributesWorkbookPath = Join-Path $basePath $AttributesWorkbookPath
    }
    if (-not (Test-Path $AttributesWorkbookPath -PathType Leaf)) {
        throw "Attribute workbook not found: $AttributesWorkbookPath"
    }
    $AttributesWorkbookPath = (Resolve-Path $AttributesWorkbookPath).Path
    if ([string]::IsNullOrWhiteSpace($AttributesWorkbookTimestamp)) {
        throw "Attributes workbook timestamp was not supplied to Invoke-ProcessIfcAttributes."
    }

    # --- Read attribute definitions ---
    $AttributeDefinitionSheet = 'PWAttributesList'
    $AttributeDataSheet = 'Attributes'

    $attributeSheet = @($AttributeDefinitionRows)
    $attributeDefs = @()
    foreach ($row in $attributeSheet) {
        if (-not $row) { continue }

        $injectFlag = Get-RowValue -Row $row -Candidates @('InjectToIFC','InjectToIfc','Inject2IFC','Inject2Ifc')
        $normalizedInject = if ($null -ne $injectFlag) { $injectFlag.Trim().ToLower() } else { '' }
        $shouldInclude = (-not [string]::IsNullOrWhiteSpace($normalizedInject)) -and ($normalizedInject -in @('yes','y','true','1'))
        if (-not $shouldInclude) { continue }

        $rowOrder = $null
        if ($row.PSObject.Properties.Name -contains 'Row') {
            [int]$parsedRow = 0
            if ([int]::TryParse($row.Row, [ref]$parsedRow)) { $rowOrder = $parsedRow }
        }
        if (-not $rowOrder) { $rowOrder = [int]$attributeDefs.Count + 1 }

        $sourceName = $null
        if ($row.PSObject.Properties.Name -contains 'AttributeName') {
            $sourceName = $row.AttributeName
        }
        elseif ($row.PSObject.Properties.Name -contains 'PWAttibuteName') {
            $sourceName = $row.PWAttibuteName
        }
        elseif ($row.PSObject.Properties.Name -contains 'PWAttributeName') {
            $sourceName = $row.PWAttributeName
        }
        $sourceName = Normalize-ExcelText -Text $sourceName
        if ([string]::IsNullOrWhiteSpace($sourceName)) { continue }

        $outputName = Get-RowValue -Row $row -Candidates @('OutputName')
        $outputName = Normalize-ExcelText -Text $outputName
        if ([string]::IsNullOrWhiteSpace($outputName)) { $outputName = $sourceName }

        $attributeDefs += [pscustomobject]@{
            RowOrder      = $rowOrder
            Name          = $outputName
            WorkbookColumn = $outputName
            SourceColumn  = $sourceName
            PropertySet   = 'PWattributes'
        }
    }
    $attributeDefs = $attributeDefs | Sort-Object RowOrder
    if ($attributeDefs.Count -eq 0) {
        throw "Attribute definition named range '$AttributeDefinitionSheet' does not contain usable rows."
    }

    # --- Load attribute workbook + build lookup by file-name column ---
    $attributeRows = @($AttributeRows)
    $matchColumn = $null
    foreach ($candidate in @('FileName', 'Filename')) {
        if ($attributeRows | Get-Member -MemberType NoteProperty | Where-Object Name -eq $candidate) {
            $matchColumn = $candidate
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($matchColumn)) {
        throw "Attribute named range '$AttributeDataSheet' must contain a file-name column ('FileName' or 'Filename')."
    }
    foreach ($row in $attributeRows) {
        Initialize-RowFieldMap -Row $row
    }
    $attributeLookup = @{}
    foreach ($row in $attributeRows) {
        $key = $row.$matchColumn
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $attributeLookup[$key] = $row
    }

    # Allow matching rows even when the spreadsheet drops the .ifc extension
    $attributeLookupNoExt = @{}
    foreach ($row in $attributeRows) {
        $key = $row.$matchColumn
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $base = [System.IO.Path]::GetFileNameWithoutExtension($key)
        if (-not [string]::IsNullOrWhiteSpace($base)) {
            $attributeLookupNoExt[$base] = $row
        }
    }

    # --- Load filename section definitions + lookup table for descriptions ---
    # The Federation range may also contain an informational Description column, which is ignored here.
    $federationSheet = @($FederationRows)
    $federationPartDefinitions = @(Get-FederationPartDefinitions -FederationRows $federationSheet)
    $federationMatchDefinitions = @($federationPartDefinitions | Where-Object { $_.GroupOrder -gt 0 -and -not $_.IsFileExtension })
    $fileNameSections = @(
        $federationPartDefinitions |
            Where-Object { $_.InjectToIFC } |
            ForEach-Object {
                [pscustomobject]@{
                    FieldName    = $_.Name
                    FileNamePart = $_.FileNamePart
                    Position     = $_.Position
                    IsFileExtension = $_.IsFileExtension
                }
            }
    )

    # Only Filename-Part, Code, and Description are read from the Lookups range.
    # Any extra columns are ignored by the current process logic.
    $lookupMap = @{}
    $lookupSheet = @($LookupRows)
    $lookupLoadFailed = $false
    $lookupWarningDetail = $null

    if (-not $lookupLoadFailed) {
        if ($lookupSheet.Count -eq 0) {
            $lookupWarningDetail = "Lookups named range is empty in '$ConfigFilePath'."
        }
        else {
            $missingColumns = @()
            if (-not ($lookupSheet | Get-Member -MemberType NoteProperty | Where-Object Name -eq 'Filename-Part')) {
                $missingColumns += "'Filename-Part'"
            }
            if (-not ($lookupSheet | Get-Member -MemberType NoteProperty | Where-Object Name -eq 'Code')) {
                $missingColumns += "'Code'"
            }
            if (-not ($lookupSheet | Get-Member -MemberType NoteProperty | Where-Object Name -eq 'Description')) {
                $missingColumns += "'Description'"
            }

            if ($missingColumns.Count -gt 0) {
                $lookupWarningDetail = ("Lookups named range is missing required column(s): {0}." -f ($missingColumns -join ", "))
            }
            else {
                foreach ($row in $lookupSheet) {
                    $partKey = Normalize-ExcelText -Text (Get-RowValue -Row $row -Candidates @('Filename-Part','FilenamePart','Filename_Part'))
                    $code = Normalize-ExcelText -Text $row.Code
                    $desc = Normalize-ExcelText -Text $row.Description
                    if ([string]::IsNullOrWhiteSpace($partKey) -or [string]::IsNullOrWhiteSpace($code) -or [string]::IsNullOrWhiteSpace($desc)) { continue }
                    $mapKey = ("{0}||{1}" -f $partKey, $code)
                    if (-not $lookupMap.ContainsKey($mapKey)) { $lookupMap[$mapKey] = $desc }
                }
            }
        }
    }

    if ($lookupWarningDetail) {
        $lookupWarning = "LOOKUPS TABLE WARNING: $lookupWarningDetail"
        $lookupFallback = "Filename section descriptions will default to 'N/A' for this run."
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
        Write-Host $lookupWarning -ForegroundColor Red
        Write-Host $lookupFallback -ForegroundColor Red
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
        Write-Warning $lookupWarning
        if (Get-Command Write-Log -ErrorAction SilentlyContinue) {
            Write-Log -Message $lookupWarning -Color 'Red' -Settings $Settings -LogsFolderOverride $LogsFolder -BasePath $basePath
            Write-Log -Message $lookupFallback -Color 'Red' -Settings $Settings -LogsFolderOverride $LogsFolder -BasePath $basePath
        }
    }

    # --- Load existing metadata (if present) ---
    $metadataPath = Join-Path $ProcessedFolder "processed-metadata.json"
    $metadata = @{}
    if (Test-Path $metadataPath -PathType Leaf) {
        # Try to rehydrate metadata, falling back to a manual conversion path when AsHashtable is unavailable
        try {
            $json = Get-Content -Path $metadataPath -Raw -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($json)) {
                $convertCmd = Get-Command ConvertFrom-Json -ErrorAction SilentlyContinue
                $supportsHashtable = $convertCmd -and $convertCmd.Parameters.ContainsKey('AsHashtable')
                if ($supportsHashtable) {
                    $metadata = ConvertFrom-Json -InputObject $json -AsHashtable
                }
                else {
                    $converted = ConvertFrom-Json -InputObject $json
                    $metadata = ConvertTo-Hashtable -InputObject $converted
                }
            }
        } catch {
            Write-Warning "Unable to parse metadata file. Rebuilding from scratch."
            $metadata = @{}
        }
    }

    $processedReport = Join-Path $ProcessedFolder "processed-attributes.csv"
    $processed = @()
    $skipped = @()
    $failures = @()

    # --- Enumerate federatable source files ---
    $federatableExtensions = if (Get-Command Get-FederatableModelExtensions -ErrorAction SilentlyContinue) {
        Get-FederatableModelExtensions
    }
    else {
        @('.ifc', '.dwg', '.dgn', '.rvt')
    }
    $passthroughExtensions = @($federatableExtensions | Where-Object { $_ -ne '.ifc' })

    # Enumerate all IFC sources we might need to revisit.
    $ifcFiles = @(Get-ChildItem -Path $localFolder -Filter "*.ifc" -File)
    $passthroughFiles = @(Get-ChildItem -Path $localFolder -File | Where-Object {
        $passthroughExtensions -contains $_.Extension.ToLowerInvariant()
    })
    $processActionTotal = $ifcFiles.Count + $passthroughFiles.Count
    $processActionCurrent = 0
    function Write-IfcProcessProgress {
        param(
            [string]$Status,
            [int]$Current,
            [int]$Total,
            [string]$FileName
        )
        $safeStatus = if ([string]::IsNullOrWhiteSpace($Status)) { 'Processing' } else { $Status }
        $safeFile = if ($FileName) { $FileName -replace '\|', ' ' } else { '' }
        Write-Host ("IFC_PROCESS_PROGRESS|Current={0}|Total={1}|Status={2}|File={3}" -f $Current, $Total, $safeStatus, $safeFile)
    }

    if ($ifcFiles.Count -eq 0 -and $passthroughFiles.Count -eq 0) {
        Write-Warning "No federatable files found in source folder: $localFolder"
        return
    }
    Write-IfcProcessProgress -Status 'Queued' -Current 0 -Total $processActionTotal -FileName ''

    foreach ($ifc in $ifcFiles) {
        Write-IfcProcessProgress -Status 'Processing' -Current $processActionCurrent -Total $processActionTotal -FileName $ifc.Name
        $relativePath = $ifc.FullName.Substring($localFolder.TrimEnd('\').Length + 1)
        $sourceTimestamp = $ifc.LastWriteTimeUtc.ToString("o")
        $destinationPath = Join-Path $ProcessedFolder $relativePath
        $fileNameMatchesPattern = Test-FederationFileNameMatch -FileName $ifc.Name -PartDefinitions $federationMatchDefinitions
        $nameParts = @(Split-ModelFileNameParts -FileName $ifc.Name)

        # Find the matching attribute row (exact match first, then by filename without extension)
        $matchRow = $null
        if ($attributeLookup.ContainsKey($ifc.Name)) {
            $matchRow = $attributeLookup[$ifc.Name]
        } else {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($ifc.Name)
            if ($attributeLookup.ContainsKey($base)) { $matchRow = $attributeLookup[$base] }
            elseif ($attributeLookupNoExt.ContainsKey($base)) { $matchRow = $attributeLookupNoExt[$base] }
        }

        if (-not $matchRow) {
            Write-Warning "No attribute workbook entry found for '$($ifc.Name)'; skipping."
            $skipped += $ifc.FullName
            $processActionCurrent++
            Write-IfcProcessProgress -Status 'Skipped' -Current $processActionCurrent -Total $processActionTotal -FileName $ifc.Name
            continue
        }

        $reportedFileSize = Convert-FileSizeValue -Value (Get-FirstRowFieldValue -Row $matchRow -ColumnNames @(
            'FileSize', 'File_Size', 'PW_File_Size'
        ))
        $reportedFileUpdate = Convert-DateValue -Value (Get-FirstRowFieldValue -Row $matchRow -ColumnNames @(
            'FileUpdateDate', 'File_UpdateDate', 'File UpdateDate', 'PW_File_UpdateDate'
        ))
        # Fallback to actual file info when spreadsheet values are missing so change detection still works
        if ($null -eq $reportedFileSize) { $reportedFileSize = $ifc.Length }
        if ($null -eq $reportedFileUpdate) { $reportedFileUpdate = $ifc.LastWriteTimeUtc.ToString("o") }

        # Build attribute payload (skip blank values)
        $attributeValues = @()
        foreach ($def in $attributeDefs) {
            $value = Get-RowFieldValue -Row $matchRow -ColumnName $def.WorkbookColumn
            if ($null -eq $value -and $def.PSObject.Properties.Name -contains 'SourceColumn' -and
                -not [string]::IsNullOrWhiteSpace($def.SourceColumn) -and $def.SourceColumn -ne $def.WorkbookColumn) {
                $value = Get-RowFieldValue -Row $matchRow -ColumnName $def.SourceColumn
            }
            if ($null -eq $value -and $def.Name -ne $def.WorkbookColumn) {
                $value = Get-RowFieldValue -Row $matchRow -ColumnName $def.Name
            }
            if ($null -eq $value) { continue }

            $stringValue = $null
            if (Should-FormatAsDate -Definition $def) {
                $stringValue = Convert-DateValue -Value $value
            }
            elseif (Should-FormatAsFileSize -Definition $def) {
                $stringValue = Convert-FileSizeToBits -Value $value
            }
            else {
                $stringValue = $value.ToString()
            }
            if ([string]::IsNullOrWhiteSpace($stringValue)) { continue }
            $attributeValues += [pscustomobject]@{
                Name        = $def.Name
                Value       = $stringValue
                DataType    = 'Text'
                PropertySet = $def.PropertySet
            }
        }

        if ($fileNameSections.Count -gt 0 -and $fileNameMatchesPattern) {
            for ($i = 0; $i -lt $fileNameSections.Count; $i++) {
                $section = $fileNameSections[$i]
                $fieldName = $section.FieldName
                $fileNamePart = $section.FileNamePart
                if ([string]::IsNullOrWhiteSpace($fieldName) -or [string]::IsNullOrWhiteSpace($fileNamePart)) { continue }

                $definition = [pscustomobject]@{
                    FileNamePart    = $fileNamePart
                    Position        = $section.Position
                    IsFileExtension = $section.IsFileExtension
                }
                $code = Get-FederationPartValue -FileName $ifc.Name -Definition $definition -NameParts $nameParts
                if ([string]::IsNullOrWhiteSpace($code) -and -not $section.IsFileExtension) {
                    $partIndex = $null
                    if (-not [int]::TryParse($fileNamePart, [ref]$partIndex)) {
                        Write-Warning "Invalid Filename-Part '$fileNamePart' for FieldNames '$fieldName'; skipping."
                        continue
                    }
                }
                $desc = $null
                if (-not [string]::IsNullOrWhiteSpace($code)) {
                    $lookupKey = ("{0}||{1}" -f $fileNamePart, $code)
                    if ($lookupMap.ContainsKey($lookupKey)) {
                        $desc = $lookupMap[$lookupKey]
                    }
                }
                if ([string]::IsNullOrWhiteSpace($desc)) { $desc = 'N/A' }
                $valueText = if ([string]::IsNullOrWhiteSpace($code)) { $desc } else { ("{0} = {1}" -f $code, $desc) }
                $attributeValues += [pscustomobject]@{
                    Name        = $fieldName
                    Value       = $valueText
                    DataType    = 'Text'
                    PropertySet = 'FileName_sections_Descriptions'
                }
            }
        }
        elseif ($fileNameSections.Count -gt 0) {
            Write-Host ("{0} >>>> filename-derived attributes skipped; name does not match Federation pattern" -f $ifc.Name) -ForegroundColor Yellow
        }

        if ($attributeValues.Count -eq 0) {
            Write-Warning "No attribute values found for '$($ifc.Name)' (all blank). Skipping."
            $skipped += $ifc.FullName
            $processActionCurrent++
            Write-IfcProcessProgress -Status 'Skipped' -Current $processActionCurrent -Total $processActionTotal -FileName $ifc.Name
            continue
        }


        # Create payload grouped by property set for JSON contract
        $propertySets = @()
        foreach ($group in ($attributeValues | Group-Object PropertySet)) {
            $props = @()
            foreach ($entry in $group.Group) {
                $props += [ordered]@{
                    name      = $entry.Name
                    value     = $entry.Value
                    data_type = $entry.DataType
                }
            }
            $propertySets += [ordered]@{
                name       = $group.Name
                attributes = $props
            }
        }
        $payload = [ordered]@{ property_sets = $propertySets }
        $payloadJson = $payload | ConvertTo-Json -Depth 10
        $payloadHash = Get-StringHash -Text $payloadJson

        $existing = $metadata[$relativePath]
        $shouldProcess = $true
        if (-not $Force -and $existing) {
            # Backfill missing legacy fields so the change detector works with older metadata files
            $backfilled = $false
            $existingReportedFileSize = Get-PropertyValue -Object $existing -Name 'ReportedFileSize'
            if ([string]::IsNullOrWhiteSpace($existingReportedFileSize)) {
                if ($null -ne $reportedFileSize) { Set-PropertyValue -Object $existing -Name 'ReportedFileSize' -Value $reportedFileSize; $backfilled = $true }
            }
            $existingReportedFileUpdate = Get-PropertyValue -Object $existing -Name 'ReportedFileUpdate'
            if ([string]::IsNullOrWhiteSpace($existingReportedFileUpdate)) {
                if ($null -ne $reportedFileUpdate) { Set-PropertyValue -Object $existing -Name 'ReportedFileUpdate' -Value $reportedFileUpdate; $backfilled = $true }
            }
            if ($backfilled) { $metadata[$relativePath] = $existing }

            $processedPath = Resolve-FEDAUTOStateFilePath `
                -StoredPath (Get-PropertyValue -Object $existing -Name 'ProcessedPath') `
                -RootFolder $ProcessedFolder `
                -RelativePath $(if (Get-PropertyValue -Object $existing -Name 'ProcessedRelativePath') { Get-PropertyValue -Object $existing -Name 'ProcessedRelativePath' } else { $relativePath })
            $processedMissing = [string]::IsNullOrWhiteSpace($processedPath) -or -not (Test-Path -LiteralPath $processedPath -PathType Leaf)
            if ($existing.ReportedFileSize -eq $reportedFileSize -and
                $existing.ReportedFileUpdate -eq $reportedFileUpdate -and
                $existing.AttributeHash -eq $payloadHash -and
                -not $processedMissing) {
                Write-Host $ifc.Name -NoNewline -ForegroundColor Green
                Write-Host " Skipping >>>> No changes detected" -ForegroundColor Red
                $skipped += $ifc.FullName
                $shouldProcess = $false
            }
        }

        if (-not $shouldProcess) {
            $processActionCurrent++
            Write-IfcProcessProgress -Status 'Skipped' -Current $processActionCurrent -Total $processActionTotal -FileName $ifc.Name
            continue
        }

        # Ensure destination path mirrors source structure
        $destinationDir = Split-Path -Parent $destinationPath
        if (-not (Test-Path $destinationDir)) {
            New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
        }

        $processStart = Get-Date
        $sourceSize = $ifc.Length
        $tabNames = @($propertySets | ForEach-Object { $_.name } | Where-Object { $_ }) | Sort-Object -Unique
        $tabsAdded = $tabNames.Count
        $processSucceeded = $false
        $processError = $null

        Write-host $ifc.Name -NoNewline -ForegroundColor Green
        Write-host " Processing >>>> " -NoNewline -ForegroundColor Red
        Write-host ($sourceSize.ToString() + " ...bytes ") -NoNewline -ForegroundColor Red
        Write-host   -ForegroundColor White

        try {
            Update-IfcProjectProperties -SourcePath $ifc.FullName -DestinationPath $destinationPath -PropertySets $propertySets -TopLevelName $ifc.Name
            $processSucceeded = $true
        }
        catch {
            $processError = $_
            Write-Warning "Failed to process '$($ifc.Name)': $_"
            $failures += $ifc.FullName
        }
        finally {
            $processEnd = Get-Date
            $processDuration = New-TimeSpan -Start $processStart -End $processEnd
            if ($processSucceeded) {
                $destInfo = if (Test-Path $destinationPath -PathType Leaf) { Get-Item $destinationPath } else { $null }
                $finalSize = if ($destInfo) { $destInfo.Length } else { $null }

                Write-Host " >>>> processed - $($ifc.Name)" -ForegroundColor White
                Write-Host (" >>>> Tabs added: {0}" -f $tabsAdded) -ForegroundColor White
                if ($tabsAdded -gt 0) {
                    Write-Host (" >>>> Tabs: {0}" -f ($tabNames -join ", ")) -ForegroundColor White
                }
                Write-Host (" >>>> Size: {0} -> {1}" -f (Format-ByteCount $sourceSize), (Format-ByteCount $finalSize)) -ForegroundColor White
            }
            else {
                Write-Host " >>>> FAILED - $($ifc.Name)" -ForegroundColor Red
                if ($processError) {
                    Write-Host (" >>>> Error: {0}" -f $processError) -ForegroundColor Red
                }
            }

            Write-Host " >>>> Duration: $([math]::Round($processDuration.TotalSeconds,2)) seconds"
        }
        $processActionCurrent++
        $processStatus = if ($processSucceeded) { 'Processed' } else { 'Failed' }
        Write-IfcProcessProgress -Status $processStatus -Current $processActionCurrent -Total $processActionTotal -FileName $ifc.Name

        if ($processSucceeded) {
            $processed += $destinationPath
            $metadata[$relativePath] = [ordered]@{
                SourcePath         = $ifc.FullName
                SourceRelativePath = $relativePath
                ProcessedPath      = $destinationPath
                ProcessedRelativePath = $relativePath
                SourceTimestamp    = $sourceTimestamp
                ReportedFileSize   = $reportedFileSize
                ReportedFileUpdate = $reportedFileUpdate
                AttributeHash      = $payloadHash
                FederationPatternMatched = $fileNameMatchesPattern
                ProcessedTimestamp = (Get-Date).ToUniversalTime().ToString("o")
                AttributesFile     = $AttributesWorkbookPath
                AttributesFileName = [System.IO.Path]::GetFileName($AttributesWorkbookPath)
                AttributesModified = $AttributesWorkbookTimestamp
            }
        }
    }

    foreach ($file in $passthroughFiles) {
        Write-IfcProcessProgress -Status 'Copying' -Current $processActionCurrent -Total $processActionTotal -FileName $file.Name
        $relativePath = $file.FullName.Substring($localFolder.TrimEnd('\').Length + 1)
        $sourceTimestamp = $file.LastWriteTimeUtc.ToString("o")
        $sourceSize = $file.Length
        $destinationPath = Join-Path $ProcessedFolder $relativePath
        $destinationDir = Split-Path -Parent $destinationPath
        if (-not (Test-Path $destinationDir)) {
            New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
        }

        $existing = $metadata[$relativePath]
        $shouldCopy = $true
        if (-not $Force -and $existing) {
            $processedPath = Resolve-FEDAUTOStateFilePath `
                -StoredPath (Get-PropertyValue -Object $existing -Name 'ProcessedPath') `
                -RootFolder $ProcessedFolder `
                -RelativePath $(if (Get-PropertyValue -Object $existing -Name 'ProcessedRelativePath') { Get-PropertyValue -Object $existing -Name 'ProcessedRelativePath' } else { $relativePath })
            $processedMissing = [string]::IsNullOrWhiteSpace($processedPath) -or -not (Test-Path -LiteralPath $processedPath -PathType Leaf)
            $existingSourceSize = Get-PropertyValue -Object $existing -Name 'SourceSize'
            if ($null -eq $existingSourceSize) {
                $existingSourceSize = Get-PropertyValue -Object $existing -Name 'ReportedFileSize'
            }
            $existingSourceTimestamp = Get-PropertyValue -Object $existing -Name 'SourceTimestamp'
            if ($existingSourceTimestamp -eq $sourceTimestamp -and
                $existingSourceSize -eq $sourceSize -and
                -not $processedMissing) {
                Write-Host $file.Name -NoNewline -ForegroundColor Green
                Write-Host " Skipping >>>> No changes detected (pass-through)" -ForegroundColor Red
                $skipped += $file.FullName
                $shouldCopy = $false
            }
        }

        if (-not $shouldCopy) {
            $processActionCurrent++
            Write-IfcProcessProgress -Status 'Skipped' -Current $processActionCurrent -Total $processActionTotal -FileName $file.Name
            continue
        }

        $copyStart = Get-Date
        $copySucceeded = $false
        $copyError = $null

        Write-Host $file.Name -NoNewline -ForegroundColor Green
        Write-Host " Copying >>>> pass-through for federation" -ForegroundColor Red

        try {
            Copy-Item -Path $file.FullName -Destination $destinationPath -Force
            $copySucceeded = $true
        }
        catch {
            $copyError = $_
            Write-Warning "Failed to copy '$($file.Name)' for federation: $_"
            $failures += $file.FullName
        }
        finally {
            $copyEnd = Get-Date
            $copyDuration = New-TimeSpan -Start $copyStart -End $copyEnd
            if ($copySucceeded) {
                Write-Host " >>>> copied - $($file.Name)" -ForegroundColor White
            }
            else {
                Write-Host " >>>> FAILED - $($file.Name)" -ForegroundColor Red
                if ($copyError) {
                    Write-Host (" >>>> Error: {0}" -f $copyError) -ForegroundColor Red
                }
            }
            Write-Host " >>>> Duration: $([math]::Round($copyDuration.TotalSeconds,2)) seconds"
        }
        $processActionCurrent++
        $copyStatus = if ($copySucceeded) { 'Copied' } else { 'Failed' }
        Write-IfcProcessProgress -Status $copyStatus -Current $processActionCurrent -Total $processActionTotal -FileName $file.Name

        if ($copySucceeded) {
            $processed += $destinationPath
            $metadata[$relativePath] = [ordered]@{
                SourcePath         = $file.FullName
                SourceRelativePath = $relativePath
                ProcessedPath      = $destinationPath
                ProcessedRelativePath = $relativePath
                SourceTimestamp    = $sourceTimestamp
                SourceSize         = $sourceSize
                ReportedFileSize   = $sourceSize
                ReportedFileUpdate = $sourceTimestamp
                AttributeHash      = ''
                Passthrough        = $true
                ProcessedTimestamp = (Get-Date).ToUniversalTime().ToString("o")
                AttributesFile     = $AttributesWorkbookPath
                AttributesFileName = [System.IO.Path]::GetFileName($AttributesWorkbookPath)
                AttributesModified = $AttributesWorkbookTimestamp
            }
        }
    }

    # --- Save metadata + summary ---
    $metadataJson = $metadata | ConvertTo-Json -Depth 10
    Set-Content -Path $metadataPath -Value $metadataJson -Encoding UTF8

    $summary = [ordered]@{
        Processed  = $processed.Count
        Skipped    = $skipped.Count
        Failed     = $failures.Count
        Timestamp  = (Get-Date).ToUniversalTime().ToString("o")
        Files      = $processed
        RelativeFiles = @($processed | ForEach-Object { ConvertTo-FEDAUTORelativePath -Path $_ -RootFolder $ProcessedFolder })
        SkippedFiles = $skipped
        RelativeSkippedFiles = @($skipped | ForEach-Object { ConvertTo-FEDAUTORelativePath -Path $_ -RootFolder $localFolder })
        FailedFiles = $failures
        RelativeFailedFiles = @($failures | ForEach-Object { ConvertTo-FEDAUTORelativePath -Path $_ -RootFolder $localFolder })
    }
    $summaryJson = $summary | ConvertTo-Json -Depth 10
    $summaryPath = Join-Path $ProcessedFolder "processed-summary.json"
    Set-Content -Path $summaryPath -Value $summaryJson -Encoding UTF8

    if ($failures.Count -gt 0) {
        $failures | ForEach-Object { Write-Host "Failed: $_" }
    }

    $ProcessEndTime = Get-Date
    $TotalDuration = New-TimeSpan -Start $ProcessStartTime -End $ProcessEndTime
    Write-Host "Process finish time: $ProcessEndTime"
    Write-Host "Total duration: $([math]::Round($TotalDuration.TotalMinutes,2)) minutes"
    Write-Host "=== END 030-ProcessIfcAttributes ==="
    if ($TranscriptStarted) {
        Stop-Transcript | Out-Null
        $global:MainTranscriptActive = $false
    }
}


