<# Builds the self-contained GUI launcher. #>
[CmdletBinding()]
param([string]$OutputFile)

if ($PSVersionTable.PSEdition -ne 'Desktop') {
    & (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') -NoProfile -File $PSCommandPath -OutputFile $OutputFile
    exit $LASTEXITCODE
}
$basePath = Split-Path -Parent $PSCommandPath
if (-not $OutputFile) { $OutputFile = Join-Path $basePath 'FA_GUI.exe' }
$inputFile = Join-Path $basePath '007-Gui.ps1'
$pipelineFile = Join-Path $basePath 'FA_Main.exe'
if (-not (Test-Path -LiteralPath $pipelineFile -PathType Leaf)) {
    throw "Pipeline executable not found: $pipelineFile. Build it with 000-2Exe.ps1 before building the GUI launcher."
}

function New-FEDAUTOBuildVersion {
    param(
        [Parameter(Mandatory = $true)][string]$ExeName,
        [Parameter(Mandatory = $true)][string]$RootPath
    )

    $statePath = Join-Path $RootPath 'BuildVersions.json'
    $today = Get-Date -Format 'yyyyMMdd'
    $state = [pscustomobject]@{ Date = $today; LastIndex = 0; Builds = @() }
    if (Test-Path -LiteralPath $statePath -PathType Leaf) {
        try {
            $existing = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json -ErrorAction Stop
            if ($existing -and $existing.Date -eq $today) { $state = $existing }
        }
        catch {
            Write-Warning "Could not read build-version state '$statePath'. Starting a new counter for today. Error: $_"
        }
    }

    $lastIndex = 0
    if ($state.PSObject.Properties.Name -contains 'LastIndex') { [void][int]::TryParse($state.LastIndex.ToString(), [ref]$lastIndex) }
    $nextIndex = $lastIndex + 1
    $version = '{0}-{1:D3}' -f $today, $nextIndex
    $builds = @()
    if ($state.PSObject.Properties.Name -contains 'Builds') { $builds = @($state.Builds) }
    $builds += [pscustomobject]@{
        ExeName = $ExeName
        Version = $version
        BuiltAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
    [pscustomobject]@{
        Date = $today
        LastIndex = $nextIndex
        Builds = @($builds | Select-Object -Last 200)
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding UTF8
    return $version
}

function New-FEDAUTOCompiledInputWithBuildMetadata {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [Parameter(Mandatory = $true)][string]$ExeName,
        [Parameter(Mandatory = $true)][string]$BuildVersion
    )

    $lines = Get-Content -LiteralPath $SourcePath
    $paramStart = $null
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*param\s*\(') { $paramStart = $i; break }
    }
    if ($null -eq $paramStart) { throw "Could not find param() block in $SourcePath" }

    $parenDepth = 0
    $paramEnd = $null
    for ($i = $paramStart; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $parenDepth += ($line.ToCharArray() | Where-Object { $_ -eq '(' }).Count
        $parenDepth -= ($line.ToCharArray() | Where-Object { $_ -eq ')' }).Count
        if ($parenDepth -le 0) { $paramEnd = $i; break }
    }
    if ($null -eq $paramEnd) { throw "Could not find end of param() block in $SourcePath" }

    $header = $lines[0..$paramEnd] -join "`r`n"
    $body = $lines[($paramEnd + 1)..($lines.Count - 1)] -join "`r`n"
    $buildTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $buildMetadata = @"
`$script:FEDAUTOExecutableName = '$ExeName'
`$script:FEDAUTOExecutableBuildVersion = '$BuildVersion'
`$script:FEDAUTOExecutableBuildTime = '$buildTime'
"@
    Set-Content -LiteralPath $DestinationPath -Value ($header + "`r`n`r`n" + $buildMetadata + "`r`n`r`n" + $body) -Encoding UTF8
}

$exeName = [IO.Path]::GetFileName($OutputFile)
$buildVersion = New-FEDAUTOBuildVersion -ExeName $exeName -RootPath $basePath
$bundleDir = Join-Path ([System.IO.Path]::GetTempPath()) 'FEDAUTO_Gui_Build'
if (Test-Path -LiteralPath $bundleDir) { Remove-Item -LiteralPath $bundleDir -Recurse -Force }
New-Item -Path $bundleDir -ItemType Directory -Force | Out-Null
$compiledInput = Join-Path $bundleDir '007-Gui.compiled.ps1'
New-FEDAUTOCompiledInputWithBuildMetadata -SourcePath $inputFile -DestinationPath $compiledInput -ExeName $exeName -BuildVersion $buildVersion
$inputFile = $compiledInput

Import-Module ps2exe -ErrorAction Stop
$resources = @{}
foreach ($file in '012-SharedFunctions.Ps1','013-ConfigFunctions.Ps1','041-FederationFunctions.Ps1') {
    $path = Join-Path $basePath $file
    if (-not (Test-Path $path)) { throw "Support file not found: $path" }
    $resources["%APPDATA%\FEDAUTO\$file"] = $path
}
$command = Get-Command Invoke-PS2EXE
$resourceParameter = if ($command.Parameters.ContainsKey('ResourceFile')) { 'ResourceFile' } elseif ($command.Parameters.ContainsKey('Include')) { 'Include' } elseif ($command.Parameters.ContainsKey('embedFiles')) { 'embedFiles' } else { $null }
$splat = @{ InputFile = $inputFile; OutputFile = $OutputFile; NoConsole = $true; RequireAdmin = $false; Nested = $true }
if ($resourceParameter) { $splat[$resourceParameter] = $resources } else { Write-Warning 'This PS2EXE version cannot embed support files; keep 012-SharedFunctions.Ps1, 013-ConfigFunctions.ps1, and 041-FederationFunctions.Ps1 beside the EXE.' }
Invoke-PS2EXE @splat
if (-not (Test-Path $OutputFile)) { throw "Build failed: $OutputFile was not created." }
$outputDirectory = Split-Path -Parent $OutputFile
$deployedPipeline = Join-Path $outputDirectory 'FA_Main.exe'
if (([IO.Path]::GetFullPath($pipelineFile)) -ne ([IO.Path]::GetFullPath($deployedPipeline))) {
    Copy-Item -LiteralPath $pipelineFile -Destination $deployedPipeline -Force
}
Write-Host "Built GUI launcher: $OutputFile" -ForegroundColor Green
Write-Host "Build version: $buildVersion" -ForegroundColor Green
