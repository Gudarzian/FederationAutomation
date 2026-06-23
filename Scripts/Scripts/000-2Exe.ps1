<#
Builds 006-Main.exe and bundles supporting scripts/resources so Config.json or Config.xlsx
must sit beside the EXE at runtime.
#>
[CmdletBinding()]
param(
    [string]$OutputFile
)

# ps2exe 1.x compilation can silently fail under pwsh; always build under Windows PowerShell.
if ($PSVersionTable.PSEdition -ne 'Desktop') {
    $desktopPwsh = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path $desktopPwsh)) {
        throw "Windows PowerShell executable not found at '$desktopPwsh'."
    }
    Write-Host "Relaunching build under Windows PowerShell (Desktop edition)..." -ForegroundColor Yellow
    $relaunchArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    if ($OutputFile) {
        $relaunchArgs += @('-OutputFile', $OutputFile)
    }
    & $desktopPwsh @relaunchArgs
    exit $LASTEXITCODE
}

$basePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$inputFile  = Join-Path $basePath '006-Main.ps1'
if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $OutputFile = Join-Path $basePath '006-Main.exe'
}
else {
    $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
}
$outputDir = Split-Path -Parent $OutputFile
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

# Files to bundle into the EXE
$includes = @(
    '012-SharedFunctions.Ps1',
    '013-ConfigFunctions.ps1',
    '021-DownloadFunctions.Ps1',
    '031-ProcessFunctions.Ps1',
    '041-FederationFunctions.Ps1',
    '011-FunctionsDepository.Ps1'
)
$includePaths = $includes | ForEach-Object { Join-Path $basePath $_ }

# Stage bundled files in a temp folder without spaces to avoid PS2EXE parsing issues
$bundleDir = Join-Path ([System.IO.Path]::GetTempPath()) "006Main_bundle"
if (Test-Path $bundleDir) {
    Remove-Item -Path $bundleDir -Recurse -Force
}
New-Item -Path $bundleDir -ItemType Directory -Force | Out-Null
$stagedIncludes = @{}
foreach ($path in $includePaths) {
    if (-not (Test-Path $path)) { throw "Support file not found: $path" }
    $dest = Join-Path $bundleDir (Split-Path $path -Leaf)
    Copy-Item -Path $path -Destination $dest -Force
    $leaf = Split-Path $dest -Leaf
    $target = "%APPDATA%\\Pythonx\\$leaf"
    $stagedIncludes[$target] = $dest
}

# Build a temporary input that inlines the function scripts so they don't need to extract at runtime.
$functionFiles = @(
    '012-SharedFunctions.Ps1',
    '013-ConfigFunctions.ps1',
    '021-DownloadFunctions.Ps1',
    '031-ProcessFunctions.Ps1',
    '041-FederationFunctions.Ps1'
)
$compiledInput = Join-Path $bundleDir '006-Main.compiled.ps1'
$functionsContent = ($functionFiles | ForEach-Object {
    $path = Join-Path $basePath $_
    if (-not (Test-Path $path)) { throw "Support file not found: $path" }
    Get-Content -Path $path -Raw
}) -join "`r`n`r`n"
$mainLines = Get-Content -Path $inputFile

# Insert functions right after the param block so CmdletBinding stays at the top.
$paramStart = $null
for ($i = 0; $i -lt $mainLines.Count; $i++) {
    if ($mainLines[$i] -match '^\s*param\s*\(') {
        $paramStart = $i
        break
    }
}
if ($null -eq $paramStart) {
    throw "Could not find param() block in $inputFile"
}
$parenDepth = 0
$paramEnd = $null
for ($i = $paramStart; $i -lt $mainLines.Count; $i++) {
    $line = $mainLines[$i]
    $parenDepth += ($line.ToCharArray() | Where-Object { $_ -eq '(' }).Count
    $parenDepth -= ($line.ToCharArray() | Where-Object { $_ -eq ')' }).Count
    if ($parenDepth -le 0) {
        $paramEnd = $i
        break
    }
}
if ($null -eq $paramEnd) {
    throw "Could not find end of param() block in $inputFile"
}

$header = $mainLines[0..$paramEnd] -join "`r`n"
$body = $mainLines[($paramEnd + 1)..($mainLines.Count - 1)] -join "`r`n"
$compiledContent = $header + "`r`n`r`n" + $functionsContent + "`r`n`r`n" + $body
Set-Content -Path $compiledInput -Value $compiledContent -Encoding utf8
$inputFile = $compiledInput

# Ensure PS2EXE is available (install if missing)
$ps2exeCmd = Get-Command -Name Invoke-PS2EXE -ErrorAction SilentlyContinue
if (-not $ps2exeCmd) {
    try {
        Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ps2exe -ErrorAction Stop
        $ps2exeCmd = Get-Command -Name Invoke-PS2EXE -ErrorAction Stop
    }
    catch {
        throw "PS2EXE module is required but could not be loaded: $_"
    }
}

# Pick whichever bundling parameter this PS2EXE version supports
$paramNames = $ps2exeCmd.Parameters.Keys
$bundleParam =
    if ($paramNames -contains 'ResourceFile') { 'ResourceFile' }
    elseif ($paramNames -contains 'Include') { 'Include' }
    elseif ($paramNames -contains 'embedFiles') { 'embedFiles' }
    else { $null }

$splat = @{
    InputFile    = $inputFile
    OutputFile   = $outputFile
    NoConsole    = $false
    RequireAdmin = $false
    Nested       = $true
}
if ($bundleParam) {
    $splat[$bundleParam] = $stagedIncludes
}
else {
    Write-Warning "This PS2EXE version does not expose a bundling parameter (Include/ResourceFile/embedFiles); dependencies will not be bundled."
}

Invoke-PS2EXE @splat
if (-not (Test-Path $outputFile)) {
    throw "Build failed: output EXE was not created at '$outputFile'."
}
Write-Host "Built: $outputFile" -ForegroundColor Green
