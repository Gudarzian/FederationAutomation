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
Import-Module ps2exe -ErrorAction Stop
$resources = @{}
foreach ($file in '012-SharedFunctions.Ps1','013-ConfigFunctions.Ps1') {
    $path = Join-Path $basePath $file
    if (-not (Test-Path $path)) { throw "Support file not found: $path" }
    $resources["%APPDATA%\FEDAUTO\$file"] = $path
}
$command = Get-Command Invoke-PS2EXE
$resourceParameter = if ($command.Parameters.ContainsKey('ResourceFile')) { 'ResourceFile' } elseif ($command.Parameters.ContainsKey('Include')) { 'Include' } elseif ($command.Parameters.ContainsKey('embedFiles')) { 'embedFiles' } else { $null }
$splat = @{ InputFile = $inputFile; OutputFile = $OutputFile; NoConsole = $true; RequireAdmin = $false; Nested = $true }
if ($resourceParameter) { $splat[$resourceParameter] = $resources } else { Write-Warning 'This PS2EXE version cannot embed support files; keep 012-SharedFunctions.Ps1 and 013-ConfigFunctions.ps1 beside the EXE.' }
Invoke-PS2EXE @splat
if (-not (Test-Path $OutputFile)) { throw "Build failed: $OutputFile was not created." }
$outputDirectory = Split-Path -Parent $OutputFile
$deployedPipeline = Join-Path $outputDirectory 'FA_Main.exe'
if (([IO.Path]::GetFullPath($pipelineFile)) -ne ([IO.Path]::GetFullPath($deployedPipeline))) {
    Copy-Item -LiteralPath $pipelineFile -Destination $deployedPipeline -Force
}
Write-Host "Built GUI launcher: $OutputFile" -ForegroundColor Green
