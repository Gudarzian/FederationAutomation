<#
Builds the Navisworks add-in used by federation runs to persist the required
Full Render and graduated-background appearance in saved output files.
#>
[CmdletBinding()]
param(
    [string]$NavisworksVersion = '2026',
    [string]$OutputDirectory
)

$basePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectPath = Join-Path $basePath 'NavisworksVisualStylePlugin\NavisworksVisualStylePlugin.csproj'
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $basePath 'NavisworksVisualStylePlugin\bin\Release\net48'
}

$navisworksApiPath = Join-Path "C:\Program Files\Autodesk\Navisworks Manage $NavisworksVersion" 'Autodesk.Navisworks.Api.dll'
if (-not (Test-Path -LiteralPath $navisworksApiPath -PathType Leaf)) {
    throw "Navisworks API assembly was not found: $navisworksApiPath"
}

& dotnet build $projectPath --configuration Release "/p:NavisworksApiPath=$navisworksApiPath"
if ($LASTEXITCODE -ne 0) { throw 'Navisworks visual-style plug-in build failed.' }

$pluginPath = Join-Path $OutputDirectory 'FederationAutomation.NavisworksVisualStyle.dll'
if (-not (Test-Path -LiteralPath $pluginPath -PathType Leaf)) {
    throw "Navisworks visual-style plug-in was not produced: $pluginPath"
}
Write-Host "Built: $pluginPath"
