<#
Compatibility wrapper for building the Navisworks visual-style add-in.
The build logic and embedded C# project/source live in 011-FunctionsDepository.Ps1.
#>
[CmdletBinding()]
param(
    [string]$NavisworksVersion = '',
    [string]$OutputDirectory
)

$basePath = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $basePath '011-FunctionsDepository.Ps1')

$splat = @{}
if (-not [string]::IsNullOrWhiteSpace($NavisworksVersion)) {
    $splat.NavisworksVersion = $NavisworksVersion
}
if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $splat.OutputDirectory = $OutputDirectory
}

Build-FEDAUTOVisualStylePlugin @splat | Out-Null
