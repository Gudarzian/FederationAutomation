<#
Legacy wrapper for Federation Automation function files.
Loads the split function files when available.
#>
$basePath = if ($PSScriptRoot) { $PSScriptRoot }
elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
elseif ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') { Split-Path -Parent $MyInvocation.MyCommand.Path }
else {
    Add-Type -AssemblyName System.Reflection
    Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location)
}

$functionFiles = @(
    "012-SharedFunctions.Ps1",
    "021-DownloadFunctions.Ps1",
    "031-ProcessFunctions.Ps1",
    "041-FederationFunctions.Ps1"
)

foreach ($file in $functionFiles) {
    $path = Join-Path $basePath $file
    if (-not (Test-Path $path)) {
        throw "Support file not found: $file"
    }
    . $path
}
