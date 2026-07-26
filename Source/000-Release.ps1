<#
Creates or verifies a consistent Federation Automation release.
Root files are the only source of truth. By default this synchronises source,
configuration, templates, and documentation without rebuilding executables.
#>
[CmdletBinding()]
param(
    [switch]$BuildExecutables,
    [switch]$Verify
)

$ErrorActionPreference = 'Stop'
$basePath = Split-Path -Parent $PSCommandPath
$manifestPath = Join-Path $basePath 'ReleaseManifest.json'

$releaseFiles = @(
    @{ Source='006-Main.ps1'; Destination='PowerShell_Run_Package\006-Main.ps1' },
    @{ Source='007-Gui.ps1'; Destination='PowerShell_Run_Package\007-Gui.ps1' },
    @{ Source='012-SharedFunctions.Ps1'; Destination='PowerShell_Run_Package\012-SharedFunctions.Ps1' },
    @{ Source='013-ConfigFunctions.ps1'; Destination='PowerShell_Run_Package\013-ConfigFunctions.ps1' },
    @{ Source='021-DownloadFunctions.ps1'; Destination='PowerShell_Run_Package\021-DownloadFunctions.ps1' },
    @{ Source='031-ProcessFunctions.Ps1'; Destination='PowerShell_Run_Package\031-ProcessFunctions.Ps1' },
    @{ Source='041-FederationFunctions.Ps1'; Destination='PowerShell_Run_Package\041-FederationFunctions.Ps1' },
    @{ Source='051-IfcDataExtractionFunctions.ps1'; Destination='PowerShell_Run_Package\051-IfcDataExtractionFunctions.ps1' },
    @{ Source='Config.json'; Destination='PowerShell_Run_Package\Config.json' },
    @{ Source='NavisworksOptions.xml'; Destination='PowerShell_Run_Package\NavisworksOptions.xml' },
    @{ Source='000-2Exe.ps1'; Destination='GitHub_Library\Source\000-2Exe.ps1' },
    @{ Source='000-Gui2Exe.ps1'; Destination='GitHub_Library\Source\000-Gui2Exe.ps1' },
    @{ Source='000-Release.ps1'; Destination='GitHub_Library\Source\000-Release.ps1' },
    @{ Source='000-BuildNavisworksVisualStylePlugin.ps1'; Destination='GitHub_Library\Source\000-BuildNavisworksVisualStylePlugin.ps1' },
    @{ Source='006-Main.ps1'; Destination='GitHub_Library\Source\006-Main.ps1' },
    @{ Source='007-Gui.ps1'; Destination='GitHub_Library\Source\007-Gui.ps1' },
    @{ Source='011-FunctionsDepository.Ps1'; Destination='GitHub_Library\Source\011-FunctionsDepository.Ps1' },
    @{ Source='012-SharedFunctions.Ps1'; Destination='GitHub_Library\Source\012-SharedFunctions.Ps1' },
    @{ Source='013-ConfigFunctions.ps1'; Destination='GitHub_Library\Source\013-ConfigFunctions.ps1' },
    @{ Source='021-DownloadFunctions.ps1'; Destination='GitHub_Library\Source\021-DownloadFunctions.ps1' },
    @{ Source='031-ProcessFunctions.Ps1'; Destination='GitHub_Library\Source\031-ProcessFunctions.Ps1' },
    @{ Source='041-FederationFunctions.Ps1'; Destination='GitHub_Library\Source\041-FederationFunctions.Ps1' },
    @{ Source='051-IfcDataExtractionFunctions.ps1'; Destination='GitHub_Library\Source\051-IfcDataExtractionFunctions.ps1' },
    @{ Source='Config.json'; Destination='GitHub_Library\Exe_Files\Config.json' },
    @{ Source='NavisworksOptions.xml'; Destination='GitHub_Library\Exe_Files\NavisworksOptions.xml' },
    @{ Source='Generic_Config.json'; Destination='GitHub_Library\Templates\Generic_Config.json' },
    @{ Source='DependencyVersions.json'; Destination='GitHub_Library\DependencyVersions.json' },
    @{ Source='Docs\UserManual.md'; Destination='GitHub_Library\Docs\UserManual.md' },
    @{ Source='Docs\README-NoExcel-JSON-CSV.md'; Destination='GitHub_Library\Docs\README-NoExcel-JSON-CSV.md' },
    @{ Source='Docs\Federation-Automation-User-Manual-Friendly.md'; Destination='GitHub_Library\Docs\Federation-Automation-User-Manual-Friendly.md' }
)

function Get-ReleaseHash {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Test-ReleaseInputs {
    $scriptFiles = @($releaseFiles.Source | Where-Object { $_ -like '*.ps1' } | Select-Object -Unique)
    foreach ($relativePath in $scriptFiles) {
        $path = Join-Path $basePath $relativePath
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required source file is missing: $relativePath" }
        $tokens = $null; $parseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$parseErrors) | Out-Null
        if ($parseErrors) { throw "PowerShell parser error in '$relativePath': $($parseErrors[0].Message)" }
    }
    foreach ($relativePath in @('Config.json','Generic_Config.json','DependencyVersions.json')) {
        $path = Join-Path $basePath $relativePath
        try { $null = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -ErrorAction Stop }
        catch { throw "Invalid JSON in '$relativePath': $($_.Exception.Message)" }
    }
}

function Get-ReleaseEntries {
    $entries = foreach ($item in $releaseFiles) {
        $sourcePath = Join-Path $basePath $item.Source
        $destinationPath = Join-Path $basePath $item.Destination
        [pscustomobject]@{
            Source = $item.Source
            Destination = $item.Destination
            SourceHash = Get-ReleaseHash -Path $sourcePath
            DestinationHash = if (Test-Path -LiteralPath $destinationPath -PathType Leaf) { Get-ReleaseHash -Path $destinationPath } else { $null }
        }
    }
    return @($entries)
}

Test-ReleaseInputs

if ($Verify) {
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Release manifest not found: $manifestPath" }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -ErrorAction Stop
    $mismatches = @()
    foreach ($entry in Get-ReleaseEntries) {
        if ($entry.SourceHash -ne $entry.DestinationHash) { $mismatches += "$($entry.Source) -> $($entry.Destination)" }
    }
    if ($mismatches.Count -gt 0) { throw ("Release verification failed:`n - " + ($mismatches -join "`n - ")) }
    Write-Host "Release verification passed: $($manifest.CreatedUtc)" -ForegroundColor Green
    return
}

foreach ($item in $releaseFiles) {
    $sourcePath = Join-Path $basePath $item.Source
    $destinationPath = Join-Path $basePath $item.Destination
    $destinationFolder = Split-Path -Parent $destinationPath
    if (-not (Test-Path -LiteralPath $destinationFolder -PathType Container)) { New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null }
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

if ($BuildExecutables) {
    & (Join-Path $basePath '000-2Exe.ps1')
    if ($LASTEXITCODE -ne 0) { throw "FA_Main.exe build failed with exit code $LASTEXITCODE." }
    & (Join-Path $basePath '000-Gui2Exe.ps1')
    if ($LASTEXITCODE -ne 0) { throw "FA_GUI.exe build failed with exit code $LASTEXITCODE." }
    foreach ($exe in 'FA_Main.exe','FA_GUI.exe') {
        Copy-Item -LiteralPath (Join-Path $basePath $exe) -Destination (Join-Path $basePath "GitHub_Library\Exe_Files\$exe") -Force
    }
}

$entries = Get-ReleaseEntries
$manifest = [ordered]@{
    CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
    ExecutableBuildIncluded = [bool]$BuildExecutables
    Files = $entries
}
[IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
Write-Host "Release synchronised and manifest written: $manifestPath" -ForegroundColor Green
