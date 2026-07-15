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

function Resolve-FEDAUTOInstalledNavisworksApiPath {
    param(
        [string]$Version
    )

    $autodeskRoot = 'C:\Program Files\Autodesk'
    if (-not [string]::IsNullOrWhiteSpace($Version)) {
        $requestedApiPath = Join-Path (Join-Path $autodeskRoot "Navisworks Manage $Version") 'Autodesk.Navisworks.Api.dll'
        if (Test-Path -LiteralPath $requestedApiPath -PathType Leaf) {
            return $requestedApiPath
        }

        throw "Navisworks API assembly was not found for Navisworks Manage $Version`: $requestedApiPath"
    }

    $detectedInstall = Get-ChildItem -LiteralPath $autodeskRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^Navisworks Manage \d{4}$' } |
        Sort-Object @{ Expression = { [int]($_.Name -replace '\D+', '') }; Descending = $true } |
        Select-Object -First 1

    if ($detectedInstall) {
        $detectedApiPath = Join-Path $detectedInstall.FullName 'Autodesk.Navisworks.Api.dll'
        if (Test-Path -LiteralPath $detectedApiPath -PathType Leaf) {
            return $detectedApiPath
        }
    }

    throw "No Navisworks Manage API assembly was found under $autodeskRoot. Install Navisworks Manage or pass -NavisworksVersion for an installed version."
}

function Get-FEDAUTOVisualStylePluginProjectXml {
    @'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <AssemblyName>FederationAutomation.NavisworksVisualStyle</AssemblyName>
    <RootNamespace>FederationAutomation.NavisworksVisualStyle</RootNamespace>
    <LangVersion>latest</LangVersion>
    <Nullable>disable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <Reference Include="Autodesk.Navisworks.Api">
      <HintPath>$(NavisworksApiPath)</HintPath>
      <Private>false</Private>
    </Reference>
  </ItemGroup>
</Project>
'@
}

function Get-FEDAUTOVisualStylePluginSourceCode {
    @'
using System;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;

namespace FederationAutomation.NavisworksVisualStyle
{
    [Plugin("FederationAutomationVisualStyle", "FEDAUTO", DisplayName = "Federation Automation Visual Style")]
    [AddInPlugin(AddInLocation.None)]
    public sealed class NavisworksVisualStylePlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (document == null || document.IsClear)
                {
                    return 1;
                }

                Viewpoint viewpoint = document.CurrentViewpoint.CreateCopy();
                viewpoint.RenderStyle = ViewpointRenderStyle.FullRender;
                document.CurrentViewpoint.CopyFrom(viewpoint);

                // Keep the standard Navisworks graduated sky background in the saved document.
                document.SetGraduatedBackground(
                    Color.FromByteRGB(128, 179, 232),
                    Color.FromByteRGB(235, 245, 255));
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("Federation Automation could not apply the Navisworks visual style: " + exception.Message);
                return 1;
            }
        }
    }
}
'@
}

function Build-FEDAUTOVisualStylePlugin {
<#
Builds the Navisworks add-in from embedded C# project/source content.
Blank NavisworksVersion uses the latest installed Navisworks Manage version.
#>
    [CmdletBinding()]
    param(
        [string]$NavisworksVersion = '',
        [string]$OutputDirectory
    )

    $navisworksApiPath = Resolve-FEDAUTOInstalledNavisworksApiPath -Version $NavisworksVersion
    if ([string]::IsNullOrWhiteSpace($NavisworksVersion)) {
        Write-Host "Using latest detected Navisworks API: $navisworksApiPath"
    }
    else {
        Write-Host "Using requested Navisworks API: $navisworksApiPath"
    }

    $buildRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("FEDAUTO-NavisworksVisualStyle-{0}" -f ([guid]::NewGuid().ToString('N')))
    $projectPath = Join-Path $buildRoot 'NavisworksVisualStylePlugin.csproj'
    $sourcePath = Join-Path $buildRoot 'NavisworksVisualStylePlugin.cs'
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $OutputDirectory = Join-Path $buildRoot 'bin\Release\net48'
    }

    New-Item -Path $buildRoot -ItemType Directory -Force | Out-Null
    [System.IO.File]::WriteAllText($projectPath, (Get-FEDAUTOVisualStylePluginProjectXml), [System.Text.UTF8Encoding]::new($false))
    [System.IO.File]::WriteAllText($sourcePath, (Get-FEDAUTOVisualStylePluginSourceCode), [System.Text.UTF8Encoding]::new($false))

    & dotnet build $projectPath --configuration Release --output $OutputDirectory "/p:NavisworksApiPath=$navisworksApiPath" 2>&1 |
        ForEach-Object { Write-Host $_ }
    if ($LASTEXITCODE -ne 0) { throw 'Navisworks visual-style plug-in build failed.' }

    $pluginPath = Join-Path $OutputDirectory 'FederationAutomation.NavisworksVisualStyle.dll'
    if (-not (Test-Path -LiteralPath $pluginPath -PathType Leaf)) {
        throw "Navisworks visual-style plug-in was not produced: $pluginPath"
    }
    Write-Host "Built: $pluginPath"
    return $pluginPath
}
