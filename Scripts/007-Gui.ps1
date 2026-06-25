<#
Federation Automation configuration editor. Run with Windows PowerShell; it can edit JSON files
and import legacy Excel configuration for migration.
#>
[CmdletBinding()]
param([string]$ConfigFile)

if ([threading.thread]::CurrentThread.ApartmentState -ne 'STA') {
    $desktopPowerShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $arguments = @('-NoProfile', '-STA', '-File', $PSCommandPath)
    if ($ConfigFile) { $arguments += @('-ConfigFile', $ConfigFile) }
    Start-Process -FilePath $desktopPowerShell -ArgumentList $arguments
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
if (-not ('FEDAUTO.ModernFolderPicker' -as [type])) {
    Add-Type -ReferencedAssemblies PresentationFramework,PresentationCore,WindowsBase,System.Xaml -TypeDefinition @'
using System;
using System.Diagnostics;
using System.Collections.Concurrent;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Data;
using System.Windows.Media;

namespace FEDAUTO {
    [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IFileOpenDialog {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes(); void SetFileTypeIndex(); void GetFileTypeIndex(); void Advise(); void Unadvise();
        void SetOptions(uint options); void GetOptions(out uint options); void SetDefaultFolder(); void SetFolder(IShellItem folder);
        void GetFolder(); void GetCurrentSelection(); void SetFileName(); void GetFileName(); void SetTitle();
        void SetOkButtonLabel(); void SetFileNameLabel(); void GetResult(out IShellItem item); void AddPlace();
        void SetDefaultExtension(); void Close(); void SetClientGuid(); void ClearClientData(); void SetFilter();
        void GetResults(); void GetSelectedItems();
    }
    [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItem {
        void BindToHandler(); void GetParent(); void GetDisplayName(uint sigdnName, out IntPtr name);
        void GetAttributes(); void Compare();
    }
    [ComImport, Guid("dc1c5a9c-e88a-4dde-a5a1-60f82a20aef7")]
    internal class FileOpenDialog { }
    public static class ModernFolderPicker {
        private const uint FOS_PICKFOLDERS = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM = 0x00000040;
        private const uint FOS_PATHMUSTEXIST = 0x00000800;
        private const uint SIGDN_FILESYSPATH = 0x80058000;
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(string path, IntPtr bindContext, ref Guid riid, out IShellItem item);
        public static string Show(string initialPath) {
            IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialog();
            dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
            if (!String.IsNullOrEmpty(initialPath) && System.IO.Directory.Exists(initialPath)) {
                Guid shellItemGuid = new Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe");
                IShellItem initialFolder;
                SHCreateItemFromParsingName(initialPath, IntPtr.Zero, ref shellItemGuid, out initialFolder);
                dialog.SetFolder(initialFolder);
            }
            if (dialog.Show(IntPtr.Zero) != 0) return null;
            IShellItem item; dialog.GetResult(out item);
            IntPtr path; item.GetDisplayName(SIGDN_FILESYSPATH, out path);
            try { return Marshal.PtrToStringUni(path); }
            finally { Marshal.FreeCoTaskMem(path); }
        }
    }
    public sealed class BackgroundRun {
        public Process Process { get; private set; }
        private readonly ConcurrentQueue<string> lines = new ConcurrentQueue<string>();
        private BackgroundRun(Process process) { Process = process; }
        public bool TryGetLine(out string line) { return lines.TryDequeue(out line); }
        public static BackgroundRun Start(string executable, string arguments) {
            ProcessStartInfo info = new ProcessStartInfo(executable, arguments);
            info.UseShellExecute = false;
            info.CreateNoWindow = true;
            info.RedirectStandardOutput = true;
            info.RedirectStandardError = true;
            Process process = new Process();
            process.StartInfo = info;
            BackgroundRun run = new BackgroundRun(process);
            process.OutputDataReceived += delegate(object sender, DataReceivedEventArgs e) { if (e.Data != null) run.lines.Enqueue(e.Data); };
            process.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs e) { if (e.Data != null) run.lines.Enqueue("ERROR: " + e.Data); };
            if (!process.Start()) throw new InvalidOperationException("Unable to start the pipeline process.");
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            return run;
        }
    }
    public sealed class GroupOrderBrushConverter : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            int order; if (!Int32.TryParse(value == null ? "0" : value.ToString(), out order) || order <= 0) return new SolidColorBrush(Color.FromRgb(224, 224, 224));
            Color[] colours = { Color.FromRgb(21, 128, 61), Color.FromRgb(101, 163, 13), Color.FromRgb(202, 138, 4), Color.FromRgb(234, 88, 12), Color.FromRgb(185, 28, 28) };
            return new SolidColorBrush(colours[Math.Min(order - 1, colours.Length - 1)]);
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) { return Binding.DoNothing; }
    }
    public sealed class GroupOrderTextConverter : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            int order; return (!Int32.TryParse(value == null ? "0" : value.ToString(), out order) || order <= 0) ? "—" : order.ToString();
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) { return Binding.DoNothing; }
    }
}
'@
}
$basePath = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { Split-Path -Parent ([Reflection.Assembly]::GetEntryAssembly().Location) }
$guiStateDirectory = Join-Path $env:LOCALAPPDATA 'Federation-Automation'
$guiStatePath = Join-Path $guiStateDirectory 'GuiState.json'

function Remove-FEDAUTOStaleTemporaryFiles {
    # PS2EXE/compiler leaves CSC*.TMP files beside builds.  Restrict cleanup to
    # that exact generated-file pattern in the application folder.
    try {
        Get-ChildItem -LiteralPath $basePath -File -Filter 'CSC*.TMP' -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^CSC[0-9A-F]+\.TMP$' } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
    catch { }
}
Remove-FEDAUTOStaleTemporaryFiles

function Get-GuiState {
    if (-not (Test-Path -LiteralPath $guiStatePath)) { return $null }
    try {
        return (Get-Content -LiteralPath $guiStatePath -Raw | ConvertFrom-Json)
    }
    catch { return $null }
}

function Save-GuiState {
    param([string]$LastConfigFile, [bool]$LastSessionCompleted)
    try {
        if (-not (Test-Path -LiteralPath $guiStateDirectory)) { New-Item -Path $guiStateDirectory -ItemType Directory -Force | Out-Null }
        [pscustomobject]@{ LastConfigFile = $LastConfigFile; LastSessionCompleted = $LastSessionCompleted } | ConvertTo-Json | Set-Content -LiteralPath $guiStatePath -Encoding UTF8
    }
    catch { Write-Warning "Could not save GUI session state. $_" }
}

function Set-LastConfigurationPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    Save-GuiState -LastConfigFile $Path -LastSessionCompleted:$false
}

function Get-DefaultConfigurationPath {
    $state = Get-GuiState
    if ($state -and $state.LastConfigFile -and (Test-Path -LiteralPath $state.LastConfigFile -PathType Leaf)) { return $state.LastConfigFile }
    $defaultConfig = Join-Path $basePath 'Config.json'
    if (Test-Path -LiteralPath $defaultConfig -PathType Leaf) { return $defaultConfig }
    $legacyDefault = Join-Path $basePath 'Config.xlsx'
    if (Test-Path -LiteralPath $legacyDefault -PathType Leaf) { return $legacyDefault }
    return $null
}
function Get-GuiSupportScriptText {
    param([string]$FileName)
    $localPath = Join-Path $basePath $FileName
    if (Test-Path -LiteralPath $localPath) {
        # Read text and execute in-memory: a compiled GUI must not depend on
        # the user's PowerShell script execution policy for its bundled logic.
        return (Get-Content -LiteralPath $localPath -Raw)
    }
    $assembly = [Reflection.Assembly]::GetEntryAssembly()
    $resourceName = $assembly.GetManifestResourceNames() | Where-Object { $_ -like "*$FileName" } | Select-Object -First 1
    if (-not $resourceName) { throw "Required support file '$FileName' was not found." }
    $stream = $assembly.GetManifestResourceStream($resourceName)
    $reader = New-Object IO.StreamReader($stream)
    try { return $reader.ReadToEnd() }
    finally { $reader.Dispose(); $stream.Dispose() }
}
# Dot-source at script scope so the functions remain available to UI callbacks.
. ([scriptblock]::Create((Get-GuiSupportScriptText '012-SharedFunctions.Ps1')))
. ([scriptblock]::Create((Get-GuiSupportScriptText '013-ConfigFunctions.ps1')))

function ConvertTo-FEDAUTOBoolean {
    param($Value)
    if ($Value -is [bool]) { return $Value }
    if ($null -eq $Value) { return $false }
    return $Value.ToString().Trim().ToLowerInvariant() -in @('yes','y','true','1')
}

function ConvertTo-FEDAUTOCleanText {
    param($Value)
    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    # Excel content occasionally contains UTF-8 non-breaking spaces decoded as
    # "Â ". Remove that artefact and trim padding from copied cells.
    $badPair = ([char]0x00C2).ToString() + ([char]0x00A0).ToString()
    $text = $text.Replace($badPair, '')
    $text = $text -replace '[\u00A0\u2007\u202F]', ' '
    return $text.Trim()
}

function New-GridRows {
    param([array]$Rows, [string[]]$FallbackColumns, [string[]]$IgnoreColumns, [string[]]$BooleanColumns)
    $columns = New-Object System.Collections.Generic.List[string]
    foreach ($column in $FallbackColumns) { if (-not $columns.Contains($column)) { $columns.Add($column) } }
    foreach ($row in $Rows) { foreach ($property in $row.PSObject.Properties) { if (($IgnoreColumns -notcontains $property.Name) -and -not $columns.Contains($property.Name)) { $columns.Add($property.Name) } } }
    $collection = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    foreach ($row in $Rows) {
        $newRow = [ordered]@{}
        foreach ($column in $columns) {
            $rawValue = if ($row.PSObject.Properties.Name -contains $column) { $row.$column } else { $null }
            $newRow[$column] = if ($BooleanColumns -contains $column) { ConvertTo-FEDAUTOBoolean $rawValue } else { ConvertTo-FEDAUTOCleanText $rawValue }
        }
        $collection.Add([pscustomobject]$newRow)
    }
    # Do not enumerate the collection on return: a one-row range must still
    # bind as a collection rather than becoming a single PSCustomObject.
    Write-Output -NoEnumerate $collection
}

function ConvertFrom-GridRows {
    param($Rows, [string[]]$BooleanColumns)
    # PowerShell variable names are case-insensitive. Do not reuse the $Rows
    # parameter name here: `$rows = @()` would erase the caller's input.
    $convertedRows = @()
    foreach ($row in $Rows) {
        $item = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) {
            $item[$property.Name] = if ($BooleanColumns -contains $property.Name) { if (ConvertTo-FEDAUTOBoolean $property.Value) { 'Yes' } else { 'No' } } else { ConvertTo-FEDAUTOCleanText $property.Value }
        }
        $convertedRows += [pscustomobject]$item
    }
    return $convertedRows
}

function Get-FEDAUTOGroupOrderValue {
    param($Value)
    $number = 0
    if ($null -ne $Value) { [void][int]::TryParse($Value.ToString(), [ref]$number) }
    return [Math]::Max($number, 0)
}

function Set-FEDAUTOGroupOrder {
    param($SelectedRow, [int]$RequestedOrder)
    if ($script:UpdatingFederationGroupOrder) { return }
    if ($null -eq $SelectedRow -or -not ($SelectedRow.PSObject.Properties.Name -contains 'GroupOrder')) { return }
    $script:UpdatingFederationGroupOrder = $true
    try {
        $allRows = @($script:FederationRows | Where-Object { $_ -and ($_.PSObject.Properties.Name -contains 'GroupOrder') })
        $activeRows = New-Object System.Collections.Generic.List[object]
        foreach ($row in $allRows) {
            if ($row -eq $SelectedRow) { continue }
            if ((Get-FEDAUTOGroupOrderValue $row.GroupOrder) -gt 0) { $activeRows.Add($row) }
        }
        $sortedRows = @($activeRows | Sort-Object @{ Expression = { Get-FEDAUTOGroupOrderValue $_.GroupOrder } }, @{ Expression = { [array]::IndexOf($allRows, $_) } })
        $activeRows = New-Object System.Collections.Generic.List[object]
        foreach ($row in $sortedRows) { $activeRows.Add($row) }
        if ($RequestedOrder -le 0) {
            $SelectedRow.GroupOrder = '0'
        }
        else {
            $insertIndex = [Math]::Min($RequestedOrder - 1, $activeRows.Count)
            $activeRows.Insert($insertIndex, $SelectedRow)
        }
        for ($index = 0; $index -lt $activeRows.Count; $index++) { $activeRows[$index].GroupOrder = ($index + 1).ToString() }
        $FederationGrid.Items.Refresh()
    }
    finally { $script:UpdatingFederationGroupOrder = $false }
}

function Normalize-FEDAUTOGroupOrders {
    param($Rows)
    $ordered = @($Rows | Where-Object { (Get-FEDAUTOGroupOrderValue $_.GroupOrder) -gt 0 } | Sort-Object @{ Expression = { Get-FEDAUTOGroupOrderValue $_.GroupOrder } })
    for ($index = 0; $index -lt $ordered.Count; $index++) { $ordered[$index].GroupOrder = ($index + 1).ToString() }
}

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Federation Automation" Height="820" Width="1280" MinHeight="620" MinWidth="960" WindowStartupLocation="CenterScreen" Background="#F5F7FA">
  <DockPanel>
    <Border DockPanel.Dock="Top" Background="#07345C" Padding="24,16"><StackPanel><TextBlock Text="Federation Automation" Foreground="White" FontSize="28" FontWeight="SemiBold"/><TextBlock Text="Configuration and model federation" Foreground="#C9D8E6" FontSize="15"/></StackPanel></Border>
    <Border DockPanel.Dock="Bottom" Background="White" Padding="16" BorderBrush="#D5DCE3" BorderThickness="0,1,0,0"><DockPanel><TextBlock Name="StatusText" Text="Ready" VerticalAlignment="Center" Foreground="#425466"/><StackPanel Orientation="Horizontal" HorizontalAlignment="Right"><Button Name="ExportExcelButton" Content="Export Excel..." Padding="16,7" Margin="0,0,10,0"/><Button Name="ValidateButton" Content="Validate configuration" Padding="16,7" Margin="0,0,10,0"/><Button Name="SaveButton" Content="Save" Padding="22,7" Margin="0,0,10,0"/><Button Name="RunButton" Content="Save and Run" Padding="22,7" Background="#0867C8" Foreground="White" FontWeight="SemiBold"/></StackPanel></DockPanel></Border>
    <Grid Margin="24"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
      <Grid Grid.Row="0" Margin="0,0,0,16"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBox Name="ConfigPathBox" FontSize="15" Padding="10" VerticalContentAlignment="Center"/><Button Name="OpenButton" Grid.Column="1" Content="Open..." Padding="18,8" Margin="10,0,0,0"/><Button Name="NewButton" Grid.Column="2" Content="New JSON" Padding="18,8" Margin="10,0,0,0"/></Grid>
      <TabControl Grid.Row="1" Name="MainTabs"><TabItem Header="Settings"><ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel Name="SettingsPanel" Margin="20"/></ScrollViewer></TabItem>
        <TabItem Header="Download"><DockPanel Margin="10"><Border Name="DownloadStatusPanel" DockPanel.Dock="Top" Background="#D9E8F5" Padding="10,7" Margin="0,0,0,10" CornerRadius="3"><TextBlock Name="DownloadStatusText" TextWrapping="Wrap"/></Border><DataGrid Name="DownloadGrid" AutoGenerateColumns="True" CanUserAddRows="True" CanUserDeleteRows="True" EnableRowVirtualization="False" VirtualizingPanel.IsVirtualizing="False" ClipboardCopyMode="ExcludeHeader"/></DockPanel></TabItem>
        <TabItem Header="Attributes"><DataGrid Name="AttributesGrid" Margin="10" AutoGenerateColumns="True" CanUserAddRows="True" CanUserDeleteRows="True" EnableRowVirtualization="False" VirtualizingPanel.IsVirtualizing="False" ClipboardCopyMode="ExcludeHeader"/></TabItem>
        <TabItem Header="Grouping"><DockPanel Margin="10"><StackPanel Name="GroupingOptionsPanel" DockPanel.Dock="Top" Margin="0,0,0,10"/><Grid><DataGrid Name="WildcardSelectionGrid" Visibility="Collapsed" AutoGenerateColumns="True" CanUserAddRows="True" CanUserDeleteRows="True" EnableRowVirtualization="False" VirtualizingPanel.IsVirtualizing="False" ClipboardCopyMode="ExcludeHeader"/><DataGrid Name="FederationGrid" AutoGenerateColumns="True" CanUserAddRows="True" CanUserDeleteRows="True" EnableRowVirtualization="False" VirtualizingPanel.IsVirtualizing="False" ClipboardCopyMode="ExcludeHeader"/></Grid></DockPanel></TabItem>
        <TabItem Name="LookupsTab" Header="Lookups"><DataGrid Name="LookupsGrid" Margin="10" AutoGenerateColumns="True" CanUserAddRows="True" CanUserDeleteRows="True" EnableRowVirtualization="False" VirtualizingPanel.IsVirtualizing="False" ClipboardCopyMode="ExcludeHeader"/></TabItem>
        <TabItem Header="Run"><Grid Margin="18"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions><TextBlock Text="Run dashboard" FontSize="22" FontWeight="SemiBold"/><TextBlock Grid.Row="1" Text="Save the configuration, then run the same pipeline used by the command-line launcher." Foreground="#566573" Margin="0,8,0,14"/><Border Grid.Row="2" Background="White" BorderBrush="#D5DCE3" BorderThickness="1" CornerRadius="4" Padding="16" Margin="0,0,0,12"><StackPanel><TextBlock Text="Overall status" Foreground="#566573"/><TextBlock Name="RunStageText" Text="Ready" FontSize="18" FontWeight="Bold" Margin="0,3,0,2"/><TextBlock Name="RunDetailText" Text="Waiting to start." Foreground="#425466" TextTrimming="CharacterEllipsis"/><ProgressBar Name="RunProgressBar" Minimum="0" Maximum="100" Value="0" Height="14" Margin="0,12,0,0"/></StackPanel></Border><Border Grid.Row="3" Background="White" BorderBrush="#D5DCE3" BorderThickness="1" CornerRadius="4" Padding="18"><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions><TextBlock Text="Live activity" FontSize="18" FontWeight="SemiBold"/><RichTextBox Name="ActivityBox" Grid.Row="1" Margin="0,12,0,0" IsReadOnly="True" VerticalScrollBarVisibility="Auto"/></Grid></Border></Grid></TabItem>
       </TabControl>
    </Grid>
  </DockPanel>
</Window>
'@
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
foreach ($name in 'ConfigPathBox','OpenButton','NewButton','SaveButton','RunButton','ValidateButton','ExportExcelButton','StatusText','ActivityBox','RunStageText','RunDetailText','RunProgressBar','MainTabs','SettingsPanel','DownloadStatusPanel','DownloadStatusText','DownloadGrid','AttributesGrid','FederationGrid','WildcardSelectionGrid','GroupingOptionsPanel','LookupsGrid','LookupsTab') { Set-Variable -Name $name -Value $window.FindName($name) }

function Set-EditorConfiguration {
    param($Configuration, [string]$Path)
    $script:SettingsRows = New-GridRows (Merge-FEDAUTOSettingsWithCatalog $Configuration.Settings) @('Section','Parameter','Value','Desc','DefaultValue','IsDefault')
    $script:DownloadRows = New-GridRows $Configuration.Download @('Run','ReadFolder','FileFilter','Exclude','SkipIfSame','CheckDateToo','MinState') @('Enabled','SourceType','Folder','Filter') @('Run','SkipIfSame','CheckDateToo')
    $script:AttributesRows = New-GridRows $Configuration.PWAttributesList @('AttributeName','OutputName','ExportToXLSX','InjectToIFC') @('Attribute','PropertySet','Enabled') @('ExportToXLSX','InjectToIFC')
    $script:FederationRows = New-GridRows $Configuration.Federation @() @() @('InjectToIFC')
    $wildcardRows = @($Configuration.WildcardSelection)
    # An empty DataGrid has no generated columns and cannot accept its first
    # row. Seed the editor with a blank rule, then omit it on save/export.
    if ($wildcardRows.Count -eq 0) {
        $wildcardRows = @([pscustomobject]@{ Inclusions=''; Exclusions=''; ExportFileName=''; ReadFromOutputFolder=$false })
    }
    $script:WildcardSelectionRows = New-GridRows $wildcardRows @('Inclusions','Exclusions','ExportFileName','ReadFromOutputFolder') @('IncludeInFinalModel') @('ReadFromOutputFolder')
    $script:FederationGroupOrderOptions = @(0..([Math]::Max($script:FederationRows.Count, 1)) | ForEach-Object { $_.ToString() })
    Normalize-FEDAUTOGroupOrders $script:FederationRows
    # Reference was a legacy Lookups column and is not used by processing or federation.
    $script:LookupsRows = New-GridRows $Configuration.Lookups @() @('Reference')
    $SettingsPanel.Tag = $script:SettingsRows
    $DownloadGrid.ItemsSource = $script:DownloadRows
    $AttributesGrid.ItemsSource = $script:AttributesRows
    $FederationGrid.ItemsSource = $script:FederationRows
    $WildcardSelectionGrid.ItemsSource = $script:WildcardSelectionRows
    $LookupsGrid.ItemsSource = $script:LookupsRows
    Show-SettingsEditor
    Show-FEDAUTOGroupingOptions
    $ConfigPathBox.Text = $Path
    Set-LastConfigurationPath $Path
    $StatusText.Text = "Loaded $($Configuration.Format) configuration."
}

function Show-FEDAUTOGroupingOptions {
    $GroupingOptionsPanel.Children.Clear()
    $getSetting = {
        param([string]$Name)
        $script:SettingsRows | Where-Object { $_.Parameter -eq $Name } | Select-Object -First 1
    }
    $methodSetting = & $getSetting 'FederationGroupingMethod'
    $finalNameSetting = & $getSetting 'FederatedFileName'
    $unmatchedSetting = & $getSetting 'IncludeUnmatchedFilesInFederatedModel'
    $namingSetting = & $getSetting 'NWDNamingMethod'
    $addRow = {
        param([string]$Label, $Control)
        $row = New-Object Windows.Controls.Grid -Property @{ Margin='0,0,0,6' }
        [void]$row.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width='230' }))
        [void]$row.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width='*' }))
        [void]$row.Children.Add((New-Object Windows.Controls.TextBlock -Property @{ Text=$Label; VerticalAlignment='Center'; FontWeight='SemiBold' }))
        [Windows.Controls.Grid]::SetColumn($Control, 1); [void]$row.Children.Add($Control)
        [void]$GroupingOptionsPanel.Children.Add($row)
    }
    $methodCombo = New-Object Windows.Controls.ComboBox -Property @{ ItemsSource=@('Naming Convention and Lookups','Wildcard Selection'); SelectedItem=$(if ($methodSetting) { $methodSetting.Value } else { 'Naming Convention and Lookups' }); Padding='5,3'; Tag=$methodSetting }
    $methodCombo.Add_SelectionChanged({
        param($sender, $eventArgs)
        if ($sender.Tag) { $sender.Tag.Value = $sender.SelectedItem.ToString() }
        $isWildcard = $sender.SelectedItem.ToString() -eq 'Wildcard Selection'
        $FederationGrid.Visibility = if ($isWildcard) { 'Collapsed' } else { 'Visible' }
        $WildcardSelectionGrid.Visibility = if ($isWildcard) { 'Visible' } else { 'Collapsed' }
        $LookupsTab.Visibility = if ($isWildcard) { 'Collapsed' } else { 'Visible' }
        $sender.Parent.Parent.Tag.Visibility = if ($isWildcard) { 'Collapsed' } else { 'Visible' }
    })
    & $addRow 'Grouping method' $methodCombo
    $finalNameBox = New-Object Windows.Controls.TextBox -Property @{ Text=$(if ($finalNameSetting) { $finalNameSetting.Value } else { 'Project Federated.nwd' }); Padding='5,3'; Tag=$finalNameSetting }
    if ($finalNameSetting) { $finalNameBox.Add_TextChanged({ param($sender,$eventArgs) $sender.Tag.Value=$sender.Text }) }
    & $addRow 'Final federated file name' $finalNameBox
    $legacyPanel = New-Object Windows.Controls.StackPanel
    $GroupingOptionsPanel.Tag = $legacyPanel
    $unmatchedCombo = New-Object Windows.Controls.ComboBox -Property @{ ItemsSource=@('Yes','No'); SelectedItem=$(if ($unmatchedSetting) { $unmatchedSetting.Value } else { 'No' }); Padding='5,3'; Tag=$unmatchedSetting }
    if ($unmatchedSetting) { $unmatchedCombo.Add_SelectionChanged({ param($sender,$eventArgs) $sender.Tag.Value=$sender.SelectedItem.ToString() }) }
    & $addRow 'Include unmatched files in final model' $unmatchedCombo
    $namingCombo = New-Object Windows.Controls.ComboBox -Property @{ ItemsSource=@('Full','OnlyCodes','OnlyDesc','Codes-Desc'); SelectedItem=$(if ($namingSetting) { $namingSetting.Value } else { 'Full' }); Padding='5,3'; Tag=$namingSetting }
    if ($namingSetting) { $namingCombo.Add_SelectionChanged({ param($sender,$eventArgs) $sender.Tag.Value=$sender.SelectedItem.ToString() }) }
    & $addRow 'Grouped NWD naming method' $namingCombo
    # The last two rows are mode-specific; move them into a wrapper so they can be hidden together.
    $legacyRows = @($GroupingOptionsPanel.Children | Select-Object -Last 3)
    foreach ($legacyRow in $legacyRows) { [void]$GroupingOptionsPanel.Children.Remove($legacyRow); [void]$legacyPanel.Children.Add($legacyRow) }
    [void]$GroupingOptionsPanel.Children.Add($legacyPanel)
    $isWildcard = $methodCombo.SelectedItem -eq 'Wildcard Selection'
    $FederationGrid.Visibility = if ($isWildcard) { 'Collapsed' } else { 'Visible' }
    $WildcardSelectionGrid.Visibility = if ($isWildcard) { 'Visible' } else { 'Collapsed' }
    $LookupsTab.Visibility = if ($isWildcard) { 'Collapsed' } else { 'Visible' }
    $legacyPanel.Visibility = if ($isWildcard) { 'Collapsed' } else { 'Visible' }
}

function Get-SettingControlType {
    param([string]$Parameter)
    if ($Parameter -in @('RunDownload','IncludeUnmatchedFilesInFederatedModel','NavisworksVisible')) { return 'YesNo' }
    if ($Parameter -in @('RunProcess','RunFederation','ReviztoPublish','NWDNamingMethod')) { return 'Choice' }
    if ($Parameter -match '(Folder|Path)$' -or $Parameter -in @('SourceFolder','ProcessedFolder','LogFolder','FederationInputFolder','FederationOutputFolder')) { return 'Folder' }
    if ($Parameter -in @('NavisworksConfigXML','NavisworksViewsImportXML')) { return 'File' }
    return 'Text'
}

function Get-SettingHelpText {
    param($Setting)
    $help = @{
        RunDownload = 'Controls source acquisition from Download rows. Choose Yes to retrieve or copy configured files; No to ignore Download rows and use existing files.'
        PWUser = 'Optional ProjectWise user name. Leave blank when your normal Bentley/IMS sign-in is used.'
        PWPass = 'Optional ProjectWise password. Do not store production passwords in a shared configuration file.'
        LogFolder = 'Folder where run logs are written. A relative path is based on the application folder.'
        SourceFolder = 'Folder where downloaded or copied source model files are staged before processing.'
        AttributesFile = 'Excel workbook, stored under SourceFolder, that contains the captured source metadata used during IFC processing.'
        RunProcess = 'Controls IFC metadata processing. Yes runs when needed, No skips it, and Force reprocesses all applicable IFC files.'
        ProcessedFolder = 'Folder where IFC files with injected attributes and process summaries are written.'
        RunFederation = 'Controls Navisworks federation. Yes runs only when changes require it, No disables it, and Force always rebuilds the federation.'
        IncludeUnmatchedFilesInFederatedModel = 'Choose Yes to add models that do not match the federation naming rules to the final federated NWD.'
        FederationInputFolder = 'Folder used as the source for federation. Leave blank to let the pipeline choose ProcessedFolder or SourceFolder automatically.'
        FederationOutputFolder = 'Folder where grouped NWD files and the final federated model are created.'
        FederatedFileName = 'Name of the final federated Navisworks model. The .nwd extension is added automatically when omitted.'
        NavisworksVersion = 'Preferred installed Navisworks version. Leave blank to let the pipeline detect a suitable installed version.'
        NavisworksConfigXML = 'Optional Navisworks XML options file used when creating federated models.'
        NavisworksViewsImportXML = 'Optional XML file containing saved views to import into the final Navisworks model.'
        NavisworksVisible = 'Choose Yes to show Navisworks while federation runs; No runs it in the background.'
        NWDNamingMethod = 'Controls names of grouped NWD files: Full, OnlyCodes, OnlyDesc, or Codes-Desc.'
        ReviztoPublish = 'Controls publishing the final federated model to Revizto. Force publishes when a valid model is available.'
        ReviztoPublishCode = 'The Revizto scheduler publish code for the target project. Required only when publishing is enabled.'
    }
    if ($help.ContainsKey($Setting.Parameter)) { return $help[$Setting.Parameter] }
    if ($Setting.Desc) { return $Setting.Desc }
    return 'No additional guidance is available for this setting.'
}

function ConvertTo-FEDAUTOStoredPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    try {
        $fullPath = [IO.Path]::GetFullPath($Path)
        $rootPath = ([IO.Path]::GetFullPath($basePath)).TrimEnd('\')
        if ($fullPath.StartsWith($rootPath + '\', [StringComparison]::OrdinalIgnoreCase)) { return $fullPath.Substring($rootPath.Length + 1) }
    }
    catch { }
    return $Path
}

function Get-FEDAUTOInitialFolder {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $basePath }
    $candidate = $Path
    if (-not [IO.Path]::IsPathRooted($candidate)) { $candidate = Join-Path $basePath $candidate }
    if (Test-Path -LiteralPath $candidate -PathType Container) { return $candidate }
    $parent = Split-Path -Parent $candidate
    if ($parent -and (Test-Path -LiteralPath $parent -PathType Container)) { return $parent }
    return $basePath
}

function Select-ModernFolder {
    param([string]$InitialPath)
    try { return [FEDAUTO.ModernFolderPicker]::Show((Get-FEDAUTOInitialFolder $InitialPath)) }
    catch { [Windows.MessageBox]::Show("Unable to open the Windows folder picker. $($_.Exception.Message)", 'Folder picker'); return $null }
}

function Test-FEDAUTOProjectWiseFolder {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $raw = $Path.Trim()
    if ($raw -match '^pw://[^/]+/' -or $raw -match '^pw:\\?[^\\]+') { return $true }
    $normalized = $raw -replace '/', '\\'
    if ($normalized -match '^[A-Za-z]:\\' -or $normalized -match '^\\\\') { return $false }
    return $normalized -match '^([A-Za-z0-9_.-]+):\\?'
}

function Get-FEDAUTOAcquisitionState {
    $runSetting = $script:SettingsRows | Where-Object { $_.Parameter -eq 'RunDownload' } | Select-Object -First 1
    $modeSetting = $script:SettingsRows | Where-Object { $_.Parameter -eq 'SourceAcquisitionMode' } | Select-Object -First 1
    $mode = if ($modeSetting -and $modeSetting.Value) { $modeSetting.Value.ToString() } else { 'Auto' }
    $isEnabled = -not $runSetting -or $runSetting.Value.ToString().Trim().ToLowerInvariant() -notin @('no','n','false','0','ignore')
    $activeRows = @($script:DownloadRows | Where-Object {
        $run = if ($_.PSObject.Properties.Name -contains 'Run') { $_.Run } else { 'Yes' }
        $folder = if ($_.PSObject.Properties.Name -contains 'ReadFolder') { $_.ReadFolder } else { $null }
        $runText = if ($null -eq $run) { '' } else { $run.ToString().Trim().ToLowerInvariant() }
        -not [string]::IsNullOrWhiteSpace($folder) -and $runText -notin @('no','n','false','0','ignore','')
    })
    $pwCount = @($activeRows | Where-Object { Test-FEDAUTOProjectWiseFolder $_.ReadFolder }).Count
    $hasProjectWise = $isEnabled -and $(if ($mode -eq 'Local') { $false } elseif ($mode -eq 'ProjectWise') { $true } else { $pwCount -gt 0 })
    [pscustomobject]@{ Enabled=$isEnabled; Mode=$mode; ActiveCount=$activeRows.Count; ProjectWiseCount=$pwCount; LocalCount=($activeRows.Count - $pwCount); HasProjectWise=$hasProjectWise }
}

function Update-FEDAUTOAcquisitionPresentation {
    $state = Get-FEDAUTOAcquisitionState
    if (-not $state.Enabled) {
        $DownloadStatusPanel.Background = '#E8ECEF'; $DownloadStatusText.Text = 'Source acquisition is disabled by RunDownload. Download rows are ignored and existing files will be used.'
    }
    elseif ($state.ActiveCount -eq 0) {
        $DownloadStatusPanel.Background = '#FFF3CD'; $DownloadStatusText.Text = 'No active Download rows. No files will be retrieved or copied.'
    }
    elseif ($state.Mode -eq 'Local') {
        $DownloadStatusPanel.Background = '#D9E8F5'; $DownloadStatusText.Text = "Local mode selected. $($state.LocalCount) local/synchronised rows will run; ProjectWise rows are excluded."
    }
    elseif ($state.Mode -eq 'ProjectWise') {
        $DownloadStatusPanel.Background = '#DFF0C8'; $DownloadStatusText.Text = "ProjectWise mode selected. $($state.ProjectWiseCount) ProjectWise rows will run; local rows are excluded."
    }
    elseif ($state.ProjectWiseCount -eq 0) {
        $DownloadStatusPanel.Background = '#D9E8F5'; $DownloadStatusText.Text = "Local/synchronised sources: $($state.LocalCount). ProjectWise sign-in is not required."
    }
    elseif ($state.LocalCount -eq 0) {
        $DownloadStatusPanel.Background = '#DFF0C8'; $DownloadStatusText.Text = "ProjectWise sources: $($state.ProjectWiseCount). ProjectWise settings are enabled."
    }
    else {
        $DownloadStatusPanel.Background = '#FFF3CD'; $DownloadStatusText.Text = "Mixed sources: $($state.ProjectWiseCount) ProjectWise and $($state.LocalCount) local/synchronised. Both source types will run."
    }
}

function Show-SettingsEditor {
    $SettingsPanel.Children.Clear()
    $acquisitionState = Get-FEDAUTOAcquisitionState
    $processSetting = $script:SettingsRows | Where-Object { $_.Parameter -eq 'RunProcess' } | Select-Object -First 1
    $processEnabled = $processSetting -and $processSetting.Value.ToString().Trim().ToLowerInvariant() -notin @('no','n','false','0','ignore','')
    $federationSetting = $script:SettingsRows | Where-Object { $_.Parameter -eq 'RunFederation' } | Select-Object -First 1
    $federationEnabled = $federationSetting -and $federationSetting.Value.ToString().Trim().ToLowerInvariant() -notin @('no','n','false','0','ignore','')
    $reviztoSetting = $script:SettingsRows | Where-Object { $_.Parameter -eq 'ReviztoPublish' } | Select-Object -First 1
    $reviztoEnabled = $reviztoSetting -and $reviztoSetting.Value.ToString().Trim().ToLowerInvariant() -notin @('no','n','false','0','ignore','')
    $lastSection = $null
    $sectionColours = @{ 'Source acquisition'='#D9E8F5'; 'Working folders & metadata'='#E8ECEF'; 'IFC processing'='#FBE4D5'; 'Federation & Navisworks'='#DFF0C8'; 'Revizto publishing'='#FFF59D'; 'Other settings'='#EDE7F6' }
    $sectionHelp = @{ 'Working folders & metadata'='These paths are used throughout the workflow. Configure them before enabling optional stages.'; 'Source acquisition'='Controls retrieval or copying of source files using the active Download rows.'; 'IFC processing'='Adds configured metadata to IFC files and writes processed copies.'; 'Federation & Navisworks'='Builds grouped and final Navisworks NWD models from the configured input files.'; 'Revizto publishing'='Optionally publishes a valid federated model to the configured Revizto target.'; 'Other settings'='Additional settings not yet assigned to a standard section.' }
    foreach ($setting in $script:SettingsRows) {
        if ($setting.Parameter -in @('PWUser','PWPass') -and -not $acquisitionState.HasProjectWise) { continue }
        if ($setting.Section -eq 'IFC processing' -and $setting.Parameter -ne 'RunProcess' -and -not $processEnabled) { continue }
        if ($setting.Section -eq 'Federation & Navisworks' -and $setting.Parameter -ne 'RunFederation' -and -not $federationEnabled) { continue }
        if ($setting.Section -eq 'Revizto publishing' -and $setting.Parameter -ne 'ReviztoPublish' -and -not $reviztoEnabled) { continue }
        if ($setting.Section -ne $lastSection) {
            $sectionColour = if ($sectionColours.ContainsKey($setting.Section)) { $sectionColours[$setting.Section] } else { '#E8ECEF' }
            $header = New-Object Windows.Controls.Border -Property @{ Background = $sectionColour; Padding = '10,7'; Margin = '0,16,0,8'; CornerRadius = '3' }
            if ($setting.Section -in @('Source acquisition','IFC processing','Federation & Navisworks','Revizto publishing')) {
                $headerGrid = New-Object Windows.Controls.Grid
                [void]$headerGrid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
                [void]$headerGrid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = 'Auto' }))
                [void]$headerGrid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = 'Auto' }))
                [void]$headerGrid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = '30' }))
                [void]$headerGrid.Children.Add((New-Object Windows.Controls.TextBlock -Property @{ Text = $setting.Section; FontWeight = 'SemiBold'; FontSize = 15; VerticalAlignment = 'Center' }))
                $isDownloadSection = $setting.Section -eq 'Source acquisition'
                $isProcessSection = $setting.Section -eq 'IFC processing'
                $isFederationSection = $setting.Section -eq 'Federation & Navisworks'
                if ($isDownloadSection) {
                    $modeSetting = $script:SettingsRows | Where-Object { $_.Parameter -eq 'SourceAcquisitionMode' } | Select-Object -First 1
                    $modeCombo = New-Object Windows.Controls.ComboBox -Property @{ ItemsSource=@('Auto','Local','ProjectWise'); SelectedItem=$modeSetting.Value; Width=110; Margin='0,0,10,0'; ToolTip='Controls which source type is allowed to run.' }
                    $modeCombo.Tag = $modeSetting
                    $modeCombo.Add_SelectionChanged({ param($sender, $eventArgs) if ($sender.SelectedItem) { $sender.Tag.Value = $sender.SelectedItem.ToString(); Show-SettingsEditor } })
                    [Windows.Controls.Grid]::SetColumn($modeCombo, 1); [void]$headerGrid.Children.Add($modeCombo)
                }
                $toggle = New-Object Windows.Controls.CheckBox -Property @{ Content = $(if ($isDownloadSection) { 'Enable source acquisition' } elseif ($isProcessSection) { 'Enable processing' } elseif ($isFederationSection) { 'Enable federation' } else { 'Enable publishing' }); IsChecked = $(if ($isDownloadSection) { $acquisitionState.Enabled } elseif ($isProcessSection) { $processEnabled } elseif ($isFederationSection) { $federationEnabled } else { $reviztoEnabled }); VerticalAlignment = 'Center' }
                $toggle.Tag = $(if ($isDownloadSection) { $script:SettingsRows | Where-Object { $_.Parameter -eq 'RunDownload' } | Select-Object -First 1 } elseif ($isProcessSection) { $processSetting } elseif ($isFederationSection) { $federationSetting } else { $reviztoSetting })
                $toggle.Add_Checked({ param($sender, $eventArgs) $sender.Tag.Value = 'Yes'; Show-SettingsEditor })
                $toggle.Add_Unchecked({ param($sender, $eventArgs) $sender.Tag.Value = 'No'; Show-SettingsEditor })
                [Windows.Controls.Grid]::SetColumn($toggle, $(if ($isDownloadSection) { 2 } else { 1 })); [void]$headerGrid.Children.Add($toggle)
                $sectionInfo = New-Object Windows.Controls.Button -Property @{ Content='i'; ToolTip=$sectionHelp[$setting.Section]; FontWeight='Bold'; FontSize=11; Width=19; Height=19; Padding=0; HorizontalAlignment='Right'; VerticalAlignment='Center' }
                $sectionInfo.Tag = $sectionHelp[$setting.Section]
                $sectionInfo.Add_Click({ param($sender, $eventArgs) [Windows.MessageBox]::Show($sender.Tag, 'Section information') })
                [Windows.Controls.Grid]::SetColumn($sectionInfo, $(if ($isDownloadSection) { 3 } else { 2 })); [void]$headerGrid.Children.Add($sectionInfo)
                $header.Child = $headerGrid
            }
            else {
                $headerGrid = New-Object Windows.Controls.Grid
                [void]$headerGrid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
                [void]$headerGrid.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = '30' }))
                [void]$headerGrid.Children.Add((New-Object Windows.Controls.TextBlock -Property @{ Text = $setting.Section; FontWeight = 'SemiBold'; FontSize = 15; VerticalAlignment = 'Center' }))
                $sectionInfo = New-Object Windows.Controls.Button -Property @{ Content='i'; ToolTip=$sectionHelp[$setting.Section]; FontWeight='Bold'; FontSize=11; Width=19; Height=19; Padding=0; HorizontalAlignment='Right'; VerticalAlignment='Center' }
                $sectionInfo.Tag = $sectionHelp[$setting.Section]
                $sectionInfo.Add_Click({ param($sender, $eventArgs) [Windows.MessageBox]::Show($sender.Tag, 'Section information') })
                [Windows.Controls.Grid]::SetColumn($sectionInfo, 1); [void]$headerGrid.Children.Add($sectionInfo)
                $header.Child = $headerGrid
            }
            [void]$SettingsPanel.Children.Add($header)
            if ($setting.Section -in @('IFC processing','Federation & Navisworks','Revizto publishing')) {
                $forceSetting = if ($setting.Section -eq 'IFC processing') { $processSetting } elseif ($setting.Section -eq 'Federation & Navisworks') { $federationSetting } else { $reviztoSetting }
                $sectionEnabled = if ($setting.Section -eq 'IFC processing') { $processEnabled } elseif ($setting.Section -eq 'Federation & Navisworks') { $federationEnabled } else { $reviztoEnabled }
                if ($sectionEnabled) {
                    $forceLabel = if ($setting.Section -eq 'IFC processing') { 'Force processing' } elseif ($setting.Section -eq 'Federation & Navisworks') { 'Force federation rebuild' } else { 'Force publish' }
                    $forcePanel = New-Object Windows.Controls.Grid -Property @{ Margin='12,0,0,8' }
                    [void]$forcePanel.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width='*' }))
                    [void]$forcePanel.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width='34' }))
                    $forceHelp = if ($setting.Section -eq 'IFC processing') { 'Reprocess applicable IFC files even when the source files and metadata have not changed.' } elseif ($setting.Section -eq 'Federation & Navisworks') { 'Rebuild federation even when no source changes are detected.' } else { 'Publish a valid federated model even when it is not newly created.' }
                    $forceToggle = New-Object Windows.Controls.CheckBox -Property @{ Content = $forceLabel; IsChecked = ($forceSetting.Value.ToString().Trim().ToLowerInvariant() -eq 'force'); ToolTip = $forceHelp }
                    $forceToggle.Tag = $forceSetting
                    $forceToggle.Add_Checked({ param($sender, $eventArgs) $sender.Tag.Value = 'Force' })
                    $forceToggle.Add_Unchecked({ param($sender, $eventArgs) $sender.Tag.Value = 'Yes' })
                    $forceInfo = New-Object Windows.Controls.Button -Property @{ Content='i'; ToolTip=$forceHelp; FontWeight='Bold'; FontSize=11; Width=19; Height=19; Padding=0; HorizontalAlignment='Right' }
                    $forceInfo.Tag = $forceHelp
                    $forceInfo.Add_Click({ param($sender, $eventArgs) [Windows.MessageBox]::Show($sender.Tag, 'Force option') })
                    [void]$forcePanel.Children.Add($forceToggle)
                    [Windows.Controls.Grid]::SetColumn($forceInfo, 1); [void]$forcePanel.Children.Add($forceInfo)
                    [void]$SettingsPanel.Children.Add($forcePanel)
                }
            }
            $lastSection = $setting.Section
        }
        if ($setting.Parameter -in @('RunDownload','SourceAcquisitionMode','RunProcess','RunFederation','ReviztoPublish','FederationGroupingMethod','FederatedFileName','IncludeUnmatchedFilesInFederatedModel','NWDNamingMethod')) { continue }
        $panel = New-Object Windows.Controls.Grid
        $panel.Margin = '0,0,0,12'
        [void]$panel.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = '260' }))
        [void]$panel.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
        [void]$panel.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition -Property @{ Width = '34' }))
        $label = New-Object Windows.Controls.StackPanel -Property @{ VerticalAlignment = 'Center' }
        $name = New-Object Windows.Controls.TextBlock -Property @{ Text = $setting.Parameter; FontWeight = 'SemiBold'; VerticalAlignment = 'Center' }
        [void]$label.Children.Add($name)
        if ($setting.IsDefault) { [void]$label.Children.Add((New-Object Windows.Controls.TextBlock -Property @{ Text = 'Default value'; Foreground = '#607D8B'; FontSize = 11 })) }
        $info = New-Object Windows.Controls.Button -Property @{ Content = 'i'; ToolTip = (Get-SettingHelpText $setting); FontWeight = 'Bold'; FontSize = 11; Width = 19; Height = 19; Padding = 0; VerticalAlignment = 'Center'; HorizontalAlignment = 'Right' }
        $info.Tag = Get-SettingHelpText $setting
        $info.Add_Click({ param($sender, $eventArgs) [Windows.MessageBox]::Show($sender.Tag, 'Setting information') })
        [Windows.Controls.Grid]::SetColumn($label, 0); [void]$panel.Children.Add($label)
        [Windows.Controls.Grid]::SetColumn($info, 2); [void]$panel.Children.Add($info)
        $type = Get-SettingControlType $setting.Parameter
        if ($type -eq 'YesNo') {
            $holder = New-Object Windows.Controls.StackPanel -Property @{ Orientation = 'Horizontal'; VerticalAlignment = 'Center' }
            foreach ($choice in 'Yes','No') {
                $radio = New-Object Windows.Controls.RadioButton -Property @{ Content = $choice; GroupName = ('setting_' + $setting.Parameter); Margin = '0,0,16,0'; IsChecked = ($setting.Value -eq $choice) }
                $radio.Tag = $setting
                $radio.Add_Checked({ param($sender, $eventArgs) $sender.Tag.Value = $sender.Content.ToString(); if ($sender.Tag.Parameter -eq 'RunDownload') { [void]$window.Dispatcher.BeginInvoke([Action]{ Show-SettingsEditor }) } })
                [void]$holder.Children.Add($radio)
            }
            [Windows.Controls.Grid]::SetColumn($holder, 1); [void]$panel.Children.Add($holder)
        }
        elseif ($type -eq 'Choice') {
            $choices = if ($setting.Parameter -in @('RunProcess','RunFederation','ReviztoPublish')) { @('Yes','No','Force') } else { @('Full','OnlyCodes','OnlyDesc','Codes-Desc') }
            $combo = New-Object Windows.Controls.ComboBox -Property @{ ItemsSource = $choices; SelectedItem = $setting.Value; MinWidth = 220 }
            $combo.Tag = $setting; $combo.Add_SelectionChanged({ param($sender, $eventArgs) if ($sender.SelectedItem) { $sender.Tag.Value = $sender.SelectedItem.ToString() } })
            [Windows.Controls.Grid]::SetColumn($combo, 1); [void]$panel.Children.Add($combo)
        }
        elseif ($type -eq 'Folder') {
            $holder = New-Object Windows.Controls.DockPanel
            $browse = New-Object Windows.Controls.Button -Property @{ Content = '…'; ToolTip = 'Browse for folder'; Padding = '7,2'; MinWidth = 30; Margin = '8,0,0,0' }
            $textBox = New-Object Windows.Controls.TextBox -Property @{ Text = $setting.Value; Padding = '7,4' }
            $textBox.Tag = $setting; $textBox.Add_TextChanged({ param($sender, $eventArgs) $sender.Tag.Value = $sender.Text })
            $browse.Tag = $textBox; $browse.Add_Click({ param($sender, $eventArgs) $selectedPath = Select-ModernFolder $sender.Tag.Text; if ($selectedPath) { $sender.Tag.Text = ConvertTo-FEDAUTOStoredPath $selectedPath } })
            [Windows.Controls.DockPanel]::SetDock($browse, 'Right'); [void]$holder.Children.Add($browse); [void]$holder.Children.Add($textBox)
            [Windows.Controls.Grid]::SetColumn($holder, 1); [void]$panel.Children.Add($holder)
        }
        elseif ($type -eq 'File') {
            $holder = New-Object Windows.Controls.DockPanel
            $browse = New-Object Windows.Controls.Button -Property @{ Content = '…'; ToolTip = 'Browse for file'; Padding = '7,2'; MinWidth = 30; Margin = '8,0,0,0' }
            $textBox = New-Object Windows.Controls.TextBox -Property @{ Text = $setting.Value; Padding = '7,4' }
            $textBox.Tag = $setting; $textBox.Add_TextChanged({ param($sender, $eventArgs) $sender.Tag.Value = $sender.Text })
            $browse.Tag = [pscustomobject]@{ TextBox = $textBox; Parameter = $setting.Parameter }
            $browse.Add_Click({
                param($sender, $eventArgs)
                $dialog = New-Object Microsoft.Win32.OpenFileDialog
                $dialog.CheckFileExists = $true
                $dialog.Filter = if ($sender.Tag.Parameter -eq 'AttributesFile') { 'Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*' } else { 'XML files (*.xml)|*.xml|All files (*.*)|*.*' }
                $initialPath = $sender.Tag.TextBox.Text
                $dialog.InitialDirectory = Get-FEDAUTOInitialFolder $initialPath
                if ($initialPath) { $dialog.FileName = Split-Path -Leaf $initialPath }
                if ($dialog.ShowDialog()) { $sender.Tag.TextBox.Text = ConvertTo-FEDAUTOStoredPath $dialog.FileName }
            })
            [Windows.Controls.DockPanel]::SetDock($browse, 'Right'); [void]$holder.Children.Add($browse); [void]$holder.Children.Add($textBox)
            [Windows.Controls.Grid]::SetColumn($holder, 1); [void]$panel.Children.Add($holder)
        }
        else {
            $textBox = New-Object Windows.Controls.TextBox -Property @{ Text = $setting.Value; Padding = '7,4' }
            $textBox.Tag = $setting; $textBox.Add_TextChanged({ param($sender, $eventArgs) $sender.Tag.Value = $sender.Text })
            if ($setting.Parameter -eq 'AttributesFile') {
                $holder = New-Object Windows.Controls.StackPanel
                [void]$holder.Children.Add($textBox)
                [void]$holder.Children.Add((New-Object Windows.Controls.TextBlock -Property @{ Text='Stored automatically inside SourceFolder'; Foreground='#607D8B'; FontSize=11; Margin='0,3,0,0' }))
                [Windows.Controls.Grid]::SetColumn($holder, 1); [void]$panel.Children.Add($holder)
            }
            else { [Windows.Controls.Grid]::SetColumn($textBox, 1); [void]$panel.Children.Add($textBox) }
        }
        [void]$SettingsPanel.Children.Add($panel)
    }
    Update-FEDAUTOAcquisitionPresentation
}

function Paste-GridData {
    param($Grid)
    if (-not [Windows.Clipboard]::ContainsText() -or -not $Grid.CurrentCell) { return }
    $rows = @([Windows.Clipboard]::GetText() -split "`r?`n" | Where-Object { $_ -ne '' })
    if ($rows.Count -eq 0) { return }
    $columns = @($Grid.Columns | Sort-Object DisplayIndex)
    $startColumn = $Grid.CurrentCell.Column.DisplayIndex
    $startRow = [array]::IndexOf(@($Grid.ItemsSource), $Grid.CurrentItem)
    if ($startRow -lt 0) { $startRow = $Grid.ItemsSource.Count }
    $requiredRows = $startRow + $rows.Count
    while ($Grid.ItemsSource.Count -lt $requiredRows) {
        $newRow = [ordered]@{}
        foreach ($column in $columns) {
            $property = $column.SortMemberPath
            if (-not $property) { $property = $column.Header.ToString() }
            $newRow[$property] = if ($property -in @('Run','SkipIfSame','CheckDateToo','ExportToXLSX','InjectToIFC','ReadFromOutputFolder')) { $false } else { '' }
        }
        [void]$Grid.ItemsSource.Add([pscustomobject]$newRow)
    }
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $values = @($rows[$r] -split "`t", -1)
        # A row copied from the grid/Excel contains every table column.  Let it
        # paste as a full row even when the user selected a cell further right.
        $targetStartColumn = if ($values.Count -eq $columns.Count) { 0 } else { $startColumn }
        while (($targetStartColumn + $values.Count) -gt $columns.Count -and $values.Count -gt 0 -and [string]::IsNullOrWhiteSpace($values[-1])) {
            $values = if ($values.Count -eq 1) { @() } else { @($values[0..($values.Count - 2)]) }
        }
        if (($targetStartColumn + $values.Count) -gt $columns.Count) { throw 'The pasted data has more columns than this table.' }
        $item = $Grid.ItemsSource[$startRow + $r]
        for ($c = 0; $c -lt $values.Count; $c++) {
            $property = $columns[$targetStartColumn + $c].SortMemberPath
            if (-not $property) { $property = $columns[$targetStartColumn + $c].Header.ToString() }
            $item.$property = if ($property -in @('Run','SkipIfSame','CheckDateToo','ExportToXLSX','InjectToIFC','ReadFromOutputFolder')) { ConvertTo-FEDAUTOBoolean $values[$c] } else { ConvertTo-FEDAUTOCleanText $values[$c] }
        }
    }
    # Pasting can run while DataGrid is committing its current edit. Refreshing
    # immediately aborts that transaction, so schedule it for the UI queue.
    [void]$Grid.Dispatcher.BeginInvoke([Action]{ $Grid.Items.Refresh() })
}

function New-FEDAUTOGridRow {
    param([Parameter(Mandatory = $true)]$Grid)
    $row = [ordered]@{}
    foreach ($column in @($Grid.Columns | Sort-Object DisplayIndex)) {
        $property = $column.SortMemberPath
        if (-not $property) { $property = $column.Header.ToString() }
        if ([string]::IsNullOrWhiteSpace($property)) { continue }
        $row[$property] = if ($property -in @('Run','SkipIfSame','CheckDateToo','ExportToXLSX','InjectToIFC','ReadFromOutputFolder')) { $false } else { '' }
    }
    return [pscustomobject]$row
}

function Enable-GridClipboard {
    param($Grid)
    $Grid.Add_PreviewKeyDown({ param($sender, $eventArgs) if ($eventArgs.Key -eq [Windows.Input.Key]::V -and ([Windows.Input.Keyboard]::Modifiers -band [Windows.Input.ModifierKeys]::Control)) { try { Paste-GridData $sender; $eventArgs.Handled = $true } catch { [Windows.MessageBox]::Show($_.Exception.Message, 'Paste failed') } } })
    $Grid.Add_AddingNewItem({ param($sender, $eventArgs) $eventArgs.NewItem = New-FEDAUTOGridRow $sender })
}

$downloadFolderTemplate = [Windows.Markup.XamlReader]::Parse(@'
<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
  <DockPanel>
    <Button DockPanel.Dock="Right" Tag="BrowseReadFolder" Content="…" ToolTip="Browse for folder" Width="26" Height="22" Padding="0" Margin="4,0,0,0"/>
    <TextBox Text="{Binding ReadFolder, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" BorderThickness="0" VerticalContentAlignment="Center"/>
  </DockPanel>
</DataTemplate>
'@)
$DownloadGrid.Add_AutoGeneratingColumn({
    param($sender, $eventArgs)
    if ($eventArgs.PropertyName -in @('Run','SkipIfSame','CheckDateToo')) {
        $column = New-Object Windows.Controls.DataGridCheckBoxColumn
        $column.Header = $eventArgs.PropertyName
        $column.SortMemberPath = $eventArgs.PropertyName
        $binding = New-Object Windows.Data.Binding($eventArgs.PropertyName)
        $binding.Mode = [Windows.Data.BindingMode]::TwoWay
        $binding.UpdateSourceTrigger = [Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $column.Binding = $binding
        $eventArgs.Column = $column
    }
    elseif ($eventArgs.PropertyName -eq 'ReadFolder') {
        $column = New-Object Windows.Controls.DataGridTemplateColumn
        $column.Header = 'ReadFolder'
        $column.SortMemberPath = 'ReadFolder'
        $column.CellTemplate = $downloadFolderTemplate
        $column.ClipboardContentBinding = New-Object Windows.Data.Binding('ReadFolder')
        $eventArgs.Column = $column
    }
})
$AttributesGrid.Add_AutoGeneratingColumn({
    param($sender, $eventArgs)
    if ($eventArgs.PropertyName -in @('ExportToXLSX','InjectToIFC')) {
        $column = New-Object Windows.Controls.DataGridCheckBoxColumn
        $column.Header = $eventArgs.PropertyName
        $column.SortMemberPath = $eventArgs.PropertyName
        $binding = New-Object Windows.Data.Binding($eventArgs.PropertyName)
        $binding.Mode = [Windows.Data.BindingMode]::TwoWay
        $binding.UpdateSourceTrigger = [Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $column.Binding = $binding
        $eventArgs.Column = $column
    }
})
$federationOrderTemplate = [Windows.Markup.XamlReader]::Parse(@'
<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="30"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Border Width="24" Height="22" CornerRadius="11" HorizontalAlignment="Center" VerticalAlignment="Center">
      <Border.Style>
        <Style TargetType="Border">
          <Setter Property="Background" Value="#E0E0E0"/>
          <Style.Triggers>
            <DataTrigger Binding="{Binding GroupOrder}" Value="1"><Setter Property="Background" Value="#15803D"/></DataTrigger>
            <DataTrigger Binding="{Binding GroupOrder}" Value="2"><Setter Property="Background" Value="#65A30D"/></DataTrigger>
            <DataTrigger Binding="{Binding GroupOrder}" Value="3"><Setter Property="Background" Value="#CA8A04"/></DataTrigger>
            <DataTrigger Binding="{Binding GroupOrder}" Value="4"><Setter Property="Background" Value="#EA580C"/></DataTrigger>
            <DataTrigger Binding="{Binding GroupOrder}" Value="5"><Setter Property="Background" Value="#B91C1C"/></DataTrigger>
          </Style.Triggers>
        </Style>
      </Border.Style>
      <TextBlock Text="{Binding GroupOrder}" HorizontalAlignment="Center" VerticalAlignment="Center" FontWeight="Bold" Foreground="White"/>
    </Border>
    <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Left">
      <Button Tag="GroupUp" Content="&#x25B2;" ToolTip="Add to the group or move earlier" Width="20" Height="20" Padding="0"/>
      <Button Tag="GroupDown" Content="&#x25BC;" ToolTip="Move later" Width="20" Height="20" Padding="0" Margin="2,0,0,0"/>
      <Button Tag="GroupRemove" Content="&#x00D7;" ToolTip="Remove from grouping" Width="20" Height="20" Padding="0" Margin="2,0,0,0"/>
    </StackPanel>
  </Grid>
</DataTemplate>
'@)
$FederationGrid.Add_AutoGeneratingColumn({
    param($sender, $eventArgs)
    if ($eventArgs.PropertyName -eq 'InjectToIFC') {
        $column = New-Object Windows.Controls.DataGridCheckBoxColumn
        $column.Header = 'InjectToIFC'
        $column.SortMemberPath = 'InjectToIFC'
        $binding = New-Object Windows.Data.Binding('InjectToIFC')
        $binding.Mode = [Windows.Data.BindingMode]::TwoWay
        $binding.UpdateSourceTrigger = [Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $column.Binding = $binding
        $eventArgs.Column = $column
    }
    elseif ($eventArgs.PropertyName -eq 'GroupOrder') {
        $column = New-Object Windows.Controls.DataGridTemplateColumn
        $column.Header = 'GroupOrder'
        $column.SortMemberPath = 'GroupOrder'
        $column.CellTemplate = $federationOrderTemplate
        $eventArgs.Column = $column
    }
})
$WildcardSelectionGrid.Add_AutoGeneratingColumn({
    param($sender, $eventArgs)
    if ($eventArgs.PropertyName -eq 'ReadFromOutputFolder') {
        $column = New-Object Windows.Controls.DataGridCheckBoxColumn
        $column.Header = $eventArgs.PropertyName
        $column.SortMemberPath = $eventArgs.PropertyName
        $column.MinWidth = 190
        $column.Width = [Windows.Controls.DataGridLength]::new(200)
        $binding = New-Object Windows.Data.Binding($eventArgs.PropertyName)
        $binding.Mode = [Windows.Data.BindingMode]::TwoWay
        $binding.UpdateSourceTrigger = [Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $column.Binding = $binding
        $eventArgs.Column = $column
    }
    elseif ($eventArgs.PropertyName -in @('Inclusions','Exclusions')) {
        $eventArgs.Column.Width = [Windows.Controls.DataGridLength]::new(1, [Windows.Controls.DataGridLengthUnitType]::Star)
    }
    elseif ($eventArgs.PropertyName -eq 'ExportFileName') {
        $eventArgs.Column.MinWidth = 170
        $eventArgs.Column.Width = [Windows.Controls.DataGridLength]::new(190)
    }
})
$DownloadGrid.AddHandler([Windows.Controls.Button]::ClickEvent, [Windows.RoutedEventHandler]{
    param($sender, $eventArgs)
    $button = $eventArgs.OriginalSource
    if (-not ($button -is [Windows.Controls.Button]) -or $button.Tag -ne 'BrowseReadFolder') { return }
    $row = $button.DataContext
    if (-not $row) { return }
    $selectedPath = Select-ModernFolder $row.ReadFolder
    if ($selectedPath) { $row.ReadFolder = ConvertTo-FEDAUTOStoredPath $selectedPath; $DownloadGrid.Items.Refresh() }
    $eventArgs.Handled = $true
})
$FederationGrid.AddHandler([Windows.Controls.Button]::ClickEvent, [Windows.RoutedEventHandler]{
    param($sender, $eventArgs)
    $button = $eventArgs.OriginalSource
    if (-not ($button -is [Windows.Controls.Button]) -or $button.Tag -notin @('GroupUp','GroupDown','GroupRemove')) { return }
    $row = $button.DataContext
    if ($null -eq $row -or -not ($row.PSObject.Properties.Name -contains 'GroupOrder')) { return }
    $current = Get-FEDAUTOGroupOrderValue $row.GroupOrder
    $activeCount = @($script:FederationRows | Where-Object { (Get-FEDAUTOGroupOrderValue $_.GroupOrder) -gt 0 }).Count
    switch ($button.Tag) {
        'GroupUp' { Set-FEDAUTOGroupOrder -SelectedRow $row -RequestedOrder $(if ($current -le 0) { $activeCount + 1 } else { [Math]::Max(1, $current - 1) }) }
        'GroupDown' { Set-FEDAUTOGroupOrder -SelectedRow $row -RequestedOrder $(if ($current -le 0) { $activeCount + 1 } else { $current + 1 }) }
        'GroupRemove' { Set-FEDAUTOGroupOrder -SelectedRow $row -RequestedOrder 0 }
    }
    $eventArgs.Handled = $true
})
foreach ($grid in $DownloadGrid,$AttributesGrid,$FederationGrid,$WildcardSelectionGrid,$LookupsGrid) { Enable-GridClipboard $grid }
$DownloadGrid.Add_CellEditEnding({ param($sender, $eventArgs) [void]$window.Dispatcher.BeginInvoke([Action]{ Show-SettingsEditor }) })
function Get-FEDAUTOEditorRows {
    param(
        [Parameter(Mandatory = $true)]$Window,
        [Parameter(Mandatory = $true)][string]$ControlName
    )
    $control = $Window.FindName($ControlName)
    if ($null -eq $control) { throw "$ControlName is not loaded." }
    $source = if ($control -is [Windows.Controls.DataGrid]) { $control.ItemsSource } else { $control.Tag }
    if ($null -eq $source) { throw "$ControlName has no editor data." }
    return @($source | ForEach-Object { $_ })
}

function Get-FEDAUTOSettingsRowsFromWindow {
    param([Parameter(Mandatory = $true)]$Window)
    $panel = $Window.FindName('SettingsPanel')
    if ($null -eq $panel -or $null -eq $panel.Tag) { throw 'Settings editor data is not loaded.' }
    return @($panel.Tag)
}

function Commit-FEDAUTOEditorChanges {
    param([Parameter(Mandatory = $true)]$Window)
    foreach ($controlName in 'DownloadGrid','AttributesGrid','FederationGrid','WildcardSelectionGrid','LookupsGrid') {
        $grid = $Window.FindName($controlName)
        if ($grid) {
            [void]$grid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell, $true)
            [void]$grid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row, $true)
        }
    }
}

function Save-FEDAUTOConfiguration {
    param([Parameter(Mandatory = $true)]$Window)
    Commit-FEDAUTOEditorChanges -Window $Window
    $pathBox = $Window.FindName('ConfigPathBox')
    $status = $Window.FindName('StatusText')
    $path = $pathBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($path)) { throw 'Choose a JSON configuration path before saving.' }
    if ([IO.Path]::GetExtension($path).ToLowerInvariant() -ne '.json') {
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'JSON configuration (*.json)|*.json'
        $dialog.FileName = ([IO.Path]::GetFileNameWithoutExtension($path) + '.json')
        if (-not $dialog.ShowDialog()) { return }
        $path = $dialog.FileName
        $pathBox.Text = $path
    }

    $settingsRowsForSave = Get-FEDAUTOSettingsRowsFromWindow -Window $Window
    $downloadRowsForSave = Get-FEDAUTOEditorRows -Window $Window -ControlName 'DownloadGrid'
    $attributeRowsForSave = Get-FEDAUTOEditorRows -Window $Window -ControlName 'AttributesGrid'
    $federationRowsForSave = Get-FEDAUTOEditorRows -Window $Window -ControlName 'FederationGrid'
    $wildcardSelectionRowsForSave = Get-FEDAUTOEditorRows -Window $Window -ControlName 'WildcardSelectionGrid'
    $lookupRowsForSave = Get-FEDAUTOEditorRows -Window $Window -ControlName 'LookupsGrid'
    if ($settingsRowsForSave.Count -eq 0) { throw 'Settings are not loaded; refusing to save an empty configuration.' }
    $attributesFile = $settingsRowsForSave | Where-Object { $_.Parameter -eq 'AttributesFile' } | Select-Object -ExpandProperty Value -First 1
    if ($attributesFile -and ([IO.Path]::GetFileName($attributesFile) -ne $attributesFile -or $attributesFile.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0)) {
        throw 'AttributesFile must be a file name only, without a folder path. It is always stored inside SourceFolder.'
    }
    $settingsToSave = @(ConvertFrom-GridRows $settingsRowsForSave)
    $downloadToSave = @(ConvertFrom-GridRows $downloadRowsForSave @('Run','SkipIfSame','CheckDateToo'))
    $attributesToSave = @(ConvertFrom-GridRows $attributeRowsForSave @('ExportToXLSX','InjectToIFC'))
    $federationToSave = @(ConvertFrom-GridRows $federationRowsForSave @('InjectToIFC'))
    $wildcardSelectionToSave = @(ConvertFrom-GridRows $wildcardSelectionRowsForSave @('ReadFromOutputFolder') | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.Inclusions) -or -not [string]::IsNullOrWhiteSpace($_.Exclusions) -or
        -not [string]::IsNullOrWhiteSpace($_.ExportFileName) -or $_.ReadFromOutputFolder -eq 'Yes'
    })
    $lookupsToSave = @(ConvertFrom-GridRows $lookupRowsForSave)
    if ($settingsToSave.Count -eq 0) { throw 'No settings were collected from the editor; configuration was not changed.' }
    Save-PipelineJsonConfiguration -Path $path -Settings $settingsToSave -Download $downloadToSave -PWAttributesList $attributesToSave -Federation $federationToSave -WildcardSelection $wildcardSelectionToSave -Lookups $lookupsToSave
    Set-LastConfigurationPath $path
    $status.Text = "Saved $path"
    return $path
}

function Export-EditorConfigurationToExcel {
    param([Parameter(Mandatory = $true)]$Window)
    Commit-FEDAUTOEditorChanges -Window $Window
    Ensure-ModuleAvailable -Name ImportExcel
    Import-Module ImportExcel -ErrorAction Stop
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = 'Excel workbook (*.xlsx)|*.xlsx'
    $dialog.FileName = 'Federation-Automation-Config.xlsx'
    if (-not $dialog.ShowDialog()) { return }
    $path = $dialog.FileName
    $settingsRows = Get-FEDAUTOEditorRows -Window $Window -ControlName 'SettingsPanel'
    $settings = @(ConvertFrom-GridRows $settingsRows | ForEach-Object { [pscustomobject]@{ Parameter=$_.Parameter; Value=$_.Value; Desc=$_.Desc } })
    $sheets = @(
        @{ Name='Settings'; Table='Settings'; Rows=$settings },
        @{ Name='Download'; Table='Download'; Rows=(ConvertFrom-GridRows (Get-FEDAUTOEditorRows -Window $Window -ControlName 'DownloadGrid') @('Run','SkipIfSame','CheckDateToo')) },
        @{ Name='PWAttributesList'; Table='PWAttributesList'; Rows=(ConvertFrom-GridRows (Get-FEDAUTOEditorRows -Window $Window -ControlName 'AttributesGrid') @('ExportToXLSX','InjectToIFC')) },
        @{ Name='Federation'; Table='Federation'; Rows=(ConvertFrom-GridRows (Get-FEDAUTOEditorRows -Window $Window -ControlName 'FederationGrid') @('InjectToIFC')) },
        @{ Name='WildcardSelection'; Table='WildcardSelection'; Rows=(ConvertFrom-GridRows (Get-FEDAUTOEditorRows -Window $Window -ControlName 'WildcardSelectionGrid') @('ReadFromOutputFolder') | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Inclusions) -or -not [string]::IsNullOrWhiteSpace($_.Exclusions) -or -not [string]::IsNullOrWhiteSpace($_.ExportFileName) -or $_.ReadFromOutputFolder -eq 'Yes' }) },
        @{ Name='Lookups'; Table='Lookups'; Rows=(ConvertFrom-GridRows (Get-FEDAUTOEditorRows -Window $Window -ControlName 'LookupsGrid')) }
    )
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $first = $true
    foreach ($sheet in $sheets) {
        $params = @{ Path=$path; WorksheetName=$sheet.Name; TableName=$sheet.Table; AutoSize=$true }
        if (-not $first) { $params.Append = $true }
        @($sheet.Rows) | Export-Excel @params
        $first = $false
    }
    $package = Open-ExcelPackage -Path $path
    try {
        $worksheet = $package.Workbook.Worksheets['Settings']
        $colours = @{ 'Source acquisition'='#D9E8F5'; 'Working folders & metadata'='#E8ECEF'; 'IFC processing'='#FBE4D5'; 'Federation & Navisworks'='#DFF0C8'; 'Revizto publishing'='#FFF59D'; 'Other settings'='#EDE7F6' }
        for ($i = 0; $i -lt $settingsRows.Count; $i++) {
            $colour = if ($colours.ContainsKey($settingsRows[$i].Section)) { $colours[$settingsRows[$i].Section] } else { '#E8ECEF' }
            $range = $worksheet.Cells[($i + 2), 1, ($i + 2), 3]
            $range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $range.Style.Fill.BackgroundColor.SetColor([Drawing.ColorTranslator]::FromHtml($colour))
        }
    }
    finally { Close-ExcelPackage $package }
    $Window.FindName('StatusText').Text = "Exported Excel configuration: $path"
}

function Set-FEDAUTORunStatus {
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)][string]$Detail,
        [ValidateRange(0,100)][int]$Progress
    )
    $RunStageText.Text = $Stage
    $RunDetailText.Text = $Detail
    $RunProgressBar.Value = $Progress
}

function Update-FEDAUTORunStatusFromLine {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    $text = $Line.Trim()
    if ($text -match '=== Download ===|=== START 020-PWDownload ===|=== START Download ===') { Set-FEDAUTORunStatus -Stage '1 of 4 — Source acquisition' -Detail 'Retrieving or copying source model files.' -Progress 15; return }
    if ($text -match '=== Process IFC Attributes ===|=== START.*Process') { Set-FEDAUTORunStatus -Stage '2 of 4 — IFC processing' -Detail 'Reading and updating IFC metadata.' -Progress 40; return }
    if ($text -match '=== Group/Federate Files ===|=== START.*(Group|Federate)') { Set-FEDAUTORunStatus -Stage '3 of 4 — Federation' -Detail 'Grouping models and creating the federated model.' -Progress 65; return }
    if ($text -match '=== Publish Revizto ===|=== START.*Publish Revizto') { Set-FEDAUTORunStatus -Stage '4 of 4 — Revizto publishing' -Detail 'Publishing the federated model.' -Progress 85; return }
    if ($text -match '=== Pipeline Totals ===') { Set-FEDAUTORunStatus -Stage 'Finishing' -Detail 'Writing run totals and finalising output.' -Progress 95; return }
    if ($text -match '^(ERROR:|WARNING:|WARN:)') { $RunDetailText.Text = $text; return }
    if ($text -notmatch '^(===|Logging to |ConfigFile:|Process start time|------)$') { $RunDetailText.Text = $text }
}

function Add-FEDAUTOActivityLine {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    $colour = '#1F2933'
    if ($Line -match '(?i)\b(error|failed|failure|exception)\b') { $colour = '#C62828' }
    elseif ($Line -match '(?i)\b(warning|warn|skipped)\b') { $colour = '#E67E22' }
    $paragraph = New-Object Windows.Documents.Paragraph
    $paragraph.Margin = '0'
    $run = New-Object Windows.Documents.Run $Line
    $run.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($colour)
    [void]$paragraph.Inlines.Add($run)
    [void]$ActivityBox.Document.Blocks.Add($paragraph)
    $ActivityBox.ScrollToEnd()
    Update-FEDAUTORunStatusFromLine $Line
}

function Start-FEDAUTOBackgroundRun {
    param([string]$ConfigPath)
    $MainTabs.SelectedIndex = $MainTabs.Items.Count - 1
    $ActivityBox.Document.Blocks.Clear()
    Set-FEDAUTORunStatus -Stage 'Preparing run' -Detail 'Saving configuration and starting the pipeline.' -Progress 0
    Add-FEDAUTOActivityLine 'Starting Federation Automation...'
    Add-FEDAUTOActivityLine ("Configuration: {0}" -f $ConfigPath)
    $StatusText.Text = 'Pipeline is running...'
    $RunButton.IsEnabled = $false

    # The GUI distribution is paired with the compiled pipeline.  Do not invoke
    # the development .ps1 entry point here: it is not part of an EXE deployment.
    $pipelinePath = Join-Path $basePath '006-Main.exe'
    if (-not (Test-Path -LiteralPath $pipelinePath -PathType Leaf)) {
        throw "Pipeline executable not found: $pipelinePath"
    }
    $arguments = ('-ConfigFile "{0}"' -f $ConfigPath)
    $script:activeBackgroundRun = [FEDAUTO.BackgroundRun]::Start($pipelinePath, $arguments)
    $script:runOutputTimer = New-Object Windows.Threading.DispatcherTimer
    $script:runOutputTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:runOutputTimer.Add_Tick({
        $line = $null
        while ($script:activeBackgroundRun.TryGetLine([ref]$line)) {
            Add-FEDAUTOActivityLine $line
        }
        if ($script:activeBackgroundRun.Process.HasExited) {
            # Drain anything queued immediately before the process exited.
            while ($script:activeBackgroundRun.TryGetLine([ref]$line)) { Add-FEDAUTOActivityLine $line }
            $exitCode = $script:activeBackgroundRun.Process.ExitCode
            Add-FEDAUTOActivityLine ("Pipeline finished with exit code {0}." -f $exitCode)
            $script:runOutputTimer.Stop()
            $RunButton.IsEnabled = $true
            $StatusText.Text = if ($exitCode -eq 0) { 'Pipeline completed successfully.' } else { "Pipeline failed (exit code $exitCode)." }
            if ($exitCode -eq 0) { Set-FEDAUTORunStatus -Stage 'Completed' -Detail 'The pipeline completed successfully.' -Progress 100 }
            else { Set-FEDAUTORunStatus -Stage 'Run failed' -Detail "The pipeline ended with exit code $exitCode. Review the live activity log." -Progress ([int]$RunProgressBar.Value) }
        }
    })
    $script:runOutputTimer.Start()
}

$OpenButton.Add_Click({ $dialog = New-Object Microsoft.Win32.OpenFileDialog; $dialog.Filter = 'Pipeline configuration (*.json;*.xlsx;*.xlsm)|*.json;*.xlsx;*.xlsm|All files|*.*'; if ($dialog.ShowDialog()) { try { Set-EditorConfiguration (Get-PipelineConfiguration -ConfigPath $dialog.FileName -BasePath $basePath) $dialog.FileName } catch { [System.Windows.MessageBox]::Show(($_.Exception.Message + "`r`n`r`n" + $_.ScriptStackTrace), 'Unable to open configuration') } } })
$NewButton.Add_Click({ $dialog = New-Object Microsoft.Win32.SaveFileDialog; $dialog.Filter = 'JSON configuration (*.json)|*.json'; $dialog.FileName = 'Config.json'; if ($dialog.ShowDialog()) { Set-EditorConfiguration ([pscustomobject]@{ Format='Json'; Settings=@(); Download=@(); PWAttributesList=@(); Federation=@(); WildcardSelection=@(); Lookups=@() }) $dialog.FileName } })
$SaveButton.Add_Click({
    param($sender, $eventArgs)
    try { Save-FEDAUTOConfiguration -Window ([Windows.Window]::GetWindow($sender)) | Out-Null }
    catch { [System.Windows.MessageBox]::Show($_.Exception.Message, 'Unable to save configuration') }
})
$ExportExcelButton.Add_Click({
    param($sender, $eventArgs)
    try { Export-EditorConfigurationToExcel -Window ([Windows.Window]::GetWindow($sender)) }
    catch { [System.Windows.MessageBox]::Show($_.Exception.Message, 'Unable to export Excel configuration') }
})
$ValidateButton.Add_Click({
    param($sender, $eventArgs)
    try {
        $owner = [Windows.Window]::GetWindow($sender)
        $savedPath = Save-FEDAUTOConfiguration -Window $owner
        $null = Get-PipelineConfiguration -ConfigPath $savedPath -BasePath $basePath
        $owner.FindName('StatusText').Text = 'Configuration is valid.'
        [System.Windows.MessageBox]::Show('Configuration is valid.', 'Federation Automation')
    }
    catch { [System.Windows.MessageBox]::Show($_.Exception.Message, 'Validation failed') }
})
$RunButton.Add_Click({
    param($sender, $eventArgs)
    try {
        $savedPath = Save-FEDAUTOConfiguration -Window ([Windows.Window]::GetWindow($sender))
        Start-FEDAUTOBackgroundRun -ConfigPath $savedPath
    }
    catch { $sender.IsEnabled = $true; [System.Windows.MessageBox]::Show($_.Exception.Message, 'Unable to start pipeline') }
})

$previousState = Get-GuiState
$previousSessionCompleted = -not $previousState -or $previousState.LastSessionCompleted -eq $true
# Mark this session as unclean until the window reaches its normal Closing event.
Save-GuiState -LastConfigFile $(if ($previousState) { $previousState.LastConfigFile } else { $null }) -LastSessionCompleted:$false
$window.Add_Closing({
    $path = $ConfigPathBox.Text.Trim()
    Save-GuiState -LastConfigFile $path -LastSessionCompleted:$true
})

$startupConfig = if ($ConfigFile) { $ConfigFile } elseif ($previousSessionCompleted) { Get-DefaultConfigurationPath } else { $null }
if ($startupConfig) { try { Set-EditorConfiguration (Get-PipelineConfiguration -ConfigPath $startupConfig -BasePath $basePath) $startupConfig } catch { [System.Windows.MessageBox]::Show(($_.Exception.Message + "`r`n`r`n" + $_.ScriptStackTrace), 'Unable to open configuration') } }
elseif (-not $ConfigFile -and -not $previousSessionCompleted) { $StatusText.Text = 'Previous session did not close cleanly. Choose a configuration to continue.' }
[void]$window.ShowDialog()
