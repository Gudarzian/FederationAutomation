<#
Download-related functions for Federation Automation.
- ProjectWise/ACC download routines and helpers.
#>

if (-not (Get-Variable -Name PWAttributeDefinitions -Scope Script -ErrorAction SilentlyContinue)) {
    $script:PWAttributeDefinitions = @()
}

function Initialize-PWAttributeList {
<#
Reads the PWAttributesList named range and stores attribute definitions for Add2List.
- `AttributeName` is the source ProjectWise field/path.
- `OutputName` is used for the exported workbook column name and the IFC property name.
- Rows are exported when `InjectToIFC` or `ExportToXLSX` is enabled.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][array]$Rows,
        [string]$AttributeColumn = 'AttributeName',
        [string]$OutputColumn = 'OutputName'
    )

    function Get-NormalizedText {
        param($Value)
        if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            return Normalize-ExcelText -Text $Value
        }
        if ($null -eq $Value) { return $null }
        return $Value.ToString().Trim()
    }

    function Test-TrueLike {
        param($Value)
        $text = Get-NormalizedText -Value $Value
        if ([string]::IsNullOrWhiteSpace($text)) { return $false }
        return ($text.ToLowerInvariant() -in @('yes', 'y', 'true', '1'))
    }

    $definitions = @()
    foreach ($row in $Rows) {
        if (-not $row) { continue }
        if (-not ($row.PSObject.Properties.Name -contains $AttributeColumn)) { continue }
        $label = Get-NormalizedText -Value $row.$AttributeColumn
        if ([string]::IsNullOrWhiteSpace($label)) { continue }

        $injectEnabled = Test-TrueLike -Value (Get-RowValue -Row $row -Candidates @('InjectToIFC'))
        $exportEnabled = Test-TrueLike -Value (Get-RowValue -Row $row -Candidates @('ExportToXLSX'))
        $hasExplicitSelection = ($row.PSObject.Properties.Name | Where-Object {
            $_ -in @('InjectToIFC','ExportToXLSX')
        }).Count -gt 0
        if ($hasExplicitSelection -and -not ($injectEnabled -or $exportEnabled)) { continue }

        $outputLabel = $null
        if ($row.PSObject.Properties.Name -contains $OutputColumn) {
            $outputLabel = Get-NormalizedText -Value $row.$OutputColumn
        }
        if ([string]::IsNullOrWhiteSpace($outputLabel)) { $outputLabel = $label }

        $definitions += [pscustomobject]@{
            Label = $outputLabel
            Field = $label
        }
    }

    $script:PWAttributeDefinitions = $definitions
}

function Resolve-PropertyPath {
<#
Resolves a dotted path against objects or dictionaries.
#>
    param(
        $Target,
        [string]$Path
    )
    if ($null -eq $Target -or [string]::IsNullOrWhiteSpace($Path)) { return $null }
    $current = $Target
    foreach ($segment in ($Path -split '\.')) {
        if ([string]::IsNullOrWhiteSpace($segment)) { continue }
        if ($current -is [System.Collections.IDictionary]) {
            if ($current.Contains($segment)) { $current = $current[$segment] }
            else { return $null }
        }
        elseif ($current.PSObject.Properties.Name -contains $segment) {
            $current = $current.$segment
        }
        else {
            return $null
        }
    }
    return $current
}

function Get-DocAttributeValue {
<#
Gets a document attribute by field name, including CustomAttributes fallback.
#>
    param(
        $Document,
        [string]$Field
    )
    if ($null -eq $Document -or [string]::IsNullOrWhiteSpace($Field)) { return $null }

    $value = Resolve-PropertyPath -Target $Document -Path $Field
    if ($null -eq $value) {
        $value = Resolve-PropertyPath -Target $Document.CustomAttributes -Path $Field
    }
    return $value
}

function Invoke-AutoCADScript {
<#
Runs AutoCAD scripts against a DWG and writes the output file.
#>
    param (
        [string]$InputDWG,
        [string]$ScriptFiles,
        [string]$OutputDWG
    )
   if ($null -eq (Get-Variable -Name "acadApp" -ErrorAction SilentlyContinue)) {
        Write-Host "Running AutoCAD"
        $global:acadApp = New-Object -ComObject "AutoCAD.Application"
        #$acadApp.Visible = $true
        do{
            $Failed = $false
            Try{
                $acadApp.ActiveDocument.SendCommand("XLOADCTL 0`n")    # Disable DWG Xrefs
                $acadApp.ActiveDocument.SendCommand("IMAGEFRAME 0`n")  # Disable image loading
                $acadApp.ActiveDocument.SendCommand("DWFFRAME 0`n")    # Disable DWF loading
                $acadApp.ActiveDocument.SendCommand("FILEDIA 0`n")     # Disable file dialog
            } catch {     
                Start-Sleep -Milliseconds 500
                $Failed = $true }
            } while ($Failed)
            foreach ($doc in @($acadApp.Documents)) {
                $doc.Close($false)  # false = do not save
            }
    } 
    $Counter=0
    do{
        $Failed = $false
        Try{
            $document = $acadApp.Documents.Open($InputDWG)
            Start-Sleep -Milliseconds 5000
        } 
        catch {
            $Counter++
            Start-Sleep -Milliseconds 5000
            $Failed = $true 
            Write-Host "Error Openning File - Trying again > $Counter"
            }
    } while ($Failed -and $Counter -le 5 )
    if( $Counter -eq 5){
        Write-Host "Error Openning File - Exiting after $Counter Try"
    }
    $ScriptFilesList = $ScriptFiles -split ","
    Foreach ($ScriptFile in $ScriptFilesList) 
    {
    $Counter=0
    $ScriptFile = Join-Path $basePath $ScriptFile
    if (!(Test-Path -Path $ScriptFile)) 
        {Write-Output "Script file does not exist: $ScriptFile"; return}
    do{
        $Failed = $false
        Try{
            $acadApp.ActiveDocument.SendCommand("FILEDIA 0`n")
            $acadApp.ActiveDocument.SendCommand('_.SCRIPT "' + $ScriptFile + '"' + [Environment]::NewLine)
            Start-Sleep -Milliseconds 5000
        } 
        catch {
            $Counter++
            Start-Sleep -Milliseconds 5000
            $Failed = $true 
            Write-Host "Error Running the Script File - Trying again> $Counter"
            }
    } while ($Failed -and $Counter -le 5 )
    if( $Counter -eq 5)
    {
        Write-Host "Error Running the Script File - Exiting after $Counter Try"
        }
    }    
    $Counter=0
    do{
        $Failed = $false
        Try{
            $document.SaveAs($OutputDWG)
           } 
        catch {
            $Counter++
            Start-Sleep -Milliseconds 2000
            $Failed = $true 
            Write-Host "Error Saving the updated File - Trying again> $Counter"
            }
    } while ($Failed -and $Counter -le 5 )
    if( $Counter -eq 5)
    {
        Write-Host "Error Saving the updated File - Exiting after $Counter Try"
        }
    $Counter=0
    do{
        $Failed = $false
        Try{
            $document.Close($false)
           } 
        catch {
            $Counter++
            Start-Sleep -Milliseconds 2000
            $Failed = $true 
            Write-Host "Error Closing the Saved File - Trying again> $Counter"
            }
    } while ($Failed -and $Counter -le 5 )
    if( $Counter -eq 5){
        Write-Host "Error Closing the Saved File - Exiting after $Counter Try"
    }
}

Function DPU($Message) {
# Download Process Update
<#
Downloads a PW document and optionally runs AutoCAD processing.
#>
    $downloadStart = Get-Date
    $displayName = if ($TempD.FileName) { $TempD.FileName } else { $TempD.Name }
    Write-host $displayName -NoNewline -ForegroundColor Green
    Write-host " Downloading >>>> " $TempD.FileSize " ...bytes " -NoNewline -ForegroundColor Red
    Write-host   -ForegroundColor White
    try {
        $null= CheckOut-PWDocuments -InputDocument $TempD -copyout -NoReferences -ExportFolder $LocalFolder -ErrorAction Stop
    }
    catch {
        Write-Warning "Download failed for $($TempD.FileName): $_"
        return
    }
    if (-not (Get-Variable -Name PWDownloadedCount -Scope Global -ErrorAction SilentlyContinue)) {
        $global:PWDownloadedCount = 0
    }
    $global:PWDownloadedCount++
    Write-host " >>>> downloaded - $Message"
    $DownloadedFile = Join-Path $LocalFolder $TempD.FileName
    $ProcessedFile = Join-Path $Tempfolder $TempD.FileName
    $TargetFile = Join-Path $F.TargetFolder $TempD.FileName
    if ($null -eq $F.ScriptFiles)
        {   
        Write-host " >>>> $($TempD.FileName) - No script files provided, copying without processing"
        Copy-Item -Path $DownloadedFile -Destination $TargetFile -Force
        Write-host " >>>> $($TempD.FileName) - saved to $($F.TargetFolder)"
        }
    else 
        {
        Write-host " >>>> Running $($F.ScriptFiles) on $($TempD.FileName) - Processing the file"
        Invoke-AutoCADScript -InputDWG $DownloadedFile -ScriptFiles $F.ScriptFiles -OutputDWG $ProcessedFile 
        Copy-Item -Path $ProcessedFile -Destination $TargetFile -Force
        Get-ChildItem -Path $ProcessedFile -File -Recurse | Remove-Item -Force                    
        Write-host " >>>> $($TempD.FileName) - Processed and saved to $($F.TargetFolder)"
        }
    $downloadEnd = Get-Date
    $downloadDuration = New-TimeSpan -Start $downloadStart -End $downloadEnd
    Write-Host " >>>> Duration: $([math]::Round($downloadDuration.TotalSeconds,2)) seconds"
}

Function DU($Message) {
# Download Process Update
 <#
Downloads a PW document without processing.
#>
   $downloadStart = Get-Date
    $displayName = if ($TempD.FileName) { $TempD.FileName } else { $TempD.Name }
    Write-host $displayName -NoNewline -ForegroundColor Green
    Write-host " Downloading >>>> " $TempD.FileSize " ...bytes " -NoNewline -ForegroundColor Red
    Write-host   -ForegroundColor White
    try {
        $null= CheckOut-PWDocuments -InputDocument $TempD -copyout -NoReferences -ExportFolder $LocalFolder -ErrorAction Stop
    }
    catch {
        Write-Warning "Download failed for $($TempD.FileName): $_"
        return
    }
    if (-not (Get-Variable -Name PWDownloadedCount -Scope Global -ErrorAction SilentlyContinue)) {
        $global:PWDownloadedCount = 0
    }
    $global:PWDownloadedCount++
    $downloadEnd = Get-Date
    $downloadDuration = New-TimeSpan -Start $downloadStart -End $downloadEnd
    Write-host " >>>> downloaded - $Message"
    Write-host " >>>> $($TempD.FileName) - saved to $($LocalFolder)"
    Write-Host " >>>> Duration: $([math]::Round($downloadDuration.TotalSeconds,2)) seconds"
}

function Add2List($doc){
<#
Adds a document row to the global list (dynamic attributes if configured).
#>
    if ($script:PWAttributeDefinitions -and $script:PWAttributeDefinitions.Count -gt 0) {
        $row = [ordered]@{}
        foreach ($def in $script:PWAttributeDefinitions) {
            $label = $def.Label
            $field = if ($def.Field) { $def.Field } else { $def.Label }
            $row[$label] = Get-DocAttributeValue -Document $doc -Field $field
        }
        $global:list += [pscustomobject]$row
        return
    }

    # Fallback to original hard-coded main attributes if no definitions are loaded
    $global:list += [PSCustomObject]@{
        Title              = $doc.Name
        FileName           = $doc.FileName
        PW_Description     = $doc.Description
        PW_Status          = $doc.DocumentStatus
        PW_File_Size       = $doc.FileSize
        PW_Document_Version= $doc.Version
        PW_Created_By      = $doc.DocumentCreatorName
        PW_Creation_Time   = $doc.CreateDate
        PW_Updated_By      = $doc.DocumentUpdater.Email
        PW_Doc_GUID        = $doc.DocumentGUID
        PW_Doc_URN         = $doc.DocumentURN
        File_UpdaterEmail  = $doc.FileUpdater.Email
        File_UpdateDate    = $doc.FileUpdateDate
        PW_Doc_Modification_Time = $doc.DocumentUpdateDate
        }
}

function Invoke-PWDownload {
<#
Downloads IFCs from ProjectWise and writes the attributes workbook.
#>
    [CmdletBinding()]
    param(
        [string]$InputExcelFile = "Config.json",
        [array]$Settings = @(),
        [Parameter(Mandatory = $true)][array]$DownloadRows,
        [Parameter(Mandatory = $true)][array]$PWAttributeListRows,
        [string]$LogsFolder,
        [string]$PWUser,
        [string]$PWPass
    )

    $basePath = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
    elseif ($null -ne $MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne '') {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    else {
        Add-Type -AssemblyName System.Reflection
        Split-Path -Parent ([System.Reflection.Assembly]::GetEntryAssembly().Location)
    }

    $global:list = @()
    $global:PWDeletedCount = 0
    $global:PWDownloadedCount = 0
    $global:PWDownloadStatus = 'Started'
    $deletedCount = 0
    $searchAttemptCount = 0
    $searchFailedCount = 0
    $searchFailureDetails = New-Object System.Collections.Generic.List[string]

    # JSON/CSV build: no runtime module installation or execution-policy changes.
    # ProjectWise support is loaded only if a ProjectWise row is used.

    $ConfigFilePath = if ([System.IO.Path]::IsPathRooted($InputExcelFile)) {
        $InputExcelFile
    }
    else {
        Join-Path $basePath $InputExcelFile
    }
    $Config = @($Settings)
    $LogsFolder     = if ($LogsFolder) { $LogsFolder } else { CureFolderPath($Config | Where-Object { $_.Parameter -eq "LogFolder" } | Select-Object -ExpandProperty Value) }
    $AttributesFile = $Config | Where-Object { $_.Parameter -eq "AttributesFile" } | Select-Object -ExpandProperty Value
    $LocalFolder     = CureFolderPath ($Config | Where-Object { $_.Parameter -eq "SourceFolder" } | Select-Object -ExpandProperty Value)
    $ProcessedFolder = CureFolderPath ($Config | Where-Object { $_.Parameter -eq "ProcessedFolder" } | Select-Object -ExpandProperty Value)
    if ([string]::IsNullOrWhiteSpace($LocalFolder)) { $LocalFolder = "SourceFiles" }
    if ([string]::IsNullOrWhiteSpace($ProcessedFolder)) { $ProcessedFolder = "ProcessedIFC" }
    if ([string]::IsNullOrWhiteSpace($AttributesFile)) { $AttributesFile = "PWAttributes.csv" }

    if (-not $PWUser) {
        $PWUser = $Config | Where-Object { $_.Parameter -eq "PWUser" } | Select-Object -ExpandProperty Value
    }
    if (-not $PWPass) {
        $PWPass = $Config | Where-Object { $_.Parameter -eq "PWPass" } | Select-Object -ExpandProperty Value
    }
    if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
        $PWUser = Normalize-ExcelText -Text $PWUser
        $PWPass = Normalize-ExcelText -Text $PWPass
    }
    else {
        if ($null -ne $PWUser) { $PWUser = $PWUser.ToString().Trim() }
        if ($null -ne $PWPass) { $PWPass = $PWPass.ToString().Trim() }
    }

    $LogsFolder = Resolve-RelativePath -PathValue $LogsFolder -Root $basePath
    $LocalFolder = Resolve-RelativePath -PathValue $LocalFolder -Root $basePath
    $ProcessedFolder = Resolve-RelativePath -PathValue $ProcessedFolder -Root $basePath
    try {
        if (-not (Test-Path $LocalFolder -PathType Container)) {
            New-Item -Path $LocalFolder -ItemType Directory -Force | Out-Null
        }
        $LocalFolder = (Resolve-Path $LocalFolder).Path
    }
    catch {
        throw "Could not create or access SourceFolder '$LocalFolder'. Error: $_"
    }
    $PWAttributes = Join-Path $LocalFolder $AttributesFile
    $previousAttributeRows = @()
    if (Test-Path $PWAttributes -PathType Leaf) {
        try {
            $previousAttributeRows = Get-ExcelDataSafe -Path $PWAttributes -NamedRange 'Attributes'
        }
        catch {
            Write-Warning "Unable to read previous attributes workbook '$PWAttributes' for deleted-file metadata. Error: $_"
            $previousAttributeRows = @()
        }
    }

    if (-not $LogsFolder) {
        $LogsFolder = Join-Path $basePath 'Logs'
    }
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
    Write-Host "Logging to $LogFilePath"
    Write-Host "=== START 020-PWDownload ==="
    Write-Host "Config file: $ConfigFilePath"
    Write-Host "Source folder: $LocalFolder"

    Initialize-PWAttributeList -Rows $PWAttributeListRows -AttributeColumn 'AttributeName' -OutputColumn 'OutputName'

    # Download.ReadFolder accepts ProjectWise addresses or local/synced filesystem folders.
    # Local/synced folders include absolute paths, relative paths, ACC Desktop Connector,
    # OneDrive, SharePoint, UNC, and normal local folders.
    $FS = @($DownloadRows) | Where-Object {
        $runValue = $null
        if ($_.PSObject.Properties.Name -contains 'Run') {
            $runValue = $_.Run
        }
        if ($null -eq $runValue) { return $false }
        $runText = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $runValue
        }
        else {
            $runValue.ToString().Trim()
        }
        if ([string]::IsNullOrWhiteSpace($runText)) { return $false }
        $runText.ToLowerInvariant() -in @('yes','y','true','1')
    }

    if (-not $FS -or $FS.Count -eq 0) {
        throw "No active rows found in the Download named range. Set Run to Yes for at least one row."
    }

    <#
    Parses a PW folder value from Config.json (usually copy/pasted from PWE)
    into datasource + searchable folder path.
    #>
    function Resolve-ProjectWiseReadFolderInfo {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
        $raw = $Path.Trim()
        if ($raw -match '^pw://(?<ds>[^/]+)/(?<rest>.*)$') {
            $ds = $Matches['ds']
            $rest = $Matches['rest']
            $decoded = $rest -replace '&space;', ' ' -replace '&amp;', '&' -replace '&&', '&'
            $decoded = $decoded -replace '/', '\'
            $decoded = $decoded.TrimEnd('\')
            # For pw:// style paths copied from PWE, remove a leading "Documents\"
            # because Get-PWDocumentsBySearch in this workflow expects project-relative path.
            if ($decoded -match '^(?i)Documents\\') { $decoded = $decoded.Substring(10) }
            return [pscustomobject]@{
                DataSource = $ds
                FolderPath = $decoded
                Raw        = $raw
            }
        }
        if ($raw -match '^pw:\\\\?(?<ds>[^\\]+)\\?(?<rest>.*)$') {
            $ds = $Matches['ds']
            $rest = $Matches['rest']
            $decoded = $rest -replace '&space;', ' ' -replace '&amp;', '&' -replace '&&', '&'
            $decoded = $decoded -replace '/', '\'
            $decoded = $decoded.TrimEnd('\')
            # Same normalization as pw:// form.
            if ($decoded -match '^(?i)Documents\\') { $decoded = $decoded.Substring(10) }
            return [pscustomobject]@{
                DataSource = $ds
                FolderPath = $decoded
                Raw        = $raw
            }
        }
        $normalized = $raw -replace '/', '\'
        if ($normalized -match '^[A-Za-z]:\\') {
            return $null
        }
        if ($normalized -match '^\\\\') {
            return $null
        }
        if ($normalized -match '^([A-Za-z0-9_.-]+):\\?(?<rest>.*)$') {
            # "datasource:\Project\Folder" form: datasource is taken from the prefix.
            return [pscustomobject]@{
                DataSource = $Matches[1]
                FolderPath = ($Matches['rest'] -replace '^\\', '')
                Raw        = $raw
            }
        }
        return $null
    }

    function Test-ProjectWiseReadFolder {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
        return $null -ne (Resolve-ProjectWiseReadFolderInfo -Path $Path)
    }

    <#
    Converts Windows wildcards (*, ?) into SQL LIKE patterns.
    #>
    function Convert-WindowsWildcardToSqlLike {
        param([string]$Pattern)
        if ([string]::IsNullOrWhiteSpace($Pattern)) { return $Pattern }
        return ($Pattern -replace '\*', '%' -replace '\?', '_')
    }

    function Test-PWNotLoggedError {
        param($ErrorRecord)
        if (-not $ErrorRecord) { return $false }
        $message = $ErrorRecord.ToString()
        return ($message -match 'Not logged into a ProjectWise datasource')
    }

    function Get-ReadFolderValue {
        param($Row)
        if (-not ($Row.PSObject.Properties.Name -contains 'ReadFolder')) { return $null }
        return Normalize-ExcelText -Text $Row.ReadFolder
    }

    function Test-YesLike {
        param($Value)
        if ($null -eq $Value) { return $false }
        $text = if (Get-Command Normalize-ExcelText -ErrorAction SilentlyContinue) {
            Normalize-ExcelText -Text $Value
        }
        else {
            $Value.ToString().Trim()
        }
        return (-not [string]::IsNullOrWhiteSpace($text)) -and $text.ToLowerInvariant() -in @('yes','y','true','1')
    }

    function Get-ExcludeTerms {
        param($Value)
        if ($null -eq $Value) { return @() }
        return @($Value.ToString() -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    function Get-FileFilterPatterns {
        param($Value)
        if ($null -eq $Value) { return @('*') }
        $patterns = @($Value.ToString() -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($patterns.Count -eq 0) { return @('*') }
        return $patterns
    }

    function Write-DownloadProgress {
        param(
            [string]$Status,
            [string]$FileName
        )
        $safeStatus = if ([string]::IsNullOrWhiteSpace($Status)) { 'Processing' } else { $Status }
        $safeFile = if ($FileName) { $FileName -replace '\|', ' ' } else { '' }
        Write-Host ("DOWNLOAD_PROGRESS|FolderCurrent={0}|FolderTotal={1}|Current={2}|Total={3}|Status={4}|File={5}" -f `
            $script:FEDAUTODownloadFolderCurrent,
            $script:FEDAUTODownloadFolderTotal,
            $script:FEDAUTODownloadFileCurrent,
            $script:FEDAUTODownloadFileTotal,
            $safeStatus,
            $safeFile)
    }

    function Add-LocalFileToList {
        param(
            [System.IO.FileInfo]$File,
            [string]$ReadFolder
        )

        $commonValues = @{
            Title = $File.Name
            FileName = $File.Name
            FileSize = $File.Length
            File_Size = $File.Length
            PW_File_Size = $File.Length
            FileUpdateDate = $File.LastWriteTime
            File_UpdateDate = $File.LastWriteTime
            PW_File_UpdateDate = $File.LastWriteTime
            SourceFolder = $LocalFolder
            SourcePath = $File.FullName
            ReadFolder = $ReadFolder
        }

        if ($script:PWAttributeDefinitions -and $script:PWAttributeDefinitions.Count -gt 0) {
            $row = [ordered]@{}
            function Test-RowKey {
                param([string]$Name)
                foreach ($key in $row.Keys) {
                    if ($key -eq $Name) { return $true }
                }
                return $false
            }
            foreach ($def in $script:PWAttributeDefinitions) {
                $label = $def.Label
                $field = if ($def.Field) { $def.Field } else { $def.Label }
                if ($commonValues.ContainsKey($label)) {
                    $row[$label] = $commonValues[$label]
                }
                elseif ($commonValues.ContainsKey($field)) {
                    $row[$label] = $commonValues[$field]
                }
                else {
                    $row[$label] = $null
                }
            }
            if (-not (Test-RowKey -Name 'FileName')) { $row['FileName'] = $File.Name }
            if (-not (Test-RowKey -Name 'FileSize')) { $row['FileSize'] = $File.Length }
            if (-not (Test-RowKey -Name 'FileUpdateDate')) { $row['FileUpdateDate'] = $File.LastWriteTime }
            if (-not (Test-RowKey -Name 'SourcePath')) { $row['SourcePath'] = $File.FullName }
            if (-not (Test-RowKey -Name 'ReadFolder')) { $row['ReadFolder'] = $ReadFolder }
            $global:list += [pscustomobject]$row
            return
        }

        $global:list += [pscustomobject]@{
            Title          = $File.Name
            FileName       = $File.Name
            FileSize       = $File.Length
            FileUpdateDate = $File.LastWriteTime
            SourceFolder   = $LocalFolder
            SourcePath     = $File.FullName
            ReadFolder     = $ReadFolder
        }
    }

    function Copy-LocalSourceFile {
        param(
            [System.IO.FileInfo]$File,
            $Row,
            [string]$ReadFolder
        )
        Write-DownloadProgress -Status 'Copying' -FileName $File.Name

        $destination = Join-Path $LocalFolder $File.Name
        $skipIfSame = -not ($Row.PSObject.Properties.Name -contains 'SkipIfSame') -or -not ($Row.SkipIfSame -eq 'No')
        $checkDateToo = $Row.PSObject.Properties.Name -contains 'CheckDateToo' -and (Test-YesLike -Value $Row.CheckDateToo)
        $copyNeeded = $true
        $reason = "file was not copied before"

        if (Test-Path $destination -PathType Leaf) {
            $existing = Get-Item $destination
            $sameSize = $existing.Length -eq $File.Length
            $sameDate = $existing.LastWriteTimeUtc -eq $File.LastWriteTimeUtc
            if (-not $skipIfSame) {
                $copyNeeded = $true
                $reason = "Force copy as requested"
            }
            elseif ($sameSize -and (-not $checkDateToo -or $sameDate)) {
                $copyNeeded = $false
            }
            else {
                $reason = "The read file and source file were different"
            }
        }

        if ($copyNeeded) {
            $copyStart = Get-Date
            Write-Host $File.Name -NoNewline -ForegroundColor Green
            Write-Host " Copying >>>> " $File.Length " ...bytes " -NoNewline -ForegroundColor Red
            Write-Host -ForegroundColor White
            Copy-Item -Path $File.FullName -Destination $destination -Force
            (Get-Item $destination).LastWriteTimeUtc = $File.LastWriteTimeUtc
            if (-not (Get-Variable -Name PWDownloadedCount -Scope Global -ErrorAction SilentlyContinue)) {
                $global:PWDownloadedCount = 0
            }
            $global:PWDownloadedCount++
            $copyEnd = Get-Date
            $copyDuration = New-TimeSpan -Start $copyStart -End $copyEnd
            Write-Host " >>>> copied - $reason"
            Write-Host " >>>> $($File.Name) - saved to $LocalFolder"
            Write-Host " >>>> Duration: $([math]::Round($copyDuration.TotalSeconds,2)) seconds"
        }
        else {
            Write-Host $File.Name -NoNewline -ForegroundColor Green
            if ($checkDateToo) {
                Write-Host " NOT copied >>>> Same date and size" -ForegroundColor Red
            }
            else {
                Write-Host " NOT copied >>>> Same size" -ForegroundColor Red
            }
        }

        Add-LocalFileToList -File $File -ReadFolder $ReadFolder
        $script:FEDAUTODownloadFileCurrent++
        $downloadStatus = if ($copyNeeded) { 'Copied' } else { 'Skipped' }
        Write-DownloadProgress -Status $downloadStatus -FileName $File.Name
    }

    function Ensure-ProjectWiseCommands {
        Ensure-ModuleAvailable -Name PWPS_DAB -InstallName pwps_dab
        $requiredPWCommands = @('New-PWLogin', 'Undo-PWLogin', 'Get-PWDocumentsBySearch', 'Get-PWDocumentsByGUIDs', 'CheckOut-PWDocuments')
        foreach ($cmd in $requiredPWCommands) {
            if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
                throw "Required ProjectWise command '$cmd' is unavailable. Ensure ProjectWise PowerShell tooling is properly registered on this machine."
            }
        }
    }

    $rowsMissingReadFolder = @($FS | Where-Object { [string]::IsNullOrWhiteSpace((Get-ReadFolderValue -Row $_)) })
    if ($rowsMissingReadFolder.Count -gt 0) {
        throw "Every active Download row must have a ReadFolder value."
    }

    # SourceAcquisitionMode is a user-declared filter.  Auto preserves the
    # existing per-row mixed-source behaviour; Local and ProjectWise exclude
    # incompatible active rows without altering the saved configuration.
    $sourceModeValue = $Config | Where-Object { $_.Parameter -eq 'SourceAcquisitionMode' } | Select-Object -ExpandProperty Value -First 1
    $sourceMode = if ($sourceModeValue) { $sourceModeValue.ToString().Trim().ToLowerInvariant() } else { 'auto' }
    if ($sourceMode -notin @('auto','local','projectwise')) { $sourceMode = 'auto' }
    if ($sourceMode -in @('local','projectwise')) {
        $originalRows = @($FS)
        $FS = @($FS | Where-Object {
            $isProjectWise = Test-ProjectWiseReadFolder -Path (Get-ReadFolderValue -Row $_)
            if ($sourceMode -eq 'local') { return -not $isProjectWise }
            return $isProjectWise
        })
        $excludedCount = $originalRows.Count - $FS.Count
        if ($excludedCount -gt 0) {
            Write-Warning ("SourceAcquisitionMode '{0}' excluded {1} incompatible active Download row(s)." -f $sourceMode, $excludedCount)
        }
        if ($FS.Count -eq 0) {
            throw "No active Download rows are compatible with SourceAcquisitionMode '$sourceMode'."
        }
    }

    $hasProjectWiseRows = $false
    foreach ($row in $FS) {
        if (Test-ProjectWiseReadFolder -Path (Get-ReadFolderValue -Row $row)) {
            $hasProjectWiseRows = $true
            break
        }
    }

    if ($hasProjectWiseRows) {
        Ensure-ProjectWiseCommands
    }

    $ProcessStartTime = Get-Date
    Write-host "Process start time"
    $ProcessStartTime

    $currentDatasource = $null
    $projectWiseLoggedIn = $false
    $hasCredentials = -not [string]::IsNullOrWhiteSpace($PWUser) -and -not [string]::IsNullOrWhiteSpace($PWPass)
    $script:FEDAUTODownloadFolderTotal = $FS.Count
    $script:FEDAUTODownloadFolderCurrent = 0
    $script:FEDAUTODownloadFileTotal = 0
    $script:FEDAUTODownloadFileCurrent = 0

    foreach($F in $FS) {
        $script:FEDAUTODownloadFolderCurrent++
        $script:FEDAUTODownloadFileTotal = 0
        $script:FEDAUTODownloadFileCurrent = 0
        $readFolder = Get-ReadFolderValue -Row $F
        $checkDateValue = $null
        if ($F.PSObject.Properties.Name -contains 'CheckDateToo') {
            $checkDateValue = $F.CheckDateToo
        }
        Write-Host "------"
        Write-Host "Read folder: $readFolder"
        Write-Host "File filter: $($F.FileFilter); SkipIfSame: $($F.SkipIfSame); CheckDateToo: $checkDateValue; Exclude: $($F.Exclude)"

        $folderInfo = Resolve-ProjectWiseReadFolderInfo -Path $readFolder
        if ($folderInfo) {
            $DatasourceName = $folderInfo.DataSource
            $pwFolderPath = $folderInfo.FolderPath
            if ($currentDatasource -ne $DatasourceName) {
                if ($projectWiseLoggedIn) {
                    Undo-PWLogin -ErrorAction SilentlyContinue
                    $projectWiseLoggedIn = $false
                }

                $loginMode = "Bentley IMS"
                $credentialLoginFailed = $false
                $credentialLoginError = $null
                if ($hasCredentials) {
                    try {
                        $loginMode = "PWUser/PWPass"
                        Write-Host "ProjectWise login: using PWUser/PWPass from Settings." -ForegroundColor Yellow
                        $securePassword = ConvertTo-SecureString -String $PWPass -AsPlainText -Force
                        New-PWLogin -DatasourceName $DatasourceName -UserName $PWUser -Password $securePassword -NonAdminLogin -ErrorAction Stop
                        $projectWiseLoggedIn = $true
                    }
                    catch {
                        $credentialLoginFailed = $true
                        $credentialLoginError = $_
                        Write-Warning "ProjectWise login with PWUser/PWPass failed. Attempting Bentley IMS fallback. Error: $_"
                    }
                }

                if (-not $projectWiseLoggedIn) {
                    try {
                        $loginMode = if ($credentialLoginFailed) { "Bentley IMS (fallback)" } else { "Bentley IMS" }
                        Write-Host "ProjectWise login: using Bentley IMS." -ForegroundColor Yellow
                        New-PWLogin -DatasourceName $DatasourceName -BentleyIMS -NonAdminLogin -ErrorAction Stop
                        $projectWiseLoggedIn = $true
                    }
                    catch {
                        if ($credentialLoginFailed) {
                            $message = "ProjectWise login failed with PWUser/PWPass and Bentley IMS fallback. PWUser/PWPass error: $credentialLoginError ; IMS error: $_"
                            Write-Host "ERROR: $message" -ForegroundColor Red
                            throw $message
                        }
                        Write-Host "ERROR: ProjectWise login failed: $_" -ForegroundColor Red
                        throw
                    }
                }

                $sessionEstablished = $true
                try {
                    if (-not [string]::IsNullOrWhiteSpace($pwFolderPath)) {
                        $null = Get-PWDocumentsBySearch -FolderPath $pwFolderPath -FileName '%' -JustThisFolder -ErrorAction Stop |
                            Select-Object -First 1
                    }
                }
                catch {
                    if (Test-PWNotLoggedError -ErrorRecord $_) {
                        $sessionEstablished = $false
                    }
                    else {
                        Write-Warning "Login probe could not validate folder access for '$pwFolderPath'. Continuing because session appears logged in. Error: $_"
                    }
                }

                if (-not $sessionEstablished) {
                    $msg = "ProjectWise login did not create an active session for datasource '$DatasourceName' (method: $loginMode). Stopping download."
                    Write-Host "ERROR: $msg" -ForegroundColor Red
                    throw $msg
                }
                $currentDatasource = $DatasourceName
            }

            Write-Host ("Datasource: pw:\\{0}" -f $DatasourceName)
            Write-Host "ProjectWise folder: $pwFolderPath"
            Write-Host "Starting the work on " $pwFolderPath
            $searchAttemptCount++
            try {
                $PWdocs = @()
                $seenDocumentKeys = @{}
                foreach ($fileFilterPattern in (Get-FileFilterPatterns -Value $F.FileFilter)) {
                    $fileFilterSql = Convert-WindowsWildcardToSqlLike -Pattern $fileFilterPattern
                    $matchedDocs = @(Get-PWDocumentsBySearch -FolderPath $pwFolderPath -FileName $fileFilterSql -JustThisFolder -ErrorAction Stop)
                    foreach ($doc in $matchedDocs) {
                        $documentKey = if ($doc.DocumentGUID) { $doc.DocumentGUID.ToString() } elseif ($doc.FileName) { $doc.FileName.ToString().ToLowerInvariant() } else { $doc.Name.ToString().ToLowerInvariant() }
                        if ([string]::IsNullOrWhiteSpace($documentKey) -or $seenDocumentKeys.ContainsKey($documentKey)) { continue }
                        $seenDocumentKeys[$documentKey] = $true
                        $PWdocs += $doc
                    }
                }
            }
            catch {
                if (Test-PWNotLoggedError -ErrorRecord $_) {
                    $msg = "ProjectWise session dropped or was not established for datasource '$DatasourceName'. Stopping download."
                    Write-Host "ERROR: $msg" -ForegroundColor Red
                    throw $msg
                }
                $searchFailedCount++
                if ($searchFailureDetails.Count -lt 5) {
                    $searchFailureDetails.Add(("{0} -> {1}" -f $readFolder, $_.Exception.Message)) | Out-Null
                }
                Write-Warning "Search failed for folder ${readFolder}: $_"
                continue
            }
            $ExclusionList = Get-ExcludeTerms -Value $F.Exclude
            foreach ($ExclusionItem in $ExclusionList) {
                $PWdocs = $PWdocs | Where-Object { $_.Name -notlike "*$ExclusionItem*" }
            }
            $PWdocs = @($PWdocs)
            $j = $PWdocs.Count
            Write-Host "Documents found after exclusion: $j"
            $script:FEDAUTODownloadFileTotal = $j
            $script:FEDAUTODownloadFileCurrent = 0
            Write-DownloadProgress -Status 'Queued' -FileName $readFolder
            For ($i = 0 ; $i -lt $j ; $i++)
            {
                try {
                    $TempD = Get-PWDocumentsByGUIDs -DocumentGUID $PWdocs[$i].DocumentGUID -ErrorAction Stop
                }
                catch {
                    Write-Warning "Failed to get document by GUID for index $i in folder ${readFolder}: $_"
                    $script:FEDAUTODownloadFileCurrent++
                    Write-DownloadProgress -Status 'Failed' -FileName ("Document index {0}" -f ($i + 1))
                    continue
                }
                $displayNameForProgress = if ($TempD.FileName) { $TempD.FileName } else { $TempD.Name }
                Write-DownloadProgress -Status 'Downloading' -FileName $displayNameForProgress
                $FilePath = Join-Path $LocalFolder $TempD.FileName
                if (Test-Path $FilePath)
                    {
                    $LocalFile = Get-ChildItem -Path $LocalFolder $TempD.FileName
                    if ($F.SkipIfSame -eq "No")
                        {
                        DU("Force download as requested")
                        }
                    else
                        {
                        if(!($TempD.FileSize -eq $LocalFile.Length) -or `
                           (!($TempD.FileUpdateDate.ToString("yyyy-MM-dd HH:mm tt") -ge $LocalFile.LastWriteTime.ToString("yyyy-MM-dd HH:mm tt")) `
                            -and ($checkDateValue -eq "Yes")))
                            {
                            DU("The PW and the local files were different")
                            }
                        else
                            {
                            $displayName = if ($TempD.FileName) { $TempD.FileName } else { $TempD.Name }
                            Write-host $displayName -NoNewline -ForegroundColor Green
                            if  ($checkDateValue -eq "Yes") {
                                Write-host " NOT downloaded >>>> Same date and size" -ForegroundColor Red
                            }
                            else {
                                Write-host " NOT downloaded >>>> Same size" -ForegroundColor Red
                            }
                        }
                        }
                    }
                Else
                    {
                    DU("file on PW was not downloaded before")
                    }
                Add2List($TempD)
                $script:FEDAUTODownloadFileCurrent++
                Write-DownloadProgress -Status 'Actioned' -FileName $displayNameForProgress
            }
            continue
        }

        $resolvedReadFolder = Resolve-RelativePath -PathValue (CureFolderPath $readFolder) -Root $basePath
        if (-not (Test-Path $resolvedReadFolder -PathType Container)) {
            $searchAttemptCount++
            $searchFailedCount++
            if ($searchFailureDetails.Count -lt 5) {
                $searchFailureDetails.Add(("{0} -> Local folder not found" -f $readFolder)) | Out-Null
            }
            Write-Warning "Local read folder not found: '$resolvedReadFolder'."
            continue
        }

        Write-Host "Local/synced folder: $resolvedReadFolder"
        $searchAttemptCount++
        $localFiles = @()
        $seenLocalPaths = @{}
        foreach ($fileFilter in (Get-FileFilterPatterns -Value $F.FileFilter)) {
            foreach ($file in @(Get-ChildItem -Path $resolvedReadFolder -File -Filter $fileFilter -ErrorAction Stop)) {
                $pathKey = $file.FullName.ToLowerInvariant()
                if ($seenLocalPaths.ContainsKey($pathKey)) { continue }
                $seenLocalPaths[$pathKey] = $true
                $localFiles += $file
            }
        }
        $ExclusionList = Get-ExcludeTerms -Value $F.Exclude
        foreach ($ExclusionItem in $ExclusionList) {
            $localFiles = $localFiles | Where-Object { $_.Name -notlike "*$ExclusionItem*" }
        }
        $localFiles = @($localFiles)
        Write-Host "Local files found after exclusion: $($localFiles.Count)"
        $script:FEDAUTODownloadFileTotal = $localFiles.Count
        $script:FEDAUTODownloadFileCurrent = 0
        Write-DownloadProgress -Status 'Queued' -FileName $resolvedReadFolder
        foreach ($file in $localFiles) {
            Copy-LocalSourceFile -File $file -Row $F -ReadFolder $resolvedReadFolder
        }
    }

    if ($projectWiseLoggedIn) {
        Undo-PWLogin -ErrorAction SilentlyContinue
        $projectWiseLoggedIn = $false
    }
    #$Now = Get-Date -format "yyyyMMdd-hhmmss"
    #$LogFile = $LogsFolder+'PWDownloadLog' + $Now + '.csv' # File name with date and time
    $latestNames = @()
    foreach ($row in $list) {
        if (-not $row) { continue }
        $name = $row.FileName
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $latestNames += [System.IO.Path]::GetFileName($name)
    }
    $latestNames = $latestNames | Where-Object { $_ } | Sort-Object -Unique

    $allSearchesFailed = ($searchAttemptCount -gt 0 -and $searchFailedCount -eq $searchAttemptCount)
    $partialSearchFailure = ($searchAttemptCount -gt 0 -and $searchFailedCount -gt 0 -and $searchFailedCount -lt $searchAttemptCount)

    if ($allSearchesFailed) {
        $global:PWDownloadStatus = 'AllSearchFailed'
        $sampleErrors = if ($searchFailureDetails.Count -gt 0) { ($searchFailureDetails -join " | ") } else { "No detailed errors captured." }
        Write-Error ("All read-folder searches failed ({0}/{1}). Check ReadFolder path format, datasource access, and permissions. Sample errors: {2}" -f $searchFailedCount, $searchAttemptCount, $sampleErrors)
    }
    elseif ($partialSearchFailure) {
        $global:PWDownloadStatus = 'PartialSearchFailed'
        $sampleErrors = if ($searchFailureDetails.Count -gt 0) { ($searchFailureDetails -join " | ") } else { "No detailed errors captured." }
        Write-Error ("Partial read-folder search failure ({0}/{1}). Cause: one or more configured folders failed. Sample errors: {2} Response: skipped deleted-file cleanup and orphan NWC cleanup to prevent false local deletions." -f $searchFailedCount, $searchAttemptCount, $sampleErrors)
    }

    if ($list.Count -eq 0) {
        if ($global:PWDownloadStatus -ne 'AllSearchFailed') {
            $global:PWDownloadStatus = 'Empty'
            Write-Error "Download returned zero files. Treating as skipped download; no source-side files were updated."
        }
    }
    else {
        if (-not $partialSearchFailure) {
            $global:PWDownloadStatus = 'Success'
        }
        $list | Export-Csv -LiteralPath $PWAttributes -NoTypeInformation -Encoding UTF8

        if ($partialSearchFailure) {
            Write-Host "Safety response active: skipping deleted-file and orphan NWC cleanup because some read-folder searches failed." -ForegroundColor Yellow
        }

        if ($latestNames.Count -gt 0 -and (Test-Path $LocalFolder -PathType Container) -and -not $partialSearchFailure) {
            $deletedRoot = Join-Path (Split-Path -Parent $LocalFolder) 'deleted'
            if (-not (Test-Path $deletedRoot)) { New-Item -Path $deletedRoot -ItemType Directory -Force | Out-Null }

            $deletedLogRoot = if ($LogsFolder) { $LogsFolder } else { Split-Path -Parent $LocalFolder }
            $deletedLogPath = Join-Path $deletedLogRoot 'Deleted_files.csv'
            $deletedRows = @()
            $previousLookup = @{}
            if ($previousAttributeRows) {
                foreach ($row in $previousAttributeRows) {
                    $fileName = $null
                    if ($row.PSObject.Properties.Name -contains 'FileName') { $fileName = $row.FileName }
                    elseif ($row.PSObject.Properties.Name -contains 'Filename') { $fileName = $row.Filename }
                    if ([string]::IsNullOrWhiteSpace($fileName)) { continue }
                    $previousLookup[[System.IO.Path]::GetFileName($fileName)] = $row
                }
            }

            $federatableExtensions = if (Get-Command Get-FederatableModelExtensions -ErrorAction SilentlyContinue) {
                Get-FederatableModelExtensions
            }
            else {
                @('.ifc', '.dwg', '.dgn', '.rvt')
            }
            $localFiles = @(Get-ChildItem -Path $LocalFolder -File | Where-Object {
                $federatableExtensions -contains $_.Extension.ToLowerInvariant()
            })
            foreach ($file in $localFiles) {
                if ($latestNames -contains $file.Name) { continue }
                if ($file.Extension.ToLowerInvariant() -eq '.nwc' -and
                    (Test-GeneratedNwcForSourceNames -NwcName $file.Name -SourceNames $latestNames)) {
                    Write-Host "Keeping generated NWC cache: $($file.Name)"
                    continue
                }

                $deletedPath = Join-Path $deletedRoot $file.Name
                try {
                    Move-Item -Path $file.FullName -Destination $deletedPath -Force
                }
                catch {
                    Write-Warning "Failed to relocate deleted file '$($file.FullName)': $_"
                }

                $processedPath = $null
                if ($ProcessedFolder -and (Test-Path $ProcessedFolder -PathType Container)) {
                    $candidate = Join-Path $ProcessedFolder $file.Name
                    if (Test-Path $candidate -PathType Leaf) {
                        try {
                            Remove-Item -Path $candidate -Force
                        }
                        catch {
                            Write-Warning "Failed to delete processed file '$candidate': $_"
                        }
                        $processedPath = $candidate
                    }
                    $nwcCandidate = Join-Path $ProcessedFolder ([System.IO.Path]::GetFileNameWithoutExtension($file.Name) + '.nwc')
                    if (Test-Path $nwcCandidate -PathType Leaf) {
                        try {
                            Remove-Item -Path $nwcCandidate -Force
                        }
                        catch {
                            Write-Warning "Failed to delete processed NWC '$nwcCandidate': $_"
                        }
                    }
                }

                $prevRow = $null
                $pwGuid = $null
                $pwUrn = $null
                $readFolderValue = $null
                if ($previousLookup.ContainsKey($file.Name)) {
                    $prevRow = $previousLookup[$file.Name]
                    if ($prevRow.PSObject.Properties.Name -contains 'PW_Doc_GUID') { $pwGuid = $prevRow.PW_Doc_GUID }
                    if ($prevRow.PSObject.Properties.Name -contains 'PW_Doc_URN') { $pwUrn = $prevRow.PW_Doc_URN }
                    if ($prevRow.PSObject.Properties.Name -contains 'ReadFolder') { $readFolderValue = $prevRow.ReadFolder }
                }

                $deletedRows += [pscustomobject]@{
                    DeletedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm")
                    FileName    = $file.Name
                    SourcePath  = $file.FullName
                    DeletedPath = $deletedPath
                    ProcessedPath = $processedPath
                    PW_Doc_GUID = $pwGuid
                    PW_Doc_URN  = $pwUrn
                    ReadFolder  = $readFolderValue
                }
            }

            if ($deletedRows.Count -gt 0) {
                $deletedCount = $deletedRows.Count
                $existingRows = @()
                if (Test-Path $deletedLogPath -PathType Leaf) {
                    try {
                        $existingRows = Get-ExcelDataSafe -Path $deletedLogPath -NamedRange 'Deleted'
                    }
                    catch {
                        Write-Warning "Unable to read existing deleted log '$deletedLogPath'. Overwriting."
                        $existingRows = @()
                    }
                }
                $allRows = @()
                if ($existingRows) { $allRows += $existingRows }
                $allRows += $deletedRows
                $allRows | Export-Csv -LiteralPath $deletedLogPath -NoTypeInformation -Encoding UTF8
            }
        }

        if ($ProcessedFolder -and (Test-Path $ProcessedFolder -PathType Container) -and $latestNames.Count -gt 0 -and -not $partialSearchFailure) {
            $federatableExtensions = if (Get-Command Get-FederatableModelExtensions -ErrorAction SilentlyContinue) {
                Get-FederatableModelExtensions
            }
            else {
                @('.ifc', '.dwg', '.dgn', '.rvt')
            }
            $processedSourceNames = @(Get-ChildItem -Path $ProcessedFolder -File | Where-Object {
                $_.Extension.ToLowerInvariant() -in @('.ifc', '.dwg', '.dgn', '.rvt')
            } | Select-Object -ExpandProperty Name)
            $nwcFiles = Get-ChildItem -Path $ProcessedFolder -Filter "*.nwc" -File
            foreach ($nwc in $nwcFiles) {
                $hasSource = Test-GeneratedNwcForSourceNames -NwcName $nwc.Name -SourceNames $latestNames
                $hasProcessedSource = Test-GeneratedNwcForSourceNames -NwcName $nwc.Name -SourceNames $processedSourceNames
                if (-not $hasSource -and -not $hasProcessedSource) {
                    try {
                        Remove-Item -Path $nwc.FullName -Force
                    }
                    catch {
                        Write-Warning "Failed to delete orphan NWC '$($nwc.FullName)': $_"
                    }
                }
            }
        }
    }
    $global:PWDeletedCount = $deletedCount
    $ProcessEndTime = Get-Date
    Write-Host 'Process finish time'
    $ProcessEndTime
    $TotalDuration = New-TimeSpan -Start $ProcessStartTime -End $ProcessEndTime
    Write-Host "Total duration: $([math]::Round($TotalDuration.TotalMinutes,2)) minutes"
    Write-Host "=== END 020-PWDownload ==="
    if ($TranscriptStarted) {
        Stop-Transcript | Out-Null
        $global:MainTranscriptActive = $false
    }
}



