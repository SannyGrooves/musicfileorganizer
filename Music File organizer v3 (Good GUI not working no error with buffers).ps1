# PowerShell script to move music files with improved stability and responsiveness for large file loads
# Requires .NET Framework 4.5 or later

# Initialize debug log file
$debugLogPath = Join-Path $env:TEMP "MoveMusicFiles_Debug.log"
function Write-DebugLog {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $debugLogPath -Append -Encoding UTF8
}

# Log startup diagnostics
Write-DebugLog "Script starting..."
Write-DebugLog "PowerShell Version: $PSVersionTable.PSVersion"
Write-DebugLog "NET Framework: $([System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)"
Write-DebugLog "Script Path: $PSCommandPath"
Write-DebugLog "Current User: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"

# Wrap entire script in try/catch for unhandled exceptions
try {
    # Dynamically load required assemblies
    Write-DebugLog "Loading assemblies..."
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    } catch {
        $errorMessage = "Error: Failed to load required .NET assemblies. Ensure .NET Framework 4.5 or later is installed.`nDetails: $_"
        Write-DebugLog $errorMessage
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) 2>$null
        Write-Error $errorMessage
        exit 1
    }

    # Check execution policy
    Write-DebugLog "Checking execution policy..."
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
    Write-DebugLog "Current Execution Policy: $currentPolicy"
    if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "AllSigned") {
        try {
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force -ErrorAction Stop
            Write-DebugLog "Set execution policy to RemoteSigned"
        } catch {
            $errorMessage = "Error setting execution policy: $_`nPlease run PowerShell as Administrator and set the execution policy to RemoteSigned."
            Write-DebugLog $errorMessage
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) 2>$null
            Write-Error $errorMessage
            exit 1
        }
    }

    # Request admin privileges if not already elevated
    Write-DebugLog "Checking for admin privileges..."
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-DebugLog "Not running as admin, attempting to elevate..."
        try {
            Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -ErrorAction Stop
            Write-DebugLog "Elevation successful, exiting current instance"
            exit
        } catch {
            $errorMessage = "Failed to elevate to admin: $_"
            Write-DebugLog $errorMessage
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) 2>$null
            Write-Error $errorMessage
            exit 1
        }
    }
    Write-DebugLog "Running as admin"

    # Function to show conflict resolution dialog
    function Show-ConflictDialog {
        param (
            [string]$Path,
            [string]$Type
        )
        $dialog = New-Object System.Windows.Forms.Form
        try {
            $dialog.Text = "$Type Conflict"
            $dialog.Size = New-Object System.Drawing.Size(400, 200)
            $dialog.StartPosition = "CenterScreen"
            $dialog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
            $dialog.ForeColor = [System.Drawing.Color]::White
            $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $dialog.MaximizeBox = $false
            $dialog.MinimizeBox = $false
            $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(20, 20)
            $label.Size = New-Object System.Drawing.Size(360, 60)
            $label.Text = "The $Type '$Path' already exists.`r`nPlease choose an action:"
            $label.ForeColor = [System.Drawing.Color]::White
            $dialog.Controls.Add($label)

            $overwriteButton = New-Object System.Windows.Forms.Button
            $overwriteButton.Location = New-Object System.Drawing.Point(20, 100)
            $overwriteButton.Size = New-Object System.Drawing.Size(100, 30)
            $overwriteButton.Text = "Overwrite"
            $overwriteButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
            $overwriteButton.ForeColor = [System.Drawing.Color]::White
            $overwriteButton.Add_Click({ $dialog.Tag = "Overwrite"; $dialog.Close() })
            $dialog.Controls.Add($overwriteButton)

            $keepBothButton = New-Object System.Windows.Forms.Button
            $keepBothButton.Location = New-Object System.Drawing.Point(140, 100)
            $keepBothButton.Size = New-Object System.Drawing.Size(100, 30)
            $keepBothButton.Text = "Keep Both"
            $keepBothButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
            $keepBothButton.ForeColor = [System.Drawing.Color]::White
            $keepBothButton.Add_Click({ $dialog.Tag = "KeepBoth"; $dialog.Close() })
            $dialog.Controls.Add($keepBothButton)

            $skipButton = New-Object System.Windows.Forms.Button
            $skipButton.Location = New-Object System.Drawing.Point(260, 100)
            $skipButton.Size = New-Object System.Drawing.Size(100, 30)
            $skipButton.Text = "Skip"
            $skipButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
            $skipButton.ForeColor = [System.Drawing.Color]::White
            $skipButton.Add_Click({ $dialog.Tag = "Skip"; $dialog.Close() })
            $dialog.Controls.Add($skipButton)

            $dialog.ShowDialog() | Out-Null
            return $dialog.Tag
        } finally {
            $dialog.Dispose()
        }
    }

    # Function to sanitize filenames
    function Sanitize-Filename {
        param ([string]$Filename)
        $sanitized = $Filename -replace '\u2019', "'"
        $sanitized = $sanitized -replace '\s+', ' '
        $sanitized = $sanitized -replace '\.+$', ''
        $sanitized = $sanitized -replace '[<>:"/\\|?*\[\]\(\)]', ''
        $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() -join ''
        $sanitized = $sanitized -replace "[$invalidChars]", ''
        $sanitized = $sanitized.Trim().Trim('.')
        if ($sanitized.Length -gt 200) {
            $sanitized = $sanitized.Substring(0, 200).Trim()
        }
        if ([string]::IsNullOrEmpty($sanitized)) {
            return "unnamed"
        }
        return $sanitized
    }

    # Function to get unique destination path for files
    function Get-UniquePath {
        param ([string]$Path)
        $basePath = [System.IO.Path]::GetDirectoryName($Path)
        $filename = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $ext = [System.IO.Path]::GetExtension($Path)
        $counter = 0
        $newPath = $Path
        while (Test-Path -LiteralPath $newPath) {
            $counter++
            $newPath = Join-Path $basePath "$filename_$counter$ext"
        }
        if ($newPath.Length -gt 260) {
            $shortenedFilename = $filename.Substring(0, [Math]::Min($filename.Length, 200 - $basePath.Length - $ext.Length - 10))
            $newPath = Join-Path $basePath "$shortenedFilename$ext"
            $newPath = Get-UniquePath -Path $newPath
        }
        return $newPath
    }

    # Function to get unique destination folder path
    function Get-UniqueFolderPath {
        param ([string]$Path)
        $basePath = [System.IO.Path]::GetDirectoryName($Path)
        $folderName = [System.IO.Path]::GetFileName($Path)
        $counter = 0
        $newPath = $Path
        while (Test-Path -LiteralPath $newPath) {
            $counter++
            $newPath = Join-Path $basePath "$folderName_$counter"
        }
        if ($newPath.Length -gt 260) {
            $shortenedFolderName = $folderName.Substring(0, [Math]::Min($folderName.Length, 200 - $basePath.Length - 10))
            $newPath = Join-Path $basePath "$shortenedFolderName"
            $newPath = Get-UniqueFolderPath -Path $newPath
        }
        return $newPath
    }

    # Function to test if file is accessible
    function Test-FileAccessible {
        param ([string]$Path)
        try {
            $fileStream = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
            $fileStream.Close()
            $fileStream.Dispose()
            return $true
        } catch {
            return $false
        }
    }

    # Function to copy large files with buffering
    function Copy-LargeFile {
        param (
            [string]$SourcePath,
            [string]$DestPath
        )
        try {
            $bufferSize = 300MB
            $buffer = New-Object byte[] $bufferSize
            $sourceStream = [System.IO.File]::OpenRead($SourcePath)
            $destStream = [System.IO.File]::Create($DestPath)
            $bytesRead = 0
            while (($bytesRead = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $destStream.Write($buffer, 0, $bytesRead)
                if ($script:cancelOperation) {
                    throw "Operation cancelled by user."
                }
            }
            $sourceStream.Close()
            $destStream.Close()
            $sourceStream.Dispose()
            $destStream.Dispose()
            return $true
        } catch {
            Write-Log -Message "Error copying file ${SourcePath} to ${DestPath}: $_" -Status "Error"
            return $false
        }
    }

    # Function to create backup of files
    function Create-Backup {
        param (
            [string]$Source,
            [string]$BackupDir
        )
        try {
            if (-not (Test-Path -LiteralPath $BackupDir)) {
                New-Item -ItemType Directory -Path $BackupDir -ErrorAction Stop | Out-Null
            }
            $patterns = @("*.mp3", "*.wav", "*.flac", "*.aiff")
            $fileCount = 0
            foreach ($pattern in $patterns) {
                $fileCount += ([System.IO.Directory]::EnumerateFiles($Source, $pattern, [System.IO.SearchOption]::AllDirectories) | Measure-Object).Count
            }
            $currentFile = 0
            $batchSize = 1000
            $fileBatch = @()
            foreach ($pattern in $patterns) {
                $files = [System.IO.Directory]::EnumerateFiles($Source, $pattern, [System.IO.SearchOption]::AllDirectories)
                foreach ($file in $files) {
                    if ($script:cancelOperation) {
                        Write-Log -Message "Backup operation cancelled by user." -Status "Info"
                        return
                    }
                    $fileBatch += $file
                    if ($fileBatch.Count -ge $batchSize -or $file -eq $files[-1]) {
                        foreach ($batchFile in $fileBatch) {
                            $currentFile++
                            $progressBar.Value = if ($fileCount -gt 0) { [Math]::Min(1000, ($currentFile / $fileCount) * 1000) } else { 0 }
                            $relativePath = $batchFile.Substring($Source.Length).TrimStart('\')
                            $backupPath = Join-Path $BackupDir $relativePath
                            $backupFolder = [System.IO.Path]::GetDirectoryName($backupPath)
                            if (-not (Test-Path -LiteralPath $backupFolder)) {
                                New-Item -ItemType Directory -Path $backupFolder -ErrorAction Stop | Out-Null
                            }
                            $fileName = [System.IO.Path]::GetFileName($batchFile)
                            $fileExt = [System.IO.Path]::GetExtension($batchFile).TrimStart('.')
                            if (Copy-LargeFile -SourcePath $batchFile -DestPath $backupPath) {
                                Write-Log -Message "Backed up: $batchFile to $backupPath" -Status "Success" -Action "BackedUp" -OldName $fileName -NewName $fileName -FileType $fileExt -OpenFileInFolder $batchFile -BackupLocation $backupPath
                            }
                        }
                        $fileBatch = @()
                        [System.GC]::Collect()
                    }
                }
            }
            Write-Log -Message "Backup completed to: $BackupDir" -Status "Success"
        } catch {
            Write-Log -Message "Error creating backup: $_" -Status "Error"
        } finally {
            $progressBar.Value = 0
            [System.GC]::Collect()
        }
    }

    # Function to log messages and write to temp file
    function Write-Log {
        param (
            [string]$Message,
            [string]$Status = "Info",
            [string]$Action = "",
            [string]$WriteNewFolder = "",
            [string]$FileType = "",
            [string]$OldName = "",
            [string]$NewName = "",
            [string]$OpenFileInFolder = "",
            [string]$BackupLocation = ""
        )
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = @{
            Timestamp = $timestamp
            Message   = $Message
            Status    = $Status
            Action    = $Action
            WriteNewFolder = $WriteNewFolder
            FileType  = $FileType
            OldName   = $OldName
            NewName   = $NewName
            OpenFileInFolder = $OpenFileInFolder
            BackupLocation = $BackupLocation
        }
        $script:logEntries += $logEntry
        Write-DebugLog "Log Entry: [$timestamp] $Message ($Status)"
        if ($script:logEntries.Count -ge 1000) {
            Write-TempLogEntries
        }
        if ($statusLabel -and $statusLabel.IsHandleCreated) {
            try {
                $statusLabel.BeginInvoke([Action]{
                    $statusColor = switch ($Status) {
                        "Success" { "Green" }
                        "Error"   { "Red" }
                        "Info"    { "White" }
                        default   { "Yellow" }
                    }
                    $statusLabel.Text = "[$timestamp] $Message"
                    $statusLabel.ForeColor = [System.Drawing.Color]::$statusColor
                }) | Out-Null
            } catch {
                Write-DebugLog "Failed to update statusLabel: $_"
                Write-Output "[$timestamp] $Message ($Status)"
            }
        } else {
            Write-DebugLog "statusLabel not available, writing to console"
            Write-Output "[$timestamp] $Message ($Status)"
        }
    }

    # Function to write log entries to temp file
    function Write-TempLogEntries {
        try {
            $tempLogPath = $script:tempLogPath
            $jsonEntries = $script:logEntries | ConvertTo-Json -Compress
            [System.IO.File]::AppendAllText($tempLogPath, $jsonEntries + "`n", [System.Text.Encoding]::UTF8)
            $script:logEntries = @()
            [System.GC]::Collect()
        } catch {
            Write-DebugLog "Error writing temp log entries: $_"
            Write-Output "Error writing temp log entries: $_"
        }
    }

    # Function to read temp log entries
    function Read-TempLogEntries {
        $entries = @()
        try {
            $tempLogPath = $script:tempLogPath
            if (Test-Path -LiteralPath $tempLogPath) {
                $lines = [System.IO.File]::ReadAllLines($tempLogPath, [System.Text.Encoding]::UTF8)
                foreach ($line in $lines) {
                    if ($line.Trim()) {
                        $entries += $line | ConvertFrom-Json
                    }
                }
            }
        } catch {
            Write-DebugLog "Error reading temp log entries: $_"
            Write-Output "Error reading temp log entries: $_"
        }
        return $entries
    }

    # Function to generate enhanced HTML log file
    function Generate-HtmlLog {
        param ([string]$LogFilePath)
        try {
            Write-TempLogEntries
            $allEntries = Read-TempLogEntries
            $successCount = ($allEntries | Where-Object { $_.Status -eq "Success" }).Count
            $errorCount = ($allEntries | Where-Object { $_.Status -eq "Error" }).Count
            $infoCount = ($allEntries | Where-Object { $_.Status -eq "Info" }).Count
            $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Music File Organizer Log</title>
    <style>
        body {
            background: #222222;
            color: #f0f0f0;
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 40px;
            line-height: 1.8;
            font-size: 16px;
        }
        h1 {
            text-align: center;
            color: #ffffff;
            font-size: 2.5em;
            margin-bottom: 20px;
        }
        .summary {
            background: #333333;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            text-align: center;
        }
        .summary span {
            font-size: 1.3em;
            margin: 0 20px;
            cursor: pointer;
        }
        .summary span:hover {
            color: #17a2b8;
        }
        .filter {
            margin-bottom: 30px;
            text-align: center;
            display: flex;
            justify-content: center;
            gap: 15px;
            flex-wrap: wrap;
        }
        .filter label {
            font-size: 1.1em;
            margin-right: 10px;
        }
        .filter select, .filter button {
            padding: 10px;
            font-size: 1em;
            background: #444444;
            color: #ffffff;
            border: none;
            border-radius: 5px;
            cursor: pointer;
        }
        .filter select:hover, .filter button:hover {
            background: #555555;
        }
        .file-section {
            background: #2a2a2a;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        .file-section summary {
            cursor: pointer;
            font-size: 1.3em;
            font-weight: bold;
            color: #ffffff;
        }
        .file-details {
            display: grid;
            grid-template-columns: 1fr 3fr;
            gap: 15px;
            margin-top: 15px;
        }
        .file-details div {
            padding: 8px;
        }
        .file-details .label {
            font-weight: bold;
            color: #cccccc;
        }
        .file-details .value {
            word-break: break-all;
        }
        .file-details a {
            color: #17a2b8;
            text-decoration: none;
        }
        .file-details a:hover {
            text-decoration: underline;
        }
        .status-success::before { content: '✅ '; }
        .status-error::before { content: '❌ '; }
        .status-info::before { content: 'ℹ️ '; }
        .status-success { color: #28a745; }
        .status-error { color: #dc3545; }
        .status-info { color: #17a2b8; }
        footer {
            text-align: center;
            margin-top: 40px;
            color: #aaaaaa;
            font-size: 0.9em;
        }
        #back-to-top {
            position: fixed;
            bottom: 20px;
            right: 20px;
            padding: 10px 20px;
            background: #17a2b8;
            color: #ffffff;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            display: none;
        }
        #back-to-top:hover {
            background: #2cd4ef;
        }
    </style>
    <script>
        function filterLogs(status = 'all', folder = 'all', fileType = 'all') {
            const sections = document.getElementsByClassName('file-section');
            for (let section of sections) {
                const sectionStatus = section.getAttribute('data-status').toLowerCase();
                const sectionFolder = section.getAttribute('data-folder').toLowerCase();
                const sectionFileType = section.getAttribute('data-filetype').toLowerCase();
                const statusMatch = (status === 'all' || status === sectionStatus);
                const folderMatch = (folder === 'all' || folder === sectionFolder);
                const fileTypeMatch = (fileType === 'all' || fileType === sectionFileType);
                section.style.display = (statusMatch && folderMatch && fileTypeMatch) ? 'block' : 'none';
            }
        }
        function resetFilters() {
            document.getElementById('statusFilter').value = 'all';
            document.getElementById('folderFilter').value = 'all';
            document.getElementById('fileTypeFilter').value = 'all';
            filterLogs();
        }
        window.addEventListener('scroll', () => {
            const backToTop = document.getElementById('back-to-top');
            backToTop.style.display = window.scrollY > 200 ? 'block' : 'none';
        });
        function applySummaryFilter(status) {
            document.getElementById('statusFilter').value = status;
            filterLogs(status, document.getElementById('folderFilter').value, document.getElementById('fileTypeFilter').value);
        }
    </script>
</head>
<body>
    <h1>Music File Organizer Log</h1>
    <div class="summary">
        <span onclick="applySummaryFilter('success')">Successful Operations: $successCount</span>
        <span onclick="applySummaryFilter('error')">Failed Operations: $errorCount</span>
        <span onclick="applySummaryFilter('info')">Info Operations: $infoCount</span>
    </div>
    <div class="filter">
        <label for="statusFilter">Filter by Status:</label>
        <select id="statusFilter" onchange="filterLogs(this.value, document.getElementById('folderFilter').value, document.getElementById('fileTypeFilter').value)">
            <option value="all">All Statuses</option>
            <option value="success">Successful Operations</option>
            <option value="error">Failed Operations</option>
            <option value="info">Info Operations</option>
        </select>
        <label for="folderFilter">Filter by Folder:</label>
        <select id="folderFilter" onchange="filterLogs(document.getElementById('statusFilter').value, this.value, document.getElementById('fileTypeFilter').value)">
            <option value="all">All Folders</option>
            <option value="yes">Changed Folders</option>
            <option value="no">Unchanged Folders</option>
        </select>
        <label for="fileTypeFilter">Filter by File Type:</label>
        <select id="fileTypeFilter" onchange="filterLogs(document.getElementById('statusFilter').value, document.getElementById('folderFilter').value, this.value)">
            <option value="all">All File Types</option>
            <option value="mp3">MP3</option>
            <option value="wav">WAV</option>
            <option value="flac">FLAC</option>
            <option value="aiff">AIFF</option>
        </select>
        <button onclick="resetFilters()">Reset Filters</button>
    </div>
"@

            $displayNames = @{
                "Action" = "File Action"
                "WriteNewFolder" = "Folder Changed"
                "FileType" = "File Type"
                "OldName" = "Original Name"
                "NewName" = "New Name"
                "OpenFileInFolder" = "File Location"
                "BackupLocation" = "Backup Location"
            }

            $fileGroups = $allEntries | Where-Object { $_.OldName } | Group-Object -Property OldName
            foreach ($group in $fileGroups) {
                $fileEntries = $group.Group
                $primaryEntry = $fileEntries | Where-Object { $_.Action -in "Renamed", "Moved", "BackedUp" } | Select-Object -First 1
                if (-not $primaryEntry) { continue }
                $status = $primaryEntry.Status.ToLower()
                $folderStatus = $primaryEntry.WriteNewFolder.ToLower()
                $fileType = $primaryEntry.FileType.ToLower()
                $htmlContent += "<details class='file-section' data-status='$status' data-folder='$folderStatus' data-filetype='$fileType'>`n"
                $htmlContent += "<summary>File: $([System.Net.WebUtility]::HtmlEncode($primaryEntry.OldName))</summary>`n"
                $htmlContent += "<div class='file-details'>`n"
                $fields = @{
                    "Action" = $primaryEntry.Action
                    "WriteNewFolder" = $primaryEntry.WriteNewFolder
                    "FileType" = $primaryEntry.FileType
                    "OldName" = $primaryEntry.OldName
                    "NewName" = $primaryEntry.NewName
                    "OpenFileInFolder" = $primaryEntry.OpenFileInFolder
                    "BackupLocation" = $primaryEntry.BackupLocation
                }
                foreach ($field in $fields.Keys) {
                    if ($fields[$field]) {
                        $value = [System.Net.WebUtility]::HtmlEncode($fields[$field])
                        if ($field -in "OpenFileInFolder", "BackupLocation" -and $value) {
                            $encodedPath = [System.Uri]::EscapeDataString($value -replace '\\', '/')
                            $value = "<a href='file:///$encodedPath' target='_blank'>$value</a>"
                        }
                        $displayName = $displayNames[$field]
                        $htmlContent += "<div class='label'>${displayName}:</div><div class='value'>$value</div>`n"
                    }
                }
                $htmlContent += "</div></details>`n"
            }

            $generalEntries = $allEntries | Where-Object { -not $_.OldName }
            if ($generalEntries) {
                $htmlContent += "<details class='file-section' data-status='info' data-folder='no' data-filetype='none'>`n"
                $htmlContent += "<summary>General Operations</summary>`n"
                $htmlContent += "<div class='file-details'>`n"
                foreach ($entry in $generalEntries) {
                    $status = $entry.Status.ToLower()
                    $statusClass = switch ($status) {
                        "success" { "status-success" }
                        "error"   { "status-error" }
                        "info"    { "status-info" }
                        default   { "status-other" }
                    }
                    $htmlContent += "<div class='label'>$($entry.Timestamp):</div><div class='value'><span class='$statusClass'>$([System.Net.WebUtility]::HtmlEncode($entry.Message)) ($($entry.Status))</span></div>`n"
                }
                $htmlContent += "</div></details>`n"
            }

            $htmlContent += @"
    <button id="back-to-top" onclick="window.scrollTo({top: 0, behavior: 'smooth'})">Back to Top</button>
    <footer>Generated by Music File Organizer - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</footer>
</body>
</html>
"@
            [System.IO.File]::WriteAllText($LogFilePath, $htmlContent, [System.Text.Encoding]::UTF8)
            Write-DebugLog "Generated HTML log file: $LogFilePath"
            return $true
        } catch {
            Write-DebugLog "Error writing HTML log file: $_"
            Write-Log -Message "Error writing HTML log file: $_" -Status "Error"
            return $false
        } finally {
            if (Test-Path -LiteralPath $script:tempLogPath) {
                Remove-Item -LiteralPath $script:tempLogPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # Initialize log entries array and cancellation flag
    $script:logEntries = @()
    $script:cancelOperation = $false
    $script:tempLogPath = Join-Path $env:TEMP "MusicFileOrganizer_TempLog.jsonl"
    $script:logFilePath = $null

    # Create the main form
    Write-DebugLog "Creating main form..."
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Music File Organizer"
    $form.Size = New-Object System.Drawing.Size(650, 400)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $form.ForeColor = [System.Drawing.Color]::White
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $form.Add_FormClosing({
        $script:logEntries = $null
        $script:cancelOperation = $false
        if (Test-Path -LiteralPath $script:tempLogPath) {
            Remove-Item -LiteralPath $script:tempLogPath -Force -ErrorAction SilentlyContinue
        }
        [System.GC]::Collect()
    })

    # Create toolbar (MenuStrip)
    $menuStrip = New-Object System.Windows.Forms.MenuStrip
    $menuStrip.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $menuStrip.ForeColor = [System.Drawing.Color]::White
    $toolsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $toolsMenu.Text = "Tools"
    $menuStrip.Items.Add($toolsMenu) | Out-Null

    $loadToolsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $loadToolsMenuItem.Text = "Load External Script"
    $loadToolsMenuItem.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        try {
            $openFileDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1"
            $openFileDialog.Title = "Select PowerShell Script"
            if ($openFileDialog.ShowDialog() -eq "OK") {
                $scriptPath = $openFileDialog.FileName
                if ([System.IO.Path]::GetExtension($scriptPath) -ne ".ps1") {
                    Write-Log -Message "Error: Selected file is not a .ps1 script." -Status "Error"
                    return
                }
                Write-Log -Message "Loading script: $scriptPath" -Status "Info"
                . $scriptPath
                Write-Log -Message "Script executed successfully." -Status "Success"
            }
        } catch {
            Write-Log -Message "Error executing script ${scriptPath}: $_" -Status "Error"
        } finally {
            $openFileDialog.Dispose()
        }
    })
    $toolsMenu.DropDownItems.Add($loadToolsMenuItem) | Out-Null

    $aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $aboutMenuItem.Text = "About"
    $aboutMenuItem.Add_Click({
        [System.Windows.Forms.MessageBox]::Show(
            "Music File Organizer v2.0`nCreated by Sanny Grooves for music producers & DJs`nBuilt with PowerShell and .NET Framework",
            "About Music File Organizer",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    $toolsMenu.DropDownItems.Add($aboutMenuItem) | Out-Null
    $form.Controls.Add($menuStrip)

    # Source directory controls
    $yPos = 40
    $sourceLabel = New-Object System.Windows.Forms.Label
    $sourceLabel.Location = New-Object System.Drawing.Point(15, $yPos)
    $sourceLabel.Size = New-Object System.Drawing.Size(130, 20)
    $sourceLabel.Text = "Source Directory:"
    $sourceLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($sourceLabel)

    $sourceTextBox = New-Object System.Windows.Forms.TextBox
    $sourceTextBox.Location = New-Object System.Drawing.Point(155, $yPos)
    $sourceTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $sourceTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $sourceTextBox.ForeColor = [System.Drawing.Color]::White
    $sourceTextBox.BorderStyle = "FixedSingle"
    $sourceTextBox.Add_TextChanged({
        $sourceTextBox.BackColor = if ([string]::IsNullOrWhiteSpace($sourceTextBox.Text)) { [System.Drawing.Color]::FromArgb(80, 0, 0) } else { [System.Drawing.Color]::FromArgb(51, 51, 51) }
    })
    $form.Controls.Add($sourceTextBox)

    $sourceButton = New-Object System.Windows.Forms.Button
    $sourceButton.Location = New-Object System.Drawing.Point(505, $yPos)
    $sourceButton.Size = New-Object System.Drawing.Size(80, 25)
    $sourceButton.Text = "Browse"
    $sourceButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $sourceButton.ForeColor = [System.Drawing.Color]::White
    $sourceButton.FlatStyle = "Flat"
    $sourceButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        try {
            $folderBrowser.Description = "Select Source Directory"
            if ($folderBrowser.ShowDialog() -eq "OK") {
                $sourceTextBox.Text = $folderBrowser.SelectedPath
            }
        } finally {
            $folderBrowser.Dispose()
        }
    })
    $form.Controls.Add($sourceButton)

    # Destination directory controls
    $yPos += 35
    $destLabel = New-Object System.Windows.Forms.Label
    $destLabel.Location = New-Object System.Drawing.Point(15, $yPos)
    $destLabel.Size = New-Object System.Drawing.Size(130, 20)
    $destLabel.Text = "Destination Directory:"
    $destLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($destLabel)

    $destTextBox = New-Object System.Windows.Forms.TextBox
    $destTextBox.Location = New-Object System.Drawing.Point(155, $yPos)
    $destTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $destTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $destTextBox.ForeColor = [System.Drawing.Color]::White
    $destTextBox.BorderStyle = "FixedSingle"
    $destTextBox.Add_TextChanged({
        $destTextBox.BackColor = if ([string]::IsNullOrWhiteSpace($destTextBox.Text)) { [System.Drawing.Color]::FromArgb(80, 0, 0) } else { [System.Drawing.Color]::FromArgb(51, 51, 51) }
    })
    $form.Controls.Add($destTextBox)

    $destButton = New-Object System.Windows.Forms.Button
    $destButton.Location = New-Object System.Drawing.Point(505, $yPos)
    $destButton.Size = New-Object System.Drawing.Size(80, 25)
    $destButton.Text = "Browse"
    $destButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $destButton.ForeColor = [System.Drawing.Color]::White
    $destButton.FlatStyle = "Flat"
    $destButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        try {
            $folderBrowser.Description = "Select Destination Directory"
            if ($folderBrowser.ShowDialog() -eq "OK") {
                $destTextBox.Text = $folderBrowser.SelectedPath
            }
        } finally {
            $folderBrowser.Dispose()
        }
    })
    $form.Controls.Add($destButton)

    # Backup directory controls
    $yPos += 35
    $backupLabel = New-Object System.Windows.Forms.Label
    $backupLabel.Location = New-Object System.Drawing.Point(15, $yPos)
    $backupLabel.Size = New-Object System.Drawing.Size(130, 20)
    $backupLabel.Text = "Backup Directory:"
    $backupLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($backupLabel)

    $backupTextBox = New-Object System.Windows.Forms.TextBox
    $backupTextBox.Location = New-Object System.Drawing.Point(155, $yPos)
    $backupTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $backupTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $backupTextBox.ForeColor = [System.Drawing.Color]::White
    $backupTextBox.BorderStyle = "FixedSingle"
    $form.Controls.Add($backupTextBox)

    $backupButton = New-Object System.Windows.Forms.Button
    $backupButton.Location = New-Object System.Drawing.Point(505, $yPos)
    $backupButton.Size = New-Object System.Drawing.Size(80, 25)
    $backupButton.Text = "Browse"
    $backupButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $backupButton.ForeColor = [System.Drawing.Color]::White
    $backupButton.FlatStyle = "Flat"
    $backupButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        try {
            $folderBrowser.Description = "Select Backup Directory"
            if ($folderBrowser.ShowDialog() -eq "OK") {
                $backupTextBox.Text = $folderBrowser.SelectedPath
            }
        } finally {
            $folderBrowser.Dispose()
        }
    })
    $form.Controls.Add($backupButton)

    # Replacements file controls
    $yPos += 35
    $replaceLabel = New-Object System.Windows.Forms.Label
    $replaceLabel.Location = New-Object System.Drawing.Point(15, $yPos)
    $replaceLabel.Size = New-Object System.Drawing.Size(130, 20)
    $replaceLabel.Text = "Replacements File:"
    $replaceLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($replaceLabel)

    $replaceTextBox = New-Object System.Windows.Forms.TextBox
    $replaceTextBox.Location = New-Object System.Drawing.Point(155, $yPos)
    $replaceTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $replaceTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $replaceTextBox.ForeColor = [System.Drawing.Color]::White
    $replaceTextBox.BorderStyle = "FixedSingle"
    $form.Controls.Add($replaceTextBox)

    $replaceButton = New-Object System.Windows.Forms.Button
    $replaceButton.Location = New-Object System.Drawing.Point(505, $yPos)
    $replaceButton.Size = New-Object System.Drawing.Size(80, 25)
    $replaceButton.Text = "Browse"
    $replaceButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $replaceButton.ForeColor = [System.Drawing.Color]::White
    $replaceButton.FlatStyle = "Flat"
    $replaceButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        try {
            $openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            $openFileDialog.Title = "Select Replacements File"
            if ($openFileDialog.ShowDialog() -eq "OK") {
                $replaceTextBox.Text = $openFileDialog.FileName
            }
        } finally {
            $openFileDialog.Dispose()
        }
    })
    $form.Controls.Add($replaceButton)

    # Conflict resolution toggle controls
    $yPos += 35
    $conflictLabel = New-Object System.Windows.Forms.Label
    $conflictLabel.Location = New-Object System.Drawing.Point(15, $yPos)
    $conflictLabel.Size = New-Object System.Drawing.Size(130, 20)
    $conflictLabel.Text = "Conflict Resolution:"
    $conflictLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($conflictLabel)

    $overwriteRadio = New-Object System.Windows.Forms.RadioButton
    $overwriteRadio.Location = New-Object System.Drawing.Point(155, $yPos)
    $overwriteRadio.Size = New-Object System.Drawing.Size(120, 20)
    $overwriteRadio.Text = "Overwrite Existing"
    $overwriteRadio.ForeColor = [System.Drawing.Color]::White
    $overwriteRadio.Checked = $true
    $form.Controls.Add($overwriteRadio)

    $skipRadio = New-Object System.Windows.Forms.RadioButton
    $skipRadio.Location = New-Object System.Drawing.Point(285, $yPos)
    $skipRadio.Size = New-Object System.Drawing.Size(100, 20)
    $skipRadio.Text = "Skip Existing"
    $skipRadio.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($skipRadio)

    $promptRadio = New-Object System.Windows.Forms.RadioButton
    $promptRadio.Location = New-Object System.Drawing.Point(395, $yPos)
    $promptRadio.Size = New-Object System.Drawing.Size(100, 20)
    $promptRadio.Text = "Prompt"
    $promptRadio.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($promptRadio)

    # Status label
    $yPos += 35
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(15, $yPos)
    $statusLabel.Size = New-Object System.Drawing.Size(560, 20)
    $statusLabel.Text = "Ready"
    $statusLabel.ForeColor = [System.Drawing.Color]::LightGray
    $form.Controls.Add($statusLabel)

    # Progress bar
    $yPos += 25
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(15, $yPos)
    $progressBar.Size = New-Object System.Drawing.Size(560, 20)
    $progressBar.Style = "Continuous"
    $form.Controls.Add($progressBar)

    # Buttons
    $yPos += 35
    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Location = New-Object System.Drawing.Point(130, $yPos)
    $startButton.Size = New-Object System.Drawing.Size(100, 30)
    $startButton.Text = "Start"
    $startButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $startButton.ForeColor = [System.Drawing.Color]::White
    $startButton.FlatStyle = "Flat"
    $startButton.Add_Click({
        Write-DebugLog "Start button clicked"
        $startButton.Enabled = $false
        $stopButton.Enabled = $true
        $statusLabel.Text = "Initializing..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Yellow
        $script:cancelOperation = $false

        # Prompt for log file path immediately
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        try {
            $saveFileDialog.Filter = "HTML Files (*.html)|*.html"
            $saveFileDialog.Title = "Save Log File"
            $saveFileDialog.FileName = "MusicFileOrganizer_Log_$timestamp.html"
            $saveFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')
            if ($saveFileDialog.ShowDialog() -ne "OK") {
                Write-Log -Message "Log file generation cancelled by user." -Status "Info"
                $startButton.Enabled = $true
                $stopButton.Enabled = $false
                $statusLabel.Text = "Ready"
                $statusLabel.ForeColor = [System.Drawing.Color]::LightGray
                return
            }
            $script:logFilePath = $saveFileDialog.FileName
        } catch {
            Write-DebugLog "Error in SaveFileDialog: $_"
            Write-Log -Message "Error selecting log file path: $_" -Status "Error"
            $startButton.Enabled = $true
            $stopButton.Enabled = $false
            $statusLabel.Text = "Error selecting log file"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            return
        } finally {
            $saveFileDialog.Dispose()
        }

        # Create a runspace for asynchronous processing
        $runspace = [RunspaceFactory]::CreateRunspace()
        $runspace.Open()
        $powershell = [PowerShell]::Create()
        $powershell.Runspace = $runspace

        $powershell.AddScript({
            param($source, $dest, $backup, $replaceFile, $conflictAction, $statusLabel, $progressBar, $logFilePath, $form)

            Write-DebugLog "Runspace started"

            try {
                # Validate inputs
                if ([string]::IsNullOrWhiteSpace($source)) {
                    Write-Log -Message "Error: Source directory is empty." -Status "Error"
                    return
                }
                if (-not (Test-Path -LiteralPath $source -PathType Container)) {
                    Write-Log -Message "Error: Source directory does not exist: $source" -Status "Error"
                    return
                }
                if ([string]::IsNullOrWhiteSpace($dest)) {
                    Write-Log -Message "Error: Destination directory is empty." -Status "Error"
                    return
                }
                if ($backup -and -not (Test-Path -LiteralPath (Split-Path $backup -Parent) -PathType Container)) {
                    Write-Log -Message "Error: Backup directory's parent path is invalid: $backup" -Status "Error"
                    return
                }
                if ($replaceFile -and -not (Test-Path -LiteralPath $replaceFile -PathType Leaf)) {
                    Write-Log -Message "Error: Replacements file does not exist: $replaceFile" -Status "Error"
                    return
                }

                # Initialize log entries
                $script:logEntries = @()
                if (Test-Path -LiteralPath $script:tempLogPath) {
                    Remove-Item -LiteralPath $script:tempLogPath -Force -ErrorAction SilentlyContinue
                }

                # Create destination directory
                if (-not (Test-Path -LiteralPath $dest)) {
                    try {
                        New-Item -ItemType Directory -Path $dest -ErrorAction Stop | Out-Null
                        Write-Log -Message "Created destination directory: $dest" -Status "Success"
                    } catch {
                        Write-Log -Message "Error creating destination directory: $_" -Status "Error"
                        return
                    }
                }
                Write-Log -Message "Started Music File Organizer" -Status "Info"

                # Create backup
                if (-not [string]::IsNullOrWhiteSpace($backup)) {
                    $backupDir = Join-Path $backup "Backup_$($timestamp)"
                    try {
                        Create-Backup -Source $source -BackupDir $backupDir
                    } catch {
                        Write-DebugLog "Error during backup: $_"
                        Write-Log -Message "Backup failed: $_" -Status "Error"
                    }
                    if ($script:cancelOperation) {
                        Write-Log -Message "Operation stopped by user during backup." -Status "Info"
                        return
                    }
                }

                # Create extension-specific directories
                $extensions = @("mp3", "wav", "flac", "aiff")
                foreach ($ext in $extensions) {
                    $extPath = Join-Path $dest $ext
                    if (-not (Test-Path -LiteralPath $extPath)) {
                        try {
                            New-Item -ItemType Directory -Path $extPath -ErrorAction Stop | Out-Null
                            Write-Log -Message "Created directory: $extPath" -Status "Success"
                        } catch {
                            Write-Log -Message "Error creating directory ${extPath}: $_" -Status "Error"
                        }
                    }
                }

                # Load replacements
                $removeList = @("Official Video", "Official Audio", "Lyrics", "HD", "Remix")
                if ($replaceFile -and (Test-Path -LiteralPath $replaceFile)) {
                    try {
                        $removeList = (Get-Content $replaceFile -Raw -ErrorAction Stop) -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                        Write-Log -Message "Loaded replacements file with $($removeList.Count) entries." -Status "Success"
                    } catch {
                        Write-Log -Message "Error reading replacements file: $_" -Status "Error"
                        Write-Log -Message "Using default replacements list." -Status "Info"
                    }
                } else {
                    Write-Log -Message "Using default replacements list." -Status "Info"
                }

                # Process files in batches
                $patterns = @("*.mp3", "*.wav", "*.flac", "*.aiff")
                $fileCount = 0
                foreach ($pattern in $patterns) {
                    $fileCount += ([System.IO.Directory]::EnumerateFiles($source, $pattern, [System.IO.SearchOption]::AllDirectories) | Measure-Object).Count
                }
                if ($fileCount -eq 0) {
                    Write-Log -Message "No audio files found in source directory." -Status "Error"
                    return
                }
                Write-Log -Message "Found $fileCount files to process." -Status "Info"
                $currentFile = 0
                $batchSize = 500
                $fileBatch = @()
                $batchNumber = 0
                $totalBatches = [Math]::Ceiling($fileCount / $batchSize)

                foreach ($pattern in $patterns) {
                    $files = [System.IO.Directory]::EnumerateFiles($source, $pattern, [System.IO.SearchOption]::AllDirectories)
                    foreach ($file in $files) {
                        if ($script:cancelOperation) {
                            Write-Log -Message "File processing stopped by user." -Status "Info"
                            return
                        }
                        $fileBatch += $file
                        if ($fileBatch.Count -ge $batchSize -or $file -eq $files[-1]) {
                            $batchNumber++
                            $form.Invoke([Action]{
                                $statusLabel.Text = "Processing batch $batchNumber of $totalBatches"
                                $progressBar.Value = [Math]::Min(100, ($currentFile / $fileCount) * 100)
                            })
                            foreach ($batchFile in $fileBatch) {
                                $currentFile++
                                try {
                                    $filename = [System.IO.Path]::GetFileNameWithoutExtension($batchFile)
                                    $ext = [System.IO.Path]::GetExtension($batchFile).TrimStart(".")
                                    $newFilename = $filename

                                    foreach ($removeText in $removeList) {
                                        $newFilename = $newFilename -replace [regex]::Escape($removeText), ""
                                    }

                                    $newFilename = Sanitize-Filename -Filename $newFilename
                                    if ($newFilename -ne $filename) {
                                        Write-Log -Message "Renamed: $filename.$ext to $newFilename.$ext" -Status "Success" -Action "Renamed" -OldName "$filename.$ext" -NewName "$newFilename.$ext" -FileType $ext
                                    }

                                    $fileDest = Join-Path $dest "$ext\$newFilename"
                                    $writeNewFolder = "No"
                                    if (Test-Path -LiteralPath $fileDest) {
                                        $action = $conflictAction
                                        if ($action -eq "Prompt") {
                                            $action = $form.Invoke([Func[string]]{ Show-ConflictDialog -Path $fileDest -Type "Folder" })
                                        }
                                        if ($action -eq "Overwrite") {
                                            try {
                                                Remove-Item -Path $fileDest -Recurse -Force -ErrorAction Stop
                                                New-Item -ItemType Directory -Path $fileDest -ErrorAction Stop | Out-Null
                                                Write-Log -Message "Overwritten folder: $fileDest" -Status "Success" -WriteNewFolder "Yes"
                                                $writeNewFolder = "Yes"
                                            } catch {
                                                Write-Log -Message "Error overwriting folder ${fileDest}: $_" -Status "Error" -OldName "$filename.$ext"
                                                continue
                                            }
                                        } elseif ($action -eq "KeepBoth") {
                                            $fileDest = Get-UniqueFolderPath -Path $fileDest
                                            try {
                                                New-Item -ItemType Directory -Path $fileDest -ErrorAction Stop | Out-Null
                                                Write-Log -Message "Created unique folder: $fileDest" -Status "Success" -WriteNewFolder "Yes"
                                                $writeNewFolder = "Yes"
                                            } catch {
                                                Write-Log -Message "Error creating folder ${fileDest}: $_" -Status "Error" -OldName "$filename.$ext"
                                                continue
                                            }
                                        } else {
                                            Write-Log -Message "Skipped file due to existing folder: $fileDest" -Status "Info" -OldName "$filename.$ext"
                                            continue
                                        }
                                    } else {
                                        try {
                                            New-Item -ItemType Directory -Path $fileDest -ErrorAction Stop | Out-Null
                                            Write-Log -Message "Created folder: $fileDest" -Status "Success" -WriteNewFolder "Yes"
                                            $writeNewFolder = "Yes"
                                        } catch {
                                            Write-Log -Message "Error creating folder ${fileDest}: $_" -Status "Error" -OldName "$filename.$ext"
                                            continue
                                        }
                                    }

                                    $destFile = Join-Path $fileDest "$newFilename.$ext"
                                    if ($destFile.Length -gt 260) {
                                        Write-Log -Message "Destination path too long: $destFile" -Status "Error" -OldName "$filename.$ext"
                                        continue
                                    }
                                    if (Test-Path -LiteralPath $destFile) {
                                        $action = $conflictAction
                                        if ($action -eq "Prompt") {
                                            $action = $form.Invoke([Func[string]]{ Show-ConflictDialog -Path $destFile -Type "File" })
                                        }
                                        if ($action -eq "Overwrite") {
                                            # Proceed with move
                                        } elseif ($action -eq "KeepBoth") {
                                            $destFile = Get-UniquePath -Path $destFile
                                        } else {
                                            Write-Log -Message "Skipped file due to existing file: $destFile" -Status "Info" -OldName "$filename.$ext"
                                            continue
                                        }
                                    }

                                    if (-not (Test-FileAccessible -Path $batchFile)) {
                                        Write-Log -Message "Error: File $filename.$ext is locked or inaccessible." -Status "Error" -OldName "$filename.$ext"
                                        continue
                                    }

                                    $maxRetries = 3
                                    $retryCount = 0
                                    $success = $false
                                    while (-not $success -and $retryCount -lt $maxRetries) {
                                        try {
                                            Move-Item -Path $batchFile -Destination $destFile -Force -ErrorAction Stop
                                            Write-Log -Message "Moved: $filename.$ext -> $destFile" -Status "Success" -Action "Moved" -OldName "$filename.$ext" -NewName "$newFilename.$ext" -FileType $ext -OpenFileInFolder $destFile -WriteNewFolder $writeNewFolder
                                            $success = $true
                                        } catch {
                                            $retryCount++
                                            if ($retryCount -eq $maxRetries) {
                                                Write-Log -Message "Error moving $filename.$ext after $maxRetries attempts: $_" -Status "Error" -OldName "$filename.$ext"
                                            } else {
                                                Start-Sleep -Milliseconds 500
                                            }
                                        }
                                    }
                                } catch {
                                    Write-DebugLog "Error processing file $batchFile`: $($_)"
                                    Write-Log -Message "Error processing $filename.$ext`: $_" -Status "Error" -OldName "$filename.$ext"
                                }
                            }
                            $fileBatch = @()
                            [System.GC]::Collect()
                            $form.Invoke([Action]{
                                $progressBar.Value = [Math]::Min(100, ($currentFile / $fileCount) * 100)
                            })
                        }
                    }
                }

                # Generate and open log file
                if ((Read-TempLogEntries).Count -gt 0 -or $script:logEntries.Count -gt 0) {
                    try {
                        if (Generate-HtmlLog -LogFilePath $logFilePath) {
                            try {
                                Invoke-Item $logFilePath -ErrorAction Stop
                                Write-Log -Message "Opened HTML log file: $logFilePath" -Status "Success"
                            } catch {
                                Write-Log -Message "Error opening HTML log file: $_" -Status "Error"
                            }
                        } else {
                            Write-Log -Message "Failed to generate HTML log file." -Status "Error"
                        }
                    } catch {
                        Write-DebugLog "Error generating log file: $_"
                        Write-Log -Message "Error generating log file: $_" -Status "Error"
                    }
                } else {
                    Write-Log -Message "No log entries to save." -Status "Info"
                }
            } catch {
                Write-DebugLog "Unexpected error in runspace: $_"
                Write-Log -Message "Unexpected error in operation: $_" -Status "Error"
                $form.Invoke([Action]{
                    [System.Windows.Forms.MessageBox]::Show(
                        "An unexpected error occurred: $_",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) 2>$null
                })
            }
        }).AddArgument($sourceTextBox.Text.Trim()).AddArgument($destTextBox.Text.Trim()).AddArgument($backupTextBox.Text.Trim()).AddArgument($replaceTextBox.Text.Trim()).AddArgument(
            $(if ($overwriteRadio.Checked) { "Overwrite" } elseif ($skipRadio.Checked) { "Skip" } else { "Prompt" })
        ).AddArgument($statusLabel).AddArgument($progressBar).AddArgument($script:logFilePath).AddArgument($form)

        # Handle runspace completion
        $asyncResult = $powershell.BeginInvoke()
        Register-ObjectEvent -InputObject $powershell -EventName InvocationStateChanged -Action {
            param($sender, $eventArgs)
            if ($eventArgs.InvocationStateInfo.State -eq 'Completed' -or $eventArgs.InvocationStateInfo.State -eq 'Failed') {
                $form.Invoke([Action]{
                    $progressBar.Value = 0
                    $startButton.Enabled = $true
                    $stopButton.Enabled = $false
                    if ($statusLabel.Text -notmatch "Error|Cancelled") {
                        $statusLabel.Text = "Ready"
                        $statusLabel.ForeColor = [System.Drawing.Color]::LightGray
                    }
                })
                $sender.Runspace.Dispose()
                $sender.Dispose()
                Write-DebugLog "Runspace completed or failed"
            }
        } | Out-Null
    })
    $form.Controls.Add($startButton)

    $stopButton = New-Object System.Windows.Forms.Button
    $stopButton.Location = New-Object System.Drawing.Point(250, $yPos)
    $stopButton.Size = New-Object System.Drawing.Size(100, 30)
    $stopButton.Text = "Stop"
    $stopButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $stopButton.ForeColor = [System.Drawing.Color]::White
    $stopButton.FlatStyle = "Flat"
    $stopButton.Enabled = $false
    $stopButton.Add_Click({
        Write-DebugLog "Stop button clicked"
        $script:cancelOperation = $true
        $stopButton.Enabled = $false
        $startButton.Enabled = $true
        $progressBar.Value = 0
        $statusLabel.Text = "Cancelled: Operation stopped"
        $statusLabel.ForeColor = [System.Drawing.Color]::Yellow
        Write-Log -Message "Operation cancelled by user." -Status "Info"

        # Generate and open log file on cancellation
        if ($script:logFilePath -and ((Read-TempLogEntries).Count -gt 0 -or $script:logEntries.Count -gt 0)) {
            try {
                if (Generate-HtmlLog -LogFilePath $script:logFilePath) {
                    try {
                        Invoke-Item $script:logFilePath -ErrorAction Stop
                        Write-Log -Message "Opened HTML log file: $script:logFilePath" -Status "Success"
                    } catch {
                        Write-Log -Message "Error opening HTML log file: $_" -Status "Error"
                    }
                } else {
                    Write-Log -Message "Failed to generate HTML log file." -Status "Error"
                }
            } catch {
                Write-DebugLog "Error generating log file on cancellation: $_"
                Write-Log -Message "Error generating log file: $_" -Status "Error"
            }
        } else {
            Write-Log -Message "No log entries to save on cancellation." -Status "Info"
        }
    })
    $form.Controls.Add($stopButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Location = New-Object System.Drawing.Point(370, $yPos)
    $exitButton.Size = New-Object System.Drawing.Size(100, 30)
    $exitButton.Text = "Exit"
    $exitButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $exitButton.ForeColor = [System.Drawing.Color]::White
    $exitButton.FlatStyle = "Flat"
    $exitButton.Add_Click({ $form.Close() })
    $form.Controls.Add($exitButton)

    # Tooltips
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.SetToolTip($sourceTextBox, "Enter or browse to the directory containing music files")
    $toolTip.SetToolTip($destTextBox, "Enter or browse to the directory where files will be organized")
    $toolTip.SetToolTip($backupTextBox, "Optional: Enter or browse to the directory for backups")
    $toolTip.SetToolTip($replaceTextBox, "Optional: Select a text file with comma-separated terms to remove from filenames")
    $toolTip.SetToolTip($overwriteRadio, "Overwrite existing files/folders without prompting")
    $toolTip.SetToolTip($skipRadio, "Skip files/folders that already exist")
    $toolTip.SetToolTip($promptRadio, "Prompt for action when conflicts occur")
    $toolTip.SetToolTip($startButton, "Begin organizing music files")
    $toolTip.SetToolTip($stopButton, "Stop the current operation")
    $toolTip.SetToolTip($exitButton, "Close the application")

    # Show the form
    Write-DebugLog "Showing form..."
    [void]$form.ShowDialog()
    $form.Dispose()
    Write-DebugLog "Form closed"
} catch {
    $errorMessage = "Unexpected error in script: $_`nSee $debugLogPath for details."
    Write-DebugLog $errorMessage
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) 2>$null
    Write-Error $errorMessage
    exit 1
}