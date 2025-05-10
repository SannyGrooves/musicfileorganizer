# PowerShell script to move audio files into individual folders named after the files
# Supports replacements.txt, add.txt, backup, duplicate handling with highest quality selection, duplicate song checking by name,
# automatic space cleanup, adjustable similarity threshold, and batch processing by file size
# Automatically installs ffprobe (via FFmpeg) if missing
# Requires .NET Framework 4.5 or later
# Added breakpoints for debugging (marked with # BREAKPOINT and optional Set-DebugBreakpoints function)
# Fixed BeginInvoke error by checking IsHandleCreated in Write-Log
# Added "BEST" to filename of highest quality file (based on format, bitrate, and size)
# Improved $bestFile selection to prioritize lossless formats (.flac, .wav, .aiff) over lossy (.mp3, .ogg, .wma)
# Added batch processing by file size (default: 300MB) to handle large datasets (e.g., 60GB+ .flac/.wav, 30GB+ .mp3)
# Fixed inconsistent casing of "BEST" suffix to ensure it is always "BEST" (uppercase) for highest quality files
# Added automated import safety for large folders by loading files incrementally to prevent crashes
# Optimized incremental file handling for audio files between 20MB and 100MB
# Set default buffer size to 16MB and batch size to 300MB, optimized for 20MB–100MB files
# Lowered memory threshold to 300MB and added memory monitoring for stability
# Fixed variable reference error by using ${batchNumber} in Write-Log strings
# Added real-time status updates for file enumeration to show progress in UI

# Set error action preference to stop on all errors
$ErrorActionPreference = "Stop"

# Function to log messages asynchronously
function Write-Log {
    param (
        [string]$Message,
        [string]$Status = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = @{
        Timestamp = $timestamp
        Message   = $Message
        Status    = $Status
    }
    $script:logEntries += $logEntry
    if ($outputTextBox -and $outputTextBox.IsHandleCreated) {
        try {
            $outputTextBox.BeginInvoke([Action]{
                $outputTextBox.AppendText("[$timestamp] $Message ($Status)`r`n")
                $outputTextBox.ScrollToCaret()
            }) | Out-Null
        } catch {
            # Silently ignore UI update errors to prevent recursive logging
        }
    }
}

# Function to update UI status asynchronously
function Update-UIStatus {
    param (
        [string]$Message,
        [int]$ProgressValue = -1
    )
    if ($batchProgressLabel -and $batchProgressLabel.IsHandleCreated) {
        try {
            $batchProgressLabel.BeginInvoke([Action]{
                $batchProgressLabel.Text = $Message
                if ($ProgressValue -ge 0 -and $ProgressValue -le 100) {
                    $progressBar.Value = $ProgressValue
                }
            }) | Out-Null
        } catch {
            # Silently ignore UI update errors
        }
    }
}

# Function to set breakpoints programmatically for debugging
function Set-DebugBreakpoints {
    # Remove existing breakpoints to avoid duplicates
    Get-PSBreakpoint | Remove-PSBreakpoint

    # Set breakpoints at key lines (line numbers based on this script)
    $scriptPath = $PSCommandPath
    $breakpoints = @(
        @{Line = 755; Description = "Start of Start button click event"},
        @{Line = 796; Description = "After input validation"},
        @{Line = 844; Description = "Before batch creation"},
        @{Line = 856; Description = "Inside incremental file enumeration"},
        @{Line = 906; Description = "Start of batch processing loop"},
        @{Line = 917; Description = "Start of file entry creation loop"},
        @{Line = 970; Description = "Before similarity group creation"},
        @{Line = 992; Description = "Inside similarity group creation loop"},
        @{Line = 1014; Description = "Before processing similarity groups"},
        @{Line = 1024; Description = "Inside similarity group processing loop"},
        @{Line = 1128; Description = "Before processing file groups"},
        @{Line = 1138; Description = "Inside file group processing loop"},
        @{Line = 1153; Description = "Before best file selection"}
    )

    foreach ($bp in $breakpoints) {
        Write-Log -Message "Setting breakpoint at line $($bp.Line): $($bp.Description)" -Status "Info"
        Set-PSBreakpoint -Script $scriptPath -Line $bp.Line
    }
    Write-Log -Message "Breakpoints set. Run the script in a debugger (e.g., VS Code, PowerShell ISE) to pause at breakpoints." -Status "Info"
}

# Function to check and install dependencies
function Install-Dependencies {
    # Check .NET Framework 4.5 or later
    try {
        $netVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction Stop).Release
        if ($netVersion -lt 378389) { # 378389 is .NET 4.5
            $msg = "Error: .NET Framework 4.5 or later is required. Please install it via Windows Update or download from Microsoft."
            Write-Log -Message $msg -Status "Error"
            [System.Windows.Forms.MessageBox]::Show($msg, "Dependency Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            exit 1
        }
        Write-Log -Message ".NET Framework 4.5 or later detected." -Status "Info"
    } catch {
        $msg = "Error: Unable to verify .NET Framework version. Ensure .NET Framework 4.5 or later is installed."
        Write-Log -Message $msg -Status "Error"
        [System.Windows.Forms.MessageBox]::Show($msg, "Dependency Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }

    # Check for ffprobe
    if (-not (Get-Command ffprobe -ErrorAction SilentlyContinue)) {
        Write-Log -Message "ffprobe not found. Attempting to install FFmpeg..." -Status "Info"
        try {
            # Try winget first
            if (Get-Command winget -ErrorAction SilentlyContinue) {
                Write-Log -Message "Installing FFmpeg via winget..." -Status "Info"
                Start-Process winget -ArgumentList "install --id Gyan.FFmpeg -e --silent" -Wait -NoNewWindow
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
                if (Get-Command ffprobe -ErrorAction SilentlyContinue) {
                    Write-Log -Message "FFmpeg (ffprobe) installed successfully via winget." -Status "Success"
                    return
                }
            }

            # Fallback to downloading static build
            Write-Log -Message "winget unavailable or failed. Downloading FFmpeg static build..." -Status "Info"
            $ffmpegDir = Join-Path $PSScriptRoot "ffmpeg"
            $ffmpegZip = Join-Path $ffmpegDir "ffmpeg.zip"
            $ffmpegUrl = "https://github.com/GyanD/codexffmpeg/releases/download/2023-11-14-git-6c69bd04b2/ffmpeg-2023-11-14-git-6c69bd04b2-full_build.zip"
            if (-not (Test-Path $ffmpegDir)) {
                New-Item -ItemType Directory -Path $ffmpegDir -ErrorAction Stop | Out-Null
            }
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            (New-Object Net.WebClient).DownloadFile($ffmpegUrl, $ffmpegZip)
            Expand-Archive -Path $ffmpegZip -DestinationPath $ffmpegDir -Force -ErrorAction Stop
            $ffmpegExePath = Get-ChildItem -Path $ffmpegDir -Recurse -Include ffprobe.exe | Select-Object -First 1
            if ($ffmpegExePath) {
                $env:Path += ";$($ffmpegExePath.DirectoryName)"
                [System.Environment]::SetEnvironmentVariable("Path", $env:Path, [System.EnvironmentVariableTarget]::Process)
                Write-Log -Message "FFmpeg (ffprobe) installed successfully to $ffmpegDir." -Status "Success"
            } else {
                throw "ffprobe.exe not found in downloaded FFmpeg package."
            }
        } catch {
            Write-Log -Message "Error installing FFmpeg: $_" -Status "Error"
            [System.Windows.Forms.MessageBox]::Show("Failed to install ffprobe (FFmpeg). Please install FFmpeg manually and ensure ffprobe is in your PATH.", "Dependency Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            exit 1
        }
    } else {
        Write-Log -Message "ffprobe detected in PATH." -Status "Info"
    }
}

# Function to clean leading/trailing spaces from filenames
function Clean-FilenameSpaces {
    param ([string]$Filename)
    $cleaned = $Filename.Trim()
    return $cleaned
}

# Function to compute Levenshtein distance for string similarity
function Get-LevenshteinDistance {
    param (
        [string]$String1,
        [string]$String2
    )
    $len1 = $String1.Length
    $len2 = $String2.Length
    $dist = New-Object 'int[,]' ($len1 + 1), ($len2 + 1)

    # Initialize first row and column
    for ($i = 0; $i -le $len1; $i++) {
        $dist[$i, 0] = $i
    }
    for ($j = 0; $j -le $len2; $j++) {
        $dist[0, $j] = $j
    }

    # Fill distance matrix
    for ($i = 1; $i -le $len1; $i++) {
        for ($j = 1; $j -le $len2; $j++) {
            $cost = if ($String1[$i-1] -eq $String2[$j-1]) { 0 } else { 1 }
            $delete = $dist[$i-1, $j] + 1
            $insert = $dist[$i, $j-1] + 1
            $substitute = $dist[$i-1, $j-1] + $cost
            $dist[$i, $j] = [Math]::Min($delete, [Math]::Min($insert, $substitute))
        }
    }
    return $dist[$len1, $len2]
}

# Function to compute word-based similarity (for loose matching)
function Compare-WordBasedSimilarity {
    param (
        [string]$String1,
        [string]$String2,
        [string[]]$RemoveList
    )
    # Normalize strings: lowercase, remove common terms, split into words
    $norm1 = $String1.ToLower()
    $norm2 = $String2.ToLower()
    foreach ($term in $RemoveList) {
        $norm1 = $norm1 -replace [regex]::Escape($term.ToLower()), ""
        $norm2 = $norm2 -replace [regex]::Escape($term.ToLower()), ""
    }
    $words1 = ($norm1 -split '\s+' | Where-Object { $_ -and $_.Length -gt 2 }) | Sort-Object
    $words2 = ($norm2 -split '\s+' | Where-Object { $_ -and $_.Length -gt 2 }) | Sort-Object

    # Count common words
    $commonWords = ($words1 | Where-Object { $words2 -contains $_ } | Measure-Object).Count
    return $commonWords -gt 2
}

# Function to compute string similarity (0 to 1, where 1 is identical)
function Get-StringSimilarity {
    param (
        [string]$String1,
        [string]$String2,
        [int]$SliderValue,
        [string[]]$RemoveList
    )
    if ([string]::IsNullOrEmpty($String1) -or [string]::IsNullOrEmpty($String2)) { return 0 }
    
    if ($SliderValue -eq 1) {
        # Word-based similarity: true if more than 2 words match
        return if (Compare-WordBasedSimilarity -String1 $String1 -String2 $String2 -RemoveList $RemoveList) { 1 } else { 0 }
    } elseif ($SliderValue -eq 100) {
        # Exact match required
        return if ($String1 -eq $String2) { 1 } else { 0 }
    } else {
        # Levenshtein-based similarity with interpolated threshold
        $distance = Get-LevenshteinDistance -String1 $String1 -String2 $String2
        $maxLength = [Math]::Max($String1.Length, $String2.Length)
        $similarity = [Math]::Max(0, 1 - ($distance / $maxLength))
        return $similarity
    }
}

# Function to sanitize filenames (excluding the "BEST" suffix)
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

# Function to get format priority (higher is better quality)
function Get-FormatPriority {
    param ([string]$Extension)
    $losslessFormats = @("flac", "wav", "aiff")
    $lossyFormats = @("mp3", "ogg", "wma")
    $ext = $Extension.ToLower().TrimStart(".")
    if ($losslessFormats -contains $ext) {
        return 2 # Lossless formats have higher priority
    } elseif ($lossyFormats -contains $ext) {
        return 1 # Lossy formats have lower priority
    } else {
        return 0 # Unknown formats
    }
}

# Function to get unique folder path for duplicates
function Get-UniqueFolderPath {
    param ([string]$FolderPath)
    $basePath = [System.IO.Path]::GetDirectoryName($FolderPath)
    $folderName = [System.IO.Path]::GetFileName($FolderPath)
    $counter = 0
    $newPath = $FolderPath
    while (Test-Path -LiteralPath $newPath) {
        $counter++
        $newPath = Join-Path $basePath "${folderName}_duplicate$counter"
    }
    return $newPath
}

# Function to test if file is accessible
function Test-FileAccessible {
    param ([string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            Write-Log -Message "File does not exist: $Path" -Status "Error"
            return $false
        }
        return $true
    } catch {
        Write-Log -Message "File accessibility check failed for ${Path}: $_" -Status "Error"
        return $false
    }
}

# Function to get file bitrate using ffprobe, with estimation for lossless formats
function Get-FileBitrate {
    param ([string]$FilePath)
    try {
        $ffprobeOutput = & ffprobe -v error -show_entries stream=bit_rate:format=duration -of json $FilePath 2>&1
        $json = $ffprobeOutput | ConvertFrom-Json
        $bitrate = [int]($json.streams[0].bit_rate)
        $duration = [double]($json.format.duration)
        
        # For lossless formats, estimate bitrate if ffprobe returns 0 or invalid
        $ext = [System.IO.Path]::GetExtension($FilePath).ToLower().TrimStart(".")
        if ($bitrate -le 0 -and @("flac", "wav", "aiff") -contains $ext) {
            if ($duration -gt 0) {
                $fileSize = (Get-Item -LiteralPath $FilePath).Length
                # Estimate bitrate: (file size in bits) / duration in seconds
                $bitrate = [math]::Round(($fileSize * 8) / $duration)
                Write-Log -Message "Estimated bitrate for lossless file ${FilePath}: $bitrate bps" -Status "Info"
            } else {
                # Fallback to 0 if duration is unavailable
                Write-Log -Message "Could not estimate bitrate for ${FilePath}: Duration unavailable" -Status "Warning"
                $bitrate = 0
            }
        }
        return $bitrate
    } catch {
        Write-Log -Message "Error getting bitrate for ${FilePath}: $_" -Status "Error"
        return 0
    }
}

# Function to copy large files with dynamic buffer size
function Copy-LargeFile {
    param (
        [string]$SourcePath,
        [string]$DestPath,
        [int]$BufferSizeBytes
    )
    try {
        $buffer = New-Object byte[] $BufferSizeBytes
        $sourceStream = [System.IO.File]::OpenRead($SourcePath)
        $destStream = [System.IO.File]::Create($DestPath)
        $bytesRead = 0
        while (($bytesRead = $sourceStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $destStream.Write($buffer, 0, $bytesRead)
            [System.Windows.Forms.Application]::DoEvents()
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
        [string]$BackupDir,
        [int]$BufferSizeBytes
    )
    try {
        if (-not (Test-Path -LiteralPath $BackupDir)) {
            New-Item -ItemType Directory -Path $BackupDir -ErrorAction Stop | Out-Null
        }
        $files = Get-ChildItem -Path $Source -Recurse -Include *.mp3,*.wav,*.flac,*.aiff,*.ogg,*.wma -ErrorAction SilentlyContinue
        $fileCount = ($files | Measure-Object).Count
        $currentFile = 0
        foreach ($file in $files) {
            $currentFile++
            $progressBar.Value = [Math]::Min(100, ($currentFile / $fileCount) * 100)
            $relativePath = $file.FullName.Substring($Source.Length).TrimStart('\')
            $backupPath = Join-Path $BackupDir $relativePath
            $backupFolder = [System.IO.Path]::GetDirectoryName($backupPath)
            if (-not (Test-Path -LiteralPath $backupFolder)) {
                New-Item -ItemType Directory -Path $backupFolder -ErrorAction Stop | Out-Null
            }
            if (Copy-LargeFile -SourcePath $file.FullName -DestPath $backupPath -BufferSizeBytes $BufferSizeBytes) {
                Write-Log -Message "Backed up: $($file.FullName) to $backupPath" -Status "Success"
            }
            [System.GC]::Collect()
        }
        Write-Log -Message "Backup completed to: $BackupDir" -Status "Success"
    } catch {
        Write-Log -Message "Error creating backup: $_" -Status "Error"
    } finally {
        $progressBar.Value = 0
    }
}

# Function to generate enhanced HTML log file
function Generate-HtmlLog {
    param (
        [string]$LogFilePath,
        [int]$BatchNumber = 0
    )
    try {
        $successCount = ($script:logEntries | Where-Object { $_.Status -eq "Success" }).Count
        $errorCount = ($script:logEntries | Where-Object { $_.Status -eq "Error" }).Count
        $batchInfo = if ($BatchNumber -gt 0) { " (Batch $BatchNumber)" } else { "" }
        $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Music File Organizer Log$batchInfo</title>
    <style>
        body { background: linear-gradient(135deg, #1a1a1a, #2a2a2a); color: #ffffff; font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; line-height: 1.6; }
        h1 { text-align: center; color: #ffffff; text-shadow: 0 0 10px rgba(255,255,255,0.3); }
        .summary { background: #333333; padding: 20px; border-radius: 10px; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 8px rgba(0,0,0,0.3); }
        .summary span { font-size: 1.2em; margin: 0 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; background: #333333; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 8px rgba(0,0,0,0.3); }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #555555; }
        th { background: #444444; color: #ffffff; text-transform: uppercase; letter-spacing: 1px; }
        tr:hover { background: #3a3a3a; }
        .status-success { color: #28a745; font-weight: bold; }
        .status-error { color: #dc3545; font-weight: bold; }
        .status-info { color: #17a2b8; font-weight: bold; }
        footer { text-align: center; margin-top: 20px; color: #aaaaaa; font-size: 0.9em; }
    </style>
</head>
<body>
    <h1>Music File Organizer Log$batchInfo</h1>
    <div class="summary">
        <span>Successful Operations: $successCount</span>
        <span>Failed Operations: $errorCount</span>
    </div>
    <table>
        <tr><th>Timestamp</th><th>Message</th><th>Status</th></tr>
"@
        foreach ($entry in $script:logEntries) {
            $statusClass = switch ($entry.Status) {
                "Success" { "status-success" }
                "Error"   { "status-error" }
                default   { "status-info" }
            }
            $escapedMessage = [System.Net.WebUtility]::HtmlEncode($entry.Message)
            $htmlContent += "<tr><td>$($entry.Timestamp)</td><td>$escapedMessage</td><td class='$statusClass'>$($entry.Status)</td></tr>`n"
        }
        $htmlContent += @"
    </table>
    <footer>Generated by Music File Organizer - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</footer>
</body>
</html>
"@
        [System.IO.File]::WriteAllText($LogFilePath, $htmlContent, [System.Text.Encoding]::UTF8)
        return $true
    } catch {
        Write-Log -Message "Error writing HTML log file: $_" -Status "Error"
        return $false
    }
}

# Function to generate summary HTML log file
function Generate-SummaryLog {
    param (
        [string]$SummaryLogFilePath,
        [string[]]$BatchLogFiles
    )
    try {
        $totalSuccess = 0
        $totalError = 0
        foreach ($batchLog in $BatchLogFiles) {
            $logContent = [System.IO.File]::ReadAllText($batchLog)
            $successMatch = [regex]::Match($logContent, 'Successful Operations: (\d+)')
            $errorMatch = [regex]::Match($logContent, 'Failed Operations: (\d+)')
            if ($successMatch.Success) { $totalSuccess += [int]$successMatch.Groups[1].Value }
            if ($errorMatch.Success) { $totalError += [int]$errorMatch.Groups[1].Value }
        }
        $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Music File Organizer Summary Log</title>
    <style>
        body { background: linear-gradient(135deg, #1a1a1a, #2a2a2a); color: #ffffff; font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; line-height: 1.6; }
        h1 { text-align: center; color: #ffffff; text-shadow: 0 0 10px rgba(255,255,255,0.3); }
        .summary { background: #333333; padding: 20px; border-radius: 10px; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 8px rgba(0,0,0,0.3); }
        .summary span { font-size: 1.2em; margin: 0 20px; }
        ul { list-style: none; padding: 0; }
        li { margin: 10px 0; }
        a { color: #17a2b8; text-decoration: none; }
        a:hover { text-decoration: underline; }
        footer { text-align: center; margin-top: 20px; color: #aaaaaa; font-size: 0.9em; }
    </style>
</head>
<body>
    <h1>Music File Organizer Summary Log</h1>
    <div class="summary">
        <span>Total Successful Operations: $totalSuccess</span>
        <span>Total Failed Operations: $totalError</span>
    </div>
    <h2>Batch Log Files</h2>
    <ul>
"@
        foreach ($batchLog in $BatchLogFiles) {
            $batchName = [System.IO.Path]::GetFileName($batchLog)
            $htmlContent += "<li><a href=`"$batchName`">$batchName</a></li>`n"
        }
        $htmlContent += @"
    </ul>
    <footer>Generated by Music File Organizer - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</footer>
</body>
</html>
"@
        [System.IO.File]::WriteAllText($SummaryLogFilePath, $htmlContent, [System.Text.Encoding]::UTF8)
        return $true
    } catch {
        Write-Log -Message "Error writing summary HTML log file: $_" -Status "Error"
        return $false
    }
}

# Function to enumerate files incrementally, optimized for 20MB–100MB files
function Get-FilesIncrementally {
    param (
        [string]$Path,
        [string[]]$Include,
        [int64]$BatchSizeBytes,
        [scriptblock]$BatchCallback
    )
    try {
        Update-UIStatus -Message "Enumerating files in source folder..."
        $directoryInfo = New-Object System.IO.DirectoryInfo($Path)
        $extensions = $Include | ForEach-Object { $_ -replace '^\*\.', '.' }
        $currentBatch = @()
        $currentBatchSize = 0
        $fileCount = 0
        $dirCount = 0
        $batchIndex = 0

        # Recursively enumerate files
        $stack = New-Object System.Collections.Generic.Stack[System.IO.DirectoryInfo]
        $stack.Push($directoryInfo)

        while ($stack.Count -gt 0) {
            $currentDir = $stack.Pop()
            $dirCount++
            try {
                # Update UI with current progress
                Update-UIStatus -Message "Enumerating files: $fileCount files, $dirCount directories scanned"

                # Collect files and sort by size (descending) to prioritize 20MB–100MB files
                $files = $currentDir.EnumerateFiles() | Where-Object { $extensions -contains $_.Extension.ToLower() } | Sort-Object Length -Descending
                foreach ($file in $files) {
                    $fileCount++
                    $fileSize = $file.Length
                    if ($currentBatchSize + $fileSize -gt $BatchSizeBytes -and $currentBatch.Count -gt 0) {
                        # Process current batch
                        $batchIndex++
                        & $BatchCallback $currentBatch
                        Update-UIStatus -Message "Created batch $batchIndex with $($currentBatch.Count) files"
                        $currentBatch = @()
                        $currentBatchSize = 0
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()
                        Start-Sleep -Milliseconds 100 # Allow memory stabilization
                    }
                    $currentBatch += $file
                    $currentBatchSize += $fileSize
                    # Update progress bar (0–100, reset per batch)
                    $progressValue = [Math]::Min(100, ($currentBatch.Count % 100))
                    Update-UIStatus -Message "Enumerating files: $fileCount files, $dirCount directories scanned" -ProgressValue $progressValue
                }
                foreach ($subDir in $currentDir.EnumerateDirectories()) {
                    $stack.Push($subDir)
                }
            } catch {
                Write-Log -Message "Error accessing directory $($currentDir.FullName): $_" -Status "Error"
            }
        }

        # Process final batch
        if ($currentBatch.Count -gt 0) {
            $batchIndex++
            & $BatchCallback $currentBatch
            Update-UIStatus -Message "Created batch $batchIndex with $($currentBatch.Count) files"
        }

        Update-UIStatus -Message "File enumeration complete: $fileCount files in $batchIndex batches"
        return $fileCount
    } catch {
        Write-Log -Message "Error during incremental file enumeration: $_" -Status "Error"
        Update-UIStatus -Message "Error during file enumeration"
        return 0
    }
}

# Initialize log entries array
$script:logEntries = @()

# Main execution with crash handling
try {
    # Install dependencies
    Install-Dependencies

    # Dynamically load required assemblies
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Add-Type -AssemblyName System.Collections -ErrorAction Stop
    } catch {
        $errorMessage = "Error: Failed to load required .NET assemblies. Ensure .NET Framework 4.5 or later is installed.`nDetails: $_"
        Write-Log -Message $errorMessage -Status "Error"
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }

    # Create the main form
    $form = New-Object System.Windows.Forms.Form
    try {
        $form.Text = "Music File Organizer"
        $form.Size = New-Object System.Drawing.Size(620, 680)
        $form.StartPosition = "CenterScreen"
        $form.BackColor = [System.Drawing.Color]::Black
        $form.ForeColor = [System.Drawing.Color]::White
        $form.Add_FormClosing({
            $script:logEntries = $null
            [System.GC]::Collect()
        })
    } catch {
        Write-Log -Message "Error initializing form: $_" -Status "Error"
        throw
    }

    # Source directory controls
    $sourceLabel = New-Object System.Windows.Forms.Label
    $sourceLabel.Location = New-Object System.Drawing.Point(15, 40)
    $sourceLabel.Size = New-Object System.Drawing.Size(130, 20)
    $sourceLabel.Text = "Source Directory:"
    $sourceLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($sourceLabel)

    $sourceTextBox = New-Object System.Windows.Forms.TextBox
    $sourceTextBox.Location = New-Object System.Drawing.Point(155, 40)
    $sourceTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $sourceTextBox.Text = ""
    $sourceTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $sourceTextBox.ForeColor = [System.Drawing.Color]::White
        $form.Controls.Add($sourceTextBox)

    $sourceButton = New-Object System.Windows.Forms.Button
    $sourceButton.Location = New-Object System.Drawing.Point(505, 40)
    $sourceButton.Size = New-Object System.Drawing.Size(80, 25)
    $sourceButton.Text = "Browse"
    $sourceButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $sourceButton.ForeColor = [System.Drawing.Color]::White
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
    $destLabel = New-Object System.Windows.Forms.Label
    $destLabel.Location = New-Object System.Drawing.Point(15, 75)
    $destLabel.Size = New-Object System.Drawing.Size(130, 20)
    $destLabel.Text = "Destination Directory:"
    $destLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($destLabel)

    $destTextBox = New-Object System.Windows.Forms.TextBox
    $destTextBox.Location = New-Object System.Drawing.Point(155, 75)
    $destTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $destTextBox.Text = ""
    $destTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $destTextBox.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($destTextBox)

    $destButton = New-Object System.Windows.Forms.Button
    $destButton.Location = New-Object System.Drawing.Point(505, 75)
    $destButton.Size = New-Object System.Drawing.Size(80, 25)
    $destButton.Text = "Browse"
    $destButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $destButton.ForeColor = [System.Drawing.Color]::White
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
    $backupLabel = New-Object System.Windows.Forms.Label
    $backupLabel.Location = New-Object System.Drawing.Point(15, 110)
    $backupLabel.Size = New-Object System.Drawing.Size(130, 20)
    $backupLabel.Text = "Backup Directory:"
    $backupLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($backupLabel)

    $backupTextBox = New-Object System.Windows.Forms.TextBox
    $backupTextBox.Location = New-Object System.Drawing.Point(155, 110)
    $backupTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $backupTextBox.Text = ""
    $backupTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $backupTextBox.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($backupTextBox)

    $backupButton = New-Object System.Windows.Forms.Button
    $backupButton.Location = New-Object System.Drawing.Point(505, 110)
    $backupButton.Size = New-Object System.Drawing.Size(80, 25)
    $backupButton.Text = "Browse"
    $backupButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $backupButton.ForeColor = [System.Drawing.Color]::White
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
    $replaceLabel = New-Object System.Windows.Forms.Label
    $replaceLabel.Location = New-Object System.Drawing.Point(15, 145)
    $replaceLabel.Size = New-Object System.Drawing.Size(130, 20)
    $replaceLabel.Text = "Replacements File:"
    $replaceLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($replaceLabel)

    $replaceTextBox = New-Object System.Windows.Forms.TextBox
    $replaceTextBox.Location = New-Object System.Drawing.Point(155, 145)
    $replaceTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $replaceTextBox.Text = ""
    $replaceTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $replaceTextBox.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($replaceTextBox)

    $replaceButton = New-Object System.Windows.Forms.Button
    $replaceButton.Location = New-Object System.Drawing.Point(505, 145)
    $replaceButton.Size = New-Object System.Drawing.Size(80, 25)
    $replaceButton.Text = "Browse"
    $replaceButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $replaceButton.ForeColor = [System.Drawing.Color]::White
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

    # Add text file controls
    $addLabel = New-Object System.Windows.Forms.Label
    $addLabel.Location = New-Object System.Drawing.Point(15, 180)
    $addLabel.Size = New-Object System.Drawing.Size(130, 20)
    $addLabel.Text = "Add Text File:"
    $addLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($addLabel)

    $addTextBox = New-Object System.Windows.Forms.TextBox
    $addTextBox.Location = New-Object System.Drawing.Point(155, 180)
    $addTextBox.Size = New-Object System.Drawing.Size(340, 20)
    $addTextBox.Text = ""
    $addTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $addTextBox.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($addTextBox)

    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Location = New-Object System.Drawing.Point(505, 180)
    $addButton.Size = New-Object System.Drawing.Size(80, 25)
    $addButton.Text = "Browse"
    $addButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $addButton.ForeColor = [System.Drawing.Color]::White
    $addButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        try {
            $openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            $openFileDialog.Title = "Select Add Text File"
            if ($openFileDialog.ShowDialog() -eq "OK") {
                $addTextBox.Text = $openFileDialog.FileName
            }
        } finally {
            $openFileDialog.Dispose()
        }
    })
    $form.Controls.Add($addButton)

    # Highest quality duplicate checkbox
    $qualityCheckBox = New-Object System.Windows.Forms.CheckBox
    $qualityCheckBox.Location = New-Object System.Drawing.Point(15, 215)
    $qualityCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $qualityCheckBox.Text = "Move Highest Quality Duplicate (Format, Bitrate, Size)"
    $qualityCheckBox.ForeColor = [System.Drawing.Color]::White
    $qualityCheckBox.Checked = $false
    $form.Controls.Add($qualityCheckBox)

    # Duplicate song check by name checkbox
    $duplicateNameCheckBox = New-Object System.Windows.Forms.CheckBox
    $duplicateNameCheckBox.Location = New-Object System.Drawing.Point(15, 240)
    $duplicateNameCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $duplicateNameCheckBox.Text = "Check for Duplicated Songs by Name"
    $duplicateNameCheckBox.ForeColor = [System.Drawing.Color]::White
    $duplicateNameCheckBox.Checked = $false
    $form.Controls.Add($duplicateNameCheckBox)

    # Remove leading/trailing spaces checkbox
    $spaceCleanupCheckBox = New-Object System.Windows.Forms.CheckBox
    $spaceCleanupCheckBox.Location = New-Object System.Drawing.Point(15, 265)
    $spaceCleanupCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $spaceCleanupCheckBox.Text = "Remove Leading/Trailing Spaces from Filenames"
    $spaceCleanupCheckBox.ForeColor = [System.Drawing.Color]::White
    $spaceCleanupCheckBox.Checked = $true
    $form.Controls.Add($spaceCleanupCheckBox)

    # Similarity threshold slider
    $similarityLabel = New-Object System.Windows.Forms.Label
    $similarityLabel.Location = New-Object System.Drawing.Point(15, 290)
    $similarityLabel.Size = New-Object System.Drawing.Size(200, 20)
    $similarityLabel.Text = "Duplicate Name Similarity Threshold:"
    $similarityLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($similarityLabel)

    $similaritySlider = New-Object System.Windows.Forms.TrackBar
    $similaritySlider.Location = New-Object System.Drawing.Point(220, 290)
    $similaritySlider.Size = New-Object System.Drawing.Size(300, 45)
    $similaritySlider.Minimum = 1
    $similaritySlider.Maximum = 100
    $similaritySlider.Value = 90
    $similaritySlider.TickFrequency = 10
    $form.Controls.Add($similaritySlider)

    $similarityValueLabel = New-Object System.Windows.Forms.Label
    $similarityValueLabel.Location = New-Object System.Drawing.Point(530, 290)
    $similarityValueLabel.Size = New-Object System.Drawing.Size(50, 20)
    $similarityValueLabel.Text = $similaritySlider.Value
    $similarityValueLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($similarityValueLabel)

    $similaritySlider.Add_Scroll({
        $similarityValueLabel.Text = $similaritySlider.Value
    })

    # Buffer size label (fixed at 16MB)
    $bufferLabel = New-Object System.Windows.Forms.Label
    $bufferLabel.Location = New-Object System.Drawing.Point(15, 335)
    $bufferLabel.Size = New-Object System.Drawing.Size(400, 20)
    $bufferLabel.Text = "File Buffer Size: 16MB (optimized for 20MB–100MB files)"
    $bufferLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($bufferLabel)

    # Batch size textbox
    $batchLabel = New-Object System.Windows.Forms.Label
    $batchLabel.Location = New-Object System.Drawing.Point(15, 360)
    $batchLabel.Size = New-Object System.Drawing.Size(200, 20)
    $batchLabel.Text = "Batch Size (MB, 10-1000):"
    $batchLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($batchLabel)

    $batchTextBox = New-Object System.Windows.Forms.TextBox
    $batchTextBox.Location = New-Object System.Drawing.Point(220, 360)
    $batchTextBox.Size = New-Object System.Drawing.Size(80, 20)
    $batchTextBox.Text = "300"
    $batchTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $batchTextBox.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($batchTextBox)

    # Progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(15, 390)
    $progressBar.Size = New-Object System.Drawing.Size(560, 20)
    $progressBar.Style = "Continuous"
    $form.Controls.Add($progressBar)

    # Batch progress label
    $batchProgressLabel = New-Object System.Windows.Forms.Label
    $batchProgressLabel.Location = New-Object System.Drawing.Point(15, 415)
    $batchProgressLabel.Size = New-Object System.Drawing.Size(560, 20)
    $batchProgressLabel.Text = "Batch Progress: Waiting to start"
    $batchProgressLabel.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($batchProgressLabel)

    # Output textbox
    $outputTextBox = New-Object System.Windows.Forms.TextBox
    $outputTextBox.Location = New-Object System.Drawing.Point(15, 440)
    $outputTextBox.Size = New-Object System.Drawing.Size(560, 150)
    $outputTextBox.Multiline = $true
    $outputTextBox.ScrollBars = "Vertical"
    $outputTextBox.ReadOnly = $true
    $outputTextBox.BackColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $outputTextBox.ForeColor = [System.Drawing.Color]::White
    $form.Controls.Add($outputTextBox)

    # Start button
    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Location = New-Object System.Drawing.Point(190, 600)
    $startButton.Size = New-Object System.Drawing.Size(100, 30)
    $startButton.Text = "Start"
    $startButton.BackColor = [System.Drawing.Color]::FromArgb(68, 68, 68)
    $startButton.ForeColor = [System.Drawing.Color]::White
    $startButton.Add_Click({
        # BREAKPOINT: Start of Start button click event
        $startButton.Enabled = $false
        try {
            $source = $sourceTextBox.Text
            $dest = $destTextBox.Text
            $backup = $backupTextBox.Text
            $replaceFile = $replaceTextBox.Text
            $addFile = $addTextBox.Text
            $useHighestQuality = $qualityCheckBox.Checked
            $checkDuplicateNames = $duplicateNameCheckBox.Checked
            $cleanSpaces = $spaceCleanupCheckBox.Checked
            $similarityThreshold = $similaritySlider.Value
            $bufferSizeBytes = 16MB # Fixed at 16MB, optimized for 20MB–100MB files
            $batchSizeMBText = $batchTextBox.Text
            $memoryThresholdBytes = 300MB # Lowered for stability with 20MB–100MB files

            # Validate batch size
            $batchSizeMB = 300
            try {
                $batchSizeMB = [double]$batchSizeMBText
                if ($batchSizeMB -lt 10 -or $batchSizeMB -gt 1000) {
                    Write-Log -Message "Error: Batch size must be between 10 and 1000 MB." -Status "Error"
                    return
                }
            } catch {
                Write-Log -Message "Error: Invalid batch size '$batchSizeMBText'. Using default 300 MB." -Status "Warning"
                $batchSizeMB = 300
            }
            $batchSizeBytes = [int64]($batchSizeMB * 1MB)

            $script:logEntries = @()
            if ([string]::IsNullOrWhiteSpace($source)) {
                Write-Log -Message "Error: Source directory is empty." -Status "Error"
                return
            }
            if (-not (Test-Path -LiteralPath $source)) {
                Write-Log -Message "Error: Source directory does not exist: $source" -Status "Error"
                return
            }
            if ([string]::IsNullOrWhiteSpace($dest)) {
                Write-Log -Message "Error: Destination directory is empty." -Status "Error"
                return
            }

            # BREAKPOINT: After input validation
            # Validate destination directory writability
            try {
                $testFile = Join-Path $dest "test_write_permissions.txt"
                [System.IO.File]::WriteAllText($testFile, "Test")
                Remove-Item -LiteralPath $testFile -Force -ErrorAction Stop
                Write-Log -Message "Destination directory is writable: $dest" -Status "Info"
            } catch {
                Write-Log -Message "Error: Destination directory is not writable: $dest. Details: $_" -Status "Error"
                return
            }

            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $global:summaryLogFilePath = Join-Path $dest "MusicFileOrganizer_Summary_Log_$timestamp.html"
            $batchLogFiles = @()
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
            Write-Log -Message "Similarity threshold set to: $similarityThreshold" -Status "Info"
            Write-Log -Message "Buffer size set to: 16 MB ($bufferSizeBytes bytes)" -Status "Info"
            Write-Log -Message "Batch size set to: $batchSizeMB MB ($batchSizeBytes bytes)" -Status "Info"
            Write-Log -Message "Memory threshold set to: 300 MB ($memoryThresholdBytes bytes)" -Status "Info"

            # Create backup if specified
            if (-not [string]::IsNullOrWhiteSpace($backup)) {
                $backupDir = Join-Path $backup "Backup_$timestamp"
                Create-Backup -Source $source -BackupDir $backupDir -BufferSizeBytes $bufferSizeBytes
            } else {
                Write-Log -Message "No backup directory specified. Skipping backup." -Status "Info"
            }

            # Load replacements (use default list if no file provided)
            $removeList = @("Official Video", "Official Audio", "Lyrics", "HD", "Remix")
            if ($replaceFile -and (Test-Path -LiteralPath $replaceFile)) {
                try {
                    $removeList = (Get-Content $replaceFile -Raw -ErrorAction Stop) -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    Write-Log -Message "Loaded replacements file with $($removeList.Count) entries: $($removeList -join ', ')" -Status "Success"
                } catch {
                    Write-Log -Message "Error reading replacements file: $_" -Status "Error"
                    Write-Log -Message "Using default replacements list: $($removeList -join ', ')" -Status "Info"
                }
            } else {
                Write-Log -Message "No replacements file selected. Using default replacements list: $($removeList -join ', ')" -Status "Info"
            }

            # Load add text (empty list if no file provided)
            $addList = @()
            if ($addFile -and (Test-Path -LiteralPath $addFile)) {
                try {
                    $addList = (Get-Content $addFile -Raw -ErrorAction Stop) -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    Write-Log -Message "Loaded add text file with $($addList.Count) entries: $($addList -join ', ')" -Status "Success"
                } catch {
                    Write-Log -Message "Error reading add text file: $_" -Status "Error"
                    Write-Log -Message "No text will be added to filenames." -Status "Info"
                }
            } else {
                Write-Log -Message "No add text file selected. No text will be added to filenames." -Status "Info"
            }

            # BREAKPOINT: Before batch creation
            # Collect and batch audio files incrementally
            $batches = New-Object System.Collections.Generic.List[Object]
            $batchIndex = 0
            Write-Log -Message "Starting incremental file enumeration" -Status "Info"
            $fileCount = Get-FilesIncrementally -Path $source -Include @("*.mp3", "*.wav", "*.flac", "*.aiff", "*.ogg", "*.wma") -BatchSizeBytes $batchSizeBytes -BatchCallback {
                param ($batch)
                $batchIndex++
                $batches.Add($batch)
                $batchSizeMB = [math]::Round(($batch | Measure-Object -Property Length -Sum).Sum / 1MB, 2)
                Write-Log -Message "Created batch $batchIndex with $($batch.Count) files ($batchSizeMB MB)" -Status "Info"
                # BREAKPOINT: Inside incremental file enumeration
            }

            if ($fileCount -eq 0) {
                Write-Log -Message "No audio files found in source directory: $source" -Status "Error"
                return
            }
            Write-Log -Message "Found $fileCount audio files in source directory, split into $batchIndex batches" -Status "Info"

            # BREAKPOINT: Start of batch processing loop
            # Process each batch
            for ($batchIndex = 0; $batchIndex -lt $batches.Count; $batchIndex++) {
                $batch = $batches[$batchIndex]
                $batchNumber = $batchIndex + 1
                $batchSizeMB = [math]::Round(($batch | Measure-Object -Property Length -Sum).Sum / 1MB, 2)
                $batchProgressLabel.Text = "Processing batch $batchNumber of $($batches.Count) ($batchSizeMB MB)"
                Write-Log -Message "Processing batch $batchNumber of $($batches.Count) with $($batch.Count) files ($batchSizeMB MB)" -Status "Info"
                $script:logEntries = @() # Clear log entries for this batch

                # Check memory usage
                $process = [System.Diagnostics.Process]::GetCurrentProcess()
                $memoryUsage = $process.WorkingSet64
                if ($memoryUsage -gt $memoryThresholdBytes) {
                    Write-Log -Message "Memory usage ($([math]::Round($memoryUsage / 1MB, 2)) MB) exceeds threshold ($([math]::Round($memoryThresholdBytes / 1MB, 2)) MB). Triggering garbage collection." -Status "Warning"
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    Start-Sleep -Milliseconds 100 # Allow memory stabilization
                }
                Write-Log -Message "Current memory usage before batch ${batchNumber}: $([math]::Round($memoryUsage / 1MB, 2)) MB" -Status "Info"

                # BREAKPOINT: Start of file entry creation loop
                # Create file entries for the batch
                $fileEntries = @()
                $currentFile = 0
                foreach ($file in $batch) {
                    $currentFile++
                    $progressBar.Value = [Math]::Min(100, ($currentFile / $batch.Count) * 100)
                    if (-not (Test-FileAccessible -Path $file.FullName)) {
                        continue
                    }
                    $filename = [System.IO.Path]::GetFileNameWithoutExtension($file.FullName)
                    if ($cleanSpaces) {
                        $filename = Clean-FilenameSpaces -Filename $filename
                    }
                    foreach ($term in $removeList) {
                        $filename = $filename -replace [regex]::Escape($term), ""
                    }
                    $filename = $filename -replace '\s+', ' '
                    $filename = $filename.Trim()
                    if ($addList) {
                        $filename = "$filename $($addList -join ' ')"
                    }
                    $sanitizedFilename = Sanitize-Filename -Filename $filename
                    $entry = [PSCustomObject]@{
                        OriginalPath     = $file.FullName
                        Filename         = $filename
                        SanitizedFilename = $sanitizedFilename
                        Extension        = $file.Extension
                        Size             = $file.Length
                        Bitrate          = if ($useHighestQuality) { Get-FileBitrate -FilePath $file.FullName } else { 0 }
                        FormatPriority   = Get-FormatPriority -Extension $file.Extension
                    }
                    $fileEntries += $entry
                }

                if ($fileEntries.Count -eq 0) {
                    Write-Log -Message "No valid files in batch $batchNumber after filtering." -Status "Warning"
                    continue
                }

                # BREAKPOINT: Before similarity group creation
                # Create similarity groups
                $similarityGroups = @()
                if ($checkDuplicateNames) {
                    Write-Log -Message "Checking for duplicate songs by name in batch $batchNumber with similarity threshold $similarityThreshold" -Status "Info"
                    $processed = @()
                    $threshold = $similarityThreshold / 100.0
                    $fileEntriesCount = $fileEntries.Count
                    for ($i = 0; $i -lt $fileEntriesCount; $i++) {
                        if ($i -in $processed) { continue }
                        $group = @($fileEntries[$i])
                        $processed += $i
                        # BREAKPOINT: Inside similarity group creation loop
                        for ($j = $i + 1; $j -lt $fileEntriesCount; $j++) {
                            if ($j -in $processed) { continue }
                            $similarity = Get-StringSimilarity -String1 $fileEntries[$i].SanitizedFilename -String2 $fileEntries[$j].SanitizedFilename -SliderValue $similarityThreshold -RemoveList $removeList
                            if ($similarity -ge $threshold) {
                                $group += $fileEntries[$j]
                                $processed += $j
                            }
                        }
                        if ($group.Count -gt 1) {
                            $similarityGroups += ,$group
                        }
                    }
                    Write-Log -Message "Found $($similarityGroups.Count) similarity groups in batch $batchNumber" -Status "Info"
                }

                # BREAKPOINT: Before processing similarity groups
                # Process similarity groups
                $fileGroups = @{}
                if ($checkDuplicateNames -and $similarityGroups) {
                    foreach ($group in $similarityGroups) {
                        $representativeName = ($group | Sort-Object { $_.SanitizedFilename.Length } | Select-Object -First 1).SanitizedFilename
                        # BREAKPOINT: Inside similarity group processing loop
                        foreach ($file in $group) {
                            if (-not $fileGroups.ContainsKey($representativeName)) {
                                $fileGroups[$representativeName] = @()
                            }
                            $fileGroups[$representativeName] += $file
                        }
                    }
                } else {
                    foreach ($file in $fileEntries) {
                        $fileGroups[$file.SanitizedFilename] = @($file)
                    }
                }

                # BREAKPOINT: Before processing file groups
                # Process each file group
                $currentGroup = 0
                $totalGroups = $fileGroups.Keys.Count
                foreach ($groupName in $fileGroups.Keys) {
                    $currentGroup++
                    $progressBar.Value = [Math]::Min(100, ($currentGroup / $totalGroups) * 100)
                    $groupFiles = $fileGroups[$groupName]
                    $destFolder = Join-Path $dest $groupName
                    $destFolder = Get-UniqueFolderPath -FolderPath $destFolder
                    # BREAKPOINT: Inside file group processing loop
                    try {
                        if (-not (Test-Path -LiteralPath $destFolder)) {
                            New-Item -ItemType Directory -Path $destFolder -ErrorAction Stop | Out-Null
                        }
                        if ($useHighestQuality -and $groupFiles.Count -gt 1) {
                            # BREAKPOINT: Before best file selection
                            $bestFile = $groupFiles | Sort-Object -Property @{Expression={$_.FormatPriority}; Descending=$true}, @{Expression={$_.Bitrate}; Descending=$true}, @{Expression={$_.Size}; Descending=$true} | Select-Object -First 1
                            $destFileName = "$($bestFile.SanitizedFilename)_BEST$($bestFile.Extension)"
                            $destPath = Join-Path $destFolder $destFileName
                            if (Copy-LargeFile -SourcePath $bestFile.OriginalPath -DestPath $destPath -BufferSizeBytes $bufferSizeBytes) {
                                Write-Log -Message "Moved highest quality file (FormatPriority: $($bestFile.FormatPriority), Bitrate: $($bestFile.Bitrate), Size: $($bestFile.Size)): $($bestFile.OriginalPath) to $destPath" -Status "Success"
                            }
                            foreach ($file in $groupFiles) {
                                if ($file.OriginalPath -ne $bestFile.OriginalPath) {
                                    Write-Log -Message "Skipped lower quality duplicate (FormatPriority: $($file.FormatPriority), Bitrate: $($file.Bitrate), Size: $($file.Size)): $($file.OriginalPath)" -Status "Info"
                                }
                            }
                        } else {
                            foreach ($file in $groupFiles) {
                                $destFileName = "$($file.SanitizedFilename)$($file.Extension)"
                                $destPath = Join-Path $destFolder $destFileName
                                if (Copy-LargeFile -SourcePath $file.OriginalPath -DestPath $destPath -BufferSizeBytes $bufferSizeBytes) {
                                    Write-Log -Message "Moved: $($file.OriginalPath) to $destPath" -Status "Success"
                                }
                            }
                        }
                    } catch {
                        Write-Log -Message "Error processing group '$groupName': $_" -Status "Error"
                    }
                }

                # Log memory usage after batch
                $memoryUsage = [System.Diagnostics.Process]::GetCurrentProcess().WorkingSet64
                Write-Log -Message "Memory usage after batch ${batchNumber}: $([math]::Round($memoryUsage / 1MB, 2)) MB" -Status "Info"

                # Generate batch log file
                $batchLogFile = Join-Path $dest "MusicFileOrganizer_Log_${timestamp}_batch${batchNumber}.html"
                if (Generate-HtmlLog -LogFilePath $batchLogFile -BatchNumber $batchNumber) {
                    Write-Log -Message "Generated batch log file: $batchLogFile" -Status "Success"
                    $batchLogFiles += $batchLogFile
                }

                # Clear memory
                $fileEntries = $null
                $similarityGroups = $null
                $fileGroups = $null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }

            # Generate summary log file
            if ($batchLogFiles) {
                if (Generate-SummaryLog -SummaryLogFilePath $global:summaryLogFilePath -BatchLogFiles $batchLogFiles) {
                    Write-Log -Message "Generated summary log file: $global:summaryLogFilePath" -Status "Success"
                }
            }

            $batchProgressLabel.Text = "Processing complete: $($batches.Count) batches processed"
            Write-Log -Message "Processing complete: $($batches.Count) batches processed" -Status "Success"
            [System.Windows.Forms.MessageBox]::Show("Processing complete! Summary log: $global:summaryLogFilePath", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            Write-Log -Message "Unexpected error: $_" -Status "Error"
            [System.Windows.Forms.MessageBox]::Show("An unexpected error occurred: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $startButton.Enabled = $true
            $progressBar.Value = 0
            $batchProgressLabel.Text = "Batch Progress: Waiting to start"
        }
    })
    $form.Controls.Add($startButton)

    # Show the form
    [System.Windows.Forms.Application]::Run($form)
} catch {
    Write-Log -Message "Critical error in main execution: $_" -Status "Error"
    [System.Windows.Forms.MessageBox]::Show("A critical error occurred: $_", "Critical Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}