# Load required .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Verify PowerShell version
$RequiredVersion = [System.Version]"5.0"
if ($PSVersionTable.PSVersion -lt $RequiredVersion) {
    [System.Windows.Forms.MessageBox]::Show(
        "PowerShell version $($PSVersionTable.PSVersion) detected. Requires PowerShell 5.0 or higher.",
        "Version Compatibility Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Auto-install required modules (ps2exe)
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    try {
        Install-Module -Name ps2exe -Force -Scope CurrentUser -ErrorAction Stop
        Import-Module ps2exe
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to install or import 'ps2exe' module!",
            "Module Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "PS1 to EXE Compiler"
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"

# Input file selection
$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Select .ps1 File:"
$lblInput.Location = New-Object System.Drawing.Point(10, 20)
$lblInput.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($lblInput)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(120, 20)
$txtInput.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($txtInput)

$btnBrowseInput = New-Object System.Windows.Forms.Button
$btnBrowseInput.Text = "Browse"
$btnBrowseInput.Location = New-Object System.Drawing.Point(480, 18)
$btnBrowseInput.Size = New-Object System.Drawing.Size(75, 23)
$form.Controls.Add($btnBrowseInput)

# Output folder selection
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Select Output Folder:"
$lblOutput.Location = New-Object System.Drawing.Point(10, 60)
$lblOutput.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(140, 60)
$txtOutput.Size = New-Object System.Drawing.Size(330, 20)
$form.Controls.Add($txtOutput)

$btnBrowseOutput = New-Object System.Windows.Forms.Button
$btnBrowseOutput.Text = "Browse"
$btnBrowseOutput.Location = New-Object System.Drawing.Point(480, 58)
$btnBrowseOutput.Size = New-Object System.Drawing.Size(75, 23)
$form.Controls.Add($btnBrowseOutput)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 100)
$progressBar.Size = New-Object System.Drawing.Size(545, 20)
$form.Controls.Add($progressBar)

# Compile button
$btnCompile = New-Object System.Windows.Forms.Button
$btnCompile.Text = "Compile"
$btnCompile.Location = New-Object System.Drawing.Point(240, 130)
$btnCompile.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnCompile)

# Log textbox
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 170)
$txtLog.Size = New-Object System.Drawing.Size(545, 180)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$form.Controls.Add($txtLog)

# Browse input file
$btnBrowseInput.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $txtInput.Text = $openFileDialog.FileName
    }
})

# Browse output folder
$btnBrowseOutput.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $txtOutput.Text = $folderBrowser.SelectedPath
    }
})

# Compile button click event
$btnCompile.Add_Click({
    $ps1File = $txtInput.Text
    $outputFolder = $txtOutput.Text

    try {
        # Validate the input file
        if (-not (Test-Path $ps1File -PathType Leaf)) {
            throw "Invalid or missing .ps1 file!"
        }
        # Validate the output folder
        if (-not (Test-Path $outputFolder -PathType Container)) {
            throw "Invalid output folder!"
        }

        $exeFile = Join-Path $outputFolder ([System.IO.Path]::GetFileNameWithoutExtension($ps1File) + ".exe")
        $logFile = Join-Path $outputFolder "compile_log.txt"

        $txtLog.Clear()
        $progressBar.Value = 10
        $txtLog.AppendText("Starting compilation...`r`n")

        # Ensure ps2exe module is installed
        if (-not (Get-Module -ListAvailable -Name ps2exe)) {
            Try {
                Install-Module -Name ps2exe -Force -Scope CurrentUser -ErrorAction Stop
                Import-Module ps2exe
            } Catch {
                throw "Failed to install or import 'ps2exe' module: $_"
            }
        }

        # Form command
        $compileCmd = "Invoke-PS2EXE -inputFile '$ps1File' -outputFile '$exeFile' -noConsole -requireAdmin"

        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = "powershell.exe"
        $startInfo.Arguments = "-NoProfile -Command `"& { $compileCmd }`""
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $startInfo
        $proc.Start() | Out-Null

        if (-not $proc.WaitForExit(30000)) {  # Timeout: 30 seconds
            $proc.Kill()
            throw "Compilation process took too long and was terminated!"
        }

        $stdOut = $proc.StandardOutput.ReadToEnd()
        $stdErr = $proc.StandardError.ReadToEnd()

        $progressBar.Value = 80

        if (Test-Path $exeFile) {
            $txtLog.AppendText("Compilation successful! EXE created:`r`n$exeFile`r`n")
            $progressBar.Value = 100
        } else {
            throw "Compilation failed. No EXE generated!"
        }

        # Save log
        $logContent = @"
--- Compile Command ---
$compileCmd

--- STDOUT ---
$stdOut

--- STDERR ---
$stdErr

--- Time ---
$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@
        $logContent | Out-File -FilePath $logFile -Encoding UTF8
        $txtLog.AppendText("Log written to: $logFile`r`n")
        Start-Process notepad.exe $logFile
    }
    catch {
        $txtLog.AppendText("ERROR: $_`r`n")
        [System.Windows.Forms.MessageBox]::Show($_, "Compilation Error", 
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $progressBar.Value = 0
    }
})

# Show the form
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
