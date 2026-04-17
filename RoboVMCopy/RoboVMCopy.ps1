#Requires -Version 5.1
<#
.SYNOPSIS
    RoboVMCopy - GUI front-end for RoboCopy on Windows.

.DESCRIPTION
    - Browse source and destination folders
    - Configure common RoboCopy flags
    - Run and stop RoboCopy with live output log
    - Save copy jobs as quick-launch buttons
    - Left click a saved button to run it
    - Right click a saved button to load or delete it
#>

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw "Run in STA mode: powershell -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
[System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)

# Support both file execution and interactive terminal execution.
$script:BasePath = if ($PSScriptRoot) {
    $PSScriptRoot
}
elseif ($PSCommandPath) {
    Split-Path -Path $PSCommandPath -Parent
}
else {
    (Get-Location).Path
}

$script:ConfigFile = Join-Path $script:BasePath 'RoboVMCopy_Jobs.json'
$script:RoboProcess = $null
$script:OutputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$script:PollTimer = $null
$script:RunLogWriter = $null
$script:RunLogPath = $null
$script:AllowClose = $false
$script:LastFileActivity = ''
$script:CopyStartTime = $null
$script:SourceTotalBytes = 0L
$script:ProcessedBytes = 0L
$script:CopiedBytes = 0L
$script:RawOutputPath = $null
$script:RawOutputLineCount = 0
$script:CrashLogPath = Join-Path $script:BasePath 'RoboVMCopy_Crash.log'

function Write-CrashLog {
    param([string]$Message)

    try {
        $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
        Add-Content -LiteralPath $script:CrashLogPath -Value $line -Encoding UTF8
    }
    catch {
    }
}

[System.Windows.Forms.Application]::add_ThreadException({
    param($sender, $e)
    $msg = if ($e -and $e.Exception) { $e.Exception.ToString() } else { 'Unknown UI thread exception.' }
    Write-CrashLog ("UI Thread Exception: {0}" -f $msg)
    [System.Windows.Forms.MessageBox]::Show(
        "Unexpected UI error captured. Details written to:`r`n$script:CrashLogPath`r`n`r`n$msg",
        'RoboVMCopy Fatal Error',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
})

[System.AppDomain]::CurrentDomain.add_UnhandledException({
    param($sender, $e)
    $msg = if ($e -and $e.ExceptionObject) { $e.ExceptionObject.ToString() } else { 'Unknown unhandled exception.' }
    Write-CrashLog ("Unhandled Exception: {0}" -f $msg)
})

function Close-RunLog {
    if ($script:RunLogWriter) {
        try {
            $script:RunLogWriter.Flush()
            $script:RunLogWriter.Dispose()
        }
        catch {
        }
        finally {
            $script:RunLogWriter = $null
        }
    }
}

function Start-RunLog {
    Close-RunLog

    try {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $script:RunLogPath = Join-Path $script:BasePath ("RoboCopy_{0}.log" -f $stamp)
        $script:RunLogWriter = New-Object System.IO.StreamWriter($script:RunLogPath, $false, [System.Text.Encoding]::UTF8)
        $script:RunLogWriter.AutoFlush = $true
    }
    catch {
        $script:RunLogPath = $null
        $script:RunLogWriter = $null
    }
}

function Reset-RawOutputTracking {
    $script:RawOutputPath = $null
    $script:RawOutputLineCount = 0
}

function Get-NewRawOutputLines {
    if (-not $script:RawOutputPath -or -not (Test-Path -LiteralPath $script:RawOutputPath)) {
        return @()
    }

    try {
        $lines = @(Get-Content -LiteralPath $script:RawOutputPath -ErrorAction SilentlyContinue)
        if ($lines.Count -le $script:RawOutputLineCount) {
            return @()
        }

        $newLines = @($lines[$script:RawOutputLineCount..($lines.Count - 1)])
        $script:RawOutputLineCount = $lines.Count
        return $newLines
    }
    catch {
        return @()
    }
}

function Normalize-Jobs {
    param([object[]]$Jobs)

    $cleanJobs = [System.Collections.Generic.List[object]]::new()

    foreach ($job in @($Jobs)) {
        if ($null -eq $job) {
            continue
        }

        $name = [string]$job.Name
        $source = [string]$job.Source
        $destination = [string]$job.Destination

        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $cleanJobs.Add([PSCustomObject]@{
            Name = $name.Trim()
            Source = $source.Trim()
            Destination = $destination.Trim()
            Flags = (Normalize-FlagString -Flags ([string]$job.Flags))
        })
    }

    return @($cleanJobs)
}

function Normalize-FlagString {
    param([string]$Flags)

    if (-not $Flags) {
        return ''
    }

    $tokens = $Flags -split '\s+' | Where-Object {
        $_ -and $_ -notin @('/NFL', '/NDL')
    }

    return ($tokens -join ' ').Trim()
}

function Resolve-RoboCopyPermissionFlags {
    param([string]$Flags)

    if (-not $Flags) {
        return ''
    }

    $hasAuditingCopy = ($Flags -match '(?i)/COPYALL') -or ($Flags -match '(?i)/COPY:[A-Z]*U[A-Z]*')
    if (-not $hasAuditingCopy) {
        return $Flags
    }

    $choice = [System.Windows.Forms.MessageBox]::Show(
        "Selected options include auditing copy (/COPYALL or /COPY:...U).`r`n`r`nWithout the Manage Auditing user right, RoboCopy returns exit code 16.`r`n`r`nUse /COPY:DATS for this run instead?",
        'RoboCopy Permissions',
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Log 'Run cancelled by user before launch.' ([System.Drawing.Color]::FromArgb(255, 165, 0))
        return $null
    }

    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        $fixed = [System.Text.RegularExpressions.Regex]::Replace($Flags, '(?i)/COPYALL', '/COPY:DATS')
        $fixed = [System.Text.RegularExpressions.Regex]::Replace($fixed, '(?i)/COPY:[A-Z]*U[A-Z]*', '/COPY:DATS')
        Write-Log 'Using /COPY:DATS for this run to avoid auditing-right failures.' ([System.Drawing.Color]::FromArgb(255, 200, 120))
        return $fixed
    }

    Write-Log 'Continuing with auditing copy flags; run may fail without Manage Auditing user right.' ([System.Drawing.Color]::FromArgb(255, 165, 0))
    return $Flags
}

function Get-Jobs {
    if (Test-Path -LiteralPath $script:ConfigFile) {
        try {
            return @(Normalize-Jobs -Jobs @(Get-Content -LiteralPath $script:ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json))
        }
        catch {
            return @()
        }
    }
    return @()
}

function Save-AllJobs {
    param([object[]]$Jobs)

    $Jobs = @(Normalize-Jobs -Jobs $Jobs)

    if ($null -eq $Jobs -or $Jobs.Count -eq 0) {
        Set-Content -LiteralPath $script:ConfigFile -Value '[]' -Encoding UTF8
        return
    }

    $Jobs | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:ConfigFile -Encoding UTF8
}

function New-FlatButton {
    param(
        [string]$Text,
        [System.Windows.Forms.Control]$Parent,
        [int]$X,
        [int]$Y,
        [int]$W,
        [int]$H,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = [System.Drawing.Point]::new($X, $Y)
    $button.Size = [System.Drawing.Size]::new($W, $H)
    $button.FlatStyle = 'Flat'
    $button.FlatAppearance.BorderColor = $BackColor
    $button.BackColor = $BackColor
    $button.ForeColor = $ForeColor
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Parent.Controls.Add($button)
    return $button
}

function New-SectionHeader {
    param(
        [string]$Text,
        [System.Windows.Forms.Control]$Parent,
        [System.Drawing.Color]$ForeColor
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = [System.Drawing.Point]::new(10, 8)
    $label.AutoSize = $true
    $label.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = $ForeColor
    $Parent.Controls.Add($label)
}

function Select-FolderPath {
    param(
        [string]$InitialPath,
        [string]$Title
    )

    # Use OpenFileDialog in folder-select mode to better expose mapped drives and This PC.
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = $Title
    $dlg.Filter = 'Folders|*.folder'
    $dlg.CheckFileExists = $false
    $dlg.CheckPathExists = $true
    $dlg.ValidateNames = $false
    $dlg.FileName = 'Select Folder'
    $dlg.DereferenceLinks = $true

    if ($InitialPath -and (Test-Path -LiteralPath $InitialPath)) {
        $dlg.InitialDirectory = $InitialPath
    }

    try {
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $candidate = Split-Path -Path $dlg.FileName -Parent
            if (-not $candidate) {
                $candidate = $dlg.FileName
            }
            if (Test-Path -LiteralPath $candidate) {
                return $candidate
            }
        }
    }
    finally {
        $dlg.Dispose()
    }

    return $null
}

# Colors (ASCII-safe only)
$clrBg     = [System.Drawing.Color]::FromArgb(28, 28, 30)
$clrPanel  = [System.Drawing.Color]::FromArgb(38, 38, 42)
$clrBorder = [System.Drawing.Color]::FromArgb(60, 60, 65)
$clrBlue   = [System.Drawing.Color]::FromArgb(0, 122, 204)
$clrGreen  = [System.Drawing.Color]::FromArgb(35, 160, 65)
$clrRed    = [System.Drawing.Color]::FromArgb(200, 50, 50)
$clrOrange = [System.Drawing.Color]::FromArgb(190, 90, 0)
$clrText   = [System.Drawing.Color]::FromArgb(220, 220, 220)
$clrMuted  = [System.Drawing.Color]::FromArgb(140, 140, 145)
$clrInput  = [System.Drawing.Color]::FromArgb(45, 45, 48)
$clrAlt    = [System.Drawing.Color]::FromArgb(55, 55, 60)
$clrJobBtn = [System.Drawing.Color]::FromArgb(0, 85, 148)
$clrLogBg  = [System.Drawing.Color]::FromArgb(12, 12, 14)

# Main form
$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = 'RoboVMCopy - RoboCopy GUI'
$script:form.Size = [System.Drawing.Size]::new(980, 800)
$script:form.MinimumSize = [System.Drawing.Size]::new(820, 680)
$script:form.StartPosition = 'CenterScreen'
$script:form.BackColor = $clrBg
$script:form.ForeColor = $clrText
$script:form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# Paths panel
$panPaths = New-Object System.Windows.Forms.Panel
$panPaths.Location = [System.Drawing.Point]::new(10, 10)
$panPaths.Size = [System.Drawing.Size]::new(940, 108)
$panPaths.BackColor = $clrPanel
$panPaths.BorderStyle = 'FixedSingle'
$panPaths.Anchor = 'Top,Left,Right'
$script:form.Controls.Add($panPaths)
New-SectionHeader -Text 'PATHS' -Parent $panPaths -ForeColor $clrBlue

$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Text = 'Source:'
$lblSrc.Location = [System.Drawing.Point]::new(10, 36)
$lblSrc.Size = [System.Drawing.Size]::new(60, 22)
$lblSrc.ForeColor = $clrText
$panPaths.Controls.Add($lblSrc)

$script:tbSource = New-Object System.Windows.Forms.TextBox
$script:tbSource.Location = [System.Drawing.Point]::new(74, 34)
$script:tbSource.Size = [System.Drawing.Size]::new(768, 24)
$script:tbSource.BackColor = $clrInput
$script:tbSource.ForeColor = $clrText
$script:tbSource.BorderStyle = 'FixedSingle'
$script:tbSource.ReadOnly = $true
$script:tbSource.Anchor = 'Top,Left,Right'
$panPaths.Controls.Add($script:tbSource)

$btnBrowseSrc = New-FlatButton -Text 'Browse...' -Parent $panPaths -X 848 -Y 32 -W 76 -H 26 -BackColor $clrAlt -ForeColor $clrText
$btnBrowseSrc.FlatAppearance.BorderColor = $clrBorder
$btnBrowseSrc.Anchor = 'Top,Right'

$lblDst = New-Object System.Windows.Forms.Label
$lblDst.Text = 'Destination:'
$lblDst.Location = [System.Drawing.Point]::new(10, 72)
$lblDst.Size = [System.Drawing.Size]::new(74, 22)
$lblDst.ForeColor = $clrText
$panPaths.Controls.Add($lblDst)

$script:tbDest = New-Object System.Windows.Forms.TextBox
$script:tbDest.Location = [System.Drawing.Point]::new(88, 70)
$script:tbDest.Size = [System.Drawing.Size]::new(754, 24)
$script:tbDest.BackColor = $clrInput
$script:tbDest.ForeColor = $clrText
$script:tbDest.BorderStyle = 'FixedSingle'
$script:tbDest.ReadOnly = $true
$script:tbDest.Anchor = 'Top,Left,Right'
$panPaths.Controls.Add($script:tbDest)

$btnBrowseDst = New-FlatButton -Text 'Browse...' -Parent $panPaths -X 848 -Y 68 -W 76 -H 26 -BackColor $clrAlt -ForeColor $clrText
$btnBrowseDst.FlatAppearance.BorderColor = $clrBorder
$btnBrowseDst.Anchor = 'Top,Right'

# Options panel
$panOpts = New-Object System.Windows.Forms.Panel
$panOpts.Location = [System.Drawing.Point]::new(10, 128)
$panOpts.Size = [System.Drawing.Size]::new(940, 118)
$panOpts.BackColor = $clrPanel
$panOpts.BorderStyle = 'FixedSingle'
$panOpts.Anchor = 'Top,Left,Right'
$script:form.Controls.Add($panOpts)
New-SectionHeader -Text 'ROBOCOPY OPTIONS' -Parent $panOpts -ForeColor $clrBlue

$flagDefs = @(
    @{ F = '/MIR';     D = 'Mirror (sync + purge extras)' }
    @{ F = '/E';       D = 'Include empty subdirectories' }
    @{ F = '/COPYALL'; D = 'Copy all file attributes/info' }
    @{ F = '/Z';       D = 'Restartable mode' }
    @{ F = '/XA:H';    D = 'Exclude hidden files' }
    @{ F = '/PURGE';   D = 'Delete destination files not in source' }
)

$script:checkboxes = @{}
$colW = 228
$col = 0
$row = 0
foreach ($fd in $flagDefs) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = "$($fd.F) - $($fd.D)"
    $cb.Location = [System.Drawing.Point]::new(10 + $col * $colW, 30 + $row * 26)
    $cb.Size = [System.Drawing.Size]::new(220, 22)
    $cb.BackColor = $clrPanel
    $cb.ForeColor = $clrText
    $cb.Cursor = [System.Windows.Forms.Cursors]::Hand
    $panOpts.Controls.Add($cb)
    $script:checkboxes[$fd.F] = $cb
    $col++
    if ($col -ge 4) {
        $col = 0
        $row++
    }
}

function Add-Spinner {
    param(
        [string]$Label,
        [int]$X,
        [int]$Min,
        [int]$Max,
        [int]$Default,
        [System.Windows.Forms.Control]$Parent,
        [int]$Y,
        [System.Drawing.Color]$LabelColor,
        [System.Drawing.Color]$InputBack,
        [System.Drawing.Color]$InputFore
    )

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.Location = [System.Drawing.Point]::new($X, $Y)
    $lbl.Size = [System.Drawing.Size]::new(96, 22)
    $lbl.ForeColor = $LabelColor
    $Parent.Controls.Add($lbl)

    $nud = New-Object System.Windows.Forms.NumericUpDown
    $nud.Location = [System.Drawing.Point]::new($X + 100, $Y - 2)
    $nud.Size = [System.Drawing.Size]::new(54, 24)
    $nud.Minimum = $Min
    $nud.Maximum = $Max
    $nud.Value = $Default
    $nud.BackColor = $InputBack
    $nud.ForeColor = $InputFore
    $Parent.Controls.Add($nud)
    return $nud
}

$nudY = 84
$script:nudMT = Add-Spinner -Label '/MT (threads):' -X 10 -Min 1 -Max 128 -Default 8 -Parent $panOpts -Y $nudY -LabelColor $clrText -InputBack $clrInput -InputFore $clrText
$script:nudR  = Add-Spinner -Label '/R (retries):' -X 175 -Min 0 -Max 99 -Default 3 -Parent $panOpts -Y $nudY -LabelColor $clrText -InputBack $clrInput -InputFore $clrText
$script:nudW  = Add-Spinner -Label '/W (wait sec):' -X 340 -Min 0 -Max 300 -Default 5 -Parent $panOpts -Y $nudY -LabelColor $clrText -InputBack $clrInput -InputFore $clrText

$lblExtra = New-Object System.Windows.Forms.Label
$lblExtra.Text = 'Extra flags:'
$lblExtra.Location = [System.Drawing.Point]::new(510, $nudY)
$lblExtra.Size = [System.Drawing.Size]::new(76, 22)
$lblExtra.ForeColor = $clrText
$panOpts.Controls.Add($lblExtra)

$script:tbExtra = New-Object System.Windows.Forms.TextBox
$script:tbExtra.Location = [System.Drawing.Point]::new(590, $nudY - 2)
$script:tbExtra.Size = [System.Drawing.Size]::new(336, 24)
$script:tbExtra.BackColor = $clrInput
$script:tbExtra.ForeColor = $clrText
$script:tbExtra.BorderStyle = 'FixedSingle'
$script:tbExtra.Anchor = 'Top,Left,Right'
$panOpts.Controls.Add($script:tbExtra)

# Action panel
$panAction = New-Object System.Windows.Forms.Panel
$panAction.Location = [System.Drawing.Point]::new(10, 256)
$panAction.Size = [System.Drawing.Size]::new(940, 64)
$panAction.BackColor = $clrBg
$panAction.Anchor = 'Top,Left,Right'
$script:form.Controls.Add($panAction)

$script:btnRun = New-FlatButton -Text 'Run RoboCopy' -Parent $panAction -X 0 -Y 2 -W 164 -H 38 -BackColor $clrGreen -ForeColor ([System.Drawing.Color]::White)
$script:btnRun.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$script:btnStop = New-FlatButton -Text 'Stop' -Parent $panAction -X 172 -Y 2 -W 84 -H 38 -BackColor $clrRed -ForeColor ([System.Drawing.Color]::White)
$script:btnStop.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$script:btnStop.Enabled = $false

$script:btnExit = New-FlatButton -Text 'Exit' -Parent $panAction -X 844 -Y 5 -W 82 -H 34 -BackColor $clrAlt -ForeColor ([System.Drawing.Color]::White)
$script:btnExit.Anchor = 'Top,Right'

$lblSaveAs = New-Object System.Windows.Forms.Label
$lblSaveAs.Text = 'Save as button:'
$lblSaveAs.Location = [System.Drawing.Point]::new(274, 11)
$lblSaveAs.Size = [System.Drawing.Size]::new(106, 22)
$lblSaveAs.ForeColor = $clrMuted
$panAction.Controls.Add($lblSaveAs)

$script:tbJobName = New-Object System.Windows.Forms.TextBox
$script:tbJobName.Location = [System.Drawing.Point]::new(383, 9)
$script:tbJobName.Size = [System.Drawing.Size]::new(374, 24)
$script:tbJobName.BackColor = $clrInput
$script:tbJobName.ForeColor = $clrText
$script:tbJobName.BorderStyle = 'FixedSingle'
$script:tbJobName.Anchor = 'Top,Left,Right'
$panAction.Controls.Add($script:tbJobName)

$lblExitHint = New-Object System.Windows.Forms.Label
$lblExitHint.Text = 'The app stays open until you click Exit.'
$lblExitHint.Location = [System.Drawing.Point]::new(274, 38)
$lblExitHint.Size = [System.Drawing.Size]::new(250, 18)
$lblExitHint.ForeColor = $clrMuted
$panAction.Controls.Add($lblExitHint)

$btnSaveJob = New-FlatButton -Text 'Save Job' -Parent $panAction -X 678 -Y 5 -W 160 -H 34 -BackColor $clrBlue -ForeColor ([System.Drawing.Color]::White)
$btnSaveJob.Anchor = 'Top,Right'

# Jobs panel
$lblJobsHdr = New-Object System.Windows.Forms.Label
$lblJobsHdr.Text = 'SAVED JOBS (left click = run | right click = load/edit/run/delete)'
$lblJobsHdr.Location = [System.Drawing.Point]::new(10, 328)
$lblJobsHdr.AutoSize = $true
$lblJobsHdr.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$lblJobsHdr.ForeColor = $clrBlue
$script:form.Controls.Add($lblJobsHdr)

$script:panJobs = New-Object System.Windows.Forms.FlowLayoutPanel
$script:panJobs.Location = [System.Drawing.Point]::new(10, 348)
$script:panJobs.Size = [System.Drawing.Size]::new(940, 88)
$script:panJobs.BackColor = $clrPanel
$script:panJobs.AutoScroll = $true
$script:panJobs.FlowDirection = 'LeftToRight'
$script:panJobs.WrapContents = $true
$script:panJobs.BorderStyle = 'FixedSingle'
$script:panJobs.Anchor = 'Top,Left,Right'
$script:form.Controls.Add($script:panJobs)

# Log panel
$script:panCurrentFile = New-Object System.Windows.Forms.Panel
$script:panCurrentFile.Location = [System.Drawing.Point]::new(10, 444)
$script:panCurrentFile.Size = [System.Drawing.Size]::new(940, 42)
$script:panCurrentFile.BackColor = $clrPanel
$script:panCurrentFile.BorderStyle = 'FixedSingle'
$script:panCurrentFile.Anchor = 'Top,Left,Right'
$script:form.Controls.Add($script:panCurrentFile)

$script:lblCurrentFileHdr = New-Object System.Windows.Forms.Label
$script:lblCurrentFileHdr.Text = 'CURRENT FILE'
$script:lblCurrentFileHdr.Location = [System.Drawing.Point]::new(10, 11)
$script:lblCurrentFileHdr.AutoSize = $true
$script:lblCurrentFileHdr.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$script:lblCurrentFileHdr.ForeColor = $clrBlue
$script:panCurrentFile.Controls.Add($script:lblCurrentFileHdr)

$script:lblCurrentFileValue = New-Object System.Windows.Forms.Label
$script:lblCurrentFileValue.Text = 'Idle'
$script:lblCurrentFileValue.Location = [System.Drawing.Point]::new(118, 11)
$script:lblCurrentFileValue.Size = [System.Drawing.Size]::new(806, 18)
$script:lblCurrentFileValue.ForeColor = $clrText
$script:lblCurrentFileValue.AutoEllipsis = $true
$script:panCurrentFile.Controls.Add($script:lblCurrentFileValue)

$script:panProgress = New-Object System.Windows.Forms.Panel
$script:panProgress.Location = [System.Drawing.Point]::new(10, 492)
$script:panProgress.Size = [System.Drawing.Size]::new(940, 42)
$script:panProgress.BackColor = $clrPanel
$script:panProgress.BorderStyle = 'FixedSingle'
$script:panProgress.Anchor = 'Top,Left,Right'
$script:form.Controls.Add($script:panProgress)

$script:lblProgressHdr = New-Object System.Windows.Forms.Label
$script:lblProgressHdr.Text = 'OVERALL PROGRESS'
$script:lblProgressHdr.Location = [System.Drawing.Point]::new(10, 11)
$script:lblProgressHdr.AutoSize = $true
$script:lblProgressHdr.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$script:lblProgressHdr.ForeColor = $clrBlue
$script:panProgress.Controls.Add($script:lblProgressHdr)

$script:pbOverall = New-Object System.Windows.Forms.ProgressBar
$script:pbOverall.Location = [System.Drawing.Point]::new(140, 10)
$script:pbOverall.Size = [System.Drawing.Size]::new(690, 20)
$script:pbOverall.Minimum = 0
$script:pbOverall.Maximum = 1000
$script:pbOverall.Value = 0
$script:pbOverall.Style = 'Blocks'
$script:pbOverall.Anchor = 'Top,Left,Right'
$script:panProgress.Controls.Add($script:pbOverall)

$script:lblProgressValue = New-Object System.Windows.Forms.Label
$script:lblProgressValue.Text = '0.0%'
$script:lblProgressValue.Location = [System.Drawing.Point]::new(840, 11)
$script:lblProgressValue.Size = [System.Drawing.Size]::new(84, 18)
$script:lblProgressValue.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:lblProgressValue.ForeColor = $clrText
$script:lblProgressValue.Anchor = 'Top,Right'
$script:panProgress.Controls.Add($script:lblProgressValue)

$lblLogHdr = New-Object System.Windows.Forms.Label
$lblLogHdr.Text = 'OUTPUT LOG'
$lblLogHdr.Location = [System.Drawing.Point]::new(10, 542)
$lblLogHdr.AutoSize = $true
$lblLogHdr.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$lblLogHdr.ForeColor = $clrBlue
$script:form.Controls.Add($lblLogHdr)

$btnClearLog = New-FlatButton -Text 'Clear' -Parent $script:form -X 896 -Y 538 -W 54 -H 22 -BackColor $clrAlt -ForeColor $clrText
$btnClearLog.Anchor = 'Top,Right'

$script:rtLog = New-Object System.Windows.Forms.RichTextBox
$script:rtLog.Location = [System.Drawing.Point]::new(10, 564)
$script:rtLog.Size = [System.Drawing.Size]::new(940, 156)
$script:rtLog.BackColor = $clrLogBg
$script:rtLog.ForeColor = [System.Drawing.Color]::FromArgb(180, 230, 180)
$script:rtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$script:rtLog.ReadOnly = $true
$script:rtLog.ScrollBars = 'Vertical'
$script:rtLog.WordWrap = $false
$script:rtLog.Anchor = 'Top,Bottom,Left,Right'
$script:form.Controls.Add($script:rtLog)

# Status bar
$script:statusBar = New-Object System.Windows.Forms.StatusStrip
$script:statusBar.BackColor = $clrBlue

$script:statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:statusLabel.Text = 'Ready'
$script:statusLabel.ForeColor = [System.Drawing.Color]::White
$script:statusLabel.Spring = $true
$script:statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$script:statusBar.Items.Add($script:statusLabel) | Out-Null

$script:statusElapsedLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:statusElapsedLabel.Text = 'Elapsed: 00:00:00'
$script:statusElapsedLabel.ForeColor = [System.Drawing.Color]::White
$script:statusBar.Items.Add($script:statusElapsedLabel) | Out-Null

$script:statusEtaLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:statusEtaLabel.Text = 'ETA: calculating...'
$script:statusEtaLabel.ForeColor = [System.Drawing.Color]::White
$script:statusBar.Items.Add($script:statusEtaLabel) | Out-Null

$script:statusThroughputLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:statusThroughputLabel.Text = 'Throughput: 0.00 Mbps'
$script:statusThroughputLabel.ForeColor = [System.Drawing.Color]::White
$script:statusBar.Items.Add($script:statusThroughputLabel) | Out-Null

$script:form.Controls.Add($script:statusBar)

function Write-Log {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::FromArgb(180, 230, 180)
    )

    $script:rtLog.SelectionStart = $script:rtLog.TextLength
    $script:rtLog.SelectionLength = 0
    $script:rtLog.SelectionColor = $Color
    $script:rtLog.AppendText("$Text`r`n")
    $script:rtLog.ScrollToCaret()

    if ($script:RunLogWriter) {
        try {
            $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Text
            $script:RunLogWriter.WriteLine($line)
        }
        catch {
        }
    }
}

function Set-CurrentFileActivity {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::FromArgb(220, 220, 220)
    )

    $script:LastFileActivity = $Text
    $script:lblCurrentFileValue.Text = $Text
    $script:lblCurrentFileValue.ForeColor = $Color
}

function Get-CurrentFileActivityFromLine {
    param([string]$Line)

    if (-not $Line) {
        return $null
    }

    $trimmed = $Line.Trim()
    if (-not $trimmed) {
        return $null
    }

    if ($trimmed -match '^(New File|Newer|Older|Same|Tweaked|Extra File|Skipped)\s+(.+)$') {
        return "{0}: {1}" -f $Matches[1], $Matches[2].Trim()
    }

    return $null
}

function Format-DurationText {
    param([TimeSpan]$Duration)

    if ($Duration.TotalSeconds -lt 0) {
        $Duration = [TimeSpan]::Zero
    }

    return '{0:00}:{1:00}:{2:00}' -f [int]$Duration.TotalHours, $Duration.Minutes, $Duration.Seconds
}

function Get-DirectorySizeBytes {
    param([string]$Path)

    try {
        $sum = (Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $sum) {
            return 0L
        }
        return [int64]$sum
    }
    catch {
        return 0L
    }
}

function Get-RoboCopyProgressRecord {
    param([string]$Line)

    if (-not $Line) {
        return $null
    }

    $trimmed = $Line.Trim()
    if (-not $trimmed) {
        return $null
    }

    if ($trimmed -notmatch '^(?<Status>New File|Newer|Older|Same|Tweaked|Extra File|Skipped)\s+(?<Rest>.+)$') {
        return $null
    }

    $status = $Matches['Status']
    $rest = $Matches['Rest'].Trim()
    $sizeBytes = 0L
    $pathText = $rest

    if ($rest -match '^(?<Size>[0-9,]+)\s+(?<Path>.+)$') {
        $sizeText = ($Matches['Size'] -replace ',', '')
        if ([int64]::TryParse($sizeText, [ref]$sizeBytes)) {
            $pathText = $Matches['Path'].Trim()
        }
        else {
            $sizeBytes = 0L
        }
    }

    return [PSCustomObject]@{
        Status = $status
        Path = $pathText
        SizeBytes = $sizeBytes
        CountsAsProcessed = ($status -in @('New File', 'Newer', 'Older', 'Same', 'Tweaked', 'Skipped'))
        CountsAsCopied = ($status -in @('New File', 'Newer', 'Older', 'Tweaked'))
    }
}

function Reset-CopyMetrics {
    $script:CopyStartTime = $null
    $script:SourceTotalBytes = 0L
    $script:ProcessedBytes = 0L
    $script:CopiedBytes = 0L
    $script:statusElapsedLabel.Text = 'Elapsed: 00:00:00'
    $script:statusEtaLabel.Text = 'ETA: calculating...'
    $script:statusThroughputLabel.Text = 'Throughput: 0.00 Mbps'
    $script:pbOverall.Style = 'Blocks'
    $script:pbOverall.Value = 0
    $script:lblProgressValue.Text = '0.0%'
}

function Update-CopyMetrics {
    if (-not $script:CopyStartTime) {
        Reset-CopyMetrics
        return
    }

    $elapsed = (Get-Date) - $script:CopyStartTime
    if ($elapsed.TotalSeconds -lt 0.01) {
        $elapsed = [TimeSpan]::FromSeconds(0.01)
    }

    $script:statusElapsedLabel.Text = 'Elapsed: {0}' -f (Format-DurationText -Duration $elapsed)

    $mbps = 0.0
    if ($script:CopiedBytes -gt 0) {
        $mbps = ($script:CopiedBytes * 8.0) / 1000000.0 / $elapsed.TotalSeconds
    }
    $script:statusThroughputLabel.Text = 'Throughput: {0:N2} Mbps' -f $mbps

    if ($script:SourceTotalBytes -gt 0 -and $script:ProcessedBytes -gt 0) {
        $ratio = [Math]::Min(1.0, ($script:ProcessedBytes / [double]$script:SourceTotalBytes))
        $script:pbOverall.Style = 'Blocks'
        $script:pbOverall.Value = [Math]::Min($script:pbOverall.Maximum, [Math]::Max(0, [int]([Math]::Round($ratio * $script:pbOverall.Maximum))))
        $script:lblProgressValue.Text = '{0:N1}%' -f ($ratio * 100.0)
        if ($ratio -gt 0) {
            $remainingSeconds = [Math]::Max(0.0, ($elapsed.TotalSeconds / $ratio) - $elapsed.TotalSeconds)
            $eta = (Get-Date).AddSeconds($remainingSeconds)
            $script:statusEtaLabel.Text = 'ETA: {0}' -f $eta.ToString('HH:mm:ss')
            return
        }
    }

    if ($script:SourceTotalBytes -gt 0) {
        $script:pbOverall.Style = 'Blocks'
        $script:pbOverall.Value = 0
        $script:lblProgressValue.Text = '0.0%'
    }
    else {
        $script:pbOverall.Style = 'Marquee'
        $script:lblProgressValue.Text = 'Scanning...'
    }

    $script:statusEtaLabel.Text = 'ETA: calculating...'
}

function Get-SelectedFlags {
    $flags = [System.Collections.Generic.List[string]]::new()

    foreach ($kv in $script:checkboxes.GetEnumerator()) {
        if ($kv.Value.Checked) {
            $flags.Add($kv.Key)
        }
    }

    $flags.Add("/MT:$($script:nudMT.Value)")
    $flags.Add("/R:$($script:nudR.Value)")
    $flags.Add("/W:$($script:nudW.Value)")

    $extra = $script:tbExtra.Text.Trim()
    if ($extra) {
        $flags.Add($extra)
    }

    return ($flags -join ' ')
}

function Apply-FlagsToControls {
    param([string]$Flags)

    foreach ($cb in $script:checkboxes.Values) {
        $cb.Checked = $false
    }

    $script:nudMT.Value = 8
    $script:nudR.Value = 3
    $script:nudW.Value = 5
    $script:tbExtra.Text = ''

    $extras = [System.Collections.Generic.List[string]]::new()
    $tokens = @()

    if ($Flags) {
        $tokens = $Flags -split '\s+'
    }

    foreach ($tok in $tokens) {
        if (-not $tok) {
            continue
        }

        if ($script:checkboxes.ContainsKey($tok)) {
            $script:checkboxes[$tok].Checked = $true
            continue
        }

        if ($tok -match '^/MT:(\d+)$') {
            $v = [int]$Matches[1]
            if ($v -lt $script:nudMT.Minimum) { $v = [int]$script:nudMT.Minimum }
            if ($v -gt $script:nudMT.Maximum) { $v = [int]$script:nudMT.Maximum }
            $script:nudMT.Value = $v
            continue
        }

        if ($tok -match '^/R:(\d+)$') {
            $v = [int]$Matches[1]
            if ($v -lt $script:nudR.Minimum) { $v = [int]$script:nudR.Minimum }
            if ($v -gt $script:nudR.Maximum) { $v = [int]$script:nudR.Maximum }
            $script:nudR.Value = $v
            continue
        }

        if ($tok -match '^/W:(\d+)$') {
            $v = [int]$Matches[1]
            if ($v -lt $script:nudW.Minimum) { $v = [int]$script:nudW.Minimum }
            if ($v -gt $script:nudW.Maximum) { $v = [int]$script:nudW.Maximum }
            $script:nudW.Value = $v
            continue
        }

        $extras.Add($tok)
    }

    $script:tbExtra.Text = ($extras -join ' ')
}

function Load-JobIntoEditor {
    param([object]$Job)

    $script:tbJobName.Text = [string]$Job.Name
    $script:tbSource.Text = [string]$Job.Source
    $script:tbDest.Text = [string]$Job.Destination
    Apply-FlagsToControls -Flags ([string]$Job.Flags)

    Write-Log "Loaded job for editing: $($Job.Name)" ([System.Drawing.Color]::FromArgb(120, 180, 255))
}

function Refresh-JobButtons {
    $script:panJobs.Controls.Clear()
    $jobs = Get-Jobs

    if ($jobs.Count -eq 0) {
        $hint = New-Object System.Windows.Forms.Label
        $hint.Text = "No saved jobs yet. Select source/destination, set options, enter a name, then click Save Job."
        $hint.Location = [System.Drawing.Point]::new(6, 14)
        $hint.Size = [System.Drawing.Size]::new(900, 22)
        $hint.ForeColor = $clrMuted
        $script:panJobs.Controls.Add($hint)
        return
    }

    foreach ($job in $jobs) {
        $btnJob = New-Object System.Windows.Forms.Button
        $btnJob.Text = [string]$job.Name
        $btnJob.Size = [System.Drawing.Size]::new(178, 56)
        $btnJob.FlatStyle = 'Flat'
        $btnJob.BackColor = $clrJobBtn
        $btnJob.ForeColor = [System.Drawing.Color]::White
        $btnJob.FlatAppearance.BorderColor = $clrBlue
        $btnJob.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnJob.Margin = [System.Windows.Forms.Padding]::new(4, 4, 0, 0)

        $btnJob.Tag = [PSCustomObject]@{
            Name = $job.Name
            Source = $job.Source
            Destination = $job.Destination
            Flags = $job.Flags
        }

        $tip = New-Object System.Windows.Forms.ToolTip
        $tip.SetToolTip($btnJob, "Src: $($job.Source)`nDst: $($job.Destination)`nFlags: $($job.Flags)")

        $btnJob.Add_Click({
            param($s, $e)
            $j = $s.Tag
            $script:tbSource.Text = $j.Source
            $script:tbDest.Text = $j.Destination
            Start-RoboCopyJob -Src $j.Source -Dst $j.Destination -Flags $j.Flags
        })

        $ctx = New-Object System.Windows.Forms.ContextMenuStrip
        $ctx.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $ctx.ForeColor = $clrText

        $miLoad = New-Object System.Windows.Forms.ToolStripMenuItem 'Load paths'
        $miLoad.Add_Click({
            param($s, $e)
            $strip = $s.GetCurrentParent()
            if ($strip -is [System.Windows.Forms.ContextMenuStrip] -and $strip.SourceControl) {
                $j = $strip.SourceControl.Tag
                $script:tbSource.Text = $j.Source
                $script:tbDest.Text = $j.Destination
                Write-Log "Loaded job: $($j.Name) [$($j.Source) -> $($j.Destination)]" ([System.Drawing.Color]::FromArgb(120, 180, 255))
            }
        })

        $miEdit = New-Object System.Windows.Forms.ToolStripMenuItem 'Edit job'
        $miEdit.Add_Click({
            param($s, $e)
            $strip = $s.GetCurrentParent()
            if ($strip -is [System.Windows.Forms.ContextMenuStrip] -and $strip.SourceControl) {
                $j = $strip.SourceControl.Tag
                Load-JobIntoEditor -Job $j
            }
        })

        $miRun = New-Object System.Windows.Forms.ToolStripMenuItem 'Run this job'
        $miRun.Add_Click({
            param($s, $e)
            $strip = $s.GetCurrentParent()
            if ($strip -is [System.Windows.Forms.ContextMenuStrip] -and $strip.SourceControl) {
                $j = $strip.SourceControl.Tag
                $script:tbSource.Text = $j.Source
                $script:tbDest.Text = $j.Destination
                Start-RoboCopyJob -Src $j.Source -Dst $j.Destination -Flags $j.Flags
            }
        })

        $miDel = New-Object System.Windows.Forms.ToolStripMenuItem 'Delete job'
        $miDel.ForeColor = [System.Drawing.Color]::FromArgb(255, 100, 100)
        $miDel.Add_Click({
            param($s, $e)
            $strip = $s.GetCurrentParent()
            if ($strip -is [System.Windows.Forms.ContextMenuStrip] -and $strip.SourceControl) {
                $jobName = $strip.SourceControl.Tag.Name

                $ans = [System.Windows.Forms.MessageBox]::Show(
                    "Delete saved job '$jobName'?",
                    'Confirm Delete',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
                if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
                    return
                }

                $remaining = @(Get-Jobs | Where-Object { $_.Name -ne $jobName })
                Save-AllJobs -Jobs $remaining
                Refresh-JobButtons
                Write-Log "Deleted job: $jobName" ([System.Drawing.Color]::FromArgb(255, 160, 80))
            }
        })

        $ctx.Items.Add($miLoad) | Out-Null
        $ctx.Items.Add($miEdit) | Out-Null
        $ctx.Items.Add($miRun) | Out-Null
        $ctx.Items.Add([System.Windows.Forms.ToolStripSeparator]::new()) | Out-Null
        $ctx.Items.Add($miDel) | Out-Null

        $btnJob.ContextMenuStrip = $ctx
        $script:panJobs.Controls.Add($btnJob)
    }
}

function Start-RoboCopyJob {
    param(
        [string]$Src,
        [string]$Dst,
        [string]$Flags
    )

    if (-not $Src -or -not $Dst) {
        [System.Windows.Forms.MessageBox]::Show(
            'Please select both source and destination folders.',
            'Missing Paths',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $Src)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Source path does not exist:`r`n`r`n$Src",
            'Invalid Source',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    try {

    $effectiveFlags = Normalize-FlagString -Flags $Flags
    $effectiveFlags = Resolve-RoboCopyPermissionFlags -Flags $effectiveFlags
    if ($null -eq $effectiveFlags) {
        return
    }

    if ($script:RoboProcess -and -not $script:RoboProcess.HasExited) {
        $script:RoboProcess.Kill()
    }
    if ($script:PollTimer) {
        $script:PollTimer.Stop()
    }

    Start-RunLog
    Reset-RawOutputTracking

    $script:CopyStartTime = Get-Date
    $script:SourceTotalBytes = Get-DirectorySizeBytes -Path $Src
    $script:ProcessedBytes = 0L
    $script:CopiedBytes = 0L
    Update-CopyMetrics

    $script:btnRun.Enabled = $false
    $script:btnStop.Enabled = $true
    $script:statusLabel.Text = 'Running...'
    $script:statusBar.BackColor = $clrOrange
    Set-CurrentFileActivity -Text 'Starting RoboCopy...' -Color ([System.Drawing.Color]::FromArgb(120, 180, 255))

    $cmdArgs = "`"$Src`" `"$Dst`" $effectiveFlags"
    $rawStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $rawDir = if ($env:TEMP -and (Test-Path -LiteralPath $env:TEMP)) {
        $env:TEMP
    }
    else {
        $script:BasePath
    }
    $script:RawOutputPath = Join-Path $rawDir ("RoboVMCopy_raw_{0}.log" -f $rawStamp)
    if (Test-Path -LiteralPath $script:RawOutputPath) {
        Remove-Item -LiteralPath $script:RawOutputPath -Force -ErrorAction SilentlyContinue
    }
    Write-Log ('-' * 90) ([System.Drawing.Color]::FromArgb(55, 55, 70))
    Write-Log "robocopy $cmdArgs" ([System.Drawing.Color]::FromArgb(120, 180, 255))
    Write-Log "Run log file: $($script:RunLogPath)" ([System.Drawing.Color]::FromArgb(120, 180, 255))
    Write-Log "Raw output capture: $($script:RawOutputPath)" ([System.Drawing.Color]::FromArgb(120, 180, 255))

    $script:OutputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'cmd.exe'
    $psi.Arguments = "/d /c robocopy $cmdArgs 1> `"$($script:RawOutputPath)`" 2>&1"
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $script:RoboProcess = New-Object System.Diagnostics.Process
    $script:RoboProcess.StartInfo = $psi
    $script:RoboProcess.EnableRaisingEvents = $true

    $script:RoboProcess.Start() | Out-Null

    $script:PollTimer = New-Object System.Windows.Forms.Timer
    $script:PollTimer.Interval = 120

    $script:PollTimer.add_Tick({
        try {
            foreach ($line in @(Get-NewRawOutputLines)) {
                $progressRecord = Get-RoboCopyProgressRecord -Line $line
                if ($progressRecord) {
                    if ($progressRecord.CountsAsProcessed -and $progressRecord.SizeBytes -gt 0) {
                        $script:ProcessedBytes += $progressRecord.SizeBytes
                    }
                    if ($progressRecord.CountsAsCopied -and $progressRecord.SizeBytes -gt 0) {
                        $script:CopiedBytes += $progressRecord.SizeBytes
                    }
                }

                $currentFile = Get-CurrentFileActivityFromLine -Line $line
                if ($currentFile) {
                    Set-CurrentFileActivity -Text $currentFile -Color $clrText
                }

                $color = if ($line -match 'ERROR|error|STDERR:') {
                    [System.Drawing.Color]::FromArgb(255, 100, 100)
                }
                elseif ($line -match 'New File|Newer|Extra File|Extra Dir') {
                    [System.Drawing.Color]::FromArgb(100, 220, 120)
                }
                elseif ($line -match 'Skipped|Same|Tweaked') {
                    [System.Drawing.Color]::FromArgb(130, 130, 145)
                }
                else {
                    [System.Drawing.Color]::FromArgb(180, 230, 180)
                }
                Write-Log $line $color
            }

            Update-CopyMetrics

            if ($script:RoboProcess.HasExited -and @(Get-NewRawOutputLines).Count -eq 0) {
                $script:PollTimer.Stop()
                $rawPathAtEnd = $script:RawOutputPath
                $rawLineCountAtEnd = $script:RawOutputLineCount
                Reset-RawOutputTracking

                if ($rawPathAtEnd) {
                    if (Test-Path -LiteralPath $rawPathAtEnd) {
                        Write-Log "Raw output complete: $rawPathAtEnd (lines: $rawLineCountAtEnd)" ([System.Drawing.Color]::FromArgb(120, 180, 255))
                    }
                    else {
                        Write-Log "Raw output file missing: $rawPathAtEnd" ([System.Drawing.Color]::FromArgb(255, 165, 0))
                    }
                }

                if ($script:RunLogWriter) {
                    Close-RunLog
                    if ($script:RunLogPath) {
                        Write-Log "Saved run log to: $($script:RunLogPath)" ([System.Drawing.Color]::FromArgb(120, 180, 255))
                    }
                }

                if (-not $script:btnRun.Enabled) {
                    $exit = $script:RoboProcess.ExitCode
                    Write-Log ('-' * 90) ([System.Drawing.Color]::FromArgb(55, 55, 70))

                    if ($exit -lt 8) {
                        $desc = switch ($exit) {
                            0 { 'No files copied. Source and destination are already in sync.' }
                            1 { 'Files copied successfully.' }
                            2 { 'Extra files detected at destination (no copy errors).' }
                            3 { 'Files copied and extra files detected at destination.' }
                            default { 'Completed with informational status flags.' }
                        }
                        Write-Log "Completed. Exit code $exit - $desc" ([System.Drawing.Color]::FromArgb(80, 210, 120))
                        $script:statusLabel.Text = "Completed (exit $exit)"
                        $script:statusBar.BackColor = [System.Drawing.Color]::FromArgb(25, 115, 55)
                    }
                    else {
                        Write-Log "Errors detected. Exit code $exit." ([System.Drawing.Color]::FromArgb(255, 100, 100))
                        $script:statusLabel.Text = "Errors (exit $exit)"
                        $script:statusBar.BackColor = $clrRed
                    }

                    Write-Log 'Copy finished. Application remains open.' ([System.Drawing.Color]::FromArgb(120, 180, 255))
                    Set-CurrentFileActivity -Text 'Idle' -Color $clrText
                    if ($script:SourceTotalBytes -gt 0) {
                        $script:ProcessedBytes = [Math]::Max($script:ProcessedBytes, $script:SourceTotalBytes)
                    }
                    Update-CopyMetrics

                    $script:btnRun.Enabled = $true
                    $script:btnStop.Enabled = $false
                }
            }
        }
        catch {
            Write-CrashLog ("Timer runtime error: {0}" -f $_.Exception.ToString())
            $script:statusLabel.Text = 'Runtime error (see log)'
            $script:statusBar.BackColor = $clrRed
            $script:btnRun.Enabled = $true
            $script:btnStop.Enabled = $false
            Write-Log ("Runtime error: {0}" -f $_.Exception.Message) ([System.Drawing.Color]::FromArgb(255, 100, 100))
            [System.Windows.Forms.MessageBox]::Show(
                "A runtime error occurred while updating copy output:`r`n`r`n$($_.Exception.Message)",
                'RoboVMCopy Error',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    })

    $script:PollTimer.Start()
    }
    catch {
        Write-CrashLog ("Copy startup failed: {0}" -f $_.Exception.ToString())
        $script:statusLabel.Text = 'Startup error (see log)'
        $script:statusBar.BackColor = $clrRed
        $script:btnRun.Enabled = $true
        $script:btnStop.Enabled = $false
        Set-CurrentFileActivity -Text 'Idle' -Color $clrText

        Write-Log ("Copy startup failed: {0}" -f $_.Exception.Message) ([System.Drawing.Color]::FromArgb(255, 100, 100))

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to start RoboCopy job:`r`n`r`n$($_.Exception.Message)",
            'RoboVMCopy Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

# Browse source
$btnBrowseSrc.Add_Click({
    $selected = Select-FolderPath -InitialPath $script:tbSource.Text -Title 'Select source folder (mapped drives visible in This PC)'
    if ($selected) {
        $script:tbSource.Text = $selected
    }
})

# Browse destination
$btnBrowseDst.Add_Click({
    $selected = Select-FolderPath -InitialPath $script:tbDest.Text -Title 'Select destination folder (mapped drives visible in This PC)'
    if ($selected) {
        $script:tbDest.Text = $selected
    }
})

# Run
$script:btnRun.Add_Click({
    Start-RoboCopyJob -Src $script:tbSource.Text.Trim() -Dst $script:tbDest.Text.Trim() -Flags (Get-SelectedFlags)
})

$script:btnExit.Add_Click({
    $ans = [System.Windows.Forms.MessageBox]::Show(
        'Exit RoboVMCopy?',
        'Confirm Exit',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        $script:AllowClose = $true
        $script:form.Close()
    }
})

# Stop
$script:btnStop.Add_Click({
    if ($script:RoboProcess -and -not $script:RoboProcess.HasExited) {
        Stop-Process -Id $script:RoboProcess.Id -Force -ErrorAction SilentlyContinue
        Write-Log 'Stopped by user.' ([System.Drawing.Color]::FromArgb(255, 165, 0))
        $script:statusLabel.Text = 'Stopped by user'
        $script:statusBar.BackColor = [System.Drawing.Color]::FromArgb(110, 55, 0)
        Set-CurrentFileActivity -Text 'Stopped' -Color ([System.Drawing.Color]::FromArgb(255, 165, 0))
        Reset-CopyMetrics
        Reset-RawOutputTracking
        $script:btnRun.Enabled = $true
        $script:btnStop.Enabled = $false
    }
})

# Save job
$btnSaveJob.Add_Click({
    $name = $script:tbJobName.Text.Trim()
    $src = $script:tbSource.Text.Trim()
    $dst = $script:tbDest.Text.Trim()

    if (-not $name) {
        [System.Windows.Forms.MessageBox]::Show(
            'Enter a label for the job button.',
            'Name Required',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        $script:tbJobName.Focus()
        return
    }

    if (-not $src -or -not $dst) {
        [System.Windows.Forms.MessageBox]::Show(
            'Select source and destination folders before saving.',
            'Paths Required',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $jobs = @(Get-Jobs)

    if ($jobs | Where-Object { $_.Name -eq $name }) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "A job named '$name' already exists. Overwrite it?",
            'Duplicate Name',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
        $jobs = @($jobs | Where-Object { $_.Name -ne $name })
    }

    $jobs += [PSCustomObject]@{
        Name = $name
        Source = $src
        Destination = $dst
        Flags = (Normalize-FlagString -Flags (Get-SelectedFlags))
    }

    Save-AllJobs -Jobs $jobs
    $script:tbJobName.Text = ''
    Refresh-JobButtons
    Write-Log "Saved job '$name' [$src -> $dst]" ([System.Drawing.Color]::FromArgb(120, 180, 255))
})

$btnClearLog.Add_Click({
    $script:rtLog.Clear()
})

$script:form.Add_FormClosing({
    param($sender, $e)

    if (-not $script:AllowClose) {
        $e.Cancel = $true
        $script:statusLabel.Text = 'Use the Exit button to close RoboVMCopy.'
        Write-Log 'Close request ignored. Click Exit to close the application.' ([System.Drawing.Color]::FromArgb(255, 165, 0))
        return
    }

    if ($script:RoboProcess -and -not $script:RoboProcess.HasExited) {
        $script:RoboProcess.Kill()
    }
    if ($script:PollTimer) {
        $script:PollTimer.Stop()
    }
    Reset-RawOutputTracking
    Close-RunLog
})

Write-Log 'RoboVMCopy ready.' ([System.Drawing.Color]::FromArgb(120, 180, 255))
Write-Log '1. Browse for source and destination folders.' $clrMuted
Write-Log '2. Pick your RoboCopy options.' $clrMuted
Write-Log '3. Click Run RoboCopy, or save as a quick-launch job button.' $clrMuted
Write-Log '4. Left click a saved button to run it.' $clrMuted
Write-Log '5. Right click a saved button for load/edit/run/delete options.' $clrMuted
Write-Log '6. Every run is saved beside the script as RoboCopy_yyyyMMdd_HHmmss.log.' $clrMuted
Write-Log '7. Per-file RoboCopy output is always shown in the GUI.' $clrMuted
Set-CurrentFileActivity -Text 'Idle' -Color $clrText
Reset-CopyMetrics
Write-Log '' $clrMuted

Refresh-JobButtons
[System.Windows.Forms.Application]::Run($script:form)
