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

$script:ConfigFile = Join-Path $PSScriptRoot 'RoboVMCopy_Jobs.json'
$script:RoboProcess = $null
$script:OutputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$script:PollTimer = $null

function Get-Jobs {
    if (Test-Path -LiteralPath $script:ConfigFile) {
        try {
            return @(Get-Content -LiteralPath $script:ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json)
        }
        catch {
            return @()
        }
    }
    return @()
}

function Save-AllJobs {
    param([object[]]$Jobs)

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
    @{ F = '/NFL';     D = 'No file list in output' }
    @{ F = '/NDL';     D = 'No directory list in output' }
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
$panAction.Size = [System.Drawing.Size]::new(940, 44)
$panAction.BackColor = $clrBg
$panAction.Anchor = 'Top,Left,Right'
$script:form.Controls.Add($panAction)

$script:btnRun = New-FlatButton -Text 'Run RoboCopy' -Parent $panAction -X 0 -Y 2 -W 164 -H 38 -BackColor $clrGreen -ForeColor ([System.Drawing.Color]::White)
$script:btnRun.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$script:btnStop = New-FlatButton -Text 'Stop' -Parent $panAction -X 172 -Y 2 -W 84 -H 38 -BackColor $clrRed -ForeColor ([System.Drawing.Color]::White)
$script:btnStop.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$script:btnStop.Enabled = $false

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

$btnSaveJob = New-FlatButton -Text 'Save Job' -Parent $panAction -X 766 -Y 5 -W 160 -H 34 -BackColor $clrBlue -ForeColor ([System.Drawing.Color]::White)
$btnSaveJob.Anchor = 'Top,Right'

# Jobs panel
$lblJobsHdr = New-Object System.Windows.Forms.Label
$lblJobsHdr.Text = 'SAVED JOBS (left click = run | right click = options)'
$lblJobsHdr.Location = [System.Drawing.Point]::new(10, 308)
$lblJobsHdr.AutoSize = $true
$lblJobsHdr.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$lblJobsHdr.ForeColor = $clrBlue
$script:form.Controls.Add($lblJobsHdr)

$script:panJobs = New-Object System.Windows.Forms.FlowLayoutPanel
$script:panJobs.Location = [System.Drawing.Point]::new(10, 328)
$script:panJobs.Size = [System.Drawing.Size]::new(940, 88)
$script:panJobs.BackColor = $clrPanel
$script:panJobs.AutoScroll = $true
$script:panJobs.FlowDirection = 'LeftToRight'
$script:panJobs.WrapContents = $true
$script:panJobs.BorderStyle = 'FixedSingle'
$script:panJobs.Anchor = 'Top,Left,Right'
$script:form.Controls.Add($script:panJobs)

# Log panel
$lblLogHdr = New-Object System.Windows.Forms.Label
$lblLogHdr.Text = 'OUTPUT LOG'
$lblLogHdr.Location = [System.Drawing.Point]::new(10, 424)
$lblLogHdr.AutoSize = $true
$lblLogHdr.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$lblLogHdr.ForeColor = $clrBlue
$script:form.Controls.Add($lblLogHdr)

$btnClearLog = New-FlatButton -Text 'Clear' -Parent $script:form -X 896 -Y 420 -W 54 -H 22 -BackColor $clrAlt -ForeColor $clrText
$btnClearLog.Anchor = 'Top,Right'

$script:rtLog = New-Object System.Windows.Forms.RichTextBox
$script:rtLog.Location = [System.Drawing.Point]::new(10, 446)
$script:rtLog.Size = [System.Drawing.Size]::new(940, 274)
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
        $btnJob.Text = $job.Name
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
                $remaining = @(Get-Jobs | Where-Object { $_.Name -ne $jobName })
                Save-AllJobs -Jobs $remaining
                Refresh-JobButtons
                Write-Log "Deleted job: $jobName" ([System.Drawing.Color]::FromArgb(255, 160, 80))
            }
        })

        $ctx.Items.Add($miLoad) | Out-Null
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

    if ($script:RoboProcess -and -not $script:RoboProcess.HasExited) {
        $script:RoboProcess.Kill()
    }
    if ($script:PollTimer) {
        $script:PollTimer.Stop()
    }

    $script:btnRun.Enabled = $false
    $script:btnStop.Enabled = $true
    $script:statusLabel.Text = 'Running...'
    $script:statusBar.BackColor = $clrOrange

    $cmdArgs = "`"$Src`" `"$Dst`" $Flags"
    Write-Log ('-' * 90) ([System.Drawing.Color]::FromArgb(55, 55, 70))
    Write-Log "robocopy $cmdArgs" ([System.Drawing.Color]::FromArgb(120, 180, 255))

    $script:OutputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'robocopy.exe'
    $psi.Arguments = $cmdArgs
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $script:RoboProcess = New-Object System.Diagnostics.Process
    $script:RoboProcess.StartInfo = $psi
    $script:RoboProcess.EnableRaisingEvents = $true

    $script:RoboProcess.add_OutputDataReceived({
        param($proc, $data)
        if ($null -ne $data.Data) {
            $script:OutputQueue.Enqueue($data.Data)
        }
    })

    $script:RoboProcess.add_ErrorDataReceived({
        param($proc, $data)
        if ($null -ne $data.Data -and $data.Data.Trim()) {
            $script:OutputQueue.Enqueue("STDERR: $($data.Data)")
        }
    })

    $script:RoboProcess.Start() | Out-Null
    $script:RoboProcess.BeginOutputReadLine()
    $script:RoboProcess.BeginErrorReadLine()

    $script:PollTimer = New-Object System.Windows.Forms.Timer
    $script:PollTimer.Interval = 120

    $script:PollTimer.add_Tick({
        $line = ''
        while ($script:OutputQueue.TryDequeue([ref]$line)) {
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

        if ($script:RoboProcess.HasExited -and $script:OutputQueue.IsEmpty) {
            $script:PollTimer.Stop()

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

                $script:btnRun.Enabled = $true
                $script:btnStop.Enabled = $false
            }
        }
    })

    $script:PollTimer.Start()
}

# Browse source
$btnBrowseSrc.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Select the source folder to copy from'
    $dlg.UseDescriptionForTitle = $true

    if ($script:tbSource.Text -and (Test-Path -LiteralPath $script:tbSource.Text)) {
        $dlg.SelectedPath = $script:tbSource.Text
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:tbSource.Text = $dlg.SelectedPath
    }

    $dlg.Dispose()
})

# Browse destination
$btnBrowseDst.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Select the destination folder to copy to'
    $dlg.UseDescriptionForTitle = $true

    if ($script:tbDest.Text -and (Test-Path -LiteralPath $script:tbDest.Text)) {
        $dlg.SelectedPath = $script:tbDest.Text
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:tbDest.Text = $dlg.SelectedPath
    }

    $dlg.Dispose()
})

# Run
$script:btnRun.Add_Click({
    Start-RoboCopyJob -Src $script:tbSource.Text.Trim() -Dst $script:tbDest.Text.Trim() -Flags (Get-SelectedFlags)
})

# Stop
$script:btnStop.Add_Click({
    if ($script:RoboProcess -and -not $script:RoboProcess.HasExited) {
        Stop-Process -Id $script:RoboProcess.Id -Force -ErrorAction SilentlyContinue
        Write-Log 'Stopped by user.' ([System.Drawing.Color]::FromArgb(255, 165, 0))
        $script:statusLabel.Text = 'Stopped by user'
        $script:statusBar.BackColor = [System.Drawing.Color]::FromArgb(110, 55, 0)
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
        Flags = (Get-SelectedFlags)
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
    if ($script:RoboProcess -and -not $script:RoboProcess.HasExited) {
        $script:RoboProcess.Kill()
    }
    if ($script:PollTimer) {
        $script:PollTimer.Stop()
    }
})

Write-Log 'RoboVMCopy ready.' ([System.Drawing.Color]::FromArgb(120, 180, 255))
Write-Log '1. Browse for source and destination folders.' $clrMuted
Write-Log '2. Pick your RoboCopy options.' $clrMuted
Write-Log '3. Click Run RoboCopy, or save as a quick-launch job button.' $clrMuted
Write-Log '4. Left click a saved button to run it.' $clrMuted
Write-Log '5. Right click a saved button for load/delete options.' $clrMuted
Write-Log '' $clrMuted

Refresh-JobButtons
[System.Windows.Forms.Application]::Run($script:form)
