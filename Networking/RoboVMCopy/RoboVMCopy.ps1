#Requires -Version 5.1
<#
.SYNOPSIS
    RoboVMCopy — GUI front-end for Windows RoboCopy on Windows 11 Pro

.DESCRIPTION
    • Browse any source and destination folder on the system
    • Configure common RoboCopy flags via checkboxes and spinners
    • Run RoboCopy with streaming output in real time
    • Save a job (source + destination + flags) as a named quick-launch button
    • Left-click a saved button → instantly starts that copy job
    • Right-click a saved button → load paths to form  OR  delete the job
    • Jobs are persisted to RoboVMCopy_Jobs.json beside this script

.NOTES
    Run with:  powershell -STA -ExecutionPolicy Bypass -File RoboVMCopy.ps1
#>

# ── STA check ────────────────────────────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw "This script must run in STA mode.`nUse:  powershell -STA -File `"$PSCommandPath`""
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# ── Persistent job storage ────────────────────────────────────────────────────
$script:ConfigFile = Join-Path $PSScriptRoot "RoboVMCopy_Jobs.json"

function Get-Jobs {
    if (Test-Path -LiteralPath $script:ConfigFile) {
        try { return @(Get-Content $script:ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json) }
        catch {}
    }
    return @()
}

function Save-AllJobs {
    param([object[]]$Jobs)
    if ($null -eq $Jobs -or $Jobs.Count -eq 0) {
        Set-Content -Path $script:ConfigFile -Value "[]" -Encoding UTF8
    }
    else {
        $Jobs | ConvertTo-Json -Depth 5 | Set-Content -Path $script:ConfigFile -Encoding UTF8
    }
}

# ── Color palette ─────────────────────────────────────────────────────────────
$clrBg     = [System.Drawing.Color]::FromArgb( 28,  28,  30)
$clrPanel  = [System.Drawing.Color]::FromArgb( 38,  38,  42)
$clrBorder = [System.Drawing.Color]::FromArgb( 60,  60,  65)
$clrBlue   = [System.Drawing.Color]::FromArgb(  0, 122, 204)
$clrGreen  = [System.Drawing.Color]::FromArgb( 35, 160,  65)
$clrRed    = [System.Drawing.Color]::FromArgb(200,  50,  50)
$clrOrange = [System.Drawing.Color]::FromArgb(190,  90,   0)
$clrText   = [System.Drawing.Color]::FromArgb(220, 220, 220)
$clrSub    = [System.Drawing.Color]::FromArgb(140, 140, 145)
$clrInput  = [System.Drawing.Color]::FromArgb( 45,  45,  48)
$clrAlt    = [System.Drawing.Color]::FromArgb( 55,  55,  60)
$clrJobBtn = [System.Drawing.Color]::FromArgb(  0,  85, 148)
$clrLogBg  = [System.Drawing.Color]::FromArgb( 12,  12,  14)

# ── Helper: flat button ───────────────────────────────────────────────────────
function New-FlatBtn {
    param($Text, $Parent, $X, $Y, $W, $H, $Bg, $Fg)
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Location  = [System.Drawing.Point]::new($X, $Y)
    $b.Size      = [System.Drawing.Size]::new($W, $H)
    $b.FlatStyle = 'Flat'
    $b.BackColor = $Bg
    $b.ForeColor = $Fg
    $b.FlatAppearance.BorderColor = $Bg
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $Parent.Controls.Add($b)
    return $b
}

# ── Helper: section header label ──────────────────────────────────────────────
function New-SectionHdr {
    param($Text, $Parent)
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $Text
    $l.Location  = [System.Drawing.Point]::new(10, 8)
    $l.AutoSize  = $true
    $l.Font      = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
    $l.ForeColor = $clrBlue
    $Parent.Controls.Add($l)
}

# ════════════════════════════════════════════════════════════════════════════
#  MAIN FORM
# ════════════════════════════════════════════════════════════════════════════
$script:form             = New-Object System.Windows.Forms.Form
$script:form.Text        = 'RoboVMCopy  —  RoboCopy GUI'
$script:form.Size        = [System.Drawing.Size]::new(980, 800)
$script:form.MinimumSize = [System.Drawing.Size]::new(820, 680)
$script:form.StartPosition = 'CenterScreen'
$script:form.BackColor   = $clrBg
$script:form.ForeColor   = $clrText
$script:form.Font        = New-Object System.Drawing.Font('Segoe UI', 9)

# ════════════════════════════════════════════════════════════════════════════
#  PANEL 1 — PATHS
# ════════════════════════════════════════════════════════════════════════════
$panPaths              = New-Object System.Windows.Forms.Panel
$panPaths.Location     = [System.Drawing.Point]::new(10, 10)
$panPaths.Size         = [System.Drawing.Size]::new(940, 108)
$panPaths.BackColor    = $clrPanel
$panPaths.BorderStyle  = 'FixedSingle'
$panPaths.Anchor       = 'Top,Left,Right'
$script:form.Controls.Add($panPaths)
New-SectionHdr 'PATHS' $panPaths

# Source
$lblSrc          = New-Object System.Windows.Forms.Label
$lblSrc.Text     = 'Source:'
$lblSrc.Location = [System.Drawing.Point]::new(10, 36)
$lblSrc.Size     = [System.Drawing.Size]::new(60, 22)
$lblSrc.ForeColor = $clrText
$panPaths.Controls.Add($lblSrc)

$script:tbSource              = New-Object System.Windows.Forms.TextBox
$script:tbSource.Location     = [System.Drawing.Point]::new(74, 34)
$script:tbSource.Size         = [System.Drawing.Size]::new(768, 24)
$script:tbSource.BackColor    = $clrInput
$script:tbSource.ForeColor    = $clrText
$script:tbSource.BorderStyle  = 'FixedSingle'
$script:tbSource.ReadOnly     = $true
$script:tbSource.Anchor       = 'Top,Left,Right'
$panPaths.Controls.Add($script:tbSource)

$btnBrowseSrc = New-FlatBtn 'Browse…' $panPaths 848 32 76 26 $clrAlt $clrText
$btnBrowseSrc.FlatAppearance.BorderColor = $clrBorder
$btnBrowseSrc.Anchor = 'Top,Right'

# Destination
$lblDst          = New-Object System.Windows.Forms.Label
$lblDst.Text     = 'Destination:'
$lblDst.Location = [System.Drawing.Point]::new(10, 72)
$lblDst.Size     = [System.Drawing.Size]::new(74, 22)
$lblDst.ForeColor = $clrText
$panPaths.Controls.Add($lblDst)

$script:tbDest              = New-Object System.Windows.Forms.TextBox
$script:tbDest.Location     = [System.Drawing.Point]::new(88, 70)
$script:tbDest.Size         = [System.Drawing.Size]::new(754, 24)
$script:tbDest.BackColor    = $clrInput
$script:tbDest.ForeColor    = $clrText
$script:tbDest.BorderStyle  = 'FixedSingle'
$script:tbDest.ReadOnly     = $true
$script:tbDest.Anchor       = 'Top,Left,Right'
$panPaths.Controls.Add($script:tbDest)

$btnBrowseDst = New-FlatBtn 'Browse…' $panPaths 848 68 76 26 $clrAlt $clrText
$btnBrowseDst.FlatAppearance.BorderColor = $clrBorder
$btnBrowseDst.Anchor = 'Top,Right'

# ════════════════════════════════════════════════════════════════════════════
#  PANEL 2 — ROBOCOPY OPTIONS
# ════════════════════════════════════════════════════════════════════════════
$panOpts             = New-Object System.Windows.Forms.Panel
$panOpts.Location    = [System.Drawing.Point]::new(10, 128)
$panOpts.Size        = [System.Drawing.Size]::new(940, 118)
$panOpts.BackColor   = $clrPanel
$panOpts.BorderStyle = 'FixedSingle'
$panOpts.Anchor      = 'Top,Left,Right'
$script:form.Controls.Add($panOpts)
New-SectionHdr 'ROBOCOPY OPTIONS' $panOpts

# Flag checkboxes  (4 columns × 2 rows)
$flagDefs = @(
    @{ F = '/MIR';     D = 'Mirror (sync + purge extras)' }
    @{ F = '/E';       D = 'Include empty subdirectories' }
    @{ F = '/COPYALL'; D = 'Copy all file attributes/info' }
    @{ F = '/Z';       D = 'Restartable mode' }
    @{ F = '/XA:H';    D = 'Exclude hidden files' }
    @{ F = '/PURGE';   D = 'Delete dest files not in src' }
    @{ F = '/NFL';     D = 'No file list in output' }
    @{ F = '/NDL';     D = 'No dir list in output' }
)

$script:checkboxes = @{}
$colW = 228
$col  = 0
$row  = 0

foreach ($fd in $flagDefs) {
    $cb           = New-Object System.Windows.Forms.CheckBox
    $cb.Text      = "$($fd.F)  —  $($fd.D)"
    $cb.Location  = [System.Drawing.Point]::new(10 + $col * $colW, 30 + $row * 26)
    $cb.Size      = [System.Drawing.Size]::new(220, 22)
    $cb.BackColor = $clrPanel
    $cb.ForeColor = $clrText
    $cb.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $panOpts.Controls.Add($cb)
    $script:checkboxes[$fd.F] = $cb
    $col++
    if ($col -ge 4) { $col = 0; $row++ }
}

# Numeric spinners + extra flags row
$nudY = 84

function Add-SpinnerRow {
    param($Label, $X, $Min, $Max, $Default, $Parent)
    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = $Label
    $lbl.Location = [System.Drawing.Point]::new($X, $nudY)
    $lbl.Size     = [System.Drawing.Size]::new(96, 22)
    $lbl.ForeColor = $clrText
    $Parent.Controls.Add($lbl)

    $nud              = New-Object System.Windows.Forms.NumericUpDown
    $nud.Location     = [System.Drawing.Point]::new($X + 100, $nudY - 2)
    $nud.Size         = [System.Drawing.Size]::new(54, 24)
    $nud.Minimum      = $Min
    $nud.Maximum      = $Max
    $nud.Value        = $Default
    $nud.BackColor    = $clrInput
    $nud.ForeColor    = $clrText
    $Parent.Controls.Add($nud)
    return $nud
}

$script:nudMT = Add-SpinnerRow '/MT (threads):'  10  1  128  8  $panOpts
$script:nudR  = Add-SpinnerRow '/R (retries):'  175  0   99  3  $panOpts
$script:nudW  = Add-SpinnerRow '/W (wait sec):' 340  0  300  5  $panOpts

$lblExtra          = New-Object System.Windows.Forms.Label
$lblExtra.Text     = 'Extra flags:'
$lblExtra.Location = [System.Drawing.Point]::new(510, $nudY)
$lblExtra.Size     = [System.Drawing.Size]::new(76, 22)
$lblExtra.ForeColor = $clrText
$panOpts.Controls.Add($lblExtra)

$script:tbExtra                   = New-Object System.Windows.Forms.TextBox
$script:tbExtra.Location          = [System.Drawing.Point]::new(590, $nudY - 2)
$script:tbExtra.Size              = [System.Drawing.Size]::new(336, 24)
$script:tbExtra.BackColor         = $clrInput
$script:tbExtra.ForeColor         = $clrText
$script:tbExtra.BorderStyle       = 'FixedSingle'
$script:tbExtra.PlaceholderText   = '/XF *.tmp /XD Temp Cache'
$script:tbExtra.Anchor            = 'Top,Left,Right'
$panOpts.Controls.Add($script:tbExtra)

# ════════════════════════════════════════════════════════════════════════════
#  ACTION ROW — Run / Stop / Save Job
# ════════════════════════════════════════════════════════════════════════════
$panAction           = New-Object System.Windows.Forms.Panel
$panAction.Location  = [System.Drawing.Point]::new(10, 256)
$panAction.Size      = [System.Drawing.Size]::new(940, 44)
$panAction.BackColor = $clrBg
$panAction.Anchor    = 'Top,Left,Right'
$script:form.Controls.Add($panAction)

$script:btnRun      = New-FlatBtn '▶  Run RoboCopy' $panAction 0 2 164 38 $clrGreen ([System.Drawing.Color]::White)
$script:btnRun.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$script:btnRun.FlatAppearance.BorderColor = $clrGreen

$script:btnStop      = New-FlatBtn '■  Stop' $panAction 172 2 84 38 $clrRed ([System.Drawing.Color]::White)
$script:btnStop.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$script:btnStop.FlatAppearance.BorderColor = $clrRed
$script:btnStop.Enabled = $false

$lblSaveAs          = New-Object System.Windows.Forms.Label
$lblSaveAs.Text     = 'Save as button:'
$lblSaveAs.Location = [System.Drawing.Point]::new(274, 11)
$lblSaveAs.Size     = [System.Drawing.Size]::new(106, 22)
$lblSaveAs.ForeColor = $clrSub
$panAction.Controls.Add($lblSaveAs)

$script:tbJobName                 = New-Object System.Windows.Forms.TextBox
$script:tbJobName.Location        = [System.Drawing.Point]::new(383, 9)
$script:tbJobName.Size            = [System.Drawing.Size]::new(374, 24)
$script:tbJobName.BackColor       = $clrInput
$script:tbJobName.ForeColor       = $clrText
$script:tbJobName.BorderStyle     = 'FixedSingle'
$script:tbJobName.PlaceholderText = 'Button label  (e.g. Backup Documents)'
$script:tbJobName.Anchor          = 'Top,Left,Right'
$panAction.Controls.Add($script:tbJobName)

$btnSaveJob = New-FlatBtn '💾  Save Job' $panAction 766 5 160 34 $clrBlue ([System.Drawing.Color]::White)
$btnSaveJob.FlatAppearance.BorderColor = $clrBlue
$btnSaveJob.Anchor = 'Top,Right'

# ════════════════════════════════════════════════════════════════════════════
#  SAVED JOBS PANEL (FlowLayout)
# ════════════════════════════════════════════════════════════════════════════
$lblJobsHdr          = New-Object System.Windows.Forms.Label
$lblJobsHdr.Text     = 'SAVED JOBS    (left-click = run  |  right-click = options)'
$lblJobsHdr.Location = [System.Drawing.Point]::new(10, 308)
$lblJobsHdr.AutoSize = $true
$lblJobsHdr.Font     = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$lblJobsHdr.ForeColor = $clrBlue
$lblJobsHdr.Anchor   = 'Top,Left'
$script:form.Controls.Add($lblJobsHdr)

$script:panJobs                = New-Object System.Windows.Forms.FlowLayoutPanel
$script:panJobs.Location       = [System.Drawing.Point]::new(10, 328)
$script:panJobs.Size           = [System.Drawing.Size]::new(940, 88)
$script:panJobs.BackColor      = $clrPanel
$script:panJobs.AutoScroll     = $true
$script:panJobs.FlowDirection  = 'LeftToRight'
$script:panJobs.WrapContents   = $true
$script:panJobs.BorderStyle    = 'FixedSingle'
$script:panJobs.Anchor         = 'Top,Left,Right'
$script:form.Controls.Add($script:panJobs)

# ════════════════════════════════════════════════════════════════════════════
#  OUTPUT LOG
# ════════════════════════════════════════════════════════════════════════════
$lblLogHdr          = New-Object System.Windows.Forms.Label
$lblLogHdr.Text     = 'OUTPUT LOG'
$lblLogHdr.Location = [System.Drawing.Point]::new(10, 424)
$lblLogHdr.AutoSize = $true
$lblLogHdr.Font     = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$lblLogHdr.ForeColor = $clrBlue
$lblLogHdr.Anchor   = 'Top,Left'
$script:form.Controls.Add($lblLogHdr)

$btnClearLog = New-FlatBtn 'Clear' $script:form 896 420 54 22 $clrAlt $clrText
$btnClearLog.FlatAppearance.BorderColor = $clrBorder
$btnClearLog.Anchor = 'Top,Right'

$script:rtLog              = New-Object System.Windows.Forms.RichTextBox
$script:rtLog.Location     = [System.Drawing.Point]::new(10, 446)
$script:rtLog.Size         = [System.Drawing.Size]::new(940, 274)
$script:rtLog.BackColor    = $clrLogBg
$script:rtLog.ForeColor    = [System.Drawing.Color]::FromArgb(180, 230, 180)
$script:rtLog.Font         = New-Object System.Drawing.Font('Consolas', 9)
$script:rtLog.ReadOnly     = $true
$script:rtLog.ScrollBars   = 'Vertical'
$script:rtLog.WordWrap     = $false
$script:rtLog.Anchor       = 'Top,Bottom,Left,Right'
$script:form.Controls.Add($script:rtLog)

# ── Status bar ────────────────────────────────────────────────────────────────
$script:statusBar  = New-Object System.Windows.Forms.StatusStrip
$script:statusBar.BackColor = $clrBlue

$script:statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:statusLabel.Text      = 'Ready'
$script:statusLabel.ForeColor = [System.Drawing.Color]::White
$script:statusLabel.Spring    = $true
$script:statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$script:statusBar.Items.Add($script:statusLabel) | Out-Null
$script:form.Controls.Add($script:statusBar)

# ════════════════════════════════════════════════════════════════════════════
#  UTILITY FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════

function Write-Log {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::FromArgb(180, 230, 180)
    )
    $script:rtLog.SelectionStart  = $script:rtLog.TextLength
    $script:rtLog.SelectionLength = 0
    $script:rtLog.SelectionColor  = $Color
    $script:rtLog.AppendText("$Text`n")
    $script:rtLog.ScrollToCaret()
}

function Get-SelectedFlags {
    $flags = [System.Collections.Generic.List[string]]::new()
    foreach ($kv in $script:checkboxes.GetEnumerator()) {
        if ($kv.Value.Checked) { $flags.Add($kv.Key) }
    }
    $flags.Add("/MT:$($script:nudMT.Value)")
    $flags.Add("/R:$($script:nudR.Value)")
    $flags.Add("/W:$($script:nudW.Value)")
    $extra = $script:tbExtra.Text.Trim()
    if ($extra) { $flags.Add($extra) }
    return ($flags -join ' ')
}

# ════════════════════════════════════════════════════════════════════════════
#  SAVED JOBS — populate FlowLayoutPanel
# ════════════════════════════════════════════════════════════════════════════
function Refresh-JobButtons {
    $script:panJobs.Controls.Clear()

    $jobs = Get-Jobs
    if ($jobs.Count -eq 0) {
        $hint           = New-Object System.Windows.Forms.Label
        $hint.Text      = "No saved jobs yet.  Fill in source/destination, choose options, enter a name, then click '💾 Save Job'."
        $hint.Location  = [System.Drawing.Point]::new(6, 14)
        $hint.Size      = [System.Drawing.Size]::new(860, 22)
        $hint.ForeColor = $clrSub
        $script:panJobs.Controls.Add($hint)
        return
    }

    foreach ($job in $jobs) {
        $btnJob           = New-Object System.Windows.Forms.Button
        $btnJob.Text      = $job.Name
        $btnJob.Size      = [System.Drawing.Size]::new(178, 56)
        $btnJob.FlatStyle = 'Flat'
        $btnJob.BackColor = $clrJobBtn
        $btnJob.ForeColor = [System.Drawing.Color]::White
        $btnJob.FlatAppearance.BorderColor = $clrBlue
        $btnJob.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $btnJob.Margin    = [System.Windows.Forms.Padding]::new(4, 4, 0, 0)

        # Store job snapshot in Tag — avoids closure-over-loop-variable bug
        $btnJob.Tag = [PSCustomObject]@{
            Name        = $job.Name
            Source      = $job.Source
            Destination = $job.Destination
            Flags       = $job.Flags
        }

        # Tooltip showing full paths
        $tip = New-Object System.Windows.Forms.ToolTip
        $tip.SetToolTip($btnJob, "Src:   $($job.Source)`nDest:  $($job.Destination)`nFlags: $($job.Flags)")

        # Left-click: run the job immediately (sender Tag is safe — no loop capture)
        $btnJob.Add_Click({
            param($s, $e)
            $j = $s.Tag
            $script:tbSource.Text = $j.Source
            $script:tbDest.Text   = $j.Destination
            Start-RoboCopyJob -Src $j.Source -Dst $j.Destination -Flags $j.Flags
        })

        # Right-click context menu — SourceControl.Tag resolves at click time
        $ctx           = New-Object System.Windows.Forms.ContextMenuStrip
        $ctx.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $ctx.ForeColor = $clrText

        $miLoad           = New-Object System.Windows.Forms.ToolStripMenuItem '📂  Load paths to form'
        $miLoad.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $miLoad.ForeColor = $clrText
        $miLoad.Add_Click({
            param($s, $e)
            $strip = $s.GetCurrentParent()
            if ($strip -is [System.Windows.Forms.ContextMenuStrip] -and $strip.SourceControl) {
                $j = $strip.SourceControl.Tag
                $script:tbSource.Text = $j.Source
                $script:tbDest.Text   = $j.Destination
                Write-Log "Loaded: $($j.Name)  [$($j.Source)  →  $($j.Destination)]" `
                          ([System.Drawing.Color]::FromArgb(120, 180, 255))
            }
        })

        $miRun           = New-Object System.Windows.Forms.ToolStripMenuItem '▶  Run this job'
        $miRun.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $miRun.ForeColor = $clrText
        $miRun.Add_Click({
            param($s, $e)
            $strip = $s.GetCurrentParent()
            if ($strip -is [System.Windows.Forms.ContextMenuStrip] -and $strip.SourceControl) {
                $j = $strip.SourceControl.Tag
                $script:tbSource.Text = $j.Source
                $script:tbDest.Text   = $j.Destination
                Start-RoboCopyJob -Src $j.Source -Dst $j.Destination -Flags $j.Flags
            }
        })

        $miDel           = New-Object System.Windows.Forms.ToolStripMenuItem '🗑  Delete job'
        $miDel.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $miDel.ForeColor = [System.Drawing.Color]::FromArgb(255, 100, 100)
        $miDel.Add_Click({
            param($s, $e)
            $strip = $s.GetCurrentParent()
            if ($strip -is [System.Windows.Forms.ContextMenuStrip] -and $strip.SourceControl) {
                $jobName   = $strip.SourceControl.Tag.Name
                $remaining = @(Get-Jobs | Where-Object { $_.Name -ne $jobName })
                Save-AllJobs -Jobs $remaining
                Refresh-JobButtons
                Write-Log "Deleted job: $jobName" ([System.Drawing.Color]::FromArgb(255, 160, 80))
            }
        })

        $ctx.Items.Add($miLoad) | Out-Null
        $ctx.Items.Add($miRun)  | Out-Null
        $ctx.Items.Add([System.Windows.Forms.ToolStripSeparator]::new()) | Out-Null
        $ctx.Items.Add($miDel)  | Out-Null

        $btnJob.ContextMenuStrip = $ctx
        $script:panJobs.Controls.Add($btnJob)
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  ROBOCOPY RUNNER
#  Uses OutputDataReceived + ConcurrentQueue + UI Timer for live streaming
# ════════════════════════════════════════════════════════════════════════════
$script:RoboProcess  = $null
$script:OutputQueue  = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$script:PollTimer    = $null

function Start-RoboCopyJob {
    param(
        [string]$Src,
        [string]$Dst,
        [string]$Flags
    )

    if (-not $Src -or -not $Dst) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select both a source and destination folder.",
            "Missing Paths",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $Src)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Source path does not exist:`n`n$Src",
            "Invalid Source",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    # Kill any still-running job
    if ($script:RoboProcess -and -not $script:RoboProcess.HasExited) {
        $script:RoboProcess.Kill()
    }
    if ($script:PollTimer) { $script:PollTimer.Stop() }

    $script:btnRun.Enabled   = $false
    $script:btnStop.Enabled  = $true
    $script:statusLabel.Text = '⏳  Running…'
    $script:statusBar.BackColor = $clrOrange

    $cmdArgs = "`"$Src`" `"$Dst`" $Flags"
    Write-Log ('─' * 90) ([System.Drawing.Color]::FromArgb(55, 55, 70))
    Write-Log "▶  robocopy $cmdArgs" ([System.Drawing.Color]::FromArgb(120, 180, 255))

    # Fresh queue for this run
    $script:OutputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

    $psi                        = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = 'robocopy.exe'
    $psi.Arguments              = $cmdArgs
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true

    $script:RoboProcess                    = New-Object System.Diagnostics.Process
    $script:RoboProcess.StartInfo          = $psi
    $script:RoboProcess.EnableRaisingEvents = $true

    # Enqueue stdout lines from thread-pool thread
    $script:RoboProcess.add_OutputDataReceived({
        param($proc, $data)
        if ($null -ne $data.Data) {
            $script:OutputQueue.Enqueue($data.Data)
        }
    })

    # Enqueue stderr lines (robocopy writes very little to stderr)
    $script:RoboProcess.add_ErrorDataReceived({
        param($proc, $data)
        if ($null -ne $data.Data -and $data.Data.Trim()) {
            $script:OutputQueue.Enqueue("STDERR: $($data.Data)")
        }
    })

    $script:RoboProcess.Start()            | Out-Null
    $script:RoboProcess.BeginOutputReadLine()
    $script:RoboProcess.BeginErrorReadLine()

    # UI timer: drains queue and watches for completion on the UI thread
    $script:PollTimer          = New-Object System.Windows.Forms.Timer
    $script:PollTimer.Interval = 120

    $script:PollTimer.add_Tick({
        $line = ''
        while ($script:OutputQueue.TryDequeue([ref]$line)) {
            $color = if ($line -match 'ERROR|error|STDERR:') {
                [System.Drawing.Color]::FromArgb(255, 100, 100)
            }
            elseif ($line -match 'New File|Newer|Extra File|Extra Dir|Newer') {
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

        # Finish when process exited AND queue is fully drained
        if ($script:RoboProcess.HasExited -and $script:OutputQueue.IsEmpty) {
            $script:PollTimer.Stop()

            # Guard: only update if UI is still in "running" state
            if (-not $script:btnRun.Enabled) {
                $exit = $script:RoboProcess.ExitCode

                Write-Log ('─' * 90) ([System.Drawing.Color]::FromArgb(55, 55, 70))

                # RoboCopy exit codes: 0-7 = success/info, 8+ = errors
                if ($exit -lt 8) {
                    $desc = switch ($exit) {
                        0 { 'No files copied — source and destination are identical.' }
                        1 { 'Copy completed — all files copied successfully.' }
                        2 { 'Extra files found in destination (no errors).' }
                        3 { 'Files copied and extra files found in destination.' }
                        default { "Copy completed with status flags." }
                    }
                    Write-Log "✔  Exit code $exit — $desc" ([System.Drawing.Color]::FromArgb(80, 210, 120))
                    $script:statusLabel.Text = "✔  Completed  (exit $exit)"
                    $script:statusBar.BackColor = [System.Drawing.Color]::FromArgb(25, 115, 55)
                }
                else {
                    Write-Log "✖  Exit code $exit — Errors occurred. Some files may not have been copied." `
                              ([System.Drawing.Color]::FromArgb(255, 100, 100))
                    $script:statusLabel.Text = "✖  Errors  (exit $exit)"
                    $script:statusBar.BackColor = $clrRed
                }

                $script:btnRun.Enabled  = $true
                $script:btnStop.Enabled = $false
            }
        }
    })

    $script:PollTimer.Start()
}

# ════════════════════════════════════════════════════════════════════════════
#  EVENT HANDLERS
# ════════════════════════════════════════════════════════════════════════════

# Browse source
$btnBrowseSrc.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description         = 'Select the SOURCE folder to copy FROM'
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
    $dlg.Description         = 'Select the DESTINATION folder to copy TO'
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
    Start-RoboCopyJob -Src   $script:tbSource.Text.Trim() `
                      -Dst   $script:tbDest.Text.Trim() `
                      -Flags (Get-SelectedFlags)
})

# Stop
$script:btnStop.Add_Click({
    if ($script:RoboProcess -and -not $script:RoboProcess.HasExited) {
        Stop-Process -Id $script:RoboProcess.Id -Force -ErrorAction SilentlyContinue
        Write-Log '■  Stopped by user.' ([System.Drawing.Color]::FromArgb(255, 165, 0))
        $script:statusLabel.Text = 'Stopped by user'
        $script:statusBar.BackColor = [System.Drawing.Color]::FromArgb(110, 55, 0)
        $script:btnRun.Enabled  = $true
        $script:btnStop.Enabled = $false
    }
})

# Save job
$btnSaveJob.Add_Click({
    $name = $script:tbJobName.Text.Trim()
    $src  = $script:tbSource.Text.Trim()
    $dst  = $script:tbDest.Text.Trim()

    if (-not $name) {
        [System.Windows.Forms.MessageBox]::Show(
            "Enter a label for the job button.",
            "Name Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        $script:tbJobName.Focus()
        return
    }
    if (-not $src -or -not $dst) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select source and destination folders before saving.",
            "Paths Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $jobs = @(Get-Jobs)

    if ($jobs | Where-Object { $_.Name -eq $name }) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "A job named '$name' already exists.  Overwrite it?",
            "Duplicate Name",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $jobs = @($jobs | Where-Object { $_.Name -ne $name })
    }

    $jobs += [PSCustomObject]@{
        Name        = $name
        Source      = $src
        Destination = $dst
        Flags       = (Get-SelectedFlags)
    }
    Save-AllJobs -Jobs $jobs
    $script:tbJobName.Text = ''
    Refresh-JobButtons
    Write-Log "💾  Saved job '$name'  [$src  →  $dst]" ([System.Drawing.Color]::FromArgb(120, 180, 255))
})

# Clear log
$btnClearLog.Add_Click({ $script:rtLog.Clear() })

# Clean up on form close
$script:form.Add_FormClosing({
    if ($script:RoboProcess -and -not $script:RoboProcess.HasExited) {
        $script:RoboProcess.Kill()
    }
    if ($script:PollTimer) { $script:PollTimer.Stop() }
})

# ════════════════════════════════════════════════════════════════════════════
#  LAUNCH
# ════════════════════════════════════════════════════════════════════════════
Write-Log 'RoboVMCopy ready.' ([System.Drawing.Color]::FromArgb(120, 180, 255))
Write-Log '  1. Click Browse… to select source and destination folders.' $clrSub
Write-Log '  2. Tick the RoboCopy options you need (hover for descriptions).' $clrSub
Write-Log '  3. Click  ▶ Run RoboCopy  to start, or save as a quick-launch button.' $clrSub
Write-Log '  4. Left-click a saved button to run it instantly.' $clrSub
Write-Log '  5. Right-click a saved button to load its paths or delete it.' $clrSub
Write-Log '' $clrSub

Refresh-JobButtons

[System.Windows.Forms.Application]::Run($script:form)
