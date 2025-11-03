<#
.SYNOPSIS
    Interactive NIC diagnostic report generator (TXT/CSV/HTML)
    HTML highlights anything with "Disabled" anywhere in yellow.
#>

#Requires -Modules NetTCPIP, NetAdapter

[CmdletBinding()]
Param(
    [string]$OutFolder = (Join-Path $env:USERPROFILE 'Documents\NIC_Diagnostics'),
    [switch]$NonInteractive
)

# ==============================
# 0. Enable ANSI
# ==============================
function Enable-ConsoleAnsi {
    $ansiOk = $false
    try {
        $signature = @"
using System;
using System.Runtime.InteropServices;
public static class AnsiConsole {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out int lpMode);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, int dwMode);
}
"@
        Add-Type -TypeDefinition $signature -ErrorAction SilentlyContinue | Out-Null

        $STD_OUTPUT_HANDLE = -11
        $hOut = [AnsiConsole]::GetStdHandle($STD_OUTPUT_HANDLE)
        if ($hOut -ne [IntPtr]::Zero) {
            $mode = 0
            if ([AnsiConsole]::GetConsoleMode($hOut, [ref]$mode)) {
                $VT_ENABLE = 0x0004
                $newMode   = $mode -bor $VT_ENABLE
                [AnsiConsole]::SetConsoleMode($hOut, $newMode) | Out-Null
                $ansiOk = $true
            }
        }
    } catch {
        $ansiOk = $false
    }
    return $ansiOk
}
$null = Enable-ConsoleAnsi

# ==============================
# 1. Config
# ==============================
if (-not (Test-Path $OutFolder)) {
    New-Item -ItemType Directory -Path $OutFolder -Force | Out-Null
}

$ESC        = [char]27 + "["
$BOLD       = "${ESC}1m"
$UNDERLINE  = "${ESC}4m"
$RESET      = "${ESC}0m"
$GREEN      = "${ESC}32m"
$YELLOW     = "${ESC}33m"
$LIGHTBLUE  = "${ESC}94m"

# =====================================================
# 2. HTML writer -- single <pre>, no div spacing
# =====================================================
function Write-NICHtmlReport {
    param(
        [string[]]$ReportLines,
        [string]  $NicName,
        [string]  $HtmlPath
    )


    $css = @'
  <style>
    body {
      background:#111;
      color:#eee;
      font-family:Consolas, "Courier New", monospace;
      padding:20px;
    }
    h1 {
      margin-top:0;
      margin-bottom:10px;
    }
    pre {
      white-space:pre;
      line-height:1.05;
      margin:0;
    }
    .b { font-weight:bold; }
    .u { text-decoration:underline; }
    .g { color:#7CFC00; }
    .y { color:#FFD700; }
    .bblue { color:#87CEFA; }
    .r { color:#FF5555; }
    .dim { color:#555; }
  </style>
'@

    $html = New-Object System.Collections.Generic.List[string]
    $html.Add('<!DOCTYPE html>')
    $html.Add('<html>')
    $html.Add('<head>')
    $html.Add('  <meta charset="UTF-8" />')
    $html.Add('  <title>NIC Diagnostic Report</title>')
    $html.Add($css)
    $html.Add('</head>')
    $html.Add('<body>')

    $encNic = [System.Net.WebUtility]::HtmlEncode($NicName)
    $html.Add(("  <h1>NIC Diagnostic Report - {0}</h1>" -f $encNic))
    $html.Add('<pre>')

    $inAdv = $false
    foreach ($l in $ReportLines) {
        $enc = [System.Net.WebUtility]::HtmlEncode($l)

        if ($l -eq 'Advanced Properties:') {
            $html.Add("  <span class=\"b\">$enc</span>")
            $inAdv = $true
            continue
        }

        if ($inAdv -and $l -eq '------------------------------------') {
            $html.Add("  $enc")
            $inAdv = $false
            continue
        }

        if ($inAdv -and $l -match '^  (.+?):\\s*(.*)$') {
            $name  = [System.Net.WebUtility]::HtmlEncode($matches[1])
            $value = [System.Net.WebUtility]::HtmlEncode($matches[2])

            if ([string]::IsNullOrWhiteSpace($matches[2])) {
                $valueHtml = '<span class="bblue">&lt;null&gt;</span>'
            } elseif ($matches[2] -match '(?i)disabled' -or $matches[2] -match '^(?i)off$') {
                $valueHtml = "<span class=\"y\">$value</span>"
            } else {
                $valueHtml = "<span class=\"g\">$value</span>"
            }

            $html.Add(("  <span class=\"u\">{0}</span>: {1}" -f $name, $valueHtml))
            continue
        }

        if ($l -match 'Errors' -and $l -notmatch 'Driver') {
            $class = if ($l -match '0$') { 'g' } else { 'r' }
            $html.Add(("  <span class=\"{0}\">{1}</span>" -f $class, $enc))
            continue
        }

        if ($l -match 'Discards') {
            $class = if ($l -match '0$') { 'g' } else { 'y' }
            $html.Add(("  <span class=\"{0}\">{1}</span>" -f $class, $enc))
            continue
        }

        $html.Add("  $enc")
    }

    $html.Add('</pre>')
    $html.Add('</body>')
    $html.Add('</html>')

    $html | Out-File -FilePath $HtmlPath -Encoding UTF8

# =====================================================
# 3. New-NICReport
# =====================================================
function New-NICReport {
    param(
        [Parameter(Mandatory)]        $Nic,
        [Parameter(Mandatory)][string]$OutFolder,
        [string] $BOLD,
        [string] $UNDERLINE,
        [string] $RESET,
        [string] $GREEN,
        [string] $YELLOW,
        [string] $LIGHTBLUE
    )

    # filenames
    $ts       = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $CsvFile  = Join-Path $OutFolder ("NIC_Report_{0}.csv"  -f $ts)
    $TxtFile  = Join-Path $OutFolder ("NIC_Report_{0}.txt"  -f $ts)
    $HtmlFile = Join-Path $OutFolder ("NIC_Report_{0}.html" -f $ts)

    # network info
    $ip4Obj = Get-NetIPAddress -InterfaceIndex $Nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
    $ip6Obj = Get-NetIPAddress -InterfaceIndex $Nic.ifIndex -AddressFamily IPv6 -ErrorAction SilentlyContinue
    $mtuObj = Get-NetIPInterface -InterfaceIndex $Nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
    $ip4 = if ($ip4Obj) { ($ip4Obj | Select-Object -ExpandProperty IPAddress -ErrorAction SilentlyContinue) -join ", " } else { "" }
    $ip6 = if ($ip6Obj) { ($ip6Obj | Select-Object -ExpandProperty IPAddress -ErrorAction SilentlyContinue) -join ", " } else { "" }
    $mtu = if ($mtuObj) { $mtuObj.NlMtu } else { $null }

    $driver = Get-NetAdapterAdvancedProperty -Name $Nic.Name -ErrorAction SilentlyContinue
    $stats  = Get-NetAdapterStatistics     -Name $Nic.Name -ErrorAction SilentlyContinue

    # pre-calc so editor doesn't wrap inside hashtable
    $rxPackets   = if ($stats) { $stats.ReceivedUnicastPackets } else { 0 }
    $txPackets   = if ($stats) { $stats.SentUnicastPackets } else { 0 }
    $rxErrors    = if ($stats) { $stats.ReceivedErrors } else { 0 }
    $txErrors    = if ($stats) { $stats.OutboundErrors } else { 0 }
    $rxDiscards  = if ($stats) { $stats.ReceivedDiscardedPackets } else { 0 }
    $txDiscards  = if ($stats) { $stats.OutboundDiscardedPackets } else { 0 }
    $jumboSetting = if ($driver) { (($driver | Where-Object { $_.DisplayName -match 'Jumbo' } | Select-Object -ExpandProperty DisplayValue -ErrorAction SilentlyContinue) -join ', ') } else { '' }

    # CSV
    $csvObj = [PSCustomObject]@{
        Name          = $Nic.Name
        Description   = $Nic.InterfaceDescription
        Status        = $Nic.Status
        MAC           = $Nic.MacAddress
        Speed         = ($Nic.LinkSpeed -as [string])
        MTU           = $mtu
        IPv4          = $ip4
        IPv6          = $ip6
        JumboSetting  = $jumboSetting
        DriverVersion = $Nic.DriverVersion
        DriverDate    = ($Nic.DriverInformation -as [string])
        RxPackets     = $rxPackets
        TxPackets     = $txPackets
        RxErrors      = $rxErrors
        TxErrors      = $txErrors
        RxDiscards    = $rxDiscards
        TxDiscards    = $txDiscards
    }
    $csvObj | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8

    # TEXT report
    $report = New-Object System.Collections.Generic.List[string]
    $report.Add('===============================')
    $report.Add(' NIC Diagnostic Report')
    $report.Add((" Generated: {0}" -f (Get-Date)))
    $report.Add('===============================')
    $report.Add('')
    $report.Add(("Adapter     : {0}" -f $Nic.Name))
    $report.Add(("Description : {0}" -f $Nic.InterfaceDescription))
    $report.Add(("Status      : {0}" -f $Nic.Status))
    $report.Add(("MAC Address : {0}" -f $Nic.MacAddress))
    $report.Add(("Speed       : {0}" -f $Nic.LinkSpeed))
    $report.Add(("MTU         : {0}" -f $mtu))
    $report.Add(("IPv4        : {0}" -f $ip4))
    $report.Add(("IPv6        : {0}" -f $ip6))
    $report.Add(("Driver Ver  : {0}" -f $Nic.DriverVersion))
    $report.Add(("Driver Info : {0}" -f $Nic.DriverInformation))
    $report.Add('')
    $report.Add('Statistics:')
    $report.Add(("  RX Packets : {0}" -f $rxPackets))
    $report.Add(("  TX Packets : {0}" -f $txPackets))
    $report.Add(("  RX Errors  : {0}" -f $rxErrors))
    $report.Add(("  TX Errors  : {0}" -f $txErrors))
    $report.Add(("  RX Discards: {0}" -f $rxDiscards))
    $report.Add(("  TX Discards: {0}" -f $txDiscards))
    $report.Add('')
    $report.Add('Advanced Properties:')
    if ($driver) {
        foreach ($prop in $driver) {
            $report.Add(("  {0}: {1}" -f $prop.DisplayName, $prop.DisplayValue))
        }
    } else {
        $report.Add('  (No advanced properties found or access denied.)')
    }
    $report.Add('------------------------------------')

    # write TXT
    $report | Out-File -FilePath $TxtFile -Encoding UTF8

    # console output
    Write-Host ""
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host " NIC Diagnostic Report" -ForegroundColor Cyan
    Write-Host (" Generated: {0}" -f (Get-Date)) -ForegroundColor Cyan
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host ""

    $inAdv2 = $false
    foreach ($line in $report) {
        if ($line -eq 'Advanced Properties:') {
            Write-Host ("{0}{1}{2}" -f $BOLD, $line, $RESET)
            $inAdv2 = $true
            continue
        }
        if ($inAdv2 -and $line -eq '------------------------------------') {
            Write-Host $line
            $inAdv2 = $false
            continue
        }

        if ($inAdv2) {
            if ($line -match '^  (.+?):\s*(.*)$') {
                $name  = $matches[1]
                $value = $matches[2]

                if ([string]::IsNullOrWhiteSpace($value)) {
                    $valOut = ("{0}<null>{1}" -f $LIGHTBLUE, $RESET)
                }
                elseif ($value -match '(?i)disabled') {
                    $valOut = ("{0}{1}{2}" -f $YELLOW, $value, $RESET)
                }
                elseif ($value -match '^(?i)off$') {
                    $valOut = ("{0}{1}{2}" -f $YELLOW, $value, $RESET)
                }
                else {
                    $valOut = ("{0}{1}{2}" -f $GREEN, $value, $RESET)
                }

                $nameU = ("{0}{1}:{2}" -f $UNDERLINE, $name, $RESET)
                Write-Host ("  {0} {1}" -f $nameU, $valOut)
            } else {
                Write-Host $line
            }
            continue
        }

        if ($line -match 'Errors' -and $line -notmatch 'Driver') {
            if ($line -match '0$') {
                Write-Host $line -ForegroundColor Green
            } else {
                Write-Host $line -ForegroundColor Red
            }
            continue
        }

        if ($line -match 'Discards') {
            if ($line -match '0$') {
                Write-Host $line -ForegroundColor Green
            } else {
                Write-Host $line -ForegroundColor Yellow
            }
            continue
        }

        Write-Host $line
    }

    # HTML
    Write-NICHtmlReport -ReportLines $report.ToArray() -NicName $Nic.Name -HtmlPath $HtmlFile

    return [PSCustomObject]@{
        Csv  = $CsvFile
        Txt  = $TxtFile
        Html = $HtmlFile
    }
} # end New-NICReport

# =====================================================
# 4. MAIN LOOP
# =====================================================
if ($NonInteractive.IsPresent) {
    $allAdaptersNI = Get-NetAdapter | Sort-Object -Property Name
    if (-not $allAdaptersNI) {
        Write-Host "No adapters found." -ForegroundColor Red
        return
    }

    foreach ($adapter in $allAdaptersNI) {
        try {
            [void](New-NICReport -Nic $adapter -OutFolder $OutFolder `
                -BOLD $BOLD -UNDERLINE $UNDERLINE -RESET $RESET `
                -GREEN $GREEN -YELLOW $YELLOW -LIGHTBLUE $LIGHTBLUE)
        } catch {
            Write-Host ("Failed to generate report for {0}: {1}" -f $adapter.Name, $_.Exception.Message) -ForegroundColor Red
        }
    }
    return
}

$script:__quit = $false
while (-not $script:__quit) {
    $allAdapters = Get-NetAdapter | Sort-Object -Property Name
    if (-not $allAdapters) {
        Write-Host "No adapters found." -ForegroundColor Red
        break
    }

    Write-Host ""
    Write-Host "Available Network Adapters:"
    $i = 1
    foreach ($adapter in $allAdapters) {
        $statusText = "(Status: {0})" -f $adapter.Status
        Write-Host ("[{0}] {1} - {2} " -f $i, $adapter.Name, $adapter.InterfaceDescription) -NoNewline
        switch ($adapter.Status) {
            'Up'           { Write-Host $statusText -ForegroundColor Green }
            'Disconnected' { Write-Host $statusText -ForegroundColor Yellow }
            'Not Present'  { Write-Host $statusText -ForegroundColor DarkYellow }
            default        { Write-Host $statusText -ForegroundColor Gray }
        }
        $i++
    }

    $sel = Read-Host "Enter adapter number (Q to quit)"
    if ($sel -match '^[Qq]$') { break }
    if ($sel -notmatch '^\d+$') {
        Write-Host "Invalid selection." -ForegroundColor Red
        continue
    }

    $idx = [int]$sel
    if ($idx -lt 1 -or $idx -gt $allAdapters.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        continue
    }

    $nic = $allAdapters[$idx - 1]

    while (-not $script:__quit) {
        $files = New-NICReport -Nic $nic -OutFolder $OutFolder `
            -BOLD $BOLD -UNDERLINE $UNDERLINE -RESET $RESET `
            -GREEN $GREEN -YELLOW $YELLOW -LIGHTBLUE $LIGHTBLUE

        Write-Host ""
        Write-Host "Reports written to:" -ForegroundColor Cyan
        Write-Host ("1) CSV : {0}" -f $files.Csv)
        Write-Host ("2) TXT : {0}" -f $files.Txt)
        Write-Host ("3) HTML: {0}" -f $files.Html)

        $open = Read-Host "Open? (1=CSV,2=TXT,3=HTML,F=folder,N=none,Q=quit)"
        switch -Regex ($open) {
            '^[1]$' { try { Start-Process "excel.exe" -ArgumentList ("`"{0}`"" -f $files.Csv) } catch { Write-Host "Failed to open CSV in Excel: $($_.Exception.Message)" -ForegroundColor Red } }
            '^[2]$' { try { Start-Process "notepad.exe" -ArgumentList ("`"{0}`"" -f $files.Txt) } catch { Write-Host "Failed to open TXT in Notepad: $($_.Exception.Message)" -ForegroundColor Red } }
            '^[3]$' { try { Start-Process $files.Html } catch { Write-Host "Failed to open HTML in default browser: $($_.Exception.Message)" -ForegroundColor Red } }
            '^[Ff]$' { try { Start-Process "explorer.exe" -ArgumentList ("`"{0}`"" -f $OutFolder) } catch { Write-Host "Failed to open folder in Explorer: $($_.Exception.Message)" -ForegroundColor Red } }
            '^[Qq]$' { $script:__quit = $true; break }
        }

        if ($script:__quit) { break }

        $next = Read-Host "R=run again on this NIC, P=pick another, Q=quit"
        if ($next -match '^[Rr]$') {
            continue
        } elseif ($next -match '^[Pp]$') {
            break
        } elseif ($next -match '^[Qq]$') {
            $script:__quit = $true
            break
        } else {
            Write-Host "Invalid selection. Returning to adapter menu." -ForegroundColor Yellow
            break
        }
    }

    if ($script:__quit) { break }
}
