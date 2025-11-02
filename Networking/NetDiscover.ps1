<#
.SYNOPSIS
    Interactive NIC diagnostic report generator.
#>

# =========================
# 0. Enable ANSI
# =========================
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
                $newMode = $mode -bor $VT_ENABLE
                [AnsiConsole]::SetConsoleMode($hOut, $newMode) | Out-Null
                $ansiOk = $true
            }
        }
    } catch {
        $ansiOk = $false
    }
    return $ansiOk
} # end Enable-ConsoleAnsi

$null = Enable-ConsoleAnsi

# =========================
# 1. Config + ANSI strings
# =========================
$OutFolder = "C:\Users\Thoma\OneDrive\Documents\!_DIAGNOSTICS"
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

# =========================
# 2. Function: New-NICReport
# =========================
function New-NICReport {
    param(
        [Parameter(Mandatory)] $Nic,
        [Parameter(Mandatory)][string] $OutFolder,
        [string] $BOLD,
        [string] $UNDERLINE,
        [string] $RESET,
        [string] $GREEN,
        [string] $YELLOW,
        [string] $LIGHTBLUE
    )

    # --- filenames ---
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $CsvFile  = Join-Path $OutFolder ("NIC_Report_{0}.csv"  -f $timestamp)
    $TxtFile  = Join-Path $OutFolder ("NIC_Report_{0}.txt"  -f $timestamp)
    $HtmlFile = Join-Path $OutFolder ("NIC_Report_{0}.html" -f $timestamp)

    # --- gather data ---
    $ip4 = (Get-NetIPAddress -InterfaceIndex $Nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress -join ", "
    $ip6 = (Get-NetIPAddress -InterfaceIndex $Nic.ifIndex -AddressFamily IPv6 -ErrorAction SilentlyContinue).IPAddress -join ", "
    $mtu = (Get-NetIPInterface -InterfaceIndex $Nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).NlMtu

    $driver = Get-NetAdapterAdvancedProperty -Name $Nic.Name -ErrorAction SilentlyContinue
    $stats  = Get-NetAdapterStatistics     -Name $Nic.Name -ErrorAction SilentlyContinue

    # --- CSV object ---
    $csvObj = [PSCustomObject]@{
        Name          = $Nic.Name
        Description   = $Nic.InterfaceDescription
        Status        = $Nic.Status
        MAC           = $Nic.MacAddress
        Speed         = $Nic.LinkSpeed
        MTU           = $mtu
        IPv4          = $ip4
        IPv6          = $ip6
        JumboSetting  = ($driver | Where-Object { $_.DisplayName -match 'Jumbo' }).DisplayValue
        DriverVersion = $Nic.DriverVersion
        DriverDate    = $Nic.DriverInformation
        RxPackets     = $stats.ReceivedUnicastPackets
        TxPackets     = $stats.SentUnicastPackets
        RxErrors      = $stats.ReceivedErrors
        TxErrors      = $stats.OutboundErrors
        RxDiscards    = $stats.ReceivedDiscardedPackets
        TxDiscards    = $stats.OutboundDiscardedPackets
    }
    $csvObj | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8

    # --- TXT / HTML content ---
    $report = @()
    $report += "==============================="
    $report += " NIC Diagnostic Report"
    $report += " Generated: $(Get-Date)"
    $report += "==============================="
    $report += ""
    $report += "Adapter     : $($Nic.Name)"
    $report += "Description : $($Nic.InterfaceDescription)"
    $report += "Status      : $($Nic.Status)"
    $report += "MAC Address : $($Nic.MacAddress)"
    $report += "Speed       : $($Nic.LinkSpeed)"
    $report += "MTU         : $mtu"
    $report += "IPv4        : $ip4"
    $report += "IPv6        : $ip6"
    $report += "Driver Ver  : $($Nic.DriverVersion)"
    $report += "Driver Info : $($Nic.DriverInformation)"
    $report += ""
    $report += "Statistics:"
    $report += "  RX Packets : $($stats.ReceivedUnicastPackets)"
    $report += "  TX Packets : $($stats.SentUnicastPackets)"
    $report += "  RX Errors  : $($stats.ReceivedErrors)"
    $report += "  TX Errors  : $($stats.OutboundErrors)"
    $report += "  RX Discards: $($stats.ReceivedDiscardedPackets)"
    $report += "  TX Discards: $($stats.OutboundDiscardedPackets)"
    $report += ""
    $report += "Advanced Properties:"
    if ($driver) {
        foreach ($prop in $driver) {
            $report += ("  {0}: {1}" -f $prop.DisplayName, $prop.DisplayValue)
        }
    } else {
        $report += "  (No advanced properties found or access denied.)"
    }
    $report += "------------------------------------"

    # write TXT
    $report | Out-File -FilePath $TxtFile -Encoding UTF8

    # --- console output with ANSI ---
    Write-Host ""
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host " NIC Diagnostic Report" -ForegroundColor Cyan
    Write-Host (" Generated: {0}" -f (Get-Date)) -ForegroundColor Cyan
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host ""

    $inAdvanced = $false
    foreach ($line in $report) {

        if ($line -eq "Advanced Properties:") {
            Write-Host ("{0}{1}{2}" -f $BOLD, $line, $RESET)
            $inAdvanced = $true
            continue
        }

        if ($inAdvanced -and $line -eq "------------------------------------") {
            Write-Host $line
            $inAdvanced = $false
            continue
        }

        if ($inAdvanced) {
            if ($line -match '^  (.+?):\s*(.*)$') {
                $name  = $matches[1]
                $value = $matches[2]

                if ([string]::IsNullOrWhiteSpace($value)) {
                    $valueOut = ("{0}<null>{1}" -f $LIGHTBLUE, $RESET)
                }
                elseif ($value -match '^(Disabled|Off)$') {
                    $valueOut = ("{0}{1}{2}" -f $YELLOW, $value, $RESET)
                }
                else {
                    $valueOut = ("{0}{1}{2}" -f $GREEN, $value, $RESET)
                }

                $underlined = ("{0}{1}:{2}" -f $UNDERLINE, $name, $RESET)
                Write-Host ("  {0} {1}" -f $underlined, $valueOut)
            }
            else {
                Write-Host $line
            }
            continue
        }

        if ($line -match "Errors" -and $line -notmatch "Driver") {
            if ($line -match "0$") {
                Write-Host $line -ForegroundColor Green
            } else {
                Write-Host $line -ForegroundColor Red
            }
        }
        elseif ($line -match "Discards") {
            if ($line -match "0$") {
                Write-Host $line -ForegroundColor Green
            } else {
                Write-Host $line -ForegroundColor Yellow
            }
        }
        else {
            Write-Host $line
        }
    } # end foreach $report

    # --- HTML ---
    $html = @()
    $html += "<html><head><title>NIC Diagnostic Report</title><meta charset=""UTF-8""></head><body><pre>"
    foreach ($l in $report) { $html += $l }
    $html += "</pre></body></html>"
    $html -join "`r`n" | Out-File -FilePath $HtmlFile -Encoding UTF8

    # return
    return [PSCustomObject]@{
        Csv  = $CsvFile
        Txt  = $TxtFile
        Html = $HtmlFile
    }
} # end function New-NICReport

# =========================
# 3. MAIN LOOP
# =========================
while ($true) {

    $allAdapters = Get-NetAdapter | Sort-Object -Property Name
    if (-not $allAdapters) {
        Write-Host "No network adapters found."
        break
    }

    Write-Host ""
    Write-Host "Available Network Adapters:"
    $i = 1
    foreach ($adapter in $allAdapters) {
        $statusText = "(Status: {0})" -f $adapter.Status
        Write-Host ("[{0}] {1} - {2} " -f $i, $adapter.Name, $adapter.InterfaceDescription) -NoNewline
        switch ($adapter.Status) {
            "Up"          { Write-Host $statusText -ForegroundColor Green }
            "Disconnected"{ Write-Host $statusText -ForegroundColor Yellow }
            "Not Present" { Write-Host $statusText -ForegroundColor DarkYellow }
            default       { Write-Host $statusText -ForegroundColor Gray }
        }
        $i++
    }

    $sel = Read-Host "Enter the number of the adapter to run the report on (or Q to quit)"
    if ($sel -match '^[Qq]$') { break }

    if ($sel -notmatch '^\d+$' -or [int]$sel -lt 1 -or [int]$sel -gt $allAdapters.Count) {
        Write-Host "Invalid selection." -ForegroundColor Red
        continue
    }

    $nic = $allAdapters[[int]$sel - 1]

    $stay = $true
    while ($stay) {

        $files = New-NICReport -Nic $nic -OutFolder $OutFolder `
            -BOLD $BOLD -UNDERLINE $UNDERLINE -RESET $RESET -GREEN $GREEN -YELLOW $YELLOW -LIGHTBLUE $LIGHTBLUE

        Write-Host ""
        Write-Host "Reports written to:" -ForegroundColor Cyan
        Write-Host ("1) CSV : {0}" -f $files.Csv)
        Write-Host ("2) TXT : {0}" -f $files.Txt)
        Write-Host ("3) HTML: {0}" -f $files.Html)
        Write-Host ""

        :OpenReportMenu while ($true) {
            $openChoice = Read-Host "Enter report number to open (1=CSV, 2=TXT, 3=HTML, GoTo Report Dir = F, Q=quit)"

            switch -Regex ($openChoice) {

                '^[1]$' {
                    try { Start-Process "excel.exe" -ArgumentList ("`"{0}`"" -f $files.Csv) } catch {}
                    break OpenReportMenu
                }

                '^[2]$' {
                    $npp = $null
                    $npp1 = "$($env:ProgramFiles)\Notepad++\notepad++.exe"
                    $pf86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
                    $npp2 = if ($pf86) { "$pf86\Notepad++\notepad++.exe" } else { $null }

                    if ($npp1 -and (Test-Path $npp1)) {
                        $npp = $npp1
                    } elseif ($npp2 -and (Test-Path $npp2)) {
                        $npp = $npp2
                    }

                    if ($npp) {
                        Start-Process $npp -ArgumentList ("`"{0}`"" -f $files.Txt)
                    } else {
                        Start-Process "notepad.exe" -ArgumentList ("`"{0}`"" -f $files.Txt)
                    }
                    break OpenReportMenu
                }

                '^[3]$' {
                    Start-Process $files.Html
                    break OpenReportMenu
                }

                '^[Ff]$' {
                    Start-Process "explorer.exe" -ArgumentList "`"$OutFolder`""
                    continue OpenReportMenu   # stay here
                }

                '^[Qq]$' {
                    exit
                }

                default {
                    # bad input → re-prompt
                    continue OpenReportMenu
                }
            } # end switch
        } # end :OpenReportMenu

        $next = Read-Host "Enter R to run again on this adapter, P to pick another adapter, or Q to quit"
        switch -Regex ($next) {
            '^[Rr]$' { continue }
            '^[Pp]$' { $stay = $false }
            '^[Qq]$' { exit }
            default  { exit }
        }
    } # end while ($stay)

} # end while ($true)
