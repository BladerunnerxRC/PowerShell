<#
.SYNOPSIS
  Interactive NIC diagnostic report (TXT/CSV/HTML + ZIP + JSON).
  - Includes hidden Advanced properties from Registry for the selected NIC.
  - Adds: stack/offload/power, driver package details, connectivity checks, routes, ARP,
          optional RSS mapping, LLDP neighbor, recent NIC events.
  - HTML uses <pre> (no extra line gaps); "Disabled" in Advanced -> yellow.
#>

# ==============================
# 0) Enable ANSI in console
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
    } catch { $ansiOk = $false }
    return $ansiOk
}
$null = Enable-ConsoleAnsi

# ==============================
# 1) Paths & ANSI strings
# ==============================
$OutFolder = "C:\Users\Thoma\OneDrive\Documents\!_DIAGNOSTICS"
if (-not (Test-Path $OutFolder)) { New-Item -ItemType Directory -Path $OutFolder -Force | Out-Null }

$ESC       = [char]27 + "["
$BOLD      = "${ESC}1m"
$UNDERLINE = "${ESC}4m"
$RESET     = "${ESC}0m"
$GREEN     = "${ESC}32m"
$YELLOW    = "${ESC}33m"
$LIGHTBLUE = "${ESC}94m"

# =====================================================
# 2) Registry discovery of hidden Advanced properties
# =====================================================
function Get-NICHiddenAdvancedProperties {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $classGuidPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}'
    $nic = Get-NetAdapter -Name $Name -ErrorAction Stop
    $guidB = $nic.InterfaceGuid.ToString("B")

    $devKey = Get-ChildItem $classGuidPath -ErrorAction SilentlyContinue |
              Where-Object {
                  try { (Get-ItemProperty $_.PsPath -ErrorAction Stop).NetCfgInstanceId -eq $guidB } catch { $false }
              } | Select-Object -First 1
    if (-not $devKey) { return @() }

    $devProps  = Get-ItemProperty -Path $devKey.PSPath
    $paramsKey = Join-Path $devKey.PSPath 'Ndi\Params'
    $out = @()

    $shown = @(Get-NetAdapterAdvancedProperty -Name $nic.Name -ErrorAction SilentlyContinue)
    $shownKeywords = $shown.RegistryKeyword | Where-Object { $_ } | Select-Object -Unique

    if (Test-Path $paramsKey) {
        foreach ($p in Get-ChildItem $paramsKey) {
            $pp       = Get-ItemProperty $p.PSPath
            $valName  = $p.PSChildName
            $desc     = $pp.ParamDesc
            $type     = $pp.type
            $default  = $pp.Default

            $curr = $null
            if ($null -ne $devProps.$valName)          { $curr = $devProps.$valName }
            elseif ($null -ne $devProps.("*$valName")) { $curr = $devProps.("*$valName") }

            $enumMap = @{}
            $enumKey = Join-Path $p.PSPath 'enum'
            if (Test-Path $enumKey) {
                foreach ($e in Get-ChildItem $enumKey) {
                    $ep = Get-ItemProperty $e.PSPath
                    $enumMap[$e.PSChildName] = $ep.'(default)'
                }
            }
            $currStr  = if ($null -eq $curr) { $null } else { [string]$curr }
            $currDisp = if ($currStr -and $enumMap.ContainsKey($currStr)) { $enumMap[$currStr] } else { $currStr }

            $shownHere = ($shownKeywords -contains $valName) -or ($shownKeywords -contains "*$valName")
            $out += [PSCustomObject]@{
                ParamKey     = $valName
                ParamDesc    = $desc
                Type         = $type
                Default      = $default
                CurrentValue = $currDisp
                Hidden       = -not $shownHere
            }
        }
    }

    # Heuristic extras on device key with no Params entry
    foreach ($propName in (Get-Item $devKey.PSPath).Property) {
        if ($propName -match '^\*' -or $propName -cmatch '^(Jumbo|RSS|Flow|Vlan|Speed|Duplex|Interrupt|EEE|Lso|Checksum|Recv|Send|ARPOffload|NSOffload|LsoV2|IPv4|IPv6)') {
            if (-not ($out.ParamKey -contains $propName.TrimStart('*'))) {
                $out += [PSCustomObject]@{
                    ParamKey     = $propName.TrimStart('*')
                    ParamDesc    = $null
                    Type         = $null
                    Default      = $null
                    CurrentValue = [string]$devProps.$propName
                    Hidden       = $true
                }
            }
        }
    }

    $out | Where-Object Hidden | Sort-Object ParamKey
}

# =====================================================
# 3) Extra diagnostics helpers
# =====================================================
function Get-NICStackCaps {
    param([string]$Name)

    $rsc = Get-NetAdapterRsc -Name $Name -ErrorAction SilentlyContinue
    $rss = Get-NetAdapterRss -Name $Name -ErrorAction SilentlyContinue
    $off = Get-NetOffloadGlobalSetting -ErrorAction SilentlyContinue
    $pwr = Get-NetAdapterPowerManagement -Name $Name -ErrorAction SilentlyContinue
    $tcp = Get-NetTCPSetting -SettingName InternetCustom -ErrorAction SilentlyContinue

    [PSCustomObject]@{
        RSCIPv4Enabled   = $rsc.IPv4Enabled
        RSCIPv6Enabled   = $rsc.IPv6Enabled
        RSS              = $rss.Enabled
        Chimney          = $off.Chimney
        IPsecOffload     = $off.IPsecOffload
        RscGlobal        = $off.ReceiveSegmentCoalescing
        ECN              = $tcp.EcnCapability
        AutoTuning       = $tcp.AutoTuningLevelLocal
        DCA              = $off.DirectCacheAccess
        PM_WakeOnMagic   = $pwr.WakeOnMagicPacket
        PM_WakePattern   = $pwr.WakeOnPattern
        PM_DeviceSleep   = $pwr.DeviceSleepOnDisconnect
    }
}

# FIXED: Escapes DeviceID for WQL and falls back if needed
function Get-NICDriverDetail {
    param([string]$PnpDeviceID)

    $escaped = $PnpDeviceID -replace '\\','\\\\' -replace "'","''"
    try {
        Get-CimInstance -ClassName Win32_PnPSignedDriver `
            -Filter "DeviceID='$escaped'" -ErrorAction Stop |
          Select-Object DeviceName, DriverVersion, DriverDate, DriverProviderName,
                        DriverName, InfName, IsSigned, Signer, Manufacturer
    } catch {
        Get-CimInstance -ClassName Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
            Where-Object { $_.DeviceID -eq $PnpDeviceID } |
          Select-Object DeviceName, DriverVersion, DriverDate, DriverProviderName,
                        DriverName, InfName, IsSigned, Signer, Manufacturer
    }
}

function Test-NICBasics {
    param([int]$IfIndex)

    $gwv4 = (Get-NetRoute -InterfaceIndex $IfIndex -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
             Sort-Object RouteMetric | Select-Object -First 1).NextHop
    $dns  = (Get-DnsClientServerAddress -InterfaceIndex $IfIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses -join ', '

    $gwOk = $null; $dnsOk = $null; $mtuNote = $null
    if ($gwv4) { $gwOk = Test-Connection -Count 1 -Quiet -Destination $gwv4 }
    if ($dns)  { $dnsOk = (Resolve-DnsName -Name microsoft.com -ErrorAction SilentlyContinue) -ne $null }

    $mtuOk = $null
    if ($gwv4) {
        try {
            $null = Test-Connection -Destination $gwv4 -Count 1 -BufferSize 1472 -DontFragment -ErrorAction Stop
            $mtuOk = $true
        } catch { $mtuOk = $false; $mtuNote = 'DF 1472 failed -> MTU/fragmentation mismatch?' }
    }

    [PSCustomObject]@{
        DefaultGateway  = $gwv4
        GatewayReachable= $gwOk
        DNSServers      = $dns
        DNSWorks        = $dnsOk
        PMTUProbe1472   = $mtuOk
        PMTUComment     = $mtuNote
    }
}

# FIXED: Filter by LogName + StartTime only; then filter ProviderName in pipeline
function Get-NICRecentEvents {
    param([string]$Name,[int]$Hours = 1)

    $start = (Get-Date).AddHours(-$Hours)
    $providers = @('e1iexpress','rt640x64','Ndis','Tcpip')

    Get-WinEvent -FilterHashtable @{ LogName='System'; StartTime=$start } -ErrorAction SilentlyContinue |
        Where-Object { $_.ProviderName -in $providers } |
        Select-Object TimeCreated, ProviderName, Id, LevelDisplayName |
        Sort-Object TimeCreated -Descending |
        Select-Object -First 50
}

function Add-NetworkTablesToReport {
    param([System.Collections.Generic.List[string]]$Report, [int]$IfIndex)

    $Report.Add('Route Table (v4, this IF):')
    foreach ($r in (Get-NetRoute -InterfaceIndex $IfIndex -AddressFamily IPv4 |
                    Sort-Object RouteMetric, DestinationPrefix |
                    Select-Object -First 12)) {
        $Report.Add( ("  {0,-18} via {1,-15} metric {2}" -f $r.DestinationPrefix, $r.NextHop, $r.RouteMetric) )
    }
    $Report.Add('------------------------------------')

    $Report.Add('ARP Table (this IF):')
    foreach ($n in (Get-NetNeighbor -InterfaceIndex $IfIndex -State Reachable,Stale,Delay,Probe,Permanent -ErrorAction SilentlyContinue |
                    Sort-Object State | Select-Object -First 15)) {
        $Report.Add( ("  {0,-15} => {1}  ({2})" -f $n.IPAddress, $n.LinkLayerAddress, $n.State) )
    }
    $Report.Add('------------------------------------')
}

function Get-NICRssMapping {
    param([string]$Name)
    try { Get-NetAdapterRss -Name $Name -ErrorAction Stop } catch { $null }
}

function Try-LLDPNeighbor {
    try { Get-NetLldpAgent -ErrorAction Stop | Out-Null; Get-NetLldpNeighbor -ErrorAction Stop } catch { $null }
}

function Save-JsonSnapshot {
    param([string]$Path,[hashtable]$Data)
    $Data | ConvertTo-Json -Depth 6 | Out-File -FilePath $Path -Encoding UTF8
}

function Zip-Run {
    param([string]$Csv,[string]$Txt,[string]$Html,[string]$Json)
    $zip = [IO.Path]::ChangeExtension($Html, '.zip')
    $tmp = Join-Path ([IO.Path]::GetDirectoryName($Html)) ([IO.Path]::GetFileNameWithoutExtension($Html) + "_bundle")
    if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
    New-Item -ItemType Directory -Path $tmp | Out-Null
    Copy-Item $Csv,$Txt,$Html,$Json -Destination $tmp -ErrorAction SilentlyContinue
    Compress-Archive -Path (Join-Path $tmp '*') -DestinationPath $zip -Force
    Remove-Item $tmp -Recurse -Force
    return $zip
}

# =====================================================
# 4) HTML writer — single <pre>, with Advanced coloring
# =====================================================
function Write-NICHtmlReport {
    param(
        [string[]]$ReportLines,
        [string]  $NicName,
        [string]  $HtmlPath
    )

    Add-Type -AssemblyName System.Web

    $css = @'
  <style>
    body { background:#111; color:#eee; font-family:Consolas, "Courier New", monospace; padding:20px; }
    h1 { margin:0 0 10px 0; }
    pre { white-space:pre; line-height:1.05; margin:0; }
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

    $encNic = [System.Web.HttpUtility]::HtmlEncode($NicName)
    $html.Add(("  <h1>NIC Diagnostic Report - {0}</h1>" -f $encNic))
    $html.Add('<pre>')

    $inAdv = $false
    foreach ($l in $ReportLines) {
        $enc = [System.Web.HttpUtility]::HtmlEncode($l)

        if ($enc -eq '===============================') { $html.Add('<span class="dim">' + $enc + '</span>'); continue }

        if ($enc -eq 'Advanced Properties:' -or $enc -eq 'Advanced Properties (hidden):') {
            $html.Add('<span class="b">' + $enc + '</span>'); $inAdv = $true; continue
        }

        if ($inAdv -and $enc -eq '------------------------------------') { $html.Add('<span class="dim">' + $enc + '</span>'); $inAdv = $false; continue }

        if ($inAdv) {
            if ($l -match '^  (.+?):\s*(.*)$') {
                $nameRaw  = $matches[1]; $valueRaw = $matches[2]
                $nameEnc  = [System.Web.HttpUtility]::HtmlEncode($nameRaw)
                $valueEnc = [System.Web.HttpUtility]::HtmlEncode($valueRaw)
                if ([string]::IsNullOrWhiteSpace($valueRaw)) {
                    $html.Add('  <span class="u">' + $nameEnc + ':</span> <span class="bblue">&lt;null&gt;</span>')
                } elseif ($valueRaw -match '(?i)disabled') {
                    $html.Add('  <span class="u">' + $nameEnc + ':</span> <span class="y">' + $valueEnc + '</span>')
                } elseif ($valueRaw -match '^(?i)off$') {
                    $html.Add('  <span class="u">' + $nameEnc + ':</span> <span class="y">' + $valueEnc + '</span>')
                } else {
                    $html.Add('  <span class="u">' + $nameEnc + ':</span> <span class="g">' + $valueEnc + '</span>')
                }
                continue
            } else { $html.Add($enc); continue }
        }

        if ($l -match 'Errors' -and $l -notmatch 'Driver') {
            if ($l -match ' 0$') { $html.Add('<span class="g">' + $enc + '</span>') } else { $html.Add('<span class="r">' + $enc + '</span>') }
            continue
        }

        if ($l -match 'Discards') {
            if ($l -match ' 0$') { $html.Add('<span class="g">' + $enc + '</span>') } else { $html.Add('<span class="y">' + $enc + '</span>') }
            continue
        }

        $html.Add($enc)
    }

    $html.Add('</pre>')
    $html.Add('</body>')
    $html.Add('</html>')

    $html -join "`r`n" | Out-File -FilePath $HtmlPath -Encoding UTF8
}

# =====================================================
# 5) Build one NIC report
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

    $ts       = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $CsvFile  = Join-Path $OutFolder ("NIC_Report_{0}.csv"  -f $ts)
    $TxtFile  = Join-Path $OutFolder ("NIC_Report_{0}.txt"  -f $ts)
    $HtmlFile = Join-Path $OutFolder ("NIC_Report_{0}.html" -f $ts)

    $ip4Obj = Get-NetIPAddress -InterfaceIndex $Nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
    $ip6Obj = Get-NetIPAddress -InterfaceIndex $Nic.ifIndex -AddressFamily IPv6 -ErrorAction SilentlyContinue
    $mtuObj = Get-NetIPInterface -InterfaceIndex $Nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue

    $ip4 = if ($ip4Obj) { $ip4Obj.IPAddress -join ", " } else { "N/A" }
    $ip6 = if ($ip6Obj) { $ip6Obj.IPAddress -join ", " } else { "N/A" }
    $mtu = if ($mtuObj) { $mtuObj.NlMtu } else { "N/A" }

    $driver = Get-NetAdapterAdvancedProperty -Name $Nic.Name -ErrorAction SilentlyContinue
    $stats  = Get-NetAdapterStatistics     -Name $Nic.Name -ErrorAction SilentlyContinue

    $rxPackets   = if ($stats) { $stats.ReceivedUnicastPackets } else { 0 }
    $txPackets   = if ($stats) { $stats.SentUnicastPackets     } else { 0 }
    $rxErrors    = if ($stats) { $stats.ReceivedErrors         } else { 0 }
    $txErrors    = if ($stats) { $stats.OutboundErrors         } else { 0 }
    $rxDiscards  = if ($stats) { $stats.ReceivedDiscardedPackets } else { 0 }
    $txDiscards  = if ($stats) { $stats.OutboundDiscardedPackets } else { 0 }
    $jumboSetting = ($driver | Where-Object { $_.DisplayName -match 'Jumbo' }).DisplayValue
    if (-not $jumboSetting) { $jumboSetting = "N/A" }

    # CSV
    [PSCustomObject]@{
        Name          = $Nic.Name
        Description   = $Nic.InterfaceDescription
        Status        = $Nic.Status
        MAC           = $Nic.MacAddress
        Speed         = $Nic.LinkSpeed
        MTU           = $mtu
        IPv4          = $ip4
        IPv6          = $ip6
        JumboSetting  = $jumboSetting
        DriverVersion = $Nic.DriverVersion
        DriverDate    = $Nic.DriverInformation
        RxPackets     = $rxPackets
        TxPackets     = $txPackets
        RxErrors      = $rxErrors
        TxErrors      = $txErrors
        RxDiscards    = $rxDiscards
        TxDiscards    = $txDiscards
    } | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8

    # TEXT report build
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

    # Statistics
    $report.Add('Statistics:')
    $report.Add(("  RX Packets : {0}" -f $rxPackets))
    $report.Add(("  TX Packets : {0}" -f $txPackets))
    $report.Add(("  RX Errors  : {0}" -f $rxErrors))
    $report.Add(("  TX Errors  : {0}" -f $txErrors))
    $report.Add(("  RX Discards: {0}" -f $rxDiscards))
    $report.Add(("  TX Discards: {0}" -f $txDiscards))
    $report.Add('')

    # Stack / offload / power
    $report.Add('Stack/Offload/Power:')
    $caps = Get-NICStackCaps -Name $Nic.Name
    if ($caps) {
        $caps.PSObject.Properties | ForEach-Object { $report.Add(("  {0}: {1}" -f $_.Name, $_.Value)) }
    } else {
        $report.Add('  (No data)')
    }
    $report.Add('------------------------------------')

    # Connectivity checks
    $report.Add('Connectivity Checks:')
    $basics = Test-NICBasics -IfIndex $Nic.ifIndex
    if ($basics) {
        $basics.PSObject.Properties | ForEach-Object { $report.Add(("  {0}: {1}" -f $_.Name, $_.Value)) }
    } else {
        $report.Add('  (No data)')
    }
    $report.Add('------------------------------------')

    # Routes + ARP
    Add-NetworkTablesToReport -Report $report -IfIndex $Nic.ifIndex

    # RSS mapping (if present)
    $rssMap = Get-NICRssMapping -Name $Nic.Name
    if ($rssMap) {
        $report.Add('RSS Mapping:')
        $rssMap.PSObject.Properties | ForEach-Object { $report.Add(("  {0}: {1}" -f $_.Name, $_.Value)) }
        $report.Add('------------------------------------')
    }

    # LLDP neighbor (if available)
    $lldp = Try-LLDPNeighbor
    if ($lldp) {
        $report.Add('LLDP Neighbor:')
        foreach ($n in $lldp | Where-Object { $_.InterfaceAlias -eq $Nic.Name }) {
            $report.Add(("  Chassis:{0}  Port:{1}  SysName:{2}" -f $n.ChassisId, $n.PortId, $n.SystemName))
        }
        $report.Add('------------------------------------')
    }

    # Driver package details
    $d = Get-NICDriverDetail -PnpDeviceID $Nic.PnPDeviceID
    if ($d) {
        $report.Add('Driver Package Detail:')
        $d.PSObject.Properties | ForEach-Object { $report.Add(("  {0}: {1}" -f $_.Name, $_.Value)) }
        $report.Add('------------------------------------')
    }

    # Visible Advanced
    $report.Add('Advanced Properties:')
    if ($driver) {
        foreach ($prop in $driver) { $report.Add(("  {0}: {1}" -f $prop.DisplayName, $prop.DisplayValue)) }
    } else {
        $report.Add('  (No advanced properties found or access denied.)')
    }
    $report.Add('------------------------------------')

    # Hidden Advanced
    try { $hidden = Get-NICHiddenAdvancedProperties -Name $Nic.Name } catch { $hidden = @() }
    $report.Add('Advanced Properties (hidden):')
    if ($hidden -and $hidden.Count -gt 0) {
        foreach ($h in $hidden) {
            $dispName = if ($h.ParamDesc) { $h.ParamDesc } else { $h.ParamKey }
            $curr     = if ($h.CurrentValue) { $h.CurrentValue } else { "" }
            $report.Add(("  {0}: {1}" -f $dispName, $curr))
        }
    } else {
        $report.Add('  (No hidden properties discovered.)')
    }
    $report.Add('------------------------------------')

    # Recent events
    $report.Add('Recent NIC/System Events (last 1h):')
    $ev = Get-NICRecentEvents -Name $Nic.Name -Hours 1
    if ($ev) {
        foreach ($e in $ev) {
            $report.Add(("  [{0}] {1}/{2} {3}" -f $e.TimeCreated, $e.ProviderName, $e.Id, $e.LevelDisplayName))
        }
    } else {
        $report.Add('  (No recent events.)')
    }
    $report.Add('------------------------------------')

    # TXT
    $report | Out-File -FilePath $TxtFile -Encoding UTF8

    # Console (colors for Advanced/Errors/Discards)
    Write-Host ""
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host " NIC Diagnostic Report" -ForegroundColor Cyan
    Write-Host (" Generated: {0}" -f (Get-Date)) -ForegroundColor Cyan
    Write-Host "===============================" -ForegroundColor Cyan
    Write-Host ""

    $inAdv2 = $false
    foreach ($line in $report) {
        if ($line -eq 'Advanced Properties:' -or $line -eq 'Advanced Properties (hidden):') {
            Write-Host ("{0}{1}{2}" -f $BOLD, $line, $RESET); $inAdv2 = $true; continue
        }
        if ($inAdv2 -and $line -eq '------------------------------------') { Write-Host $line; $inAdv2 = $false; continue }

        if ($inAdv2) {
            if ($line -match '^  (.+?):\s*(.*)$') {
                $name  = $matches[1]; $value = $matches[2]
                if ([string]::IsNullOrWhiteSpace($value))      { $valOut = ("{0}<null>{1}" -f $LIGHTBLUE, $RESET) }
                elseif ($value -match '(?i)disabled')           { $valOut = ("{0}{1}{2}" -f $YELLOW, $value, $RESET) }
                elseif ($value -match '^(?i)off$')              { $valOut = ("{0}{1}{2}" -f $YELLOW, $value, $RESET) }
                else                                            { $valOut = ("{0}{1}{2}" -f $GREEN, $value, $RESET) }
                $nameU = ("{0}{1}:{2}" -f $UNDERLINE, $name, $RESET)
                Write-Host ("  {0} {1}" -f $nameU, $valOut)
            } else { Write-Host $line }
            continue
        }

        if ($line -match 'Errors' -and $line -notmatch 'Driver') {
            if ($line -match ' 0$') { Write-Host $line -ForegroundColor Green } else { Write-Host $line -ForegroundColor Red }
            continue
        }
        if ($line -match 'Discards') {
            if ($line -match ' 0$') { Write-Host $line -ForegroundColor Green } else { Write-Host $line -ForegroundColor Yellow }
            continue
        }

        Write-Host $line
    }

    # HTML
    Write-NICHtmlReport -ReportLines $report.ToArray() -NicName $Nic.Name -HtmlPath $HtmlFile

    # JSON snapshot + ZIP bundle
    $jsonPath = [IO.Path]::ChangeExtension($HtmlFile, '.json')
    Save-JsonSnapshot -Path $jsonPath -Data @{
        Adapter = $Nic.Name
        Basics  = $basics
        Caps    = $caps
    }
    $zip = Zip-Run -Csv $CsvFile -Txt $TxtFile -Html $HtmlFile -Json $jsonPath

    # Emit file paths
    [PSCustomObject]@{ Csv = $CsvFile; Txt = $TxtFile; Html = $HtmlFile; Zip = $zip }
}

# =====================================================
# 6) Interactive menu (no flags)
# =====================================================
while ($true) {
    $allAdapters = @(Get-NetAdapter | Sort-Object -Property Name)
    if (-not $allAdapters) { Write-Host "No adapters found." -ForegroundColor Red; break }

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
    if ($sel -notmatch '^\d+$') { Write-Host "Invalid selection." -ForegroundColor Red; continue }

    $idx = [int]$sel
    if ($idx -lt 1 -or $idx -gt $allAdapters.Count) { Write-Host "Invalid selection." -ForegroundColor Red; continue }

    $nic = $allAdapters[$idx - 1]

    while ($true) {
        $files = New-NICReport -Nic $nic -OutFolder $OutFolder `
            -BOLD $BOLD -UNDERLINE $UNDERLINE -RESET $RESET `
            -GREEN $GREEN -YELLOW $YELLOW -LIGHTBLUE $LIGHTBLUE

        Write-Host ""
        Write-Host "Reports written to:" -ForegroundColor Cyan
        Write-Host ("1) CSV : {0}" -f $files.Csv)
        Write-Host ("2) TXT : {0}" -f $files.Txt)
        Write-Host ("3) HTML: {0}" -f $files.Html)
        Write-Host ("4) ZIP : {0}" -f $files.Zip)

        $open = Read-Host "Open? (1=CSV,2=TXT,3=HTML,F=folder,N=none,Q=quit)"
        switch -Regex ($open) {
            '^[1]$' { try { Start-Process "excel.exe" -ArgumentList ("`"{0}`"" -f $files.Csv) } catch {} }
            '^[2]$' { Start-Process "notepad.exe" -ArgumentList ("`"{0}`"" -f $files.Txt) }
            '^[3]$' { Start-Process $files.Html }
            '^[Ff]$' { Start-Process "explorer.exe" -ArgumentList ("`"{0}`"" -f $OutFolder) }
            '^[Qq]$' { exit }
        }

        $next = Read-Host "R=run again on this NIC, P=pick another, Q=quit"
        if ($next -match '^[Rr]$') { continue }
        elseif ($next -match '^[Pp]$') { break }
        else { exit }
    }
}
