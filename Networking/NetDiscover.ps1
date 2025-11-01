<# 
.SYNOPSIS
    Collects detailed NIC information for all active adapters.
.DESCRIPTION
    Gathers adapter name, description, status, MAC, IPv4/IPv6, MTU, speed,
    jumbo frame setting, and driver details. Outputs to console and CSV.
#>

# Output file (optional)
$OutFile = "$env:USERPROFILE\Desktop\NIC_Report.csv"

# Collect NICs that are up and not virtual/tunnel
$adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" -and $_.HardwareInterface -eq $true }

$results = foreach ($nic in $adapters) {
    $ip4 = (Get-NetIPAddress -InterfaceIndex $nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress -join ", "
    $ip6 = (Get-NetIPAddress -InterfaceIndex $nic.ifIndex -AddressFamily IPv6 -ErrorAction SilentlyContinue).IPAddress -join ", "
    $mtu = (Get-NetIPInterface -InterfaceIndex $nic.ifIndex -AddressFamily IPv4).NlMtu
    $driver = Get-NetAdapterAdvancedProperty -Name $nic.Name -ErrorAction SilentlyContinue

    [PSCustomObject]@{
        Name          = $nic.Name
        Description   = $nic.InterfaceDescription
        Status        = $nic.Status
        MAC           = $nic.MacAddress
        Speed         = $nic.LinkSpeed
        MTU           = $mtu
        IPv4          = $ip4
        IPv6          = $ip6
        JumboSetting  = ($driver | Where-Object {$_.DisplayName -match "Jumbo"}).DisplayValue
        DriverVersion = $nic.DriverVersion
        DriverDate    = $nic.DriverInformation
    }
}

# Output to console
$results | Format-Table -AutoSize

# Export to CSV
$results | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
Write-Host "NIC report saved to $OutFile"

