param([switch]$Test)

# PowerShell errors should halt execution
$ErrorActionPreference = "Stop"

# Registry prefix for finding WDS IP addresses
$registry_prefix = 'hklm:\SYSTEM\CurrentControlSet\Enum\'

# Regex to find IP addresses
$ip_regex = [regex] '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'

# List of tcp/ip printer ports and port names
$printer_ports = Get-WmiObject win32_tcpipprinterport | where {$_.HostAddress -match $ip_regex}
$port_names = $printer_ports | select -expandProperty "name"

# List of printers
$printers = Get-WmiObject win32_printer

foreach ($printer in $printers) {
    $port_name = $printer."portname"
    if ($port_name.StartsWith('WSD')) {
        # Find a key under $registry_prefix that contains $port_name, get its uuid from containerid,
        # then find a key that contains the uuid, and get its LocationInformation
        # then extract IP from LocationInformation
        $wsd_uuid = get-childItem -path $registry_prefix -recurse -erroraction silentlycontinue |
            where-object {$_.pschildname -like "*$($port_name)*"} |
            get-itemproperty -name "containerid" | select  -expandProperty "containerid"
        $wsd_uuid = $wsd_uuid.substring(0,$wsd_uuid.length-1)
        $wsd_uuid = $wsd_uuid.substring(1)
        $wsd_uuid_keys = get-childItem -path 'hklm:\system\currentcontrolset\enum\' -recurse -erroraction silentlycontinue |
            where-object {$_.pschildname -like "*$($wsd_uuid)*"} # | where-object {$_.pschildname -like "*LocationInformation*"}
        $keys = @()
        foreach ($key in $wsd_uuid_keys) {
            if ($key.psobject.properties | where-object {$_.name -eq "Property"} | where-object {$_.value -eq "LocationInformation"}) {
                $keys += $key
            }
        }
        $location = $keys | get-itemproperty -name "LocationInformation"
        $location = $location[0]."LocationInformation"
        $ip = $ip_regex.Matches($location)  |select -expandproperty "value"
        $type = "WSD"
    } elseif ( $port_names -match $port_name   ) {
        $ip = $printer_ports | where-object {$_.name -eq $port_name } | Select -ExpandProperty "HostAddress"
        $type = "TCP/IP"
    } else {
        # "Ignoring [ $($printer."name") ]"
        continue
    }
    $hostname = [Net.DNS]::GetHostEntry($ip) | Select -ExpandProperty "HostName"
    $driver = $printer | Select -ExpandProperty "DriverName"
    ""
    "$($printer."name")"
    "    Type:      $($type)"
    "    IP:        $($ip)"
    "    nslookup:  $($hostname)"
    "    Driver:    $($driver)"
    ""
    $copy = Read-Host -Prompt "    Create printer with HostAddress $($hostname)? (Y/n)"
    if ($copy -ne 'y' -and $copy -ne '') {
        continue
    }
    $preexisting_port = Get-WmiObject win32_tcpipprinterport | where {$_.Name -eq $hostname }
    if ($preexisting_port) {
        Write-Host "            Port $($hostname) already exists" -ForegroundColor yellow
    } else {
        $port = [wmiclass]"Win32_TcpIpPrinterPort"
        $newPort = $port.CreateInstance()
        $newport.Name = "$hostname"
        $newport.HostAddress = "$hostname"
        $newport.SNMPEnabled = $true
        $newport.Put() | out-null
        Write-Host "            Created port $($hostname)" -ForegroundColor green
    }
    $name = Read-Host -Prompt "        Printer name ($($hostname))"
    if ($name -eq '') {
        $name = $hostname
    }
    $preexisting_printer = Get-WmiObject win32_printer | where {$_.Name -eq $name }
    if ($preexisting_printer) {
        Write-Host "            Printer $($name) already exists" -ForegroundColor yellow
    } else {
        $printer_class = [WMICLASS]"Win32_Printer"
        $newprinter = $printer_class.createInstance()
        $newprinter.Drivername = "$driver"
        $newprinter.PortName = $hostname
        $newprinter.DeviceID = $name
        $newprinter.Name = $name
        $newprinter.Put() | out-null
        Write-Host "            Created printer $name" -ForegroundColor green
    }
    $printer_test = Read-Host -Prompt "        Send test-page to $($name)? (Y/n)"
    if ($printer_test -eq 'y' -or $printer_test -eq '') {
        #$printer.PrintTestPage() # This gives me access denied
        printui /k /n $printer."name"
    }
    $delete = Read-Host -Prompt "        Delete $($printer.name)? (Y/n)"
    if ($delete -eq 'y' -or $delete -eq '') {
        printui /dl /n $printer."name"
    }
}
