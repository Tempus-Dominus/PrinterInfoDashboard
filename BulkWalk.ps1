# Load SharpSnmpLib v11
$libPath = $PSScriptRoot+"\SharpSnmpLib_12.5.6_Net471\SharpSnmpLib.dll"
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -match "SharpSnmpLib" })) {
    Add-Type -Path $libPath
}

$ip = "192.168.209.44"
$community = "public"
#$timeout = 5000
$version = [Lextm.SharpSnmpLib.VersionCode]::V2
$endpoint = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Parse($ip), 161)
$communityStr = New-Object Lextm.SharpSnmpLib.OctetString($community)

# Initialize
Write-Host "Starting manual SNMP walk..."  
    try {

        $resp = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()

        [Lextm.SharpSnmpLib.Messaging.Messenger]::BulkWalk(
            $version,
            $endpoint,
            $communityStr,
            [Lextm.SharpSnmpLib.OctetString]::new(""),
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new("1.3.6.1.2.1.43"),
            $resp,
            10000,
            1000,
            [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree,
            $null,
            $null
        )

        $output = ""

        foreach($v in $resp){
           $output += "$($v.Id ) = $($v.Data)`n"
        }

        $output | Out-File -FilePath $PSScriptRoot"\"$ip"-BulkWalk.txt" -Encoding utf8

        break

    }
    catch {
        Write-Warning "SNMP walk failed: $($_.Exception.Message)"
        break
    }
