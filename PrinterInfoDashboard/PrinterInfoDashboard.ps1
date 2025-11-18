<#
 # @file PrinterInfoDashboard.ps1
 # @brief Displays printer supply info and alerts
 #
 # Queries printers IP Addresses using SNMP to gather supply information and alerts to display
 # for easy access and monitoring. Outputs HTML file.
 #
 # @author Keelan Hyde
 # @date 2025-11-18
 # @version 0.2
 #
 # This file is part of the PrinterInfoDashboard open-source release.
 #
 # @copyright
 # Copyright (c) 2025 Keelan Hyde.
 #
 # This software is licensed under the GNU General Public License v3.0 (GPLv3).
 # You may use, modify, and distribute this code under the terms of the GPLv3
 # as published by the Free Software Foundation.
 #
 # You can obtain a copy of the GPLv3 license at:
 # https://www.gnu.org/licenses/gpl-3.0.html
 #
 # ---------------------------------------------------------------------------
 # SPECIAL TERMS FROM THE AUTHOR (Keelan Hyde):
 # ---------------------------------------------------------------------------
 # - The author retains full copyright ownership of this software.
 #
 # - The author reserves the right to release future versions of this software
 #   under different or proprietary (closed-source) licensing terms.
 #
 # - The author retains the right to incorporate this software, in whole or in
 #   part, into commercial or proprietary products for sale or distribution.
 #
 # - The author may grant alternative or commercial licenses to specific
 #   individuals or organizations at their discretion.
 # 
 # - The author reserves the right to change or revoke this open-source
 #   license for future versions or releases of this software. This does not
 #   affect rights previously granted under existing versions distributed
 #   under the GNU GPLv3.
 #
 # - Any use, modification, or redistribution of this code must include
 #   proper credit to the original author.
 #
 # - Recipients of a license or derivative work may not sell, transfer,
 #   or sublicense their rights to this software without explicit written
 #   consent from the author.
 #
 # - Violation of these terms or of the GPLv3 automatically terminates your
 #   rights to use, modify, or distribute this software under the GPLv3.
 #   Continued use after termination constitutes copyright infringement and
 #   may result in civil or criminal liability under applicable law.
 #
 # ---------------------------------------------------------------------------
 # CONTRIBUTOR LICENSE AGREEMENT (CLA):
 # ---------------------------------------------------------------------------
 # By submitting code, documentation, or other materials to this project,
 # you agree that:
 #
 # 1. You are the original author of your contribution or have the right to
 #    submit it under the same license terms.
 #
 # 2. You grant Keelan (the project maintainer) a perpetual, worldwide,
 #    royalty-free, irrevocable right to use, modify, sublicense, and
 #    relicense your contribution as part of this project, including for
 #    commercial or proprietary distributions.
 #
 # 3. You retain ownership and authorship credit for your contribution.
 #
 # 4. You understand that this agreement allows the project maintainer to
 #    distribute future versions of this software under different terms,
 #    including proprietary licenses.
 #
 # ---------------------------------------------------------------------------
 # CONTRIBUTOR CREDIT POLICY:
 # ---------------------------------------------------------------------------
 # Contributors are encouraged and permitted to self-credit their work by:
 #
 # - Adding a comment in the relevant source file(s) in the following format:
 #
 #   Inline:  
 #         //Contributed by: [Full Name or Handle] - [Date] - ([optional email or URL])  
 #         //Contrib: [FULL NAME or HANDLE] - [Date] - ([Brief description of contribution])  
 #   
 #         Section:  
 #     /*------------------------Start Contribution------------------------
 #      * Contributed by: [FULL NAME or HANDLE] - [Date]
 #     * Description:
 #     * ------------------------------------------------------------------*\/  <-Remove the \
 #     
 #    {...}
 #
 #    /*-------------------------End Contribution------------------------
 #    * Contributed by: [FULL NAME or HANDLE] - [Date]
 #    * ------------------------------------------------------------------*\/  <-Remove the \
 #   
 #    Function/Class:
 #        All Functions and Classes must be preceeded by proper Doxygen commenting.
 #        Contributors must include the following:
 #            @contributor [FULL NAME or HANDLE] - [Date]
 #
 #  
 # - Including a brief description of their change in the project's
 #   CONTRIBUTORS.md file (if available).
 #
 # - Ensuring that commit messages clearly identify the contributor name and
 #   purpose of the change.
 #
 # - Optional: submitting an attribution line in pull request descriptions for
 #   inclusion in project documentation.
 #
 # The project maintainer reserves the right to edit formatting for
 # consistency, but not to remove proper contributor credit.
 #
 # ---------------------------------------------------------------------------
 # @note This header summarizes licensing terms but does not replace the full
 # GPLv3 text. Always refer to the official license for full details.
 #>




<#START SNMP VARIABLES & FUNCTIONS#>
# Load SharpSnmpLib v11
$libPath = $PSScriptRoot+"\SharpSnmpLib_12.5.6_Net471\SharpSnmpLib.dll"
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -match "SharpSnmpLib" })) {
    Add-Type -Path $libPath
}

#SNMP connection setup

#IP Address List
#TODO: Create IP Address query file to limit necessity of editing scrip
$ipList = @(
    #ENTER IP ADDRESS HERE AS STRING ARRAY eg: "127.0.0.1","127.0.0.2"
)


$community = "public"
$timeout = 5000
$version = [Lextm.SharpSnmpLib.VersionCode]::V2
$communityStr = New-Object Lextm.SharpSnmpLib.OctetString($community)

#Page Life for each toner supply item (PartNumber = ReportedPages/Cartridge)
$pageLife = @{
    "006R01742" = 36000     #black
    "006R01743" = 28000     #cyan
    "006R01744" = 28000     #magenta
    "006R01745" = 28000     #yellow
    "013R00681" = 180000    #Drum Cartridge (R1-R4)
    "008R08101" = 69000     #Waste Toner
    "008R08102" = 69000     #Waste Toner
    "001R00623" = 160000    #Transfer Belt Cleaner
    "008R08103" = 200000    #Second Bias Transfer Roll
    "008R08104" = 500000    #Fan Filter
    "006R04693" = 15000     #black
    "006R04694" = 10000     #cyan
    "006R04695" = 10000     #magenta
    "006R04696" = 10000     #yellow
    "008R13325" = 15000     #Waste Toner
    "Imaging Kit CRU" = 125000
    #"unknown" = 5000000
}

#Paper size conversion info
$paperSizes = @{
    "na-letter" = "8.5`"x11`""
    "na-ledger" = "11`"x17`""
    "na-tabloid" = "11`"x17`""
    "na-legal" = "8.5`"x14`""

}

#SNMP MIB return value to human readable value for training level
$trainingLevel = @{
    #Source: https://oidref.com/1.3.6.1.2.1.43.18.1.1.3
    "1" = "Other"
    "2" = "Unknown"
    "3" = "Untrained"
    "4" = "Trained"
    "5" = "Field Service"
    "6" = "Management"
}

#SNMP MIB return value to human readable value for alert severity level
$severityLevel = @{
    #Source: https://schemas.dmtf.org/wbem/cim-html/2.49.0+/CIM_PrintAlertRecord.html
    "1" = "Other"
    "2" = "Unknown"
    "3" = "Critical"
    "4" = "Warning"
    "5" = "Cleared" #Previous Critical
}

#Printer SNMP MIB OID values accessed by this script
$oids = @{
    #Source: https://mibs.observium.org/mib/Printer-MIB/#PrtMarkerCounterUnitTC
    "SystemName" = "1.3.6.1.2.1.1.5.0"
    
    "Toner" = @{
        "Descrp" = "1.3.6.1.2.1.43.11.1.1.6"
        "MaxVol" = "1.3.6.1.2.1.43.11.1.1.8"
        "CurVol" = "1.3.6.1.2.1.43.11.1.1.9"
    }
    
    "Paper" = @{
        "Descrp" = "1.3.6.1.2.1.43.8.2.1.18.1"
        "MaxVol" = "1.3.6.1.2.1.43.8.2.1.9.1"
        "CurVol" = "1.3.6.1.2.1.43.8.2.1.10.1"
        "Size" = "1.3.6.1.2.1.43.8.2.1.12.1"
        "Colour" = "1.3.6.1.2.1.43.8.2.1.21.1"
        "Type" = "1.3.6.1.2.1.43.8.2.1.22.1"
    }

    "Alert" = @{
        "Descrp" = "1.3.6.1.2.1.43.18.1.1.8.1"
        "Training" = "1.3.6.1.2.1.43.18.1.1.3.1"
        "Severity" = "1.3.6.1.2.1.43.18.1.1.2.1"
    }

}

#Function to convert OID string to the proper data type
function oid2snmpID {
    param(
        [Parameter(Mandatory)]
        [string]$Oid
    )

    # Create a one-element array of Variable objects
    return @(
        [Lextm.SharpSnmpLib.Variable]::new(
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new($Oid)
        )
    )
}

function getSnmpData{
    param(
        [Parameter(Mandatory)]
        [string]$oid
    )

    try {

        $response = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()

        [Lextm.SharpSnmpLib.Messaging.Messenger]::BulkWalk(
            $version,
            $endpoint,
            $communityStr,
            [Lextm.SharpSnmpLib.OctetString]::new(""),
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new($oid),
            $response,
            10000,
            1000,
            [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree,
            $null,
            $null
        )

    }
    catch {
        Write-Warning "SNMP walk failed: $($_.Exception.Message)"
        break
    }

    return $response
}

#Holds data on values returned by each printer
$Printers = @()

#Get printer information
foreach($ip in $ipList){
    Write-Host $ip
    $endpoint = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Parse($ip), 161)

    $TonerSupplies = @()
    $PaperSupplies = @()
    $AlertMessages = @()

    #System Name
    try {
        $target = oid2snmpID($oids.SystemName)
        $DeviceName = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get($version, $endpoint, $communityStr, ([System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]$target), $timeout)
    }catch {
        Write-Warning "SNMP walk failed: $($_.Exception.Message)"
        break
    }

    #Toner Supplies
        #Get Toner Information
        $Description = getSnmpData($oids.Toner.Descrp)
        $MaxVol = getSnmpData($oids.Toner.MaxVol)
        $CurVol = getSnmpData($oids.Toner.CurVol)

        <#Pre-process information:
         #  $Description[1..($Description.Count-1)] keeps array item 2...n, first item is the tree root and returns unneeded data
         #   | ForEach-Object {$_.Data} takes the returned key-value pair and leaves only the data
         #>
        $Description = $Description[1..($Description.Count-1)] | ForEach-Object {$_.Data}
        $MaxVol = $MaxVol[1..($MaxVol.Count-1)] | ForEach-Object {$_.Data}
        $CurVol = $CurVol[1..($CurVol.Count-1)] | ForEach-Object {$_.Data}

        #Used to compare the sizes of all the arrays
        $TonerSuppliesLength = @($Description.Count, $MaxVol.Count, $CurVol.Count)

        #Provided that all the arrays are of equal length, Continue to process data
        if(($TonerSuppliesLength | Select-Object -Unique | Measure-Object).Count -eq 1){
            $TonerCombo = [System.Linq.Enumerable]::Zip([String[]]$Description, [String[]]$MaxVol, [String[]]$CurVol)

            foreach($item in $TonerCombo){
                $TonerPercentage = ($item[2]/$item[1])
                #Accounts for @pageLife returning null
                $PagesLeftValue = if($null -ne ($pageLife.GetEnumerator() | Where-Object { $item[0] -match $_.Key }).Value){
                    (($pageLife.GetEnumerator() | Where-Object { $item[0] -match $_.Key }).Value * $TonerPercentage).ToString('N0')
                }else {
                    "N/A"
                }

                $TonerSupplies += [PSCustomObject]@{
                    Supply = ($item[0] -split ',|;')[0]
                    Percent = (($TonerPercentage*100).ToString('N0'))
                    PagesLeft = $PagesLeftValue
                }
            }
        }else{ #Arrays are not equal; return error
            #ADD-CATCH-CODE
        }

    #Paper Supplies
        #Get Paper Information
        $Description = getSnmpData($oids.Paper.Descrp)
        $MaxVol = getSnmpData($oids.Paper.MaxVol)
        $CurVol = getSnmpData($oids.Paper.CurVol)
        $Size = getSnmpData($oids.Paper.Size)
        $Type = getSnmpData($oids.Paper.Type)
        $Colour = getSnmpData($oids.Paper.Colour)

        <#Process information:
         #  $Description[1..($Description.Count-1)] keeps array item 2...n, first item is the tree root and returns unneeded data
         #   | ForEach-Object {$_.Data} takes the returned key-value pair and leaves only the data
         #>
        $Description = $Description[1..($Description.Count-1)] | ForEach-Object {$_.Data}
        $MaxVol = $MaxVol[1..($MaxVol.Count-1)] | ForEach-Object {$_.Data}
        $CurVol = $CurVol[1..($CurVol.Count-1)] | ForEach-Object {$_.Data} 
        $Size = $Size[1..($Size.Count-1)] | ForEach-Object {$_.Data} 
        $Type = $Type[1..($Type.Count-1)] | ForEach-Object {$_.Data} 
        $Colour = $Colour[1..($Colour.Count-1)] | ForEach-Object {$_.Data} 

        #Used to compare the sizes of all the arrays
        $PaperSuppliesLength = @($Description.Count, $MaxVol.Count, $CurVol.Count, $Size.Count, $Type.Count, $Colour.Count)

        #Provided that all the arrays are of equal length, Continue to process data
        if(($PaperSuppliesLength | Select-Object -Unique | Measure-Object).Count -eq 1){

            #Have to use a 2D array instead of [System.Linq.Enumerable]::Zip() as Zip can only handle a maximum of 3 parameters.
            $PaperCombo = @([string[]]$Description, [string[]]$MaxVol, [string[]]$CurVol, [string[]]$Size, [string[]]$Type, [string[]]$Colour)

            <#For-loop required due to not using [System.Linq.Enumerable]::Zip.
             #$PaperCombo[SUB-ARRAY][ITERATE-SUB-ARRAY] therefore $PaperCombo[0][$i] would iterate through $Description
             #>
            for($i = 0; $i -lt $Description.Count; $i++){
                $PaperPercentage = ($PaperCombo[2][$i]/$PaperCombo[1][$i])
                #Accounts for @pageLife returning null
                $PageSizes = if($null -ne ($paperSizes.GetEnumerator() | Where-Object { $PaperCombo[3][$i] -match $_.Key }).Value){
                (($paperSizes.GetEnumerator() | Where-Object { $PaperCombo[3][$i] -match $_.Key }).Value).ToString()
                }else {
                    $PaperCombo[3][$i]
                }

                $PaperSupplies += [PSCustomObject]@{
                    Tray = ($PaperCombo[0][$i] -split ',|;')[0]
                    Percent = (($PaperPercentage*100).ToString('N0'))
                    Remaining = ($PaperCombo[2][$i] + " of " + $PaperCombo[1][$i])
                    Size = $PageSizes
                    Finish = ($PaperCombo[5][$i] + ", " + $PaperCombo[4][$i])
                }
            }
        }else{ #Arrays are not equal; return error
            #ADD-CATCH-CODE
        }

    #Alert Messages
            $Description = getSnmpData($oids.Alert.Descrp)
            $Training = getSnmpData($oids.Alert.Training)
            $Severity = getSnmpData($oids.Alert.Severity)

        <#Pre-process information:
         #  $Description[1..($Description.Count-1)] keeps array item 2...n, first item is the tree root and returns unneeded data
         #   | ForEach-Object {$_.Data} takes the returned key-value pair and leaves only the data
         #>
        $Description = $Description[1..($Description.Count-1)] | ForEach-Object {$_.Data}
        $Training = $Training[1..($Training.Count-1)] | ForEach-Object {$_.Data}
        $Severity = $Severity[1..($Severity.Count-1)] | ForEach-Object {$_.Data}

        #Alert table can return 0 values if no alerts are present, causing [System.Linq.Enumerable]::Zip() to throw an error; this keeps zip happy.
        if(($Description.Count -eq 0) -and ($Training.Count -eq 0) -and ($Severity.Count -eq 0)){
            $Description = ">NO ALERTS AT THIS TIME<"
            $Training = " "
            $Severity = " "
        }

        #Used to compare the sizes of all the arrays
        $AlertLength = @($Description.Count, $Training.Count, $Severity.Count)

        #Provided that all the arrays are of equal length, Continue to process data
        if(($AlertLength | Select-Object -Unique | Measure-Object).Count -eq 1){
            $AlertCombo = [System.Linq.Enumerable]::Zip([String[]]$Description, [String[]]$Training, [String[]]$Severity)

            foreach($item in $AlertCombo){
                $AlertMessages += [PSCustomObject]@{
                    AlertMessage = (($item[0] -split '\.|;')[0] + '.')
                    TrainingLevel = ($trainingLevel.GetEnumerator() | Where-Object { $item[1] -match $_.Key }).Value
                    SeverityLevel = ($severityLevel.GetEnumerator() | Where-Object { $item[2] -match $_.Key }).Value
                }
            }
            
        }else{ #Arrays are not equal; return error
            #ADD-CATCH-CODE
        }

    #Stitch everything together into one nice, neat object
    $Printers += [PSCustomObject]@{
        Printer = $DeviceName.Data.ToString()
        PtrIP = $ip
        Toner = $TonerSupplies
        Paper = $PaperSupplies
        Alerts = $AlertMessages
    }

}

<#START HTML VARIABLES#>

#Start of html code: <!DOCTYPE html> to <body>
$htmlStart = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <title>Template.html</title>
    <style>
        :root {
            --cyan-500: #06b6d4;
            --yellow-500: #eab308;
            --magenta-500: #d946ef;
            --jet-black-500: #111827;

        }
        .Printer{
            border-width: 2px;
            border-style: solid;
            border-radius: 10px;
            display: flex;
            width: 1020px;
            margin: 10px;
        }
        .alertContainer{
            width: 100% !important;
        }
        .alertLabel{
            height: 50px !important;
        }
        .alertTable{
            width: 100%;
        }
        .alertTable table, tr{
            border: 1px solid black;
            border-collapse: collapse;
            width: 100%;
        }
        .alertTable table{
            width: 100%;
        }
        .Container{

            display: flex;
            flex-direction: column;
        }
        .sectionLabel{
            background-color: lightgray;
            border-radius: 10px;
            width: auto;
            height: 100px;
            margin: 1em;
            display: flex;
        }
        .svg{
            margin: 0px 10px 0px 10px;
        }
        .sectionLabelInfo{
            width: 100%;
            position: relative;
        }
        .sectionLabel-primaryInfo{
            position: absolute;
            margin: auto;
            top: 40%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-size: 2em;
        }
        .sectionLabel-secondaryInfo{
            position: absolute;
            margin: auto;
            top: 65%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-size: 1em;
        }
        .supplyContainer{
            display: inline-flex;
        }
        .supplyItem{
            width: 6em;
            height: auto;
            margin: 1em 1em 1em 1em;
        }
        .supplyIcon{
            width: 6em;
            height: 6em;
        }
        .supplyInfo{
            width: auto;
            display: block;
            text-align: center;
        }
    </style>
</head>
<body>
"@

#Everything that belongs in <body></body>
$htmlBody = @"
"@

#End of html code: </body> & </html>
$htmlEnd = @"
</body>
</html>
"@

#Populates the <body></body> with each printer
foreach($Printer in $Printers){
    $htmlInsert = @"
    <div class="Printer">
        
        <div class="Container">
            <div class="sectionLabel">
                <svg class="svg" width="100px" height="100px" viewBox="0 0 1024 1024" xmlns="http://www.w3.org/2000/svg" fill="#000000"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"><path fill="#000000" d="M256 768H105.024c-14.272 0-19.456-1.472-24.64-4.288a29.056 29.056 0 0 1-12.16-12.096C65.536 746.432 64 741.248 64 727.04V379.072c0-42.816 4.48-58.304 12.8-73.984 8.384-15.616 20.672-27.904 36.288-36.288 15.68-8.32 31.168-12.8 73.984-12.8H256V64h512v192h68.928c42.816 0 58.304 4.48 73.984 12.8 15.616 8.384 27.904 20.672 36.288 36.288 8.32 15.68 12.8 31.168 12.8 73.984v347.904c0 14.272-1.472 19.456-4.288 24.64a29.056 29.056 0 0 1-12.096 12.16c-5.184 2.752-10.368 4.224-24.64 4.224H768v192H256V768zm64-192v320h384V576H320zm-64 128V512h512v192h128V379.072c0-29.376-1.408-36.48-5.248-43.776a23.296 23.296 0 0 0-10.048-10.048c-7.232-3.84-14.4-5.248-43.776-5.248H187.072c-29.376 0-36.48 1.408-43.776 5.248a23.296 23.296 0 0 0-10.048 10.048c-3.84 7.232-5.248 14.4-5.248 43.776V704h128zm64-448h384V128H320v128zm-64 128h64v64h-64v-64zm128 0h64v64h-64v-64z"></path></g></svg>
                <div class="sectionLabelInfo">
                    <p class="sectionLabel-primaryInfo">
                        $($Printer.Printer)
                    </p>
                    <p class="sectionLabel-secondaryInfo">
                        $($Printer.PtrIP)
                    </p>
                </div>
            </div>
            <div class="supplyContainer">
                <div class="supplyItem">
                    <svg class="supplyIcon" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px" viewBox="0 0 100 100" enable-background="new 0 0 100 100" xml:space="preserve">
                        <path opacity="0.5" fill="#231F20" enable-background="new" d="m50.271 1.24c-10.731 0-19.265 3.695-25.363 10.982-5.977 7.142-9.007 17.307-9.007 30.214v15.067c0 12.931 2.928 23.105 8.701 30.24 5.915 7.31 14.468 11.017 25.422 11.017 10.344 0 18.529-2.918 24.329-8.672 5.771-5.726 8.941-14.144 9.423-25.022l.128-2.882h-2.884-18.08-2.65l-.107 2.648c-.275 6.791-1.509 9.726-2.495 10.992-1.246 1.599-3.824 2.41-7.662 2.41-3.83 0-6.405-1.19-7.871-3.639-1.182-1.973-2.591-6.508-2.591-16.908v-16.912c.114-7.181 1.06-12.34 2.811-15.287.949-1.598 2.8-3.723 7.897-3.723 3.813 0 6.368.836 7.594 2.484.982 1.321 2.192 4.388 2.377 11.493l.07 2.688h2.689 18.142 2.955l-.202-2.948c-.739-10.809-3.892-19.281-9.371-25.182-5.586-6.012-13.746-9.06-24.255-9.06z"></path>
                        <defs>
                            <filter id="Adobe_OpacityMaskFilter_c" filterUnits="userSpaceOnUse" x="0" y="0" width="100" height="100">
                                <feColorMatrix type="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 1 0"></feColorMatrix>
                            </filter>
                        </defs>
                        <mask maskUnits="userSpaceOnUse" x="0" y="0" width="100" height="100" id="SVGID_c">
                            <g filter="url(#Adobe_OpacityMaskFilter_c)">
                                <path fill="#FFFFFF" d="m72.508 88.128c-5.289 5.248-12.752 7.872-22.385 7.872-10.127 0-17.886-3.331-23.277-9.993s-8.087-16.164-8.087-28.504v-15.067c0-12.299 2.788-21.78 8.364-28.443 5.575-6.662 13.324-9.993 23.246-9.993 9.756 0 17.167 2.727 22.231 8.179 5.063 5.453 7.942 13.283 8.64 23.492h-18.141c-.164-6.313-1.138-10.67-2.921-13.068s-5.053-3.598-9.809-3.598c-4.838 0-8.262 1.691-10.27 5.074-2.009 3.382-3.075 8.948-3.198 16.697v16.912c0 8.897.994 15.005 2.983 18.326 1.988 3.321 5.402 4.981 10.239 4.981 4.756 0 8.034-1.158 9.84-3.475 1.804-2.316 2.829-6.508 3.075-12.576h18.08c-.453 10.208-3.321 17.937-8.61 23.184z"></path>
                            </g>
                        </mask>
                        <rect id="fill-level-c" y="$(100 - $printer.Toner.Where({$_.Supply -match "Cyan Cartridge" -or $_.Supply -match "Cyan Toner"}).Percent)" mask="url(#SVGID_c)" fill="var(--cyan-500)" width="100" height="100"></rect>
                    </svg>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Toner.Where({$_.Supply -match "Cyan Cartridge" -or $_.Supply -match "Cyan Toner"}).Percent)){"Volume: " + ($printer.Toner.Where({$_.Supply -match "Cyan Cartridge" -or $_.Supply -match "Cyan Toner"}).Percent).ToString() + '%'}else{"Not Installed"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Toner.Where({$_.Supply -match "Cyan Cartridge" -or $_.Supply -match "Cyan Toner"}).PagesLeft)){"Pages: " + ($printer.Toner.Where({$_.Supply -match "Cyan Cartridge" -or $_.Supply -match "Cyan Toner"}).PagesLeft).ToString()}else{"-"}
                        )
                    </span>
                </div>
                <div class="supplyItem">
                    <svg class="supplyIcon" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px" viewBox="0 0 100 100" enable-background="new 0 0 100 100" xml:space="preserve">
                        <path opacity="0.5" fill="#231F20" enable-background="new" d="m93.2553 2.4699h-2.76-23.615-2.1337l-.5373 2.065-14.1841 54.5104-14.2364-54.5128-.5387-2.0626h-2.1318-23.6136-2.76v2.76 89.5401 2.76h2.76 18.0802 2.76v-2.76-24.2299-.0614l-.0027-.0613-.6366-14.311 11.463 39.434.5783 1.9896h2.0719 12.2995 2.0719l.5784-1.9895 11.4644-39.436-.6367 14.3129-.0027.0613v.0614 24.2299 2.76h2.76 18.1417 2.76v-2.76-89.5401z"></path>
                        <defs>
                            <filter id="Adobe_OpacityMaskFilter_m" filterUnits="userSpaceOnUse" x="0" y="0" width="100" height="100">
                                <feColorMatrix type="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 1 0"></feColorMatrix>
                            </filter>
                        </defs>
                        <mask maskUnits="userSpaceOnUse" x="0" y="0" width="100" height="100" id="SVGID_m">
                            <g filter="url(#Adobe_OpacityMaskFilter_m)">
                                <path fill="#FFFFFF" d="m50.0301 69.9866 16.8503-64.7567h23.615v89.5401h-18.1418v-24.2299l1.6604-37.3289-17.8957 61.5589h-12.2994l-17.8945-61.5589 1.6604 37.3289v24.2299h-18.0801v-89.5401h23.6137z"></path>
                            </g>
                        </mask>
                        <rect id="fill-level-m" y="$(100 - $printer.Toner.Where({$_.Supply -match "Magenta Cartridge" -or $_.Supply -match "Magenta Toner"}).Percent)" mask="url(#SVGID_m)" fill="var(--magenta-500)" width="100" height="100"></rect>
                    </svg>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Toner.Where({$_.Supply -match "Magenta Cartridge" -or $_.Supply -match "Magenta Toner"}).Percent)){"Volume: " + ($printer.Toner.Where({$_.Supply -match "Magenta Cartridge" -or $_.Supply -match "Magenta Toner"}).Percent).ToString() + '%'}else{"Not Installed"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Toner.Where({$_.Supply -match "Magenta Cartridge" -or $_.Supply -match "Magenta Toner"}).PagesLeft)){"Pages: " + ($printer.Toner.Where({$_.Supply -match "Magenta Cartridge" -or $_.Supply -match "Magenta Toner"}).PagesLeft).ToString()}else{"-"}
                        )
                    </span>
                </div>
                <div class="supplyItem">
                    <svg class="supplyIcon" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px" viewBox="0 0 100 100" enable-background="new 0 0 100 100" xml:space="preserve">
                        <path opacity="0.5" fill="#231F20" enable-background="new" d="m87.3336 2.4699h-4.1537-19.6792-1.9864l-.6307 1.8836-10.8567 32.4214-10.9075-32.4249-.6324-1.88h-1.9835-19.6792-4.1579l1.6144 3.8316 23.8289 56.5556v31.9129 2.76h2.76 18.3262 2.76v-2.76-31.9142l23.7684-56.5566z"></path>
                        <defs>
                            <filter id="Adobe_OpacityMaskFilter_y" filterUnits="userSpaceOnUse" x="0" y="0" width="100" height="100">
                                <feColorMatrix type="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 1 0"></feColorMatrix>
                            </filter>
                        </defs>
                        <mask maskUnits="userSpaceOnUse" x="0" y="0" width="100" height="100" id="SVGID_y">
                            <g filter="url(#Adobe_OpacityMaskFilter_y)">
                                <path fill="#FFFFFF" d="m63.4987 5.2299h19.6791l-23.984 57.0695v32.4706h-18.3262v-32.4705l-24.0454-57.0696h19.6791l13.5294 40.2193z"></path>
                            </g>
                        </mask>
                        <rect id="fill-level-y" y="$(100 - $printer.Toner.Where({$_.Supply -match "Yellow Cartridge" -or $_.Supply -match "Yellow Toner"}).Percent)" mask="url(#SVGID_y)" fill="var(--yellow-500)" width="100" height="100"></rect>
                    </svg>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Toner.Where({$_.Supply -match "Yellow Cartridge" -or $_.Supply -match "Yellow Toner"}).Percent)){"Volume: " + ($printer.Toner.Where({$_.Supply -match "Yellow Cartridge" -or $_.Supply -match "Yellow Toner"}).Percent).ToString() + '%'}else{"Not Installed"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Toner.Where({$_.Supply -match "Yellow Cartridge" -or $_.Supply -match "Yellow Toner"}).PagesLeft)){"Pages: " + ($printer.Toner.Where({$_.Supply -match "Yellow Cartridge" -or $_.Supply -match "Yellow Toner"}).PagesLeft).ToString()}else{"-"}
                        )
                    </span>
                </div>
                <div class="supplyItem">
                    <svg class="supplyIcon" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px" viewBox="0 0 100 100" enable-background="new 0 0 100 100" xml:space="preserve">
                        <path opacity="0.5" fill="#231F20" enable-background="new" d="m89.8463 2.2299h-5.5353-22.139-1.7478l-.862 1.5204-16.6512 29.3701-.1241.2094v-28.0999-3h-3-18.0802-3v3 89.5401 3h3 18.0802 3v-3-24.5361l3.4841-4.3852 14.2962 30.2048.8125 1.7166h1.8991 21.4626 4.9396l-2.2776-4.3832-25.2674-48.6281 24.6885-37.8911z"></path>
                        <defs>
                            <filter id="Adobe_OpacityMaskFilter_k" filterUnits="userSpaceOnUse" x="0" y="0" width="100" height="100">
                                <feColorMatrix type="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 1 0"></feColorMatrix>
                            </filter>
                        </defs>
                        <mask maskUnits="userSpaceOnUse" x="0" y="0" width="100" height="100" id="SVGID_k">
                            <g filter="url(#Adobe_OpacityMaskFilter_k)">
                                <path fill="#FFFFFF" d="m39.7869 69.1872v25.5829h-18.0802v-89.5402h18.0802v39.0508l5.7193-9.6551 16.6657-29.3957h22.139l-25.6443 39.3583 26.0749 50.1818h-21.4626l-16.3583-34.5614z"></path>
                            </g>
                        </mask>
                        <rect id="fill-level-k" y="$(100 - $printer.Toner.Where({$_.Supply -match "Black Cartridge" -or $_.Supply -match "Black Toner"}).Percent)" mask="url(#SVGID_k)" fill="var(--jet-black-500)" width="100" height="100"></rect>
                    </svg>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Toner.Where({$_.Supply -match "Black Cartridge" -or $_.Supply -match "Black Toner"}).Percent)){"Volume: " + ($printer.Toner.Where({$_.Supply -match "Black Cartridge" -or $_.Supply -match "Black Toner"}).Percent).ToString() + '%'}else{"Not Installed"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Toner.Where({$_.Supply -match "Black Cartridge" -or $_.Supply -match "Black Toner"}).PagesLeft)){"Pages: " + ($printer.Toner.Where({$_.Supply -match "Black Cartridge" -or $_.Supply -match "Black Toner"}).PagesLeft).ToString()}else{"-"}
                        )
                    </span>
                </div>
            </div>

            <div class="supplyContainer">
                <div class="supplyItem">
                    <svg class="supplyIcon" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px" viewBox="0 0 100 100" enable-background="new 0 0 100 100" xml:space="preserve">
                        <path opacity="0.5" fill="#231F20" enable-background="new" d="m66.01 2.4084h-2.76-1.8449-.5422l-.5019.2052-32.6551 13.3449-1.7159.7012v1.8537 14.5134 4.0249l3.7545-1.4503 13.2802-5.1299v64.3601 2.76h2.76 17.4653 2.76v-2.76-89.6632z"></path>
                        <defs>
                            <filter id="Adobe_OpacityMaskFilter_1" filterUnits="userSpaceOnUse" x="0" y="0" width="100" height="100">
                                <feColorMatrix type="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 1 0"></feColorMatrix>
                            </filter>
                        </defs>
                        <mask maskUnits="userSpaceOnUse" x="0" y="0" width="100" height="100" id="SVGID_1">
                            <g filter="url(#Adobe_OpacityMaskFilter_1)">
                                <path fill="#FFFFFF" d="m45.7848 94.8316v-68.385l-17.0348 6.5801v-14.5133l32.6551-13.345h1.8449v89.6631h-17.4652z"></path>
                            </g>
                        </mask>
                        <rect id="fill-level-1" y="$(100 - $printer.Paper.Where({$_.Tray -match "Tray 1"}).Percent)" mask="url(#SVGID_1)" fill="#ffffff" width="100" height="100"></rect>
                    </svg>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 1"}).Percent)){($printer.Paper.Where({$_.Tray -match "Tray 1"}).Percent).ToString() + '%'}else{"Not Installed"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 1"}).Remaining)){($printer.Paper.Where({$_.Tray -match "Tray 1"}).Remaining).ToString()}else{"-"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 1"}).Size)){($printer.Paper.Where({$_.Tray -match "Tray 1"}).Size).ToString()}else{"-"}
                        )
                    </span>
                </div>
                <div class="supplyItem">
                    <svg class="supplyIcon" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px" viewBox="0 0 100 100" enable-background="new 0 0 100 100" xml:space="preserve">
                        <path opacity="0.5" fill="#231F20" enable-background="new" d="m49.877 1.855c-5.6983 0-10.8902 1.409-15.4313 4.1879-4.5486 2.7835-8.1222 6.7294-10.6216 11.7282-2.4581 4.9162-3.7045 10.4007-3.7045 16.3011v2.76h2.76 17.4652 2.76v-2.76c0-4.2028.694-7.4305 2.0629-9.5936 1.1607-1.8338 2.5183-2.6517 4.4017-2.6517 1.6285 0 2.6984.6156 3.5775 2.0584 1.0995 1.8052 1.6572 4.4672 1.6572 7.9116 0 2.5432-.652 5.2745-1.938 8.1181-1.367 3.0234-3.5489 6.6229-6.4857 10.6996l-24.1944 30.765-.5905.7509v.9552 12.2993 2.76h2.76 52.7647 2.76v-2.76-14.4518-2.76h-2.76-25.1384l8.3874-11.9756c6.4791-7.6933 11.041-14.3093 13.5602-19.666 2.6021-5.5322 3.9215-11.0897 3.9215-16.5182 0-8.8383-2.4962-15.8159-7.4193-20.739-4.9239-4.9232-11.8394-7.4194-20.5546-7.4194z"></path>
                        <defs>
                            <filter id="Adobe_OpacityMaskFilter_2" filterUnits="userSpaceOnUse" x="0" y="0" width="100" height="100">
                                <feColorMatrix type="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 1 0"></feColorMatrix>
                            </filter>
                        </defs>
                        <mask maskUnits="userSpaceOnUse" x="0" y="0" width="100" height="100" id="SVGID_2">
                            <g filter="url(#Adobe_OpacityMaskFilter_2)">
                                <path fill="#FFFFFF" d="m24.3556 95.385v-12.2994l24.2299-30.8102c3.0749-4.2638 5.34-8.0043 6.7955-11.2233 1.4554-3.2184 2.1832-6.3035 2.1832-9.2553 0-3.9762-.6867-7.0927-2.0602-9.3476-1.3734-2.2543-3.3516-3.3824-5.9345-3.3824-2.8289 0-5.0735 1.3126-6.734 3.9358-1.6604 2.6239-2.4906 6.3137-2.4906 11.0695h-17.4652c0-5.4938 1.1377-10.516 3.4131-15.0668s5.4733-8.0869 9.5936-10.6083 8.7838-3.782 13.9906-3.782c7.9947 0 14.1956 2.2043 18.6029 6.611 4.4073 4.4073 6.611 10.6698 6.611 18.7874 0 5.0428-1.2197 10.1573-3.6591 15.3436-2.4394 5.1869-6.8569 11.5717-13.2527 19.1564l-11.5 16.4198h30.4412v14.4518z"></path>
                            </g>
                        </mask>
                        <rect id="fill-level-2" y="$(100 - $printer.Paper.Where({$_.Tray -match "Tray 2"}).Percent)" mask="url(#SVGID_2)" fill="#ffffff" width="100" height="100"></rect>
                    </svg>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 2"}).Percent)){($printer.Paper.Where({$_.Tray -match "Tray 2"}).Percent).ToString() + '%'}else{"Not Installed"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 2"}).Remaining)){($printer.Paper.Where({$_.Tray -match "Tray 2"}).Remaining).ToString()}else{"-"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 2"}).Size)){($printer.Paper.Where({$_.Tray -match "Tray 2"}).Size).ToString()}else{"-"}
                        )
                    </span>
                </div>
                <div class="supplyItem">
                    <svg class="supplyIcon" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px" viewBox="0 0 100 100" enable-background="new 0 0 100 100" xml:space="preserve">
                        <path opacity="0.5" fill="#231F20" enable-background="new" d="m49.262 1.24c-5.1554 0-9.9312 1.1666-14.1945 3.4675-4.3112 2.3267-7.7149 5.6056-10.1167 9.7458-2.3896 4.12-3.6012 8.8171-3.6012 13.9611v2.76h2.76 17.4037 2.76v-2.76c0-2.2356.5513-3.9986 1.6854-5.3898 1.0085-1.2367 2.2145-1.8128 3.7952-1.8128 1.719 0 2.989.6072 3.997 1.9108 1.1476 1.4847 1.7296 3.7204 1.7296 6.6447 0 3.1568-.5982 5.6505-1.7779 7.412-.6116.9135-1.769 2.1275-4.7482 2.1275h-8.4251-2.76v2.76 14.0214 2.76h2.76 8.5481c3.4877.0209 7.5101 1.236 7.5101 10.5234 0 3.0383-.6805 5.4435-2.0225 7.1487-1.2177 1.5472-2.7856 2.2678-4.9341 2.2678-1.7192 0-3.0578-.6953-4.2127-2.1882-1.3084-1.6913-1.9444-3.7741-1.9444-6.3674v-2.76h-2.76-17.4036-2.76v2.76c0 8.4349 2.6853 15.3739 7.9812 20.6244 5.2894 5.244 12.2228 7.9029 20.6076 7.9029 8.9319 0 16.2847-2.6753 21.854-7.9515 5.6115-5.3162 8.4568-12.4251 8.4568-21.1293 0-5.385-1.1681-10.1508-3.4718-14.1652-1.6089-2.8032-3.7963-5.1443-6.5304-6.9961 2.0973-1.7373 3.9-3.836 5.3865-6.2755 2.369-3.887 3.5703-8.167 3.5703-12.721 0-8.6486-2.6395-15.6039-7.8453-20.6726-5.1853-5.0487-12.3506-7.6086-21.2971-7.6086z"></path>
                        <defs>
                            <filter id="Adobe_OpacityMaskFilter_3" filterUnits="userSpaceOnUse" x="0" y="0" width="100" height="100">
                                <feColorMatrix type="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 1 0"></feColorMatrix>
                            </filter>
                        </defs>
                        <mask maskUnits="userSpaceOnUse" x="0" y="0" width="100" height="100" id="SVGID_3">
                            <g filter="url(#Adobe_OpacityMaskFilter_3)">
                                <path fill="#FFFFFF" d="m48.9545 42.0668c3.1979 0 5.545-1.1166 7.0414-3.3516 1.4964-2.2344 2.2447-5.217 2.2447-8.9479 0-3.5668-.7687-6.3438-2.3062-8.3329-1.5374-1.9884-3.5976-2.9826-6.1805-2.9826-2.4189 0-4.3971.9436-5.9345 2.8289-1.5374 1.8859-2.3061 4.2645-2.3061 7.1337h-17.4037c0-4.6738 1.0762-8.8652 3.2286-12.5762 2.1524-3.7103 5.1658-6.611 9.0401-8.7019s8.1689-3.1363 12.8837-3.1363c8.2406 0 14.6979 2.2754 19.3717 6.8262s7.0107 10.7832 7.0107 18.6952c0 4.0588-1.0557 7.8204-3.1671 11.2848-2.1114 3.465-4.889 6.119-8.3329 7.9639 4.2228 1.8046 7.3694 4.5104 9.4398 8.1176 2.0704 3.6078 3.1056 7.8717 3.1056 12.7914 0 7.9537-2.5316 14.3289-7.5949 19.1257s-11.7152 7.1952-19.9559 7.1952c-7.6667 0-13.8881-2.3676-18.6644-7.1029s-7.1644-10.9568-7.1644-18.6644h17.4037c0 3.1979.8405 5.8832 2.5214 8.0562s3.8128 3.2594 6.3957 3.2594c2.9929 0 5.3605-1.1069 7.1029-3.3209 1.7424-2.2139 2.6136-5.1658 2.6136-8.8556 0-8.8146-3.4234-13.2424-10.2701-13.2834h-8.5481v-14.0216z"></path>
                            </g>
                        </mask>
                        <rect id="fill-level-3" y="$(100 - $printer.Paper.Where({$_.Tray -match "Tray 3"}).Percent)" mask="url(#SVGID_3)" fill="#ffffff" width="100" height="100"></rect>
                    </svg>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 3"}).Percent)){($printer.Paper.Where({$_.Tray -match "Tray 3"}).Percent).ToString() + '%'}else{"Not Installed"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 3"}).Remaining)){($printer.Paper.Where({$_.Tray -match "Tray 3"}).Remaining).ToString()}else{"-"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 3"}).Size)){($printer.Paper.Where({$_.Tray -match "Tray 3"}).Size).ToString()}else{"-"}
                        )
                    </span>
                </div>
                <div class="supplyItem">
                    <svg class="supplyIcon" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px" viewBox="0 0 100 100" enable-background="new 0 0 100 100" xml:space="preserve">
                        <path opacity="0.5" fill="#231F20" enable-background="new" d="m69.8936 2.4699h-2.76-17.4653-1.6821l-.7709 1.495-30.3797 58.9145-.3549.6882.0549.7724.7995 11.254.1822 2.5644h2.5709 26.8817v16.6116 2.76h2.76 17.4037 2.76v-2.76-16.6116h4.8657 2.76v-2.76-14.4519-2.76h-2.76-4.8657v-52.9566zm-29.4262 55.7166 6.5024-14.4634v14.4634z"></path>
                        <defs>
                            <filter id="Adobe_OpacityMaskFilter_4" filterUnits="userSpaceOnUse" x="0" y="0" width="100" height="100">
                                <feColorMatrix type="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 1 0"></feColorMatrix>
                            </filter>
                        </defs>
                        <mask maskUnits="userSpaceOnUse" x="0" y="0" width="100" height="100" id="SVGID_4">
                            <g filter="url(#Adobe_OpacityMaskFilter_4)">
                                <path fill="#FFFFFF" d="m74.7353 60.9465v14.4519h-7.6257v19.3717h-17.4037v-19.3717h-29.6417l-.7995-11.254 30.3797-58.9145h17.4652v55.7166zm-25.0294 0v-29.9492l-.369.6765-13.1604 29.2727z"></path>
                            </g>
                        </mask>
                        <rect id="fill-level-4" y="$(100 - $printer.Paper.Where({$_.Tray -match "Tray 4"}).Percent)" mask="url(#SVGID_4)" fill="#ffffff" width="100" height="100"></rect>
                    </svg>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 4"}).Percent)){($printer.Paper.Where({$_.Tray -match "Tray 4"}).Percent).ToString() + '%'}else{"Not Installed"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 4"}).Remaining)){($printer.Paper.Where({$_.Tray -match "Tray 4"}).Remaining).ToString()}else{"-"}
                        )
                    </span>
                    <span class="supplyInfo">
                        $(
                            if($null -ne ($printer.Paper.Where({$_.Tray -match "Tray 4"}).Size)){($printer.Paper.Where({$_.Tray -match "Tray 4"}).Size).ToString()}else{"-"}
                        )
                    </span>
                </div>
            </div>
        </div>

        <div class="Container alertContainer">
            <div class="sectionLabel alertLabel">
                <svg class="svg" height="50px" width="50px" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 507.425 507.425" xml:space="preserve" fill="#000000"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <path style="fill:#ffc52f;" d="M329.312,129.112l13.6,22l150.8,242.4c22.4,36,6,65.2-36.8,65.2h-406.4c-42.4,0-59.2-29.6-36.8-65.6 l198.8-320.8c22.4-36,58.8-36,81.2,0L329.312,129.112z"></path> <g> <path style="fill:#000000;" d="M253.712,343.512c-10.8,0-20-8.8-20-20v-143.2c0-10.8,9.2-20,20-20s20,8.8,20,20v143.2 C273.712,334.312,264.512,343.512,253.712,343.512z"></path> <path style="fill:#000000;" d="M253.312,407.112c-5.2,0-10.4-2-14-6c-3.6-3.6-6-8.8-6-14s2-10.4,6-14c3.6-3.6,8.8-6,14-6 s10.4,2,14,6c3.6,3.6,6,8.8,6,14s-2,10.4-6,14C263.712,404.712,258.512,407.112,253.312,407.112z"></path> </g> <path d="M456.912,465.512h-406.4c-22,0-38.4-7.6-46-21.6s-5.6-32.8,6-51.2l198.8-321.6c11.6-18.8,27.2-29.2,44.4-29.2l0,0 c16.8,0,32.4,10,43.6,28.4l35.2,56.4l0,0l13.6,22l150.8,243.6c11.6,18.4,13.6,37.2,6,51.2 C495.312,457.912,478.912,465.512,456.912,465.512z M253.312,49.912L253.312,49.912c-14,0-27.2,8.8-37.6,25.2l-198.8,321.6 c-10,16-12,31.6-5.6,43.2s20.4,17.6,39.2,17.6h406.4c18.8,0,32.8-6.4,39.2-17.6c6.4-11.2,4.4-27.2-5.6-43.2l-150.8-243.6l-13.6-22 l-35.2-56.4C280.512,58.712,267.312,49.912,253.312,49.912z"></path> <path d="M249.712,347.512c-13.2,0-24-10.8-24-24v-143.2c0-13.2,10.8-24,24-24s24,10.8,24,24v143.2 C273.712,336.712,262.912,347.512,249.712,347.512z M249.712,164.312c-8.8,0-16,7.2-16,16v143.2c0,8.8,7.2,16,16,16s16-7.2,16-16 v-143.2C265.712,171.512,258.512,164.312,249.712,164.312z"></path> <path d="M249.712,411.112L249.712,411.112c-6.4,0-12.4-2.4-16.8-6.8c-4.4-4.4-6.8-10.8-6.8-16.8c0-6.4,2.4-12.4,6.8-16.8 c4.4-4.4,10.8-7.2,16.8-7.2c6.4,0,12.4,2.4,16.8,7.2c4.4,4.4,7.2,10.4,7.2,16.8s-2.4,12.4-7.2,16.8 C262.112,408.312,256.112,411.112,249.712,411.112z M249.712,371.112c-4,0-8.4,1.6-11.2,4.8c-2.8,2.8-4.8,7.2-4.8,11.2 c0,4.4,1.6,8.4,4.8,11.2c2.8,2.8,7.2,4.8,11.2,4.8s8.4-1.6,11.2-4.8c2.8-2.8,4.8-7.2,4.8-11.2s-1.6-8.4-4.8-11.2 C258.112,372.712,253.712,371.112,249.712,371.112z"></path> </g></svg>
                <div class="sectionLabelInfo">
                    <div class="sectionLabel-primaryInfo">
                        Alerts
                    </div>
                </div>
            </div>
            <div class="alertTable">
                <table style="width: calc(100% - 2em); margin: auto 1em">
                    <tr>
                        <th>Description</th>
                    </tr>
                    $(
                        foreach($alert in $printer.Alerts){
                            "<tr><td>$($alert.AlertMessage)<td><tr>"
                        }
                    )
                </table>
            </div>
        </div>

    </div>
"@

#Insert the skeleton into the body.
$htmlBody += $htmlInsert
}

#Add the head and feet
$htmlStitched = $htmlStart + $htmlBody + $htmlEnd

#Kick it out the door to live elsewhere
$htmlStitched | Out-File -FilePath $PSScriptRoot"\PrinterReport.html" -Encoding utf8

#Make it earn it's keep
Start-Process $PSScriptRoot"\PrinterReport.html"