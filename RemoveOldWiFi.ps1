<#
Some profiles returned by netsh.exe wlan show profile, can't seem to be deleted:
Profile "IBIS" is not found on any interface.

Sames goes for special characters:
Profile "Søren Something's iPhon" is not found on any interface.

Both appear in, and can be deleted from, GUI
#>

Param (

    [Parameter(Mandatory=$False)]
    [int]$DaysThreshold = 1200,

    [Parameter(Mandatory=$False)]
    [ValidateSet('List', 'DeleteOld')]
    $Action = 'List'

)


Function Convert-WiFiDate {

    Param (
        [int[]]$DateLittleEndian
    )

    $Year = [int]"0x$(($DateLittleEndian[1,0] | ForEach {'{0:x2}' -f $_}) -join '')"
    $Month = '{0:D2}' -f [int]"0x$(($DateLittleEndian[3,2] | ForEach {'{0:x2}' -f $_}) -join '')"
    $Day = '{0:D2}' -f [int]"0x$(($DateLittleEndian[7,6] | ForEach {'{0:x2}' -f $_}) -join '')"
    $Hour = '{0:D2}' -f [int]"0x$(($DateLittleEndian[9,8] | ForEach {'{0:x2}' -f $_}) -join '')"
    $Minute = '{0:D2}' -f [int]"0x$(($DateLittleEndian[11,10] | ForEach {'{0:x2}' -f $_}) -join '')"
    $Second = '{0:D2}' -f [int]"0x$(($DateLittleEndian[13,12] | ForEach {'{0:x2}' -f $_}) -join '')"

    Get-Date "$($Year)-$($Month)-$($Day) $($Hour):$($Minute):$($Second)"

}


Function Get-KnownWiFiNetworks {

    $Result = (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles") | ForEach {

        $WiFiProperties = Get-ItemProperty -Path "Registry::$($_.Name)"
        
        If ($WiFiProperties.Description -ne 'Network' -and $WiFiProperties.NameType -eq 71) {
        
            [pscustomobject]@{
                Guid = $WiFiProperties.PSChildName
                ProfileName = $WiFiProperties.ProfileName
                Description = $WiFiProperties.Description
                DateCreated = Convert-WiFiDate $WiFiProperties.DateCreated
                DateLastConnected = Convert-WiFiDate $WiFiProperties.DateLastConnected
            }

        }

    } | Sort-Object Description

    # Configure a default display set
    $defaultDisplaySet = 'ProfileName', 'Description', 'DateCreated', 'DateLastConnected'

    # Create the default property display set
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultDisplaySet)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

    # Give this object a unique typename (just made up)
    $Result.PSObject.TypeNames.Insert(0,'WiFiConnections.List')
    $Result | Add-Member MemberSet PSStandardMembers $PSStandardMembers

    # Return results
    $Result

}


$KnownWiFiNetworks = Get-KnownWiFiNetworks

$Now = Get-Date


$WifiAge = netsh.exe wlan show profile | Where {$_ -like "*All User Profile*"} | Sort | ForEach {

    $WifiNetwork = $_.Split(':')[1].Trim()
    $LastConnected = $KnownWiFiNetworks | Where {$_.Description -eq $WifiNetwork} | Sort DateLastConnected -Descending | Select -First 1 -Expand DateLastConnected
    $DaysSinceLastConnected = If ($LastConnected) {New-TimeSpan -Start $LastConnected -End $Now | Select -Expand Days}

    [pscustomobject]@{
        WifiNetwork = $WifiNetwork
        LastConnected = $LastConnected
        DaysSinceLastConnected = $DaysSinceLastConnected
        OlderThanTreshold = If ($DaysSinceLastConnected -ge $DaysThreshold -or -Not($LastConnected)) {$true} Else {$False}
    }

}


If ($Action -eq 'DeleteOld') {

    $WifiAge | Where {$_.OlderThanTreshold} | Select -Expand WifiNetwork | ForEach {
        netsh.exe wlan delete profile name="$_"
    }

}

$WifiAge