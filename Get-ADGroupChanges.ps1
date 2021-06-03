Function Get-ADGroupChanges
{
<#
.SYNOPSIS
    Returns change history in the membership of an Active Directory Group

    ADGroupChanges Function: Get-ADGroupChanges
    Author: Y1nTh35h311 (yossis@protonmail.com, #Yossi_Sassi)
    License: BSD 3-Clause
    Required Dependencies: None
    Optional Dependencies: None

.DESCRIPTION
    "Pure" powershell command (no dependencies, no special permissions etc') to retrieve change history in an AD group membership. relies on object metadata rather than event logs. useful for DF/IR, tracking changes in groups etc'.

.PARAMETER Domain
    [string]
    Domain to be queried (optional. defaults to current domain).

.PARAMETER PageSize
    [int]
    Optional. The PageSize to set for the LDAP searcher object. Default is 1000.

.PARAMETER Output
    [string[]]
    Output to CSV file or Grid/Open results in a form (uses PowerShell ISE). by default - outputs to screen.

.EXAMPLE
This command will get the list of historic changes for the group "Domain Admins", sorted by Last change date.

PS C:\temp> Get-ADGroupChanges -GroupName 'domain admins'

GroupName            : Domain Admins
GroupDN              : CN=Domain Admins,CN=Users,DC=Adatum,DC=com
MemberSamAccountName : AdatumAdmin
LastChange           : Added
MemberDN             : CN=AdatumAdmin,CN=Users,DC=Adatum,DC=com
MemberAdminCount     : 1
DateTimeAdded            : 10/18/2016 1:00:30 PM
DateTimeRemoved          : N/A
LastChangeDateTime   : 10/18/2016 1:00:30 PM
DaysSinceLastChange  : 1554

GroupName            : Domain Admins
GroupDN              : CN=Domain Admins,CN=Users,DC=Adatum,DC=com
MemberSamAccountName : Administrator
LastChange           : Added
MemberDN             : CN=Administrator,CN=Users,DC=Adatum,DC=com
MemberAdminCount     : 1
DateTimeAdded            : 10/18/2016 12:48:46 PM
DateTimeRemoved          : N/A
LastChangeDateTime   : 10/18/2016 12:48:46 PM
DaysSinceLastChange  : 1554

Found changes information for 2 accounts in group <CN=Test,OU=IT,DC=Adatum,DC=com>

.EXAMPLE
This command will get the list of historic changes for the group "Test", and save the output to a CSV file.

PS C:\temp> Get-ADGroupChanges -GroupName test -Output CSV

GroupName            : Test
GroupDN              : CN=Test,OU=IT,DC=Adatum,DC=com
MemberSamAccountName : Adam
LastChange           : Removed
MemberDN             : CN=Adam Hobbs,OU=Managers,DC=Adatum,DC=com
MemberAdminCount     : 1
DateTimeAdded            : 1/18/2021 2:24:57 AM
DateTimeRemoved          : 2021-01-16T10:26:25Z
LastChangeDateTime   : 1/18/2021 4:16:25 AM
DaysSinceLastChange  : 2

GroupName            : Test
GroupDN              : CN=Test,OU=IT,DC=Adatum,DC=com
MemberSamAccountName : Anete
LastChange           : Added
MemberDN             : CN=Anete bonnet,OU=Development,DC=Adatum,DC=com
MemberAdminCount     : 
DateTimeAdded            : 1/19/2021 2:24:57 AM
DateTimeRemoved          : N/A
LastChangeDateTime   : 1/19/2021 2:24:57 AM
DaysSinceLastAction  : 1

Found changes information for 2 accounts in group <CN=Test,OU=IT,DC=Adatum,DC=com>
Results saved to GroupMembershipChanges_ADATUM_TEST.csv.

.LINK
www.10root.com
#>
    param(
        [cmdletbinding()]
        [Parameter(Mandatory = $false)]
        [string]$Domain,

        [Parameter(Mandatory = $false)]
        [int]$PageSize = 1000,

        [Parameter(Mandatory = $true)]
        [string]$GroupName,

        [ValidateSet("Grid", "CSV", IgnoreCase = $true)]
        [String[]]$Output
    )

    if (!$Domain)
        {
            $Domain = $env:USERDOMAIN    
        }

    $DN = (Get-ADDomain -Server $Domain).distinguishedname;
    $DomainObj = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DN");

    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $DomainObj
    $ObjSearcher.PageSize = $PageSize; $objSearcher.SizeLimit = $PageSize
    $ObjSearcher.Filter = "(&(objectClass=group)(name=$GroupName))"
    $ObjSearcher.PropertiesToLoad.AddRange(("AdminCount","CanonicalName", "DistinguishedName", "Description", "GroupType","samaccountname", "SidHistory", "ManagedBy", "msDS-ReplValueMetaData", "ObjectSID", "WhenCreated", "WhenChanged"))

    $GroupObj = $ObjSearcher.FindOne()
    $ObjSearcher.dispose()

    If ($GroupObj)
    {    
        $ReplMetaData = $GroupObj.Properties.'msds-replvaluemetadata' #'msDS-ReplValueMetaData'

        $GroupMembershipChanges = @();

        $ReplMetaData | foreach {
            [xml]$ReplValue = ""
            $ReplValue.LoadXml($_.Replace("\x00", "").Replace("&","&amp;"))
    
            $LastActionDate = $ReplValue.DS_REPL_VALUE_META_DATA.ftimeLastOriginatingChange
            [int]$DaysSinceLastAction = ($(get-date) - [datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeLastOriginatingChange).Days
    
            if ($ReplValue.DS_REPL_VALUE_META_DATA.ftimeDeleted -eq "1601-01-01T00:00:00Z")
                {
                    $DateTimeRemoved = "N/A"        
                }
                else
                {
                    $DateTimeRemoved = $ReplValue.DS_REPL_VALUE_META_DATA.ftimeDeleted
                }

            if ([datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeCreated -gt [datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeDeleted)
                {
                    $LastAction = "Added"
                }
                else
                {
                    $LastAction = "Removed"
                }

            $ObjMember = [adsi]"LDAP://$($ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn)"
            $MemberSamAccountName = $ObjMember.Properties.samaccountname -join ','
            $MemberAdminCount = $ObjMember.Properties.admincount -join ','

	    $Enabled = if (([adsisearcher]"(samaccountname=$MemberSamAccountName)").FindOne().Properties.useraccountcontrol -eq 514) {"False"} else {"True"}

            $GroupChangeObj = New-Object PSObject;
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "GroupName" -Value $($GroupObj.Properties.samaccountname -join ',') -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "GroupDN" -Value $($GroupObj.Properties.distinguishedname -join ',') -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "MemberSamAccountName" -Value $MemberSamAccountName -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "Enabled" -Value $Enabled -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "LastChange" -Value $LastAction -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "MemberDN" -Value $ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "MemberAdminCount" -Value $MemberAdminCount -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DateTimeAdded" -Value $([datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeCreated) -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DateTimeRemoved" -Value $DateTimeRemoved -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "LastChangeDateTime" -Value $([datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeLastOriginatingChange) -Force
            Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DaysSinceLastChange" -Value $DaysSinceLastAction -Force

            $GroupMembershipChanges += $GroupChangeObj
            Clear-Variable GroupChangeObj
        }

    }

    else
    {
        Write-Warning "Error while getting the Group Object. Make sure you typed the group name correctly."
        break
    }

    # Wrap up
    $GroupMembershipChanges | Sort-Object LastChangeDateTime -Descending

    Write-Host "Found changes information for $($ReplMetaData.count) accounts in group <$($GroupObj.Properties.distinguishedname -join ',')>" -ForegroundColor Cyan

    if ($Output -eq "Grid") {
            $GroupMembershipChanges | Sort-Object LastChangeDateTime -Descending | Out-GridView -Title "Group membership changes for group $($($GroupObj.Properties.samaccountname -join ',').ToUpper()) in domain $(($Domain).ToUpper())"
        }
    
    if ($Output -eq "CSV") {
            $FileName = "GroupMembershipChanges_$($Domain)_$($GroupName.ToUpper()).csv"
            $GroupMembershipChanges | Export-Csv $FileName -NoTypeInformation
            "Results saved to $FileName."
        }
}