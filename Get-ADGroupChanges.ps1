Function Get-ADGroupChanges
{
<#
.SYNOPSIS
    Gets Active Directory Groups' Membership Change History, regardless of AD Security Logs.

    ADGroupChanges Function: Get-ADGroupChanges
    Author: 1nTh35h311 (yossis@protonmail.com, #Yossi_Sassi)
    Version: 1.5.6

    Required Dependencies: None
    Optional Dependencies: None
    
    License: 
	Ten Root Cyber Security Ltd. End-User License Agreement 
	DOWNLOADING, INSTALLING, ACCESSING OR USING THE TOOLS WHICH PRESENTED IN BSIDESTLV 2021 BY TEN ROOT CYBER SECURITY LTD (THE "PRODUCT" or "SOFTWARE") CONSTITUTES EXPRESS ACCEPTANCE OF THIS AGREEMENT. TEN ROOT CYBER SECURITY IS WILLING TO LICENSE SOFTWARE TO  YOU ONLY IF YOU ACCEPT ALL OF THE TERMS CONTAINED IN THIS AGREEMENT (THE “EULA" or "AGREEMENT"). BY DOWNLOADING, INSTALLING, ACCESSING, OR USING THE SOFTWARE, USING THE PRODUCT OR OTHERWISE EXPRESSING YOUR AGREEMENT TO THE TERMS CONTAINED IN THE AGREEMENT, YOU INDIVIDUALLY AND ON BEHALF OF THE BUSINESS OR OTHER ORGANIZATION THAT YOU REPRESENT (THE “USER”) EXPRESSLY CONSENT TO BE BOUND BY THIS AGREEMENT. 
	This EULA governs User’s access to and use of any Software and/or any Product (as defined below) first placed in use by User on or after the release date of this EULA (the “Release Date”).
	1.	LICENSE GRANT
	A.	Software. Subject to the terms and conditions of this EULA, during the License Term, Ten Root Cyber Security grants User, and User accepts, upon delivery of any Software,  a non-exclusive, non-transferable, non-commercisl, royalty free, and non-sublicensable license to the Software to (i) allow Authorized Users to use such Software, in executable form only, and any accompanying Documentation, only for User’s internal use in connection with the Products, and subject to the terms in this EULA. 
	i.	Non-Commercial: User may not use the Software for commercial purposes. For the pupose of thos license, commercial purposes means that a thirs party has to pay in order to download' install, access' or use the Software. 
	B.	Ten Root Cyber Security Software Provisions.
	i.     Any use or operation of the Product, including the Software, with any product and/or mobile device developed, manufactured, produced, programmed, assembled and/or otherwise maintained by any person or entity shall be permitted only after the User has obtained any consents or approvals required (to the extent required) pursuant to applicable Law.
	ii.      UNDER NO CIRCUMSTANCES SHALL TEN ROOT CYBER SECURITY, ITS OFFICERS, EMPLOYEES OR REPRESENTATIVES BE LIABLE TO USER, USER OR ANY THIRD PARTY UNDER ANY CAUSE OF ACTION (WHETHER IN CONTRACT, TORT OR OTHERWISE) FOR ANY INCIDENTAL, SPECIAL, CONSEQUENTIAL, PUNITIVE, EXEMPLARY OR OTHER INDIRECT DAMAGES UNDER ANY LEGAL THEORY ARISING OUT OF OR RELATING TO THE USE OF ANY OF THE TEN ROOT CYBER SECURITY SOFTWARE IN CONNECTION WITH ANY PRODUCT AND/OR MOBILE DEVICE DEVELOPED, MANUFACTURED, PRODUCED, PROGRAMMED, ASSEMBLED AND/OR OTHERWISE MAINTAINED BY ANY PERSON OR ENTITY, WITHOUT OBTAINING EACH APPLICABLE CONSENT AND APPROVAL.
	iii.    No Obligation. Nothing in this EULA requires Ten Root Cyber Security to provide Updates or Upgrades to User or User to accept such Updates or Upgrades.
	C.	 License Prohibitions. Notwithstanding anything to the contrary, User shall not, alone, through a User, an Affiliate or a Third Party (or allow a User, an Affiliate or a Third Party to): (a) modify any Software; (b) reverse compile, reverse assemble, reverse engineer or otherwise translate all or any portion of any Software; (c) pledge, rent, lease, share, distribute, sell or create derivative works of any Software; (d) use any Software to provide service to any Third Party including use on a time sharing, service bureau, application service provider, software as a service, cloud services, rental or other similar basis; (e) make copies of any Software, except as provided for in the license grant above; (f) remove, alter or deface (or attempt any of the foregoing) proprietary notices, labels or marks in any Software; (g) distribute any copy of any Software to any Third Party, including without limitation selling any Product in a secondhand market; (h) use any Embedded Software other than with Products prov	ided by Ten Root Cyber Security or an authorized reseller of Ten Root Cyber Security or for more than the number of Products purchased from Ten Root Cyber Security or an authorized reseller of Ten Root Cyber Security; (i) disclose any results of testing or benchmarking of any Software to any Third Party; (j) use any Update or Upgrade beyond those to which User is entitled or with any Software to which User does not have a valid, current license; (k) deactivate, modify or impair the functioning of any disabling code in any Software; (l) circumvent or disable Ten Root Cyber Security’s copyright protection mechanisms or license management mechanisms; (m) use any Software in violation of any applicable Law (including but not limited to any Law with respect to human rights or the rights of individuals) or to support any illegal activity or to support any illegal activity; (n) use any Software to violate any rights of any Third Party; (o) use any Product for any training purposes, other than for training User’s em	ployees, where User charges fees or receives other consideration for such training, except as authorized by Ten Root Cyber Security in writing; (p) combine or operate any Products or Software with other products or software, without prior written authorization of Ten Root Cyber Security or its Affiliates, including without limitation any installation of any software on any Product, or; (q) attempt any of the foregoing. Ten Root Cyber Security expressly reserves the right to seek all available legal and equitable remedies to prevent any of the foregoing and to recover any lost profits, damages or costs resulting from any of the foregoing.
	2.       OWNERSHIP  
	A.	A.     Title to Software. Notwithstanding anything to the contrary, Software furnished hereunder is provided to Licensee subject to and in accordance with the terms and conditions of the EULA.  All title and interest of the Software and and/or any related Documentation and any derivative works thereof shall remain solely and exclusively with Ten Root Cyber Security or its licensors, as applicable. Nothing in this Agreement constitutes a sale, transfer or conveyance of any right, title or interest in any Software and/or Documentation or any derivative works thereof. Therefore, any reference to a sale of Software shall be understood as a license to Software under the terms and conditions of the Agreement. In the event of any conflict between the GTC and the EULA, the EULA shall take precedence over the GTC in all matters related to the Software.
	B.      Intellectual Property. All intellectual property rights relating to the Software and/or the Products, including without limitation, all patents, trademarks, algorithms, binary codes, business methods, computer programs, copyrights, databases, know-how, logos, concepts, techniques, processes, methods, models, commercial secrets and any other intellectual property rights, including any new developments or derivative works of such intellectual property, whether registered or not, are and shall remain the sole and exclusive property of Ten Root Cyber Security or its licensors, as applicable. All right, title and interest in and to any inventions, discoveries, improvements, methods, ideas, computer and other software or other works of authorship or other forms of intellectual property which are made, created, developed, written, conceived of or first reduced to practice solely, jointly with Licensee or on behalf of Licensee shall be and remain with Ten Root Cyber Security or its licensors, as applicable. 	Any suggestions, improvements or other feedback provided by Licensee to Ten Root Cyber Security regarding any Products, Software or services shall be the exclusive property of Ten Root Cyber Security.  Licensee hereby freely assigns any intellectual property rights to Ten Root Cyber Security in accordance with this Section, including any moral rights, and appoints Ten Root Cyber Security as its attorney-in-fact to pursue any such intellectual property rights worldwide.  
	3.     NO-WARRANTY: The software is provided without any warranty. 
	For the avoidance of doubt, Ten Root Cyber Security is not responsible for any claimed breach of any warranty caused by: Licensee’s use of the Products or Software in violation of Section 2(C) (“License Prohibitions”); placement of the Products or Software in an operating environment contrary to specific written instructions and training materials provided by Ten Root Cyber Security to Licensee; Licensee’s intentional or negligent actions or omissions, including physical damage, fire, loss or theft of a Product; any damage to a third party device alleged to or actually caused by or as a result of use of a Product or Software with a device; use of Products or Software incorporated into a system, other than as authorised by Ten Root Cyber Security; or any Products or Software that has been resold or otherwise transferred to a third party by Licensee. The warranties herein do not apply to, and Ten Root Cyber Security makes no warranties with respect to the computer or other platform on which the Software is ins	talled or otherwise embedded.  
	4.      USER INDEMNITY – User shall, at its expense: (i) indemnify and hold Ten Root Cyber Security and its Affiliates and its and their directors, officers, employees, agents, representatives, shareholders, subcontractors and suppliers harmless from and against any damages, claim, liabilities and expenses (including without limitation legal expenses) (whether brought by a Third Party or an employee, consultant or agent of User’s) arising out of any (a) misuse or use of any Product or Software furnished under the Agreement in a manner other than as authorized under this EULA, including without limitation using the Product or Software in a manner that violates applicable Law including without limitation a person’s Fourth Amendment rights under the United States Constitution (or its equivalent in the Territory); b) misappropriation of any personal information, (c) failure to obtain consents and approvals required by applicable Law for the use of any of the Ten Root Cyber Security’s Products or Software, or; (u	se of any Product or Software in breach of or to violate the terms of any other agreement with a Third Party; (ii) reimburse Ten Root Cyber Security for any expenses, costs and liabilities (including without limitation legal expenses) incurred relating to such claim; and (iii) pay all settlements, damages and costs assessed against Ten Root Cyber Security and attributable to such claim.
	5.    The sale of any Product by Ten Root Cyber Security shall not in any way confer upon User, or upon anyone claiming under User, any license (expressly, by implication, by estoppel or otherwise) under any patent claim of Ten Root Cyber Security or others covering or relating to any combination, machine or process in which such Product is or might be used, or to any process or method of making such Product.
	6.  This EULA and any disputes or claims arising hereunder are governed by the Laws of, and subject to the exclusive jurisdiction of the Israeli Law. The only and sole jurisdiction shall be in Tel-Aviv, Israel.

.DESCRIPTION
"Pure" powershell command (no dependencies, no special permissions needed etc') to retrieve change history in an AD group membership, or all groups, or per user. relies on object metadata rather than event logs. useful for DF/IR, tracking changes in groups etc'. Supports querying AD Metadata either from an Online Domain Controller, or from an offline system state backup / Snapshot.

.PARAMETER GroupName
A specific domain group, to be queried for changes (optional. defaults to "Domain Admins". N/A when -AllGroups is specified).

.PARAMETER AllGroups
Query all group changes in the domain, for all users.

.PARAMETER UserName
Query all group changes for a specific user (SamAccountName) in the domain (Runs -AllGroups 'behind the scenes', and filters results)

.PARAMETER Domain
The FQDN of the Domain to be queried (Optional. defaults to current domain DNS name. N/A for offline Backup/Snapshot instances).

.PARAMETER Output
Output to CSV file or Grid/Open results in a form (uses PowerShell ISE) - or Both. By default - outputs to console only.

.PARAMETER PageSize
Optional. The PageSize to set for the LDAP searcher object. Default is 100000.

.PARAMETER UseOfflineDBBackup
Option to query data from an offline DB using a system state backup or a Snapshot taken by ntdsutil. Does not require the domain to be available, or a DC connection. Can run on a Windows 10 Workstation, as well as Windows Server member/workgroup/DC.

.PARAMETER NTDSditDBPath
Required when specifing -UseOfflineDBBackup. The location of the NTDS.dit file, from a snapshot/system backup files.

.PARAMETER BackupInstanceLDAPPort
Optional. LDAP Port of the local NTDS Backup Instance to listen to. Default it 50005, and if in use - then randomly generated.

.PARAMETER DSAMainFilePath
Optional. Location of dsamain.exe, normally should be windows\system32.

.PARAMETER KeepOfflineDBinstanceRunning
Optional. For better performance of subsequent runs of this command, when running with -UseOfflineDBBackup. keeps the offline database connected and available. Later can run the command again with -UseExistingOfflineDBInstance.

.PARAMETER UseExistingOfflineDBInstance
Optional. After choosing to keep the offline database connected and available, can run the command again using this switch, along with -UseOfflineDBBackup.

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

.EXAMPLE
This command will get the list of historic changes for the group "Domain Admins" from a local offline NTDS backup/snapshot, using port 56015 as the local intance's LDAP port.
NOTE: This command option does not require an online Domain Controller, nor the domain to be available. 
It does require the following:
1. dsamain.exe system file, e.g. through AD tools/RSAT on Workstation/Server, or AD LDS feature (Client), or AD role binaries (Server)
2. a local System State backup or NTDS Snapshot (can be acquired using vssadmin w/mklink, ntdsutil IFM option, etc')

PS C:\temp> Get-ADGroupChanges -GroupName "domain admins" -UseOfflineDBBackup -NTDSditDBPath 'C:\temp\IFM\Active Directory\ntds.dit' -Port 56015

Option for Offline DB backup/snapshot was selected.
NTDS instance Loaded on Port <50005>

distinguishedName : {DC=Adatum,DC=com}
Path              : LDAP://localhost:50005

AD Backup Instance Loaded Successfully
AD Backup Instance Date Last Modified at 06/19/2019 15:59:38

GroupName            : Domain Admins
GroupDN              : CN=Domain Admins,CN=Users,DC=Adatum,DC=com
MemberSamAccountName : AdatumAdmin
Enabled              : True
LastChange           : Added
MemberDN             : CN=AdatumAdmin,CN=Users,DC=Adatum,DC=com
MemberAdminCount     : 1
DateTimeAdded        : 10/18/2016 1:00:30 PM
DateTimeRemoved      : N/A
LastChangeDateTime   : 10/18/2016 1:00:30 PM
DaysSinceLastChange  : 1705


GroupName            : Domain Admins
GroupDN              : CN=Domain Admins,CN=Users,DC=Adatum,DC=com
MemberSamAccountName : Administrator
Enabled              : True
LastChange           : Added
MemberDN             : CN=Administrator,CN=Users,DC=Adatum,DC=com
MemberAdminCount     : 1
DateTimeAdded        : 10/18/2016 12:48:46 PM
DateTimeRemoved      : N/A
LastChangeDateTime   : 10/18/2016 12:48:46 PM
DaysSinceLastChange  : 1705

Found changes information for 2 accounts in group <CN=Domain Admins,CN=Users,DC=Adatum,DC=com>

.LINK
yossis@protonmail.com
#>

    param(
        [cmdletbinding()]
        [Parameter(Mandatory = $false)]
        [string]$GroupName = "domain admins",

        [switch]$AllGroups = $false,

        [Parameter(Mandatory = $false)]
        [string]$UserName,

        [Parameter(Mandatory = $false)]
        [string]$Domain = "$env:USERDNSDOMAIN",

        [ValidateSet("ConsoleOnly", "Grid", "CSV", "CSV+Grid", IgnoreCase = $true)]
        [Parameter(Mandatory = $false)]
        [String[]]$Output = "ConsoleOnly",

        [Parameter(Mandatory = $false)]
        [int]$PageSize = 100000,

        [switch]$UseOfflineDBBackup = $false,

        [Parameter( ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    ParameterSetName="BackupInstance",
                    HelpMessage="Location of the NTDS.dit file, from the snapshot/system backup files")]
                    [Alias('DBPath')]
        [System.String]$NTDSditDBPath,

        [Parameter( ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    ParameterSetName="BackupInstance",
                    HelpMessage="LDAP Port of the local NTDS Backup Instance to listen to")]
                    [Alias('Port')]
        [int]$BackupInstanceLDAPPort = 50005,

        [Parameter( ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    ParameterSetName="BackupInstance",
                    HelpMessage="Location of dsamain.exe, normally should be windows\system32")]
                    [Alias('DsaMain')]
        [System.String]$DSAMainFilePath = "$ENV:windir\system32\dsamain.exe",
        
        [Parameter( ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false,
                    ParameterSetName="BackupInstance",
                    HelpMessage="Keep Offline DB instance running for better performance of subsequent command runs")]
        [switch]$KeepOfflineDBinstanceRunning = $false,

        [Parameter( ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false,
                    ParameterSetName="BackupInstance",
                    HelpMessage="Use a previously opened Offline DB instance that is currently running")]
        [switch]$UseExistingOfflineDBInstance = $false
    )

## Function to query a group's replication metadata 'living off the land' (no dependencies)
Function Get-GroupReplMetadata {

    param (
        [Parameter(Mandatory = $True)]
        [string]$GroupName,

        [Parameter(Mandatory = $True)]
        [System.DirectoryServices.DirectoryEntry]$DomainObj,

        [Parameter(Mandatory = $True)]
        [int]$PageSize
    )

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $DomainObj
        $ObjSearcher.PageSize = $PageSize; $objSearcher.SizeLimit = $PageSize
        $ObjSearcher.Filter = "(&(objectClass=group)(name=$GroupName))"
        $ObjSearcher.PropertiesToLoad.AddRange(("AdminCount","CanonicalName", "DistinguishedName", "Description", "GroupType","samaccountname", "SidHistory", "ManagedBy", "msDS-ReplValueMetaData", "ObjectSID", "WhenCreated", "WhenChanged", "member"))

        $GroupObj = $ObjSearcher.FindOne()
        $ObjSearcher.dispose()

        If ($GroupObj)
        {    
            $ReplMetaData = $GroupObj.Properties.'msds-replvaluemetadata' #'msDS-ReplValueMetaData'

            if ($ReplMetaData) {

            Write-Verbose "[X] Parsing replication metadata for group $(($GroupObj.Properties.samaccountname |Out-String).ToUpper())";

            $ReplMetaData | foreach {
                [xml]$ReplValue = ""
                $ReplValue.LoadXml($_.Replace("\x00", "").Replace("&","&amp;"))
    
                $LastActionDate = $ReplValue.DS_REPL_VALUE_META_DATA.ftimeLastOriginatingChange
                [int]$DaysSinceLastAction = ($(get-date) - [datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeLastOriginatingChange).Days
    
                if ($ReplValue.DS_REPL_VALUE_META_DATA.ftimeDeleted -eq "1601-01-01T00:00:00Z")
                    {
                        [string]$DateTimeRemoved = "-"
                    }
                    else
                    {
                        [datetime]$DateTimeRemoved = $ReplValue.DS_REPL_VALUE_META_DATA.ftimeDeleted
                    }

                if ([datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeCreated -gt [datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeDeleted)
                    {
                        $LastAction = "Added"
                    }
                    else
                    {
                        $LastAction = "Removed"
                    }

                # Check to see if local DB snapshot was specified
                if ($global:UseOfflineDBBackup)
                    {
                        $ObjMember = [adsi]"LDAP://localhost:$global:BackupInstanceLDAPPort/$($ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn)"
                    }
                else
                    {
                        $ObjMember = [adsi]"LDAP://$($ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn)"
                    }

                # Get member account's SamAccountName, AdminCount, Enabled/Disabled..
                if ($ObjMember.Properties.samaccountname)
                    {
                        $MemberSamAccountName = $ObjMember.Properties.samaccountname -join ',';
                        $MemberAdminCount = $ObjMember.Properties.admincount -join ',';
                        #$Enabled = if ($ObjMember.Properties.useraccountcontrol -eq 514 -or $ObjMember.Properties.useraccountcontrol -eq 66050) {"False"} else {"True"}
	                    $Enabled = if ($ObjMember.psbase.InvokeGet("useraccountcontrol") -band "0x2") {"False"} else {"True"}
                    }
                else
                    # if error occured on retrieval (e.g. Object Deleted, or using Chars such as <> messing up the xml) -
                    {
                        if ($ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn -like "*0ADEL:*")
                            # member is a deleted Object
                            {
                                [int]$0ADELindex = $ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn.IndexOf("\0ADEL:");
                                $MemberSamAccountName = $ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn.Substring(0,$0ADELindex);
                                $MemberAdminCount = "N/A (DELETED)";
                                $Enabled = "N/A (DELETED)";
                            }
                        else
                            # other error occured (e.g. illegal chars in DN, etc)
                            {
                                $MemberSamAccountName = $MemberAdminCount = $Enabled = "(Object parsing error)"
                            }
                    }

                $GroupChangeObj = New-Object PSObject;
                $global:GroupDN = $GroupObj.Properties.distinguishedname -join ',';

                Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "GroupName" -Value $($GroupObj.Properties.samaccountname -join ',') -Force
                Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "GroupDN" -Value $global:GroupDN -Force
                Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "MemberSamAccountName" -Value $MemberSamAccountName -Force
                Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "Enabled" -Value $Enabled -Force
                Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "LastChange" -Value $LastAction -Force

                # Handle DN value
                if ($ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn)
                    {
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "MemberDN" -Value $ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn -Force
                    }
                else
                    {
                        $StartDN = $_.IndexOf("<pszObjectDn>CN=");
                        $EndDN = $_.IndexOf("</pszObjectDn>"); 
                        [int]$StringRange = $EndDN - $StartDN + 16; 
                        $MemberDN = $_.Substring($StartDN,$StringRange).replace('<pszObjectDn>CN=','').replace('</pszObjectDn>','');
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "MemberDN" -Value $MemberDN -Force
                    }
                
                Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "MemberAdminCount" -Value $MemberAdminCount -Force

                # Handle added datetime
                if ($ReplValue.DS_REPL_VALUE_META_DATA.ftimeCreated)
                    {
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DateTimeAdded" -Value $([datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeCreated) -Force
                    }
                else
                    {
                        $StartTimeCreated = $_.IndexOf("<ftimeCreated>");
                        $EndTimeCreated = $_.IndexOf("</ftimeCreated>");
                        [int]$StringRangeTimeCreated = $EndTimeCreated - $StartTimeCreated + 15;
                        [datetime]$TimeCreated = $_.Substring($StartTimeCreated,$StringRangeTimeCreated).replace('<ftimeCreated>','').replace('</ftimeCreated>','');
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DateTimeAdded" -Value $TimeCreated -Force
                    }

                # Handle removed datetime
                if ($DateTimeRemoved)
                    {
                        # handle TTL members (temporary members that were removed, but there's no Remove date with TTL! just added timestamp
                        if ($DateTimeRemoved -eq "-") {
                            if ($global:UseOfflineDBBackup) {
                                # get group members from offline backup
                                $grpObj = [ADSI]"LDAP://localhost:$global:BackupInstanceLDAPPort/$GroupDN"
                                }
                            else
                                {
                                # get group members from online DC
                                $grpObj = ([adsisearcher]"(&(objectclass=group)(name=$GroupName))").FindOne()
                            }

                            if ($grpObj.Properties.member -notcontains $($ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn)) {
                                Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "LastChange" -Value "!Temporary Member Removed (TTL)!" -Force
                            }
                        }

                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DateTimeRemoved" -Value $DateTimeRemoved -Force
                    }
                else
                    {
                        $StartTimeDeleted = $_.IndexOf("<ftimeDeleted>");
                        $EndTimeDeleted = $_.IndexOf("</ftimeDeleted>");
                        [int]$StringRangeTimeDeleted = $EndTimeDeleted - $StartTimeDeleted + 15;
                        $TimeDeleted = $_.Substring($StartTimeDeleted,$StringRangeTimeDeleted).replace('<ftimeDeleted>','').replace('</ftimeDeleted>','');
                        if ($TimeDeleted -eq "1601-01-01T00:00:00Z")
                            {
                                $TimeDeleted = "-"
                            }
                        else
                            {
                                [datetime]$TimeDeleted = $TimeDeleted
                            }
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DateTimeRemoved" -Value $TimeDeleted -Force
                    }

                # Handle last change
                if ($ReplValue.DS_REPL_VALUE_META_DATA.ftimeLastOriginatingChange)
                    {
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "LastChangeDateTime" -Value $([datetime]$ReplValue.DS_REPL_VALUE_META_DATA.ftimeLastOriginatingChange) -Force
                    }
                else
                    {
                        $StartTimeLastChange = $_.IndexOf("<ftimeLastOriginatingChange>");
                        $EndTimeLastChange = $_.IndexOf("</ftimeLastOriginatingChange>");
                        [int]$StringRangeLastChange = $EndTimeLastChange - $StartTimeLastChange + 29;
                        [datetime]$LastChange = $_.Substring($StartTimeLastChange,$StringRangeLastChange).replace('<ftimeLastOriginatingChange>','').replace('</ftimeLastOriginatingChange>','');
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "LastChangeDateTime" -Value $LastChange -Force
                    }
                
                Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DaysSinceLastChange" -Value $DaysSinceLastAction -Force

                # check for case where user was added temproraily to a group, and is STILL an active member of the group
                if ($GroupChangeObj.LastChangeDateTime -gt $GroupChangeObj.DateTimeAdded -and $GroupChangeObj.DateTimeRemoved -eq "-" -and $GroupChangeObj.LastChange -ne "!Temporary Member Removed (TTL)!")
                    {
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DateTimeAdded" -Value $GroupChangeObj.LastChangeDateTime -Force;
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "LastChange" -Value "Added (Temporarily)" -Force
                    }

                # check for case where user was added temproraily to a group, and expired its membership (reflect the correct AddedTime)
                if ($GroupChangeObj.LastChangeDateTime -gt $GroupChangeObj.DateTimeAdded -and $GroupChangeObj.LastChange -eq "!Temporary Member Removed (TTL)!")
                    {
                        Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "DateTimeAdded" -Value $GroupChangeObj.LastChangeDateTime -Force
                    }

                # v1.5.6 - check is the member is still part of the group
                if ($GroupObj.Properties.member -contains $ReplValue.DS_REPL_VALUE_META_DATA.pszObjectDn) {
                    Add-Member -InputObject $GroupChangeObj -MemberType NoteProperty -Name "LastChange" -Value "Added" -Force
                }

                $global:GroupMembershipChanges += $GroupChangeObj
                Clear-Variable GroupChangeObj
            }
            } # end of ReplMetadata handling
        
            else # no replication metadata found for this group
            {
                Write-Verbose "[X] No replication metadata found for group $(($GroupObj.Properties.samaccountname |Out-String).ToUpper()) (no activity history to parse)"
            }

        }

        else
        {
            # $GroupObj was empty
            if ($global:UseExistingOfflineDBInstance -and !(Get-Process dsamain)) {
                    Write-Warning "[X] Error: An Offline DB Instance was not found.";
                    break
                }

            if ($AllGroups -eq $false) {
                    Write-Warning "[X] Error while getting the Group Object. Make sure you typed the group name correctly.";
                    break
                }
            else
                {
                   #Write-Warning "[X] Error while getting the Group Object. Skipping to next group.";
                }
        }
    } # End of Group Query function

## Reporting functions
function Out-GroupChangesToGrid {
    if ($AllGroups)
        {
            switch ($global:UseOfflineDBBackup)
                {
                    $true {$title = "Group Membership Changes in Domain $(($global:DomainFQDN).ToUpper()) <BACKUP FROM $global:OfflineDBDateTime>"}
                    $false {$title = "Group Membership Changes in Domain $(($global:DomainFQDN).ToUpper())"}
                }
            
            if ($global:UserName) {$title = $title + " for user $($global:UserName.ToUpper())"}
        }

    else
        {
            switch ($global:UseOfflineDBBackup)
                {
                    $true {$title = "Group membership changes for group $($($GroupObj.Properties.samaccountname -join ',').ToUpper()) in domain $(($global:DomainFQDN).ToUpper()) <BACKUP FROM $global:OfflineDBDateTime>"}
                    $false {$title = "Group membership changes for group $($($GroupObj.Properties.samaccountname -join ',').ToUpper()) in domain $(($global:DomainFQDN).ToUpper())"}
                }
        }
        
        $GroupMembershipChanges | Sort-Object LastChangeDateTime -Descending | Out-GridView -Title $Title;
    }

function Out-GroupChangesToCSV {
    if ($AllGroups) 
        {
            switch ($global:UserName)
                {
                    {$_ -ne ""} {$FileName = "GroupMembershipChanges_$($global:DomainFQDN.ToUpper())_$($global:Username.ToUpper())_$(Get-Date -Format ddMMyyyy_HHmmss).csv"}
                    default {$FileName = "GroupMembershipChanges_$($global:DomainFQDN.ToUpper())_$(Get-Date -Format ddMMyyyy_HHmmss).csv"}
                }
        }
    else
        {
            $FileName = "GroupMembershipChanges_$($global:DomainFQDN.ToUpper())_$($GroupName.ToUpper())_$(Get-Date -Format ddMMyyyy_HHmmss).csv";
        }

    if ($global:UseOfflineDBBackup)
        {
            $null = $FileName.Replace(".csv","_FROM-OFFLINE-DB.csv")
        }

    $GroupMembershipChanges | Export-Csv $FileName -NoTypeInformation;
    Write-Host "[*] Results saved to $FileName." -ForegroundColor Green
}

## global variables
$global:UseOfflineDBBackup = $UseOfflineDBBackup;
$global:BackupInstanceLDAPPort = $BackupInstanceLDAPPort;
$global:UseExistingOfflineDBInstance = $UseExistingOfflineDBInstance;
$global:Username = $UserName;
[bool]$SkipOutput = $false;

if ($UserName) {
        [switch]$AllGroups = $true;
        if (!(([adsisearcher]"(samaccountname=$global:Username)").FindOne()))
            {
                Write-Warning "[x] Account $global:Username Not found in AD. please ensure you've typed it correctly and try again.`nQuiting.";
                break;
            }
    }

# Get current machine role
$Role = Get-Wmiobject -Class Win32_computersystem
 
Switch ($Role.domainrole) {
        "0"     { $HostType = "Standalone workstation"}
        "1"     { $HostType = "Domain Member workstation"}
        "2"     { $HostType = "Standalone server"}
        "3"     { $HostType = "Domain Member server"}
        "4"     { $HostType = "Domain controller"}
        "5"     { $HostType = "Domain Controller - PDC Role"}
}

Write-Host "[*] Running on $env:COMPUTERNAME ($HostType)"

if ($HostType -notlike "*domain*" -and $Domain -eq $env:USERDOMAIN) {
    Write-Host "[*] You are running from a Non-Domain joined host. Unless using an offline DB backup, consider specifying a domain name using the " -ForegroundColor Yellow -NoNewline; Write-Host "-DOMAIN" -ForegroundColor Cyan -NoNewline; Write-Host " parameter." -ForegroundColor Yellow
}

## Begin checks for pre-requisites

# Check if -UseExistingOfflineDBInstance specified, yet there is no active instance
if ($global:UseExistingOfflineDBInstance -and $global:DSAProc -eq $null)
        {
            Write-Warning "[X] Cannot find existing offline DB instance running.";
            break
        }

    # Check if DB path to NTDS.dit was specified
    if ($NTDSditDBPath -ne "" -and $global:UseOfflineDBBackup -eq $false -and $global:UseExistingOfflineDBInstance -eq $false)
        {
            Write-Host "[*] a DBPath to an NTDS.dit file was specified.`nAn offline DB backup will be used." -ForegroundColor Cyan;
            $global:UseOfflineDBBackup = $true;
        }

    if ($global:UseOfflineDBBackup -and $global:UseExistingOfflineDBInstance -eq $false) {

        if (!(New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
            {
	            Write-Warning "[X] Must be elevated to run Offline DB operations.`nPlease open an administrative shell and try again. Quiting.";
            }
    
    Write-Host "[*] Option for Offline DB backup/snapshot was selected." -ForegroundColor Cyan
    
    # Check if DB path to NTDS.dit was specified
    if ($NTDSditDBPath -eq "" -and $global:UseExistingOfflineDBInstance -eq $false)
        {
            Write-Warning "[X] a DBPath to an NTDS.dit file was not specified.`nPlease specify an NTDS.dit file from a backup/snapshot.";
            break
        }

    # Check if DB path to NTDS.dit was specified
    if (!(Test-Path $NTDSditDBPath) -and $global:UseExistingOfflineDBInstance -eq $false)
        {
            Write-Warning "[X] Cannot find the NTDS.dit file. Make sure you specified a valid path to an NTDS.dit file from a backup/snapshop.";
            break
        }

    # Check if Windows 10 AND AD LDS not Installed - opt user to choose to install required binaries
    [string]$OS = (Get-WmiObject -ClassName win32_Operatingsystem | select caption).Caption

    # if dsamain is not present, and OS is Windows 10, and the AD LDS feature is not installed and Enabled - offer to install it
    if (!(Test-Path $DSAMainFilePath) -and $OS -like "*Windows 10*" -and $(Get-WindowsOptionalFeature -Online -FeatureName DirectoryServices-ADAM-Client).State -ne "Enabled" -and $global:UseExistingOfflineDBInstance -eq $false)
        {
            # Display a choice menu to approve running with local administrator & LAPS passwords using NTLM
            $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Yes - Install AD LDS feature & dsamain required file(s)"
            $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No","No - and Continue without installing AD LDS"
            $Cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Exit Script"
            $Options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No, $Cancel)
            $Title = "Installing AD Lightweight Directory Services (dsamain.exe)" 
            $Message = "`nBy default, the dsamain.exe pre-requisite for loading an Offline DB instance is Not present on Windows 10.`nDo you want to install this Feature?`nNOTE: You can Remove the 'AD LDS' feature at a later time.`n`n"
            $ResultChoiceADLDS = $Host.ui.PromptForChoice($Title, $Message, $Options, 2)
            
            switch ($ResultChoiceADLDS) {
                0 {
                    Enable-WindowsOptionalFeature -Online -FeatureName DirectoryServices-ADAM-Client
                }
                1 {
                    # do nothing - continue without installing AD LDS
                }
                2 {
                    Write-Warning "[X] Cannot find dsamain.exe. Loading of Backup Instance cannot continue.`nMake sure you have relevant files installed (e.g. AD RSAT or AD LDS / AD Role).";
                    break
                }       
            }       
        }

    # Check if dsamain.exe path is valid (default is \WINDIR\System32) 
    if (!(Test-Path $DSAMainFilePath) -and $global:UseExistingOfflineDBInstance -eq $false)
        {
            Write-Warning "[X] Cannot find dsamain.exe. Loading of Backup Instance cannot continue.`nMake sure you have relevant files installed (e.g. AD RSAT or AD LDS / AD Role).";
            break
        }

    # Check if port is available (default is 50005, or user specified) 
    if ((New-Object System.Net.Sockets.TcpClient).ConnectAsync('localhost',$global:BackupInstanceLDAPPort).Wait(1000) -and $global:UseExistingOfflineDBInstance -eq $false) 
	    {
	        # if port in use, try a different random port
            Write-Host "[*] Specified Port <$global:BackupInstanceLDAPPort> is in use, using a different random port." -ForegroundColor Yellow
		    $global:BackupInstanceLDAPPort = Get-Random -Minimum 49152 -Maximum 65535
	    }

    if ($global:UseExistingOfflineDBInstance -eq $false)
        {
            $DSAMainArguments = "/dbpath """ + $NTDSditDBPath + """ /ldapport $global:BackupInstanceLDAPPort /allowNonAdminAccess"
            $DsaMainProc = Start-Process -FilePath $DSAMainFilePath -ArgumentList $DSAMainArguments -PassThru -RedirectStandardError $true -WindowStyle Hidden;
	    
	    # Check if process was launched successfully
            Sleep -Seconds 3;
            if (!(Get-Process -Id $DsaMainProc.Id -ErrorAction SilentlyContinue))
                {
                    Write-Warning "[X] Process was Not loaded successfully. Make sure NTDS.dit file is not corrupted.`nLoading of the NTDS Instance Failed";
		            break
                }

            Write-Host "[*] NTDS instance Loaded on Port <$global:BackupInstanceLDAPPort>" -ForegroundColor Cyan
    
            # make sure instance loaded fine
            Sleep -seconds 3;
            [ADSI]"LDAP://localhost:$global:BackupInstanceLDAPPort";
            if (!$?)
	            {
    	            #$Error[0].ErrorRecord.Exception	
                    Write-Warning "[X] Loading of the NTDS Instance Failed"
		            break
	            }		
            else
	            {
		            Write-Host "[*] AD Backup Instance Loaded Successfully" -ForegroundColor Green;
                    $global:OfflineDBDateTime = (Get-ChildItem $NTDSditDBPath).LastWriteTimeUtc;
                    Write-Host "[*] AD Backup Instance Date Last Modified at $global:OfflineDBDateTime" -ForegroundColor Yellow
	            }

            } #end of Loading Offline DB instance
    } #end of UseOfflineDBBackup

    if ($global:UseExistingOfflineDBInstance)
            {
                $global:BackupInstanceLDAPPort = $global:Port;
                $DsaMainProc = $global:DSAProc;
            }
            
    # Get the domain Distinguished name
    if ($global:UseOfflineDBBackup) {
            $DN = ([adsi]"LDAP://localhost:$global:BackupInstanceLDAPPort").distinguishedName;
            $DomainObj = New-Object System.DirectoryServices.DirectoryEntry("LDAP://localhost:$global:BackupInstanceLDAPPort/$DN");
        }
    else
        {
            if ($Domain)
                {
                    # v1.5.5 - Get the domain NC and oldest DC from the Specified domain FQDN
                    # First, get the default domain naming context
                    
                    $domainNC = "DC=" + $($domain.Replace(".",",DC="))

                    # Build LDAP path
                    $searcher = New-Object System.DirectoryServices.DirectorySearcher;
                    $searcher.SearchRoot = "LDAP://$domainNC";

                    # Only computer objects with server roles indicating a DC
                    $searcher.Filter = "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))";

                    # Ask for attributes
                    $searcher.PropertiesToLoad.Add("name")        | Out-Null;
                    $searcher.PropertiesToLoad.Add("dnsHostName") | Out-Null;
                    $searcher.PropertiesToLoad.Add("whenCreated") | Out-Null;

                    $results = $searcher.FindAll();

                    $DCs = foreach ($r in $results) {
                        $props = $r.Properties;

                        [PSCustomObject]@{
                            DCName      = $props["dnshostname"][0]
                            WhenCreated = [datetime]$props["whencreated"][0]
                        }
                    }

                    # Sort and return the oldest DC, and DN for DomainObj
                    $OldestDCfqdn = $DCs | Sort-Object WhenCreated | Select-Object -First 1 -ExpandProperty DCName;
                    $DN = "$OldestDCfqdn/$domainNC"
                }
            else
                {
                    # Get the default domain naming context
                    $rootDSE = [ADSI]"LDAP://RootDSE";
                    $domainNC = $rootDSE.defaultNamingContext;

                    # Build LDAP path
                    $searcher = New-Object System.DirectoryServices.DirectorySearcher;
                    $searcher.SearchRoot = "LDAP://$domainNC";

                    # Only computer objects with server roles indicating a DC
                    $searcher.Filter = "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))";

                    # Ask for attributes
                    $searcher.PropertiesToLoad.Add("name")        | Out-Null;
                    $searcher.PropertiesToLoad.Add("dnsHostName") | Out-Null;
                    $searcher.PropertiesToLoad.Add("whenCreated") | Out-Null;

                    $results = $searcher.FindAll();

                    $DCs = foreach ($r in $results) {
                        $props = $r.Properties;

                        [PSCustomObject]@{
                            DCName      = $props["dnshostname"][0]
                            WhenCreated = [datetime]$props["whencreated"][0]
                        }
                    }

                    # Sort and return the oldest DC, and DN for DomainObj
                    $OldestDCfqdn = $DCs | Sort-Object WhenCreated | Select-Object -First 1 -ExpandProperty DCName;
                    $DN = "$OldestDCfqdn/$domainNC"
                }

            $DomainObj = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DN");
        }
    
    $global:DomainFQDN = $DN.Split("/")[1].Substring(3).replace("DC=",".").replace(",","");

    if ($AllGroups) {
        Write-Host "[*] Queries for single users or all groups may take a while to complete. Please be patient..." -ForegroundColor Cyan;
        # Query data from all groups
        $global:GroupMembershipChanges = @();

        # Get all groups
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher -ArgumentList $DomainObj;
        $ObjSearcher.PageSize = $PageSize; $objSearcher.SizeLimit = $PageSize;
        $ObjSearcher.Filter = '(objectCategory=group)'

        $Groups = $ObjSearcher.FindAll();
        $ObjSearcher.dispose()
        
        $Groups.Properties.samaccountname | foreach {
            Get-GroupReplMetadata -GroupName $_ -DomainObj $DomainObj -PageSize $PageSize;
        }
    }

    else 
    {
        # Query a single group
        $global:GroupMembershipChanges = @();
        Get-GroupReplMetadata -GroupName $GroupName -DomainObj $DomainObj -PageSize $PageSize;
    }

    # Wrap up steps #

    # Display results
    if ($UserName)
        {
            $global:GroupMembershipChanges = $global:GroupMembershipChanges | Where-Object MemberSamAccountName -eq $UserName;
            if ($global:GroupMembershipChanges.Count -eq 0)
                {
                    Write-Warning "[X] No group changes found for user $UserName. Make sure you spelled the samAccountName correctly.";
                    [bool]$SkipOutput = $true
                }        
        }

    if (!$SkipOutput) {
            $global:GroupMembershipChanges | Sort-Object LastChangeDateTime -Descending
        }

    if ($SkipOutput -eq $false) 
        {
        switch ($AllGroups)
            {
                "True" {Write-Host "[*] Found $($global:GroupMembershipChanges.count) group membership changes." -ForegroundColor Gray}
                "False" {Write-Host "[*] Found changes information for $($global:GroupMembershipChanges.count) accounts in group <$global:GroupDN>" -ForegroundColor Green}
            }

        switch ($Output)
            {
                "Grid" {Out-GroupChangesToGrid}
                "CSV" {Out-GroupChangesToCSV}
                "CSV+Grid" {Out-GroupChangesToGrid; Out-GroupChangesToCSV}
            }
    }

    # Prompt if to keep offline instance running 
    if ($global:UseOfflineDBBackup -and $KeepOfflineDBinstanceRunning -eq $false)
        {
            $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Yes - Keep Offline DB Instance running"
            $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No","No - Terminate Offline DB Instance"
            $Options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
            $Title = "Keep Offline DB Instance loaded in memory and running?" 
            $Message = "`nFor better performance of subsequent command runs (additional query of Group Changes), it is recommended to keep the offline DB instance running.`nIf you will Not use this command again from this process, choose 'No' to terminate the Offline DB instance.`n`n"
            $ResultChoiceKeepOfflineInstance = $Host.ui.PromptForChoice($Title, $Message, $Options, 0)
            
            switch ($ResultChoiceKeepOfflineInstance) {
                0 {
                    $KeepOfflineDBinstanceRunning = $true
                }
                1 {
                    # do nothing - DB instance will be terminated
                }
            } 
        }

    # Check if need to close Offline DB instance connection or not (Default is to close)
    if ($global:UseOfflineDBBackup -and $KeepOfflineDBinstanceRunning)
        {
            Write-Host "`n[*] Offline DB instance is running in the background (Better performance for subsequent offline DB commands).`nYou can run the command again with a different group name and the " -ForegroundColor Yellow -NoNewline; Write-Host "-UseExistingOfflineDBInstance" -NoNewline -ForegroundColor Cyan;
            Write-Host "parameter.`ne.g. " -NoNewline -ForegroundColor Yellow; Write-Host "Get-ADGroupChanges -GroupName 'Enterprise Admins' -UseOfflineDBBackup -UseExistingOfflineDBInstance`n" -ForegroundColor Cyan;
            Write-Host "[*] Offline DB Instance process details:";
            $DsaMainProc | Format-Table Processname, ID, @{N='WorkingSet (MB)';E={[math]::Round($_.WS/1mb)}}, @{N='PM (MB)';E={[math]::Round($_.PM/1mb)}}, Handles;
            $global:Port = $global:BackupInstanceLDAPPort;
            $global:DSAProc = $DsaMainProc;
        }
    elseif ($global:UseOfflineDBBackup)
        {
            # Terminate offline DB instance listener
            $DsaMainProc | Stop-Process -Force;
            
            # Remove temp file opened by dsamain.exe (not cleaned by default and 'in use' during operation)
            $FileOpenTrue = Get-ChildItem "$((Get-Location).Path)\True" | Where-Object {$_.PSIsContainer -eq $false} -ErrorAction SilentlyContinue
            if ($FileOpenTrue) {
                    Remove-Item $FileOpenTrue.FullName -ErrorAction SilentlyContinue;
                    if (!$?) # Deletion of open handle file failed - likely due to multiple dsamain instances running
                        {
                            Write-Host "[X] Unable to remove the open handle file:`n$(($Error[0]).exception.Message)" -ForegroundColor DarkYellow
                            Write-Host "[X] Make sure you don't have other dsamain instances running from multiple runs of the command.`ne.g. Type " -ForegroundColor Yellow -NoNewline;
                            Write-Host -NoNewline "Get-Process dsamain" -ForegroundColor Cyan; Write-Host ", and see results. Can remove them with " -NoNewline -ForegroundColor Yellow;
                            Write-Host -NoNewline "Get-Process dsamain | Stop-Process -Force`n" -ForegroundColor Cyan;
                        }
                }

            # handle variables
            $global:Port = $null;
            $global:DSAProc = $null;

            Write-Host "`n[*] Offline DB instance terminated successfully.`n" -ForegroundColor Yellow -NoNewline; Write-Host "TIP:" -ForegroundColor Magenta -NoNewline;
            Write-Host " for better performance (to keep the Offline DB intance running next time), specify parameter " -ForegroundColor Yellow -NoNewline; Write-Host -ForegroundColor Cyan -NoNewline "-KeepOfflineDBinstanceRunning"; 
            Write-Host -NoNewline -ForegroundColor Yellow ".`nThen run the command with parameter "; Write-Host "-UseExistingOfflineDBInstance" -ForegroundColor Cyan;
        }
}