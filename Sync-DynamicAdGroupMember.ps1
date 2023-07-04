[CmdletBinding(SupportsShouldProcess)]
<#
    .SYNOPSIS
    Manages AD group members based on Get-ADUser filter query defined in an extensionAttribute

    .DESCRIPTION
    The Sync-DynamicAdGroupMember.ps1 loops thru all AD groups that have a Get-ADUser filter query defined on a speficied extensionAttribute.
    The script then fetches all AD users that match the query and syncs them with the groups members.
    This means missing members are added and obsolete members are removed.
    Manual changes to the members of the group are overwritten.

    .LINK
    https://github.com/dominikduennebacke/Sync-DynamicAdGroupMember

    .NOTES
    - Version: 0.9.0
    - License: GPL-3.0
    - Author:   Dominik Dünnebacke
        - Email:    dominik@duennebacke.com
        - GitHub:   https://github.com/dominikduennebacke
        - LinkedIn: https://www.linkedin.com/in/dominikduennebacke/

    .INPUTS
    None

    .OUTPUTS
    None or PSCustomObject
    Returns the group, its Get-ADUser filter query, user and modification type as PSCustomObject if the PassThru parameter is specified. By default, this cmdlet does not generate any output.

    .EXAMPLE
    Syncs all members for groups that have a filter query set in extensionAttribute10 and provides output.

    .\Sync-DynamicAdGroupMember.ps1 -ExtensionAttribute 10 -VERBOSE

    VERBOSE: Checking dependencies
    VERBOSE: The secure channel between the local computer and the domain is in good condition.
    VERBOSE: Fetching AD groups with a value in extensionAttribute10
    VERBOSE: Syncing group members
    VERBOSE: role-department-sales: department -eq 'sales'
    VERBOSE: role-department-sales: (+) john.doe
    VERBOSE: role-department-sales: (+) sam.smith
    VERBOSE: role-department-sales: (-) tom.tonkins

    .EXAMPLE
    Provides output of sync changes but does not actually perform them.

    .\Sync-DynamicAdGroupMember.ps1 -ExtensionAttribute 10 -WhatIf:$true

    What if: role-department-sales: (+) john.doe
    What if: role-department-sales: (+) sam.smith
    What if: role-department-sales: (-) tom.tonkins

    .EXAMPLE
    Provides output of sync changes but does not actually perform them, with additional output.

    .\Sync-DynamicAdGroupMember.ps1 -ExtensionAttribute 10 -WhatIf:$true -VERBOSE

    VERBOSE: Checking dependencies
    VERBOSE: The secure channel between the local computer and the domain is in good condition.
    VERBOSE: Fetching AD groups with a value in extensionAttribute10
    VERBOSE: Syncing group members
    VERBOSE: role-department-sales: department -eq 'sales'
    What if: role-department-sales: (+) john.doe
    What if: role-department-sales: (+) sam.smith
    What if: role-department-sales: (-) tom.tonkins

    .EXAMPLE
    Only consideres the OU "OU=groups,DC=contoso,DC=com" looking for groups with extensionAttribute10 set.
    Only consideres the OU "OU=users,DC=contoso,DC=com" looking for users when executing the Get-ADUser filter query.
    This can speed up execution.

    .\Sync-DynamicAdGroupMember.ps1 -ExtensionAttribute 10 -GroupSearchBase "OU=groups,DC=contoso,DC=com" -UserSearchBase "OU=users,DC=contoso,DC=com"
#>


# -------------------------------------------------------------------------------------------------------------- #
#region Parameters
param (
    # extensionAttribute that is used to define Get-ADUser filter queries
    [parameter(Mandatory = $true)]
    [ValidateRange(1, 15)]
    [int]$ExtensionAttribute,

    # extensionAttribute that is used to define Where-Object filter queries
    [Alias('WhereFilterEA')]
    [ValidateRange(1, 15)]
    [int]$WhereFilterExtensionAttribute,

    # Specifies an Active Directory path to search for groups
    [string]$GroupSearchBase,

    # Specifies an Active Directory group name to update
    [string]$GroupName,

    # Specifies an Active Directory OU DistinguishedName to use for Users SearchBase
    [parameter(ParameterSetName = "UserSearchBase")]
    [string]$UserSearchBase,

    # Specifies the Active Directory group attribute to use for Users SearchBase
    [Alias('UserSearchBaseEA')]
    [ValidateRange(1, 15)]
    [parameter(ParameterSetName = "UserSearchAttribute")]
    [int]$UserSearchBaseExtensionAttribute,

    # Specifies the Active Directory Domain Services instance to connect to
    [string]$Server,

    # Returns the group, its Get-ADUser filter query, user and modification type as PSCustomObject.
    [switch]$PassThru
)
#endregion Parameters


# -------------------------------------------------------------------------------------------------------------- #
#region Checking dependencies
Write-Verbose "Checking dependencies"

# Determine if PowerShell module ActiveDirectory is installed
if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    throw "PowerShell module 'ActiveDirectory' is not installed"
}

# Determine domain controller to use for AD cmdlets (if $Server is not set)
if (-not $Server) {
    $Server = [String] (Get-ADDomainController -Discover).HostName
}
if (-not $Server) {
    throw "No AD domain controller was found"
}


# Test connection to server
if (-not (Test-ComputerSecureChannel -Server $Server -WhatIf:$false)) {
    throw "Connection to $Server failed"
}

# Determine GroupSearchBase if not set (as cmdlets do not allow an empty SearchBase)
if (-not $GroupSearchBase) {
    $GroupSearchBase = (Get-ADDomain -Server $Server).DistinguishedName
}

# Determine UserSearchBase if not set (as cmdlets do not allow an empty SearchBase)
if (-not $UserSearchBase) {
    $UserSearchBase = (Get-ADDomain -Server $Server).DistinguishedName
}

# Store extensionAttribute as string
$ExtensionAttributeString = "extensionAttribute" + $ExtensionAttribute

If ($WhereFilterExtensionAttribute) {
    $WhereFilterAttributeString = "extensionAttribute" + $WhereFilterExtensionAttribute
}
If ($UserSearchBaseExtensionAttribute) {
    $UserSearchBaseAttributeString = "extensionAttribute" + $UserSearchBaseExtensionAttribute
}
$Properties = $ExtensionAttributeString, $WhereFilterAttributeString, $UserSearchBaseAttributeString | Where-Object { $_ -ne $null }

# TODO : Test extensionAttribute uniqueness
#endregion Checking dependencies


# -------------------------------------------------------------------------------------------------------------- #
#region Fetching AD groups with extensionAttribute set
Write-Verbose "Fetching AD groups with extensionAttribute set"

# Fetching AD groups
$Params = @{
    SearchBase = $GroupSearchBase
    Filter     = "$ExtensionAttributeString -like '*'"
    Server     = $Server
    Properties = $Properties | Where-Object { $_ }
}
If ($GroupName) {
    $Params.Add('Identity', $GroupName)
    $Params.Remove('Filter')
    $Params.Remove('SearchBase')
}
[array]$AdGroups = Get-ADGroup @Params | Sort-Object Name
#endregion Fetching AD groups with extensionAttribute set


# -------------------------------------------------------------------------------------------------------------- #
#region Syncing group members
Write-Verbose "Syncing group members"
If ($WhereFilterAttributeString) {
    $AllADUserProperties = Get-ADUser -Filter * -SearchBase $UserSearchBase -Properties * -ResultSetSize 1 | Select-Object -ExpandProperty PropertyNames
}

foreach ($Group in $AdGroups) {

    # Output for reference
    Write-Verbose "$($Group.Name): $($Group.$ExtensionAttributeString)"

    $Start = Get-Date

    # Fetch AD users from query
    $GetParams = @{
        Filter = $Group.$ExtensionAttributeString
        Server = $Server
        SearchBase = If ($UserSearchBaseAttributeString -and -not [String]::IsNullOrEmpty($Group.$UserSearchBaseAttributeString)) {
            $Group.$UserSearchBaseAttributeString
        } Else {
            $UserSearchBase
        }
    }
    If ($WhereFilterAttributeString -and -not [String]::IsNullOrEmpty($Group.$WhereFilterAttributeString)) {
        $ADUserProperties = $AllADUserProperties | ForEach-Object { If (Select-String -InputObject $Group.$WhereFilterAttributeString -Pattern $_ -SimpleMatch) { $_ } }
        If ($ADUserProperties) {
            $GetParams.Add('Property', $ADUserProperties)
        }
    }
    $MembersQuery = Get-ADUser @GetParams | Sort-Object SamAccountName

    If ($WhereFilterAttributeString -and $Group.$WhereFilterAttributeString) {
        $MembersQuery = $MembersQuery | Where-Object -FilterScript ([ScriptBlock]::Create($Group.$WhereFilterAttributeString))
    }
    # Fetch current members of AD group
    $MembersCurrent = $Group | Get-ADGroupMember -Server $Server | Sort-Object SamAccountName

    # Determine missing and obsolete members
    $MissingMembers = $MembersQuery | Where-Object { $MembersCurrent.SID -notcontains $_.SID }
    $ObsoleteMembers = $MembersCurrent | Where-Object { $MembersQuery.SID -notcontains $_.SID }

    # Add missing members
    foreach ($Member in $MissingMembers) {
        If ($PSCmdlet.ShouldProcess($Member.SamAccountName, "Add user to $($Group.Name)")) {
            Add-ADGroupMember -Identity $Group.Name -Members $Member -Confirm:$false -Server $Server

            # Provide output in case $PassThru is set
            if ($PassThru) {
                [PSCustomObject]@{
                    Group  = $Group.Name
                    Query  = $Group.$ExtensionAttributeString
                    User   = $Member.SamAccountName
                    Action = "Add"
                }
            }
        }
    }

    # Remove missing members
    foreach ($Member in $ObsoleteMembers) {
        If ($PSCmdlet.ShouldProcess($Member.SamAccountName, "Remove user from $($Group.Name)")) {
            Remove-ADGroupMember -Identity $Group.Name -Members $Member -Confirm:$false -Server $Server

            # Provide output in case $PassThru is set
            if ($PassThru) {
                [PSCustomObject]@{
                    Group  = $Group.Name
                    Query  = $Group.$ExtensionAttributeString
                    User   = $Member.SamAccountName
                    Action = "Remove"
                }
            }
        }
    }

    $End = Get-Date
    $TS = $End - $Start
    Write-Verbose -Message $TS.ToString()
}
#endregion Syncing group members